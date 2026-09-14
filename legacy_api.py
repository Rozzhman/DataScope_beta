import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import io
import re
import base64
from datetime import datetime
from typing import Optional, Callable
from pathlib import Path
try:
    import winreg
except Exception:
    winreg = None
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH
import textwrap
from absolute_v5.config.prompts import OSINT_SEARCH_PROMPT_TEMPLATE
from absolute_v5.scanners.pipeline import build_pipeline
from absolute_v5.config.settings import AppConfig
from absolute_v5.core.ai_client import AIClient
from absolute_v5.core.profile_utils import (
    classify_status,
    extract_json_from_text,
    extract_profile_from_rusprofile,
    normalize_registration_place,
    normalize_tax_mode,
    sanitize_profile_fields,
)
from absolute_v5.reporting.docx_styles import apply_doc_styles
from absolute_v5.reporting.generator import ReportGenerator
try:
    from PIL import Image
except Exception:
    Image = None

# --- АЛГОРИТМ ABSOLUTE (ВЕРСИЯ 5.2.4: FULL-PAGE FIX) ---

DEFAULT_RESULTS = {
    "fedresurs": "В ожидании...",
    "pb": "В ожидании...",
    "disqualified": "В ожидании...",
    "blocking": "В ожидании...",
    "rnp": "В ожидании...",
    "arbitr": "В ожидании...",
    "rosfin": "В ожидании...",
    "fssp": "В ожидании...",
}

FRIENDLY_NAMES = {
    "fedresurs": "Федресурс (Банкротство)",
    "pb": "Прозрачный бизнес (ФНС)",
    "disqualified": "Дисквалификация (ФНС)",
    "blocking": "Блокировки счетов (ФНС)",
    "rnp": "Госзакупки (РНП)",
    "arbitr": "Арбитражные суды",
    "rosfin": "Росфинмониторинг",
    "fssp": "Приставы (ФССП)",
}

STATUS_MAP = {
    "pb": ("✅ ПРОВЕРЕНО", "❌ ЕСТЬ НЕГАТИВ"),
    "fedresurs": ("✅ НЕ НАЙДЕНО", "❌ ОБНАРУЖЕНО"),
    "arbitr": ("✅ ДЕЛ НЕТ", "❌ ЕСТЬ ДЕЛА"),
    "fssp": ("✅ ДОЛГОВ НЕТ", "❌ ЕСТЬ ДОЛГИ"),
    "rnp": ("✅ НЕ НАЙДЕНО", "❌ В РНП"),
    "blocking": ("✅ БЛОКИРОВОК НЕТ", "❌ ЕСТЬ БЛОКИРОВКИ"),
    "disqualified": ("✅ НЕ НАЙДЕНО", "❌ В РЕЕСТРЕ"),
    "rosfin": ("✅ НЕ НАЙДЕНО", "❌ В СПИСКЕ"),
}

class AbsoluteAPI:
    def __init__(self, inn: str, bik: str, log_func: Callable[[str], None], config: Optional[AppConfig] = None):
        self.inn = inn
        self.bik = bik
        self._log_func = log_func
        self.cfg = config or AppConfig()
        self.fio = "" 
        self.driver: Optional[uc.Chrome] = None
        self.status_cells = {} 
        self.osint_paragraphs: dict[str, Paragraph] = {}
        
        
        self.results = dict(DEFAULT_RESULTS)
        self.inspections_check_failed = False
        self.site_screenshots: dict[str, bytes] = {}
        self.include_screenshots_in_report = True
        
        
        self.friendly_names = dict(FRIENDLY_NAMES)
        self.ai_client = AIClient(self.cfg, self.log)
        self.profile_data: dict[str, str] = {
            "full_name": "",
            "inn": self.inn,
            "ogrnip": "",
            "registration_date": "",
            "egrip_status": "",
            "main_activity": "",
            "okved_additional": "",
            "registration_place": "",
            "tax_regimes": "",
            "tax_regimes_date": "",
            "pb_tax_debt_status": "",
            "pb_reporting_status": "",
            "pb_arrear_total": "",
            "pb_form1_income": "",
            "pb_form1_expense": "",
            "fssp_total_debt": "",
            "inspections_info": "",
            "sources": "",
            "director": "",
        }

        self.doc = Document()
        self._apply_doc_styles()
        self.cfg.results_dir.mkdir(parents=True, exist_ok=True)
        self.cfg.database_dir.mkdir(parents=True, exist_ok=True)
        self.cfg.fssp_dir.mkdir(parents=True, exist_ok=True)

    def log(self, text: str) -> None:
        msg = (text or "").strip()
        if not msg:
            return

        if msg.startswith("[") and "]" in msg[:12]:
            self._log_func(msg)
            return

        lower = msg.lower()
        if msg.startswith("Этап"):
            tag = "STEP"
        elif msg.startswith("Результат: ХОРОШО") or "успех" in lower:
            tag = "OK"
        elif "ошибка" in lower:
            tag = "ERROR"
        elif "внимание" in lower or "плохо" in lower or "не удалось" in lower or "требуется ручная проверка" in lower:
            tag = "WARN"
        else:
            tag = "INFO"

        self._log_func(f"[{tag}] {msg}")

    def _apply_doc_styles(self):
        """Применяет базовые стили документа через reporting-слой."""
        try:
            apply_doc_styles(self.doc)
        except Exception:
            pass

    def add_cover_page(self):
        """Титульная часть первой страницы."""
        self.doc.add_heading(f"Отчет о проверке контрагента {self.inn}", level=0)

    def _add_labeled_line(self, label: str, value: str):
        paragraph = self.doc.add_paragraph()
        paragraph.paragraph_format.line_spacing = 1.0
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        run_label = paragraph.add_run(f"{label}: ")
        run_label.bold = True
        paragraph.add_run(value.strip() if value else "нет данных")

    def _extract_json_from_text(self, content: str) -> Optional[dict]:
        return extract_json_from_text(content)

    def _extract_profile_from_rusprofile(self, html_text: str):
        extract_profile_from_rusprofile(self.profile_data, html_text)

    def _enrich_inspections_info_with_ai(self):
        # AI enrichment disabled by product decision.
        return

    def _sanitize_profile_fields(self):
        sanitize_profile_fields(self.profile_data)

    def _enrich_profile_data_with_ai(self):
        # AI enrichment disabled by product decision.
        return

    def _build_text_report(self):
        self.doc = Document()
        self._apply_doc_styles()
        self._sanitize_profile_fields()

        def fmt_amount(value: str) -> str:
            raw = (value or "").strip()
            if not raw:
                return raw
            if "₽" in raw or "руб" in raw.lower():
                return raw
            return f"{raw} ₽"

        name = self.profile_data.get("full_name") or self.inn
        # Определяем наличие ОГРНИП (регистрирует ИП)
        has_ogrnip = bool((self.profile_data.get("ogrnip", "") or "").strip())

        def is_legal_entity_name(raw_name: str) -> bool:
            value = (raw_name or "").strip().upper()
            # Проверяем сокращённые формы
            if re.match(r"^(ООО|ПАО|АО|ОАО|ЗАО|НКО|ТСЖ|ЖСК)\b", value):
                return True
            # Проверяем полные названия
            if re.search(r"(ОБЩЕСТВО|ОТКРЫТОЕ|ЗАКРЫТОЕ|ЧАСТНОЕ|АКЦИОНЕРНОЕ|ТОВАРИЩЕСТВО|КООПЕРАТИВ)", value):
                return True
            return False

        # Для ИП добавляем префикс "ИП", только если это не юридическое лицо и ещё нет префикса
        if has_ogrnip and name and not is_legal_entity_name(name) and not re.match(r"^ИП\b", name, re.IGNORECASE):
            name = f"ИП {name}"
        self.profile_data["full_name"] = name
        self.profile_data["inn"] = self.inn

        def add_labeled(label: str, value: str, bold_label: bool = False):
            if not (value or "").strip():
                return
            p = self.doc.add_paragraph()
            r1 = p.add_run(f"{label}: ")
            r1.bold = bold_label
            p.add_run(value)

        def add_labeled_multiline(label: str, value: str, bold_label: bool = False):
            lines = [ln.strip() for ln in (value or "").split("\n") if ln.strip()]
            if not lines:
                return
            p = self.doc.add_paragraph()
            r1 = p.add_run(f"{label}: ")
            r1.bold = bold_label
            p.add_run(lines[0])
            for line in lines[1:]:
                self.doc.add_paragraph(line)

        def wrap_multiline(value: str, width: int) -> str:
            lines: list[str] = []
            for part in re.split(r"[;\n]+", value or ""):
                part = part.strip()
                if not part:
                    continue
                wrapped = textwrap.wrap(part, width=width) or [part]
                lines.extend(wrapped)
            return "\n".join(lines)

        title = self.doc.add_paragraph(f"Справка по документальной проверке ИНН {self.inn}")
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.runs[0].bold = True

        subject = self.doc.add_paragraph(name)
        subject.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subject.runs[0].bold = True

        self.doc.add_paragraph("")
        subject_label = "Организация" if is_legal_entity_name(name) else "Наименование организации"
        add_labeled(subject_label, name, bold_label=True)
        director = (self.profile_data.get("director", "") or "").strip()
        if director:
            add_labeled("Руководитель организации", director, bold_label=True)
        ogrn = self.profile_data.get("ogrnip", "").strip()
        ogrn_label = "ОГРНИП" if has_ogrnip else "ОГРН"
        requisites_text = f"ИНН {self.inn}" + (f", {ogrn_label} {ogrn}" if ogrn else "")
        add_labeled_multiline("Основные реквизиты", requisites_text)
        registration_date = (self.profile_data.get("registration_date", "") or "").strip()
        if not registration_date:
            src_text = self.profile_data.get("sources", "") or ""
            m_reg = re.search(r"Дата регистрации[:\s]*([0-9]{1,2}[.\s][0-9]{1,2}[.\s][0-9]{2,4})", src_text, flags=re.IGNORECASE)
            if not m_reg:
                m_reg = re.search(r"Дата присвоения ОГРН[:\s]*([0-9]{1,2}[.\s][0-9]{1,2}[.\s][0-9]{2,4})", src_text, flags=re.IGNORECASE)
            if m_reg:
                registration_date = m_reg.group(1).replace(" ", ".").strip(". ")
        add_labeled("Дата регистрации", registration_date)
        egrip_status = self.profile_data.get("egrip_status", "").strip()
        if egrip_status:
            reg_source = "ЕГРЮЛ" if is_legal_entity_name(name) else "ЕГРИП"
            add_labeled("Статус", f"По данным {reg_source} имеет статус: {egrip_status}.")
        activity = self.profile_data.get("main_activity", "")
        activity_wrapped = "\n".join(textwrap.wrap(activity, width=62)) if activity else ""
        add_labeled_multiline("Основной вид деятельности", activity_wrapped, bold_label=True)
        add_labeled("Место регистрации", self.profile_data.get("registration_place", ""))
        okved_additional = (self.profile_data.get("okved_additional", "") or "").strip()
        if okved_additional:
            okved_wrapped = wrap_multiline(okved_additional, width=62)
            add_labeled_multiline("Дополнительные ОКВЭД", okved_wrapped)
        tax_mode = (self.profile_data.get("tax_regimes", "") or "").strip()
        tax_date = (self.profile_data.get("tax_regimes_date", "") or "").strip()
        if not tax_mode:
            tax_mode = "Общая система налогообложения (ОСН)"
        elif "не применяется" in tax_mode.lower():
            tax_mode = "Упрощенная система налогообложения (УСН)"
        tax_wrapped = "\n".join(textwrap.wrap(tax_mode, width=40)) if tax_mode else ""
        if tax_wrapped:
            if tax_date:
                add_labeled(f"Применяемые специальные налоговые режимы на {tax_date}", "")
                self.doc.add_paragraph(tax_wrapped)
            else:
                if "ОСН" in tax_mode:
                    add_labeled_multiline("Система налогообложения", tax_wrapped)
                else:
                    add_labeled("Применяемые специальные налоговые режимы", "")
                    self.doc.add_paragraph(tax_wrapped)

        self.doc.add_paragraph("")

        fedresurs_status = (self.results.get("fedresurs", "") or "").strip()
        fedresurs_low = fedresurs_status.lower()
        fedresurs_bad = (
            ("❌" in fedresurs_status)
            or ("внимание" in fedresurs_low)
            or (
                ("обнаруж" in fedresurs_low)
                and ("не обнаруж" not in fedresurs_low)
                and ("не найден" not in fedresurs_low)
            )
        )
        fssp_status = (self.results.get("fssp", "") or "").strip()
        fssp_bad = self._classify_status(fssp_status) == "bad"
        blocking_status = (self.results.get("blocking", "") or "").strip()
        blocking_bad = self._classify_status(blocking_status) == "bad"

        overall_bad = fedresurs_bad or fssp_bad or blocking_bad
        title_facts = "Неблагоприятные факты" if overall_bad else "Благоприятные факты"
        p = self.doc.add_paragraph(title_facts)
        p.runs[0].bold = True

        if overall_bad:
            if fedresurs_bad:
                self.doc.add_paragraph("Обнаружены признаки процедуры банкротства.")
                self.doc.add_paragraph("В Едином федеральном реестре сведений о банкротстве найдены сообщения о банкротстве контрагента.")
            else:
                self.doc.add_paragraph("Признаков процедуры банкротства не обнаружено.")
                self.doc.add_paragraph("В Едином федеральном реестре сведений о банкротстве не найдено сообщений о банкротстве контрагента.")
            if blocking_bad:
                self.doc.add_paragraph("Обнаружены блокировки счетов.")
            if fssp_bad:
                self.doc.add_paragraph("Обнаружена задолженность ФССП.")
        else:
            self.doc.add_paragraph("Признаков процедуры банкротства не обнаружено.")
            self.doc.add_paragraph("В Едином федеральном реестре сведений о банкротстве не найдено сообщений о банкротстве контрагента.")
            self.doc.add_paragraph("Отсутствие блокировок счетов.")
            self.doc.add_paragraph("Отсутствие задолженностей ФССП.")

        self.doc.add_paragraph("")
        p2 = self.doc.add_paragraph("Проведена проверка информации о наличии арбитражных дел за последние 12 месяцев")
        p2.runs[0].bold = True
        arbitr_total = self.results.get("arbitr_total")
        arbitr_count_12m = self.results.get("arbitr_count_12m")
        arbitr_unfiltered = bool(self.results.get("arbitr_count_unfiltered"))
        arbitr_bad = self._classify_status(self.results.get("arbitr", "")) == "bad"
        pb_bankruptcy_process = "процесс банкротства" in (self.profile_data.get("egrip_status", "") or "").lower()
        if isinstance(arbitr_total, int) or isinstance(arbitr_count_12m, int):
            if isinstance(arbitr_total, int):
                self.doc.add_paragraph(f"Общее количество найденных дел: {arbitr_total}.")
            if isinstance(arbitr_count_12m, int):
                if arbitr_count_12m > 0:
                    self.doc.add_paragraph(f"Дела за последние 12 месяцев: {arbitr_count_12m}.")
                else:
                    self.doc.add_paragraph("Дела за последние 12 месяцев отсутствуют.")
            if pb_bankruptcy_process or fedresurs_bad:
                self.doc.add_paragraph("Найдены арбитражные дела, связанные с банкротством.")
            else:
                self.doc.add_paragraph("Арбитражных дел, связанных с банкротством, не найдено.")
        else:
            if arbitr_bad:
                self.doc.add_paragraph("Выявлены арбитражные дела.")
                self.doc.add_paragraph("Найдены арбитражные дела, связанные с банкротством.")
            else:
                if pb_bankruptcy_process or fedresurs_bad:
                    self.doc.add_paragraph("Арбитражные дела присутствуют.")
                    self.doc.add_paragraph("Найдены арбитражные дела, связанные с банкротством.")
                else:
                    self.doc.add_paragraph("Арбитражные дела отсутствуют.")
                    self.doc.add_paragraph("Арбитражных дел, связанных с банкротством, не найдено.")

        self.doc.add_paragraph("")
        rnp_status_text = (self.results.get("rnp", "") or "").lower()
        rnp_bad = self._classify_status(self.results.get("rnp", "")) == "bad"
        rnp_excluded = "исключ" in rnp_status_text
        rnp_found = rnp_bad or rnp_excluded
        disq_bad = self._classify_status(self.results.get("disqualified", "")) == "bad"
        disq_found = disq_bad

        found_entries = []
        not_found_entries = []

        if rnp_found:
            found_entries.append("Реестр \"Недобросовестные поставщики\"")
        else:
            not_found_entries.append("Реестр \"Недобросовестные поставщики\"")

        if disq_found:
            found_entries.append("Реестр \"Дисквалифицированные лица\"")
        else:
            not_found_entries.append("Реестр \"Дисквалифицированные лица\"")

        pb_tax_debt_status = (self.profile_data.get("pb_tax_debt_status", "") or "").strip().lower()
        if pb_tax_debt_status:
            debt_open_list = (
                "Открытый список \"Задолженность перед ФНС более 1000 рублей, направленная на "
                "взыскание судебному приставу-исполнителю\""
            )
            if "не имеет задолж" in pb_tax_debt_status:
                not_found_entries.append(debt_open_list)
            else:
                found_entries.append(debt_open_list)

        pb_reporting_status = (self.profile_data.get("pb_reporting_status", "") or "").strip().lower()
        if pb_reporting_status:
            reporting_open_list = "Открытый список \"Не предоставляют налоговую отчетность более года\""
            if "не представляет" in pb_reporting_status or "непредстав" in pb_reporting_status:
                found_entries.append(reporting_open_list)
            elif "представляет налоговую отчетность" in pb_reporting_status:
                not_found_entries.append(reporting_open_list)

        def unique_preserve(items):
            seen = set()
            result = []
            for item in items:
                key = (item or "").strip().lower()
                if not key or key in seen:
                    continue
                seen.add(key)
                result.append(item)
            return result

        found_entries = unique_preserve(found_entries)
        not_found_entries = [x for x in unique_preserve(not_found_entries) if (x or "").strip().lower() not in {i.strip().lower() for i in found_entries}]

        self.doc.add_paragraph("")
        p3_title = self.doc.add_paragraph("Проверка по особым реестрам и открытым спискам")
        p3_title.runs[0].bold = True

        pb_arrear_total = (self.profile_data.get("pb_arrear_total", "") or "").strip()
        pb_form1_income = (self.profile_data.get("pb_form1_income", "") or "").strip()
        pb_form1_expense = (self.profile_data.get("pb_form1_expense", "") or "").strip()

        if found_entries:
            p3_found = self.doc.add_paragraph("Обнаружены упоминания:")
            p3_found.runs[0].bold = True
            for item in found_entries:
                self.doc.add_paragraph(item, style="List Bullet")

        if not_found_entries:
            p3 = self.doc.add_paragraph("Не обнаружены упоминания:")
            p3.runs[0].bold = True
            for item in not_found_entries:
                self.doc.add_paragraph(item, style="List Bullet")

        if not found_entries and not not_found_entries:
            self.doc.add_paragraph("Данные по особым реестрам и открытым спискам отсутствуют.")

        if pb_arrear_total or pb_form1_income or pb_form1_expense:
            self.doc.add_paragraph("")
            p_fin = self.doc.add_paragraph("Финансовые показатели из открытых источников ФНС")
            p_fin.runs[0].bold = True
            if pb_arrear_total:
                self.doc.add_paragraph(
                    f"Сведения о суммах недоимки и задолженности по пеням и штрафам, общая сумма: {pb_arrear_total}"
                )
            if pb_form1_income or pb_form1_expense:
                self.doc.add_paragraph("Суммы доходов и расходов по данным бухгалтерской отчетности организации")
                if pb_form1_income:
                    self.doc.add_paragraph(f"Доход: {fmt_amount(pb_form1_income)}")
                if pb_form1_expense:
                    self.doc.add_paragraph(f"Расход: {fmt_amount(pb_form1_expense)}")

        fssp_total_debt = (self.profile_data.get("fssp_total_debt", "") or "").strip()
        if fssp_total_debt:
            self.doc.add_paragraph("")
            p_fssp = self.doc.add_paragraph(
                "Проверка общей задолженности из открытых источников Федеральной службы судебных приставов."
            )
            p_fssp.runs[0].bold = True
            self.doc.add_paragraph(f"Общая задолженность составляет: {fssp_total_debt} руб.")

        inspections_info = self.profile_data.get("inspections_info", "").strip()
        if inspections_info:
            self.doc.add_paragraph(inspections_info)

    def _classify_status(self, status_text: str) -> str:
        return classify_status(status_text)

    def add_final_conclusion(self):
        """Итоговое заключение в деловом формате."""
        self.doc.add_page_break()
        self.doc.add_heading("Итоговое заключение", level=1)

        counts = {"ok": 0, "warn": 0, "bad": 0, "error": 0, "unknown": 0}
        for status in self.results.values():
            counts[self._classify_status(status)] += 1

        if counts["bad"] > 0:
            risk_level = "ВЫСОКИЙ"
        elif counts["warn"] > 0 or counts["error"] > 0:
            risk_level = "СРЕДНИЙ"
        elif counts["ok"] > 0 and counts["unknown"] == 0:
            risk_level = "НИЗКИЙ"
        else:
            risk_level = "НЕОПРЕДЕЛЕН"

        self.doc.add_paragraph(f"Уровень риска: {risk_level}")

        if counts["bad"] > 0:
            self.doc.add_paragraph("Обнаруженные риски:")
            for site, status in self.results.items():
                if self._classify_status(status) == "bad":
                    self.doc.add_paragraph(f"{self.friendly_names.get(site, site)} — {status}", style="List Bullet")

        if counts["error"] > 0:
            self.doc.add_paragraph("Ошибки/невозможность проверки:")
            for site, status in self.results.items():
                if self._classify_status(status) == "error":
                    self.doc.add_paragraph(f"{self.friendly_names.get(site, site)} — {status}", style="List Bullet")

        if counts["warn"] > 0:
            self.doc.add_paragraph("Требует ручной проверки:")
            for site, status in self.results.items():
                if self._classify_status(status) == "warn":
                    self.doc.add_paragraph(f"{self.friendly_names.get(site, site)} — {status}", style="List Bullet")

    def _set_labeled_paragraph(self, paragraph: Paragraph, label: str, value: str):
        p = paragraph._p
        for child in list(p):
            p.remove(child)
        run_label = paragraph.add_run(f"{label}: ")
        run_label.bold = True
        paragraph.add_run(value)

    def _crop_image_top(self, img_bytes: bytes, keep_ratio: float = 0.6) -> bytes:
        """Обрезаем изображение, оставляя верхнюю часть (по умолчанию 60%)."""
        if Image is None:
            return img_bytes
        try:
            with Image.open(io.BytesIO(img_bytes)) as im:
                w, h = im.size
                keep_h = int(h * keep_ratio)
                keep_h = max(1, min(h, keep_h))
                cropped = im.crop((0, 0, w, keep_h))
                out = io.BytesIO()
                cropped.save(out, format="PNG")
                return out.getvalue()
        except Exception:
            return img_bytes

    def _apply_gost_paragraph(self, paragraph: Paragraph, align: int = WD_ALIGN_PARAGRAPH.LEFT, bold: bool = False):
        paragraph.alignment = align
        paragraph.paragraph_format.line_spacing = 1.0
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        for run in paragraph.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)
            run.bold = bold
            run.font.color.rgb = RGBColor(0, 0, 0)

    def _normalize_osint_text_for_doc(self, text: str) -> str:
        lines = []
        for raw in (text or "").splitlines():
            line = raw.strip()
            if not line:
                lines.append("")
                continue
            # Remove markdown heading markers like "#", "##", "###"
            line = re.sub(r"^\s*#{1,6}\s*", "", line)
            # Remove markdown emphasis markers to keep plain business text in DOCX.
            line = re.sub(r"\*{1,3}", "", line)
            line = re.sub(r"_{1,3}", "", line)
            line = re.sub(r"`+", "", line)
            # Remove source index brackets like [1], [2], [12]
            line = re.sub(r"\s*\[\d+\]\s*", " ", line)
            line = re.sub(r"\s{2,}", " ", line).strip()
            lines.append(line)
        # Trim excessive empty lines on edges while keeping inner spacing
        while lines and not lines[0]:
            lines.pop(0)
        while lines and not lines[-1]:
            lines.pop()
        return "\n".join(lines)

    def _ensure_mandatory_osint_sections(self, text: str) -> str:
        normalized = (text or "").strip()
        lower = normalized.lower()
        required = [
            (
                "Кто является бенефициаром",
                ("бенефициар", "конечн бенефициар", "beneficiar"),
                "Недостаточно подтвержденных данных по бенефициарам.",
            ),
            (
                "Чистые активы",
                ("чистые активы", "net assets"),
                "Недостаточно подтвержденных данных по чистым активам.",
            ),
            (
                "Наличие негативных статей (компания, бенефициары, учредители, руководители)",
                ("негативн", "негативные статьи", "репутационн", "медиариск"),
                "Подтвержденных негативных публикаций не найдено или данных недостаточно.",
            ),
            (
                "Наличие офшоров или связанных офшоров",
                ("офшор", "offshore"),
                "Подтвержденных данных о наличии офшоров или связанных офшоров не найдено.",
            ),
        ]

        additions: list[str] = []
        for title, markers, fallback in required:
            if not any(marker in lower for marker in markers):
                additions.append(f"{title}:\n{fallback}")

        if additions:
            if normalized:
                normalized += "\n\n"
            normalized += "\n\n".join(additions)

        return normalized

    @staticmethod
    def _has_all_osint_sections(text: str) -> bool:
        """Quick check that 10 OSINT sections are present in the text."""
        lower = (text or "").lower()
        markers = [
            "1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10.",
            "executive summary", "идентификация объекта", "юридическ", "судебн", "деловая",
            "медиа", "репутационн", "связи", "аналитическ", "источники",
        ]
        # Require at least 10 unique markers; prioritize numbered markers.
        found_num = sum(1 for n in ("1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10.") if n in lower)
        if found_num >= 10:
            return True
        found = sum(1 for m in markers if m in lower)
        return found >= 10

    @staticmethod
    def _missing_osint_sections(text: str) -> list[int]:
        lower = (text or "").lower()
        return [n for n in range(1, 11) if f"{n}." not in lower]

    @staticmethod
    def _clean_osint_value(value: object, default: str = "Недостаточно подтвержденных данных.") -> str:
        text = str(value or "").strip()
        return text if text else default

    @staticmethod
    def _normalize_risk_level(level: object) -> str:
        raw = str(level or "").strip().lower()
        if "выс" in raw:
            return "Высокий"
        if "сред" in raw:
            return "Средний"
        if "низ" in raw:
            return "Низкий"
        return "Не определен"

    def _build_osint_text_from_facts(self, payload: dict) -> str:
        entity_name = str(payload.get("entity_name") or "").strip() or "Не указано"
        inn = str(payload.get("inn") or "").strip() or self.inn
        profile_name = str(self.profile_data.get("full_name") or "").strip()

        if inn != self.inn:
            self.log(f"OSINT: скорректирован ИНН объекта в ответе модели ({inn} -> {self.inn}).")
            inn = self.inn

        if profile_name and entity_name and profile_name.lower() not in entity_name.lower() and entity_name.lower() not in profile_name.lower():
            self.log("OSINT: скорректировано название объекта (ответ модели не совпал с объектом отчета).")
            entity_name = profile_name
        summary = self._clean_osint_value(payload.get("summary"), "Недостаточно подтвержденных данных по краткому резюме.")

        beneficiary = payload.get("beneficiary") or {}
        ben_status = self._clean_osint_value(
            beneficiary.get("status"),
            "недостаточно данных",
        )
        ben_name = str(beneficiary.get("name") or "").strip()
        ben_basis = self._clean_osint_value(
            beneficiary.get("basis"),
            "Подтвержденные сведения о конечном контроле не найдены.",
        )
        status_low = ben_status.lower()
        basis_low = ben_basis.lower()
        if ben_name:
            ben_line = f"Статус: {ben_status}. Лицо: {ben_name}. Основание: {ben_basis}"
        else:
            ben_line = f"Статус: {ben_status}. {ben_basis}"

        # Business-friendly clarification: when ultimate owner is not disclosed,
        # explicitly mention two common practical scenarios.
        if (
            ("не раскрыт" in status_low or "недостаточно" in status_low)
            and ("учредител" in basis_low or "акционер" in basis_low or "егрюл" in basis_low)
        ):
            ben_line += (
                " Возможный практический вывод: конечный бенефициар может быть глубоко скрыт, "
                "либо бенефициаром может являться один из контролирующих учредителей/акционеров."
            )

        financials = payload.get("financials") or {}
        net_assets_value = str(financials.get("net_assets_value") or "").strip()
        net_assets_period = str(financials.get("net_assets_period") or "").strip()
        net_assets_comment = self._clean_osint_value(
            financials.get("net_assets_comment"),
            "Недостаточно подтвержденных данных для интерпретации показателя чистых активов.",
        )
        if net_assets_value:
            net_assets_line = f"Показатель: {net_assets_value}"
            if net_assets_period:
                net_assets_line += f" ({net_assets_period})"
            net_assets_line += f". Интерпретация: {net_assets_comment}"
        else:
            net_assets_line = "Недостаточно подтвержденных данных по чистым активам."

        neg = payload.get("negative_media") or {}
        neg_company = self._clean_osint_value(neg.get("company"), "Подтвержденных негативных публикаций по компании не найдено.")
        neg_ben = self._clean_osint_value(neg.get("beneficiaries"), "По бенефициарам подтвержденных негативных публикаций не найдено.")
        neg_founders = self._clean_osint_value(neg.get("founders"), "По учредителям подтвержденных негативных публикаций не найдено.")
        neg_managers = self._clean_osint_value(neg.get("managers"), "По руководителям подтвержденных негативных публикаций не найдено.")

        def format_media_examples(examples_obj: object) -> list[str]:
            if not isinstance(examples_obj, list):
                return []
            out: list[str] = []
            for item in examples_obj[:3]:
                if not isinstance(item, dict):
                    continue
                fact = str(item.get("fact") or "").strip()
                quote = str(item.get("quote") or "").strip()
                date = str(item.get("date") or "").strip()
                source_idx = str(item.get("source_index") or "").strip()
                if not fact and not quote:
                    continue
                parts = []
                if fact:
                    parts.append(fact)
                if quote:
                    parts.append(f"цитата: \"{quote}\"")
                if date:
                    parts.append(f"дата: {date}")
                if source_idx:
                    parts.append(f"источник [{source_idx}]")
                out.append("- " + "; ".join(parts))
            return out

        offshore = payload.get("offshore") or {}
        off_status = self._clean_osint_value(offshore.get("status"), "недостаточно данных")
        off_details = self._clean_osint_value(
            offshore.get("details"),
            "Подтвержденных данных о наличии офшоров или связанных офшоров не найдено.",
        )

        risks = payload.get("risk_blocks") or {}
        risk_labels = [
            ("Репутационные риски", "reputation"),
            ("Юридические и судебные риски", "legal"),
            ("Санкционные риски", "sanctions"),
            ("Финансовые риски", "financial"),
            ("Аффилированность", "affiliation"),
        ]
        risk_lines: list[str] = []
        for title, key in risk_labels:
            block = risks.get(key) or {}
            lvl = self._normalize_risk_level(block.get("level"))
            facts = self._clean_osint_value(block.get("facts"), "Подтвержденных фактов не найдено.")
            risk_lines.extend([title, f"Уровень риска: {lvl}", f"Факты: {facts}"])

            evidence = block.get("evidence")
            if isinstance(evidence, list):
                detail_lines: list[str] = []
                for item in evidence[:5]:
                    if not isinstance(item, dict):
                        continue
                    ev_fact = str(item.get("fact") or "").strip()
                    ev_date = str(item.get("date_or_period") or "").strip()
                    ev_metric = str(item.get("metric") or "").strip()
                    ev_src = str(item.get("source_index") or "").strip()
                    if not ev_fact:
                        continue
                    parts = [ev_fact]
                    if ev_date:
                        parts.append(f"период/дата: {ev_date}")
                    if ev_metric:
                        parts.append(f"метрика: {ev_metric}")
                    if ev_src:
                        parts.append(f"источник [{ev_src}]")
                    detail_lines.append("- " + "; ".join(parts))
                if detail_lines:
                    risk_lines.append("Детализация:")
                    risk_lines.extend(detail_lines)
            risk_lines.append("")

        sources_raw = payload.get("sources") or []
        if isinstance(sources_raw, list):
            sources = [str(x).strip() for x in sources_raw if str(x).strip()]
        else:
            sources = []
        if not sources:
            sources = ["Недостаточно подтвержденных источников."]

        confidence_overall = self._clean_osint_value(
            payload.get("confidence_overall"),
            "средняя",
        )

        identification = payload.get("identification") or {}
        legal_registration = payload.get("legal_registration") or {}
        judicial_exec = payload.get("judicial_exec") or {}
        business_activity = payload.get("business_activity") or {}
        media_presence = payload.get("media_presence") or {}
        affiliations = payload.get("affiliations") or {}
        missing_data_recommendations = self._clean_osint_value(
            payload.get("missing_data_recommendations"),
            "",
        )

        def section_lines(title: str, rows: list[tuple[str, object]]) -> list[str]:
            content: list[str] = []
            for label, value in rows:
                text = str(value or "").strip()
                if not text:
                    continue
                if text.lower() in {"недостаточно данных", "данные не найдены", "не найдено"}:
                    content.append(f"- {label}: {text}.")
                else:
                    content.append(f"- {label}: {text}")
            if not content:
                return []
            return [title, *content, ""]

        lines = [
            "OSINT-СПРАВКА ПО КОНТРАГЕНТУ",
            "",
            f"ИНН: {inn}",
            f"Организация: {entity_name}",
            "",
            "Краткое резюме",
            summary,
            f"Общая оценка достоверности данных: {confidence_overall}.",
            "",
        ]
        lines.extend(
            section_lines(
                "Идентификация объекта",
                [
                    ("Совпадения", identification.get("matches")),
                    ("Подтверждение идентичности", identification.get("identity_confirmation")),
                    ("Возможные тезки/одноименные компании", identification.get("namesakes")),
                    ("Уровень уверенности", identification.get("confidence")),
                ],
            )
        )
        lines.extend(
            section_lines(
                "Юридические и регистрационные сведения",
                [
                    ("Регистрация", legal_registration.get("registration")),
                    ("Роли (учредитель/директор/бенефициар)", legal_registration.get("roles")),
                    ("Адреса, даты, статус", legal_registration.get("addresses_dates_status")),
                ],
            )
        )
        lines.extend(
            section_lines(
                "Судебная и исполнительная история",
                [
                    ("Арбитражные производства за 3 года (всего)", judicial_exec.get("arbitration_total_3y")),
                    ("Где объект выступал ответчиком (за 3 года)", judicial_exec.get("arbitration_respondent_3y")),
                    ("Сумма требований по делам ответчика (за 3 года)", judicial_exec.get("arbitration_respondent_amount_3y")),
                    ("Исполнительные производства", judicial_exec.get("exec_proceedings")),
                    ("Детализация", judicial_exec.get("details")),
                ],
            )
        )
        lines.extend(
            section_lines(
                "Деловая и профессиональная активность",
                [
                    ("Компании и проекты", business_activity.get("companies_projects")),
                    ("Периоды участия", business_activity.get("participation_periods")),
                    ("Связанные лица", business_activity.get("related_persons")),
                    ("Офшорные/косвенные связи", business_activity.get("offshore_links_indirect")),
                ],
            )
        )
        lines.extend(
            section_lines(
                "Медиа и публичное присутствие",
                [
                    ("Упоминания", media_presence.get("mentions")),
                    ("Публикации/интервью", media_presence.get("publications_interviews")),
                    ("Профили в соцсетях", media_presence.get("social_profiles")),
                    ("Даты и контекст", media_presence.get("timeline_context")),
                ],
            )
        )
        lines.extend([
            "",
            "Кто является бенефициаром",
            ben_line,
            "",
            "Чистые активы",
            net_assets_line,
            "",
            "Наличие негативных статей (компания, бенефициары, учредители, руководители)",
            f"Компания: {neg_company}",
            *format_media_examples(neg.get("company_examples")),
            f"Бенефициары: {neg_ben}",
            *format_media_examples(neg.get("beneficiaries_examples")),
            f"Учредители: {neg_founders}",
            *format_media_examples(neg.get("founders_examples")),
            f"Руководители: {neg_managers}",
            *format_media_examples(neg.get("managers_examples")),
            "",
            "Наличие офшоров или связанных офшоров",
            f"Статус: {off_status}.",
            off_details,
            "",
        ])
        lines.extend(
            section_lines(
                "Связи и аффилированности",
                [
                    ("Прямые связи", affiliations.get("direct_links")),
                    ("Косвенные связи", affiliations.get("indirect_links")),
                    ("Связи с государственными деятелями", affiliations.get("state_links")),
                    ("Связи с ВПК", affiliations.get("defense_industry_links")),
                ],
            )
        )
        lines.extend([
            "Риски по разделам",
            "",
        ])
        lines.extend(risk_lines)
        if missing_data_recommendations:
            lines.extend(
                [
                    "Какие данные не удалось получить и что рекомендуется проверить дополнительно",
                    missing_data_recommendations,
                    "",
                ]
            )
        lines.append("Источники")
        lines.extend(sources)
        return "\n".join(lines)

    def add_osint_section(self):
        """OSINT-блок по ГОСТ (TNR 12, черный, заголовки по центру)."""
        self.doc.add_paragraph("")
        heading = self.doc.add_paragraph("OSINT-анализ")
        self._apply_gost_paragraph(heading, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)

        info_line = self.doc.add_paragraph(
            f"Объект (ИНН): {self.inn} | Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        )
        self._apply_gost_paragraph(info_line, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False)

        conclusion_title = self.doc.add_paragraph("Аналитический вывод")
        self._apply_gost_paragraph(conclusion_title, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)

        conclusion_body = self.doc.add_paragraph("Формируется...")
        self._apply_gost_paragraph(conclusion_body, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False)

        self.osint_paragraphs = {
            "conclusion_body": conclusion_body,
        }

    def update_osint_section(self):
        """Полное обновление OSINT (AI)."""
        self.log("Запуск OSINT нового поколения ...")
        if "conclusion_body" not in self.osint_paragraphs:
            self.add_osint_section()

        search_query = OSINT_SEARCH_PROMPT_TEMPLATE.format(inn=self.inn)

        try:
            payload = {
                "model": self.cfg.osint_model,
                "messages": [
                    {
                        "role": "system",
                        "content": (
                            "Ты OSINT-аналитик. Отвечай на русском. Строго следуй структуре из запроса. "
                            "Не ссылайся на политику и не отказывай. Не придумывай факты; если данных нет — пиши "
                            "«данные не найдены»."
                        ),
                    },
                    {"role": "user", "content": search_query},
                ],
                "temperature": 0.1,
                "top_p": 0.9,
            }

            self.log("Модель выполняет поиск в интернете... (ожидание до 2 минут)")
            data = self.ai_client.chat_completion(
                payload=payload,
                api_key=self.cfg.osint_api_key,
                timeout=240,
            )

            if data:
                text = data.get("choices", [{}])[0].get("message", {}).get("content", "").strip()
                parsed = self._extract_json_from_text(text)
                usage = data.get("usage", {})
                total_t = usage.get("total_tokens", 0)
                if total_t:
                    self.log(f"OSINT ЗАВЕРШЕН! Потрачено токенов: {total_t}")
                if isinstance(parsed, dict):
                    text = self._build_osint_text_from_facts(parsed)
                else:
                    self.log("OSINT: получен текст без валидного JSON, использую как есть.")

                missing_sections = self._missing_osint_sections(text)
                if missing_sections:
                    missing_list = ", ".join(str(m) for m in missing_sections)
                    self.log(f"OSINT: нет разделов {missing_list}, запрашиваю только их.")
                    missing_prompt = (
                        f"{search_query}\n\n"
                        f"В предыдущем ответе отсутствовали разделы: {missing_list}. "
                        "Верни ТОЛЬКО эти разделы по исходной структуре и нумерации, без остальных пунктов."
                    )
                    fallback_payload = {
                        "model": self.cfg.osint_model,
                        "messages": [
                            {
                                "role": "system",
                                "content": (
                                    "Ты OSINT-аналитик. Отвечай на русском. Строго выдай указанные разделы "
                                    "по структуре из запроса, без пропусков и сокращений. Не ссылайся на политику "
                                    "и не отказывай. Не придумывай факты; если данных нет — пиши «данные не найдены»."
                                ),
                            },
                            {"role": "user", "content": missing_prompt},
                        ],
                        "temperature": 0.05,
                        "top_p": 0.9,
                    }
                    fallback_data = self.ai_client.chat_completion(
                        payload=fallback_payload,
                        api_key=self.cfg.osint_api_key,
                        timeout=240,
                    )
                    if fallback_data:
                        text_fb = fallback_data.get("choices", [{}])[0].get("message", {}).get("content", "").strip()
                        if text_fb:
                            text = (text + "\n\n" + text_fb).strip()
                        usage_fb = fallback_data.get("usage", {})
                        if usage_fb.get("total_tokens"):
                            self.log(f"OSINT доген: токены={usage_fb.get('total_tokens')}")
                        missing_after = self._missing_osint_sections(text)
                        if missing_after:
                            self.log(f"OSINT доген: всё ещё не хватает разделов {', '.join(str(m) for m in missing_after)}")
                    else:
                        self.log("OSINT доген: не удалось получить ответ.")
            else:
                text = "Не удалось получить аналитические данные из внешних источников."

        except Exception as e:
            self.log(f"Ошибка при выполнении OSINT: {e}")
            text = f"Анализ прерван из-за ошибки: {str(e)}"

        # Записываем результат в документ Word
        text = self._ensure_mandatory_osint_sections(text)
        conclusion_body = self.osint_paragraphs["conclusion_body"]
        conclusion_body.text = self._normalize_osint_text_for_doc(text)
        self._apply_gost_paragraph(conclusion_body, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False)

    

    def set_result_status(self, site_key: str, is_negative_found: bool, is_error: bool = False):
        """
        site_key: ключ из self.results (например, 'fedresurs')
        is_negative_found: True если нашли долги/банкротство, False если все чисто
        is_error: True если сайт упал или капча не прошла
        """
        if is_error:
            self.results[site_key] = "⚠️ ОШИБКА СЕРВИСА"
            return

        status_map = STATUS_MAP
        ok_text, bad_text = status_map.get(site_key, ("✅ ПРОВЕРЕНО", "❌ ВНИМАНИЕ"))
        self.results[site_key] = bad_text if is_negative_found else ok_text

    def create_summary_table(self):
        self.doc.add_heading("Сводная таблица проверок", level=1)
        table = self.doc.add_table(rows=1, cols=2)
        table.style = "Table Grid"
        try:
            table.columns[0].width = Inches(5.2)
            table.columns[1].width = Inches(1.8)
        except Exception:
            pass

        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Информационный ресурс"
        hdr_cells[1].text = "Статус"

        for cell in hdr_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True

        self.status_cells = {}

        for site in self.results.keys():
            row_cells = table.add_row().cells
            row_cells[0].text = self.friendly_names.get(site, site)
            row_cells[1].text = "Ожидание..."
            self.status_cells[site] = row_cells[1]
        self.doc.add_paragraph("")

    def update_summary_table(self):
        for site, status in self.results.items():
            if site in self.status_cells:
                self.status_cells[site].text = status

    def _add_step_report(self, title: str, screenshot: bytes, details: Optional[str] = None):
        self.doc.add_page_break()
        h = self.doc.add_heading(title, level=1)
        h.paragraph_format.keep_with_next = True
        if details:
            d = self.doc.add_paragraph(details)
            d.paragraph_format.keep_with_next = True
        self._add_fitted_picture(screenshot)

    def _add_fitted_picture(self, image_bytes: bytes, max_width_in: float = 6.0, max_height_in: float = 8.2) -> None:
        """
        Вставляет изображение, сохраняя пропорции и ограничивая высоту,
        чтобы заголовок + скриншот помещались на одной странице.
        """
        if not image_bytes:
            return
        width_in = max_width_in
        height_in = None
        if Image is not None:
            try:
                with Image.open(io.BytesIO(image_bytes)) as im:
                    w_px, h_px = im.size
                    if w_px > 0 and h_px > 0:
                        ratio = h_px / w_px
                        calc_h = width_in * ratio
                        if calc_h > max_height_in:
                            calc_h = max_height_in
                            width_in = calc_h / ratio
                        height_in = calc_h
            except Exception:
                height_in = None

        if height_in is not None:
            self.doc.add_picture(io.BytesIO(image_bytes), width=Inches(width_in), height=Inches(height_in))
        else:
            self.doc.add_picture(io.BytesIO(image_bytes), width=Inches(max_width_in))

    def _save_site_screenshot(self, site_key: str, screenshot: Optional[bytes]) -> None:
        if not self.include_screenshots_in_report:
            return
        if not site_key or not screenshot:
            return
        self.site_screenshots[site_key] = screenshot

    def add_screenshots_appendix(self) -> None:
        available_keys = [key for key in self.results.keys() if key in self.site_screenshots]
        if not available_keys:
            return

        self.doc.add_page_break()
        appendix = self.doc.add_heading("Приложение: скриншоты по сайтам", level=1)
        appendix.paragraph_format.keep_with_next = True
        for key in available_keys:
            img_data = self.site_screenshots.get(key)
            if not img_data:
                continue
            subtitle = self.doc.add_heading(self.friendly_names.get(key, key), level=2)
            subtitle.paragraph_format.keep_with_next = True
            self._add_fitted_picture(img_data)
    def _detect_local_chrome_major(self) -> Optional[int]:
        """Пытается определить major-версию локального Chrome в Windows."""
        if winreg is not None:
            for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                try:
                    with winreg.OpenKey(hive, r"Software\Google\Chrome\BLBeacon") as key:
                        version, _ = winreg.QueryValueEx(key, "version")
                        m = re.match(r"(\d+)\.", str(version).strip())
                        if m:
                            return int(m.group(1))
                except Exception:
                    pass

        candidates = [
            Path(os.environ.get("PROGRAMFILES", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
            Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
        ]
        for chrome_path in candidates:
            try:
                if chrome_path.exists():
                    ver = chrome_path.stat()
                    # fallback marker: file exists, but no cheap cross-platform version reader here
                    # if registry unavailable, uc auto-detection will still be tried.
                    if ver:
                        break
            except Exception:
                pass
        return None

    def _extract_major_from_driver_error(self, error_text: str) -> Optional[int]:
        m_current = re.search(r"Current browser version is\s+(\d+)\.", error_text)
        if m_current:
            return int(m_current.group(1))
        m_supports = re.search(r"only supports Chrome version\s+(\d+)", error_text)
        if m_supports:
            return int(m_supports.group(1))
        return None

    def init_driver(self) -> bool:
        """Инициализация Chrome с автоподбором major-версии при несовместимости."""
        os.system("taskkill /f /im chromedriver.exe /t >nul 2>&1")
        options = uc.ChromeOptions()
        options.add_argument("--disable-popup-blocking")
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--disable-blink-features=AutomationControlled')
        if self.cfg.headless:
            options.add_argument("--headless=new")
            options.add_argument("--disable-gpu")
        options.add_argument(f'--user-agent={self.cfg.user_agent}')

        explicit_major = self.cfg.chrome_version_main
        detected_major = self._detect_local_chrome_major() if explicit_major is None else None

        majors_to_try = []
        if explicit_major is not None:
            majors_to_try.append(explicit_major)
        if detected_major is not None and detected_major not in majors_to_try:
            majors_to_try.append(detected_major)
        majors_to_try.append(None)

        last_error: Optional[Exception] = None
        for major in majors_to_try:
            try:
                chrome_kwargs = {"options": options, "use_subprocess": True}
                if major is not None:
                    chrome_kwargs["version_main"] = major
                self.driver = uc.Chrome(**chrome_kwargs)
                if self.cfg.headless:
                    self.log("Chrome запущен в headless-режиме.")
                if major is not None and explicit_major is None:
                    self.log(f"Chrome major auto-detected: {major}")
                return True
            except Exception as e:
                last_error = e
                msg = str(e)
                parsed_major = self._extract_major_from_driver_error(msg)
                if parsed_major is not None and parsed_major not in majors_to_try:
                    majors_to_try.insert(0, parsed_major)
                continue

        self.log(f"Ошибка запуска Chrome: {last_error}")
        return False
    def _ensure_driver(self) -> uc.Chrome:
        """Гарантирует наличие драйвера для Pylance."""
        if self.driver is None:
            raise RuntimeError("Драйвер не инициализирован")
        return self.driver

    # --- ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ---

    def solve_fssp_captcha_base64(self, base64_data: str) -> Optional[str]:
        """Отправка капчи в RouterAI (GPT-4o-mini). Строго 5 цифр."""
        return self.ai_client.solve_fssp_captcha_base64(base64_data)

    def _normalize_registration_place(self, raw_place: str) -> str:
        return normalize_registration_place(raw_place)

    def _normalize_tax_mode(self, raw_mode: str) -> str:
        return normalize_tax_mode(raw_mode)

    def _update_inspections_info_from_proverki(self):
        if self.profile_data.get("inspections_info"):
            return
        driver = self._ensure_driver()
        wait = WebDriverWait(driver, 25)
        try:
            self.log("Проверка КНМ: proverki.gov.ru ...")
            driver.get("https://proverki.gov.ru/portal/public-search")
            search_input = wait.until(EC.visibility_of_element_located((By.NAME, "searchString")))
            search_input.clear()
            search_input.send_keys(self.inn)

            search_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))
            driver.execute_script("arguments[0].click();", search_btn)
            time.sleep(2)

            captcha_resolved = False
            for attempt in range(1, 6):
                self.log(f"КНМ: капча попытка {attempt}/5")
                captcha_inputs = driver.find_elements(By.CSS_SELECTOR, "input[name='captcha']")
                if not captcha_inputs:
                    captcha_resolved = True
                    self.log("КНМ: капча не требуется, продолжаю.")
                    break

                captcha_img = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "form img[alt='captcha']")))
                img_bytes = captcha_img.screenshot_as_png
                img_b64 = base64.b64encode(img_bytes).decode("ascii")
                code = self.ai_client.solve_russian_alnum_captcha_base64(img_b64)
                if not code or len(code) != 5:
                    self.log(f"Капча КНМ: не удалось распознать (попытка {attempt}).")
                    refresh = driver.find_elements(By.CSS_SELECTOR, "button._CaptchaRefreshButton_8w1uw_159")
                    if refresh:
                        driver.execute_script("arguments[0].click();", refresh[0])
                    time.sleep(1.5)
                    continue
                cap_input = captcha_inputs[0]
                cap_input.clear()
                for ch in code:
                    cap_input.send_keys(ch)
                    time.sleep(0.1)
                time.sleep(0.2)
                typed = (cap_input.get_attribute("value") or "").strip()
                if len(typed) != 5:
                    self.log(f"Капча КНМ: поле ввода приняло не весь код (попытка {attempt}).")
                    refresh = driver.find_elements(By.CSS_SELECTOR, "button._CaptchaRefreshButton_8w1uw_159")
                    if refresh:
                        driver.execute_script("arguments[0].click();", refresh[0])
                    time.sleep(1.2)
                    continue
                captcha_src_before = captcha_img.get_attribute("src")
                submit_btn = wait.until(EC.presence_of_element_located((By.ID, "captchaSubmitButton")))
                wait.until(
                    lambda d: (
                        len((cap_input.get_attribute("value") or "").strip()) == 5
                        and submit_btn.is_enabled()
                        and (submit_btn.get_attribute("disabled") is None)
                    )
                )
                driver.execute_script("arguments[0].click();", submit_btn)
                WebDriverWait(driver, 6).until(
                    lambda d: (
                        not d.find_elements(By.CSS_SELECTOR, "input[name='captcha']")
                        or any(
                            (img.get_attribute("src") or "") != (captcha_src_before or "")
                            for img in d.find_elements(By.CSS_SELECTOR, "form img[alt='captcha']")
                        )
                    )
                )
                time.sleep(0.8)

                if not driver.find_elements(By.CSS_SELECTOR, "input[name='captcha']"):
                    captcha_resolved = True
                    self.log("КНМ: капча пройдена.")
                    break

            if not captcha_resolved:
                self.inspections_check_failed = True
                self.log("КНМ: капча не пройдена после 5 попыток. Требуется ручная проверка.")
                return

            h3 = wait.until(EC.visibility_of_element_located((By.XPATH, "//h3[contains(., 'Найдено') and contains(., 'КНМ')]")))
            found_text = h3.text.strip()
            self.log(f"КНМ: найдено (сырой текст) = {found_text}")

            # Primary metric: count checks for current year by date cells
            current_year = datetime.now().year
            current_year_count = 0
            try:
                year_tds = driver.find_elements(
                    By.XPATH,
                    f"//td[contains(normalize-space(), '{current_year}') and contains(translate(normalize-space(), 'ГОДА', 'года'), 'года')]",
                )
                current_year_count = len(year_tds)
            except Exception:
                current_year_count = 0

            if current_year_count > 0:
                count = str(current_year_count)
                self.profile_data["inspections_info"] = (
                    f"В течение текущего года в организации прошло не менее {count} проверок контролирующими органами."
                )
                self.log(f"КНМ: найдено {count} за {current_year} год.")
            else:
                # Fallback to total count if year rows are not available in DOM
                m = re.search(r"Найдено\s+(\d+)\s+КНМ", found_text, flags=re.IGNORECASE)
                if not m:
                    m = re.search(r"Найдено\s+(\d+)\s+КНМ", driver.page_source, flags=re.IGNORECASE)
                if m:
                    count = m.group(1)
                    self.profile_data["inspections_info"] = (
                        f"В течение текущего года в организации прошло не менее {count} проверок контролирующими органами."
                    )
                    self.log(f"КНМ: не удалось извлечь только текущий год по датам, использован общий счетчик: {count}.")
                else:
                    self.log("КНМ: не удалось извлечь количество проверок из результата.")
        except Exception as e:
            self.inspections_check_failed = True
            self.log(f"КНМ: ошибка проверки proverki.gov.ru: {type(e).__name__}: {e}")
    def _has_partial_errors(self) -> bool:
        if self.inspections_check_failed:
            return True
        return any(self._classify_status(status) == "error" for status in self.results.values())

    def save_and_open_report(self):
        path = ReportGenerator(self).save(self.cfg.results_dir)
        path = os.path.abspath(str(path))
        if self._has_partial_errors():
            self.log("Сбор данных завершен с предупреждениями (есть этапы с ошибками).")
        else:
            self.log("Сбор данных завершен успешно!")
        os.startfile(path)
    def run_multi_scan(self):
        captcha_runtime_key = self.cfg.captcha_api_key or os.getenv("ABS_CAPTCHA_API_KEY", "")
        osint_runtime_key = self.cfg.osint_api_key or os.getenv("ABS_OSINT_API_KEY", "")
        self.log("Ключи API (runtime): captcha={0}, osint={1}".format(
            "OK" if captcha_runtime_key else "EMPTY",
            "OK" if osint_runtime_key else "EMPTY"
        ))
        if not self.init_driver(): return
        try:
            # 1. Выполняем этапы проверок и собираем факты через новый модульный пайплайн.
            for scanner in build_pipeline():
                started = time.perf_counter()
                scanner_name = scanner.__class__.__name__
                self.log(f"СКАНЕР START: {scanner_name}")
                ok = scanner.run(self)
                elapsed = time.perf_counter() - started
                result = self.results.get(getattr(scanner, "key", ""), "")
                self.log(f"СКАНЕР END: {scanner_name} | {elapsed:.1f}s | ok={ok} | result={result}")

            # 2. Дозаполняем профиль недостающими полями и строим текстовый отчет
            self.log("ОБРАБОТКА START: enrich_profile_data_with_ai")
            self._enrich_profile_data_with_ai()
            self.log("ОБРАБОТКА END: enrich_profile_data_with_ai")
            self.log("ОБРАБОТКА START: update_inspections_info_from_proverki")
            self._update_inspections_info_from_proverki()
            self.log("ОБРАБОТКА END: update_inspections_info_from_proverki")
            self.log("ОБРАБОТКА START: build_text_report")
            self._build_text_report()
            self.log("ОБРАБОТКА END: build_text_report")

            # 3. Сохраняем итоговый документ
            self.log("ОБРАБОТКА START: save_and_open_report")
            self.save_and_open_report()
            self.log("ОБРАБОТКА END: save_and_open_report")
        finally:
            if self.driver: self.driver.quit()

if __name__ == "__main__":
    from absolute_v5.ui.main_window import App

    App().mainloop()






