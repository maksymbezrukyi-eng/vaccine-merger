import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import io
from datetime import datetime

# ─────────────────────────────────────────────
# НАЛАШТУВАННЯ СТОРІНКИ
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Зведення щеплень",
    page_icon="💉",
    layout="wide"
)

st.title("💉 Зведення звітів про виконання щеплень")
st.markdown("Завантажте файли Excel від ЗОЗ — програма перевірить їх і зведе в один звіт.")
st.divider()


# ─────────────────────────────────────────────
# ДОПОМІЖНІ ФУНКЦІЇ
# ─────────────────────────────────────────────

def safe_num(val):
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        return val
    try:
        return float(str(val).replace(",", ".").strip())
    except Exception:
        return 0


def extract_facility_data(file_bytes, name, edrpou):
    """Витягує структуровані дані з файлу для дашбордів."""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws_exec = wb["Виконання"]
    ws_rem  = wb["Залишки"]
    ws_zvit = wb["Зведений звіт"]

    # Охоплення: з Зведеного звіту (рядки 11-119)
    coverage = []
    for row in range(11, 120):
        vaccine  = ws_zvit.cell(row=row, column=1).value
        age      = ws_zvit.cell(row=row, column=2).value
        plan     = ws_zvit.cell(row=row, column=3).value
        executed = ws_zvit.cell(row=row, column=4).value
        pct      = ws_zvit.cell(row=row, column=6).value
        if vaccine and isinstance(plan, (int, float)) and plan > 0:
            label = str(vaccine).strip()
            if age:
                label += f" ({str(age).strip()})"
            coverage.append({
                "label":    label,
                "vaccine":  str(vaccine).strip(),
                "age":      str(age or "").strip(),
                "plan":     safe_num(plan),
                "executed": safe_num(executed),
                "pct":      safe_num(pct),
            })

    # Залишки вакцин
    stocks = []
    for row in range(11, 38):
        vaccine = ws_rem.cell(row=row, column=1).value
        if not vaccine:
            continue
        closing = safe_num(ws_rem.cell(row=row, column=4).value)
        used    = safe_num(ws_rem.cell(row=row, column=6).value)
        opening = safe_num(ws_rem.cell(row=row, column=2).value)
        received= safe_num(ws_rem.cell(row=row, column=3).value)
        stocks.append({
            "vaccine": str(vaccine).strip(),
            "closing": closing,
            "used":    used,
            "opening": opening,
            "received": received,
        })

    # Відмови (рядки 8-13, col T=20)
    refusal_map = {
        8: "Туберкульоз", 9: "Поліомієліт", 10: "Гепатит В",
        11: "Кашлюк, дифтерія, правець", 12: "Гемофільна інфекція",
        13: "Кір, паротит, краснуха"
    }
    refusals = []
    for row, disease in refusal_map.items():
        count = safe_num(ws_exec.cell(row=row, column=20).value)
        refusals.append({"disease": disease, "count": count})

    # Протипокази (рядки 8-10, col P=16, Q=17)
    temp = sum(safe_num(ws_exec.cell(row=r, column=16).value) for r in range(8, 11))
    perm = sum(safe_num(ws_exec.cell(row=r, column=17).value) for r in range(8, 11))

    return {
        "name":     name,
        "edrpou":   edrpou,
        "coverage": coverage,
        "stocks":   stocks,
        "refusals": refusals,
        "temp_contraindications": temp,
        "perm_contraindications": perm,
    }


def validate_file(file_bytes, filename):
    errors = []
    warnings = []
    name = "—"
    edrpou = "—"
    period = "—"
    wb = None

    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)

        required = ["План", "Виконання", "Залишки", "Зведений звіт", "Аркуш1"]
        missing = [s for s in required if s not in wb.sheetnames]
        if missing:
            errors.append(f"Відсутні аркуші: {', '.join(missing)}")
            return dict(file=filename, name=name, edrpou=edrpou, period=period,
                        status="🔴 Помилки", errors=errors, warnings=warnings, wb=wb)

        ws_plan = wb["План"]
        ws_exec = wb["Виконання"]
        ws_rem  = wb["Залишки"]

        raw_name = ws_plan["D8"].value
        name = str(raw_name).strip() if raw_name else "—"
        if not raw_name or name == "":
            errors.append("Порожня назва закладу (План!D8)")

        raw_edrpou = ws_plan["E8"].value
        edrpou = str(raw_edrpou).strip().lstrip("'") if raw_edrpou else "—"
        if not raw_edrpou:
            errors.append("Порожній код ЄДРПОУ (План!E8)")
        elif not edrpou.isdigit() or len(edrpou) not in (7, 8, 9, 10):
            warnings.append(f"Код ЄДРПОУ '{edrpou}' має нестандартну довжину")

        raw_period = ws_exec["F6"].value
        period = str(raw_period) if raw_period else "—"
        if not raw_period:
            errors.append("Відсутній звітний період (Виконання!F6)")

        name_exec = ws_exec["C4"].value
        if name_exec and raw_name and str(name_exec).strip() != name:
            warnings.append("Назва закладу різниться між аркушами «План» і «Виконання»")

        for row in range(8, 105):
            val = ws_exec.cell(row=row, column=5).value
            if isinstance(val, (int, float)) and val < 0:
                vac = ws_exec.cell(row=row, column=3).value
                age = ws_exec.cell(row=row, column=4).value
                errors.append(f"Від'ємна кількість щеплень: {vac} / {age} = {val}")

        for row in range(11, 38):
            vaccine = ws_rem.cell(row=row, column=1).value
            if not vaccine:
                continue
            b     = safe_num(ws_rem.cell(row=row, column=2).value)
            c     = safe_num(ws_rem.cell(row=row, column=3).value)
            d     = ws_rem.cell(row=row, column=4).value
            f_val = safe_num(ws_rem.cell(row=row, column=6).value)
            g     = safe_num(ws_rem.cell(row=row, column=7).value)
            h     = safe_num(ws_rem.cell(row=row, column=8).value)
            expected = b + c + g + h - f_val
            if isinstance(d, (int, float)):
                if abs(d - expected) > 0.5:
                    errors.append(f"Залишки — помилка балансу для «{str(vaccine).strip()}»"
                                  f" (є {d}, має бути {expected})")
                if d < 0:
                    errors.append(f"Залишки — від'ємний залишок для «{str(vaccine).strip()}»")

        for row in range(8, 11):
            l_val = safe_num(ws_exec.cell(row=row, column=12).value)
            m_val = safe_num(ws_exec.cell(row=row, column=13).value)
            if m_val > l_val > 0:
                errors.append(f"Рядок {row}: «Отримали КДП-3» ({m_val}) > «Народилося за 7 міс.» ({l_val})")

        for row in range(8, 11):
            p = safe_num(ws_exec.cell(row=row, column=16).value)
            q = safe_num(ws_exec.cell(row=row, column=17).value)
            r = ws_exec.cell(row=row, column=18).value
            if isinstance(r, (int, float)) and abs(r - (p + q)) > 0.5:
                errors.append(f"Рядок {row}: протипокази ВСЬОГО ({r}) ≠ Тимчасові+Постійні ({p + q})")

    except Exception as e:
        errors.append(f"Не вдалось прочитати файл: {e}")

    status = "🔴 Помилки" if errors else ("🟡 Попередження" if warnings else "🟢 OK")
    return dict(file=filename, name=name, edrpou=edrpou, period=period,
                status=status, errors=errors, warnings=warnings, wb=wb)


def aggregate_files(file_bytes_list, org_name, org_edrpou, report_period):
    workbooks = [load_workbook(io.BytesIO(fbytes), data_only=True)
                 for _, fbytes in file_bytes_list]
    template_wb = load_workbook(io.BytesIO(file_bytes_list[0][1]))

    ws_out = template_wb["Виконання"]
    for row in range(8, 105):
        ws_out.cell(row=row, column=5).value = 0
    for row in range(8, 11):
        for col in [10, 11, 12, 13, 16, 17, 18]:
            ws_out.cell(row=row, column=col).value = 0
    for row in range(8, 14):
        ws_out.cell(row=row, column=20).value = 0

    for wb in workbooks:
        ws = wb["Виконання"]
        for row in range(8, 105):
            cur = ws_out.cell(row=row, column=5).value or 0
            ws_out.cell(row=row, column=5).value = cur + safe_num(ws.cell(row=row, column=5).value)
        for row in range(8, 11):
            for col in [10, 11, 12, 13, 16, 17]:
                cur = ws_out.cell(row=row, column=col).value or 0
                ws_out.cell(row=row, column=col).value = cur + safe_num(ws.cell(row=row, column=col).value)
        for row in range(8, 14):
            cur = ws_out.cell(row=row, column=20).value or 0
            ws_out.cell(row=row, column=20).value = cur + safe_num(ws.cell(row=row, column=20).value)

    for row in range(8, 11):
        p = ws_out.cell(row=row, column=16).value or 0
        q = ws_out.cell(row=row, column=17).value or 0
        ws_out.cell(row=row, column=18).value = p + q

    group_sums = {11: list(range(8,12)), 23: list(range(12,24)), 35: list(range(24,36)),
                  42: list(range(36,43)), 48: list(range(43,49)), 61: list(range(49,62))}
    for sum_row, rows in group_sums.items():
        ws_out.cell(row=sum_row, column=7).value = sum(safe_num(ws_out.cell(row=r, column=5).value) for r in rows)
    ws_out.cell(row=99,  column=7).value = sum(safe_num(ws_out.cell(row=r, column=5).value) for r in [62,64,66,68,70,72,74,76,78])
    ws_out.cell(row=100, column=7).value = sum(safe_num(ws_out.cell(row=r, column=5).value) for r in [63,65,67,69,71,73,75,77,79,81])
    ws_out.cell(row=101, column=7).value = safe_num(ws_out.cell(row=80, column=5).value)
    ws_out.cell(row=102, column=7).value = sum(safe_num(ws_out.cell(row=r, column=5).value) for r in [82,84,86,88,90,92,94,96,98,100,102])
    ws_out.cell(row=103, column=7).value = sum(safe_num(ws_out.cell(row=r, column=5).value) for r in [83,85,87,89,91,93,95,97,99,101])
    ws_out.cell(row=104, column=7).value = sum(safe_num(ws_out.cell(row=r, column=5).value) for r in [103, 104])

    ws_rem_out = template_wb["Залишки"]
    for row in range(11, 38):
        for col in [2, 3, 5, 6, 7, 8]:
            ws_rem_out.cell(row=row, column=col).value = 0
    for wb in workbooks:
        ws = wb["Залишки"]
        for row in range(11, 38):
            if not ws_rem_out.cell(row=row, column=1).value:
                continue
            for col in [2, 3, 5, 6, 7, 8]:
                cur = ws_rem_out.cell(row=row, column=col).value or 0
                ws_rem_out.cell(row=row, column=col).value = cur + safe_num(ws.cell(row=row, column=col).value)
    for row in range(11, 38):
        if not ws_rem_out.cell(row=row, column=1).value:
            continue
        b = safe_num(ws_rem_out.cell(row=row, column=2).value)
        c = safe_num(ws_rem_out.cell(row=row, column=3).value)
        f = safe_num(ws_rem_out.cell(row=row, column=6).value)
        g = safe_num(ws_rem_out.cell(row=row, column=7).value)
        h = safe_num(ws_rem_out.cell(row=row, column=8).value)
        ws_rem_out.cell(row=row, column=4).value = b + c + g + h - f

    ws_plan_out = template_wb["План"]
    for row in range(11, 47):
        ws_plan_out.cell(row=row, column=6).value = 0
    for wb in workbooks:
        ws = wb["План"]
        for row in range(11, 47):
            cur = ws_plan_out.cell(row=row, column=6).value or 0
            ws_plan_out.cell(row=row, column=6).value = cur + safe_num(ws.cell(row=row, column=6).value)

    ws_zvit_out = template_wb["Зведений звіт"]
    for row in range(11, 120):
        for col in [3, 4, 5]:
            ws_zvit_out.cell(row=row, column=col).value = 0
    for wb in workbooks:
        ws = wb["Зведений звіт"]
        for row in range(11, 120):
            for col in [3, 4, 5]:
                cur = ws_zvit_out.cell(row=row, column=col).value or 0
                ws_zvit_out.cell(row=row, column=col).value = cur + safe_num(ws.cell(row=row, column=col).value)
    for row in range(11, 120):
        c_val = ws_zvit_out.cell(row=row, column=3).value
        e_val = ws_zvit_out.cell(row=row, column=5).value
        if c_val and isinstance(c_val, (int, float)) and c_val > 0 and isinstance(e_val, (int, float)):
            ws_zvit_out.cell(row=row, column=6).value = round(e_val / c_val * 100, 1)
        else:
            ws_zvit_out.cell(row=row, column=6).value = None

    template_wb["План"]["D8"].value          = org_name
    template_wb["План"]["E8"].value          = org_edrpou
    template_wb["Виконання"]["C4"].value     = org_name
    template_wb["Виконання"]["F4"].value     = org_edrpou
    template_wb["Виконання"]["F6"].value     = report_period
    template_wb["Залишки"]["A4"].value       = org_name
    template_wb["Залишки"]["D4"].value       = org_edrpou
    template_wb["Залишки"]["D6"].value       = report_period
    template_wb["Зведений звіт"]["A3"].value = org_name
    template_wb["Зведений звіт"]["D3"].value = org_edrpou
    template_wb["Зведений звіт"]["D5"].value = report_period

    output = io.BytesIO()
    template_wb.save(output)
    output.seek(0)
    return output.getvalue()


def build_coverage_excel(facility_data_list, selected_facilities, selected_indicators):
    """Генерує таблицю охоплення по ЗОЗ у вигляді xlsx."""
    # Збираємо всі унікальні індикатори (label вакцини+вік)
    all_labels = []
    seen = set()
    for fd in facility_data_list:
        for item in fd["coverage"]:
            if item["label"] not in seen:
                all_labels.append(item["label"])
                seen.add(item["label"])

    # Фільтруємо за вибраними індикаторами
    if selected_indicators:
        all_labels = [l for l in all_labels if l in selected_indicators]

    # Фільтруємо заклади
    facilities = [fd for fd in facility_data_list if fd["name"] in selected_facilities]

    wb = Workbook()
    ws = wb.active
    ws.title = "Охоплення"

    # Стилі
    header_fill   = PatternFill("solid", fgColor="1F4E79")
    header_font   = Font(color="FFFFFF", bold=True, size=10)
    subhead_fill  = PatternFill("solid", fgColor="2E75B6")
    subhead_font  = Font(color="FFFFFF", bold=True, size=10)
    green_fill    = PatternFill("solid", fgColor="C6EFCE")
    yellow_fill   = PatternFill("solid", fgColor="FFEB9C")
    red_fill      = PatternFill("solid", fgColor="FFC7CE")
    center        = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin          = Side(style="thin", color="BFBFBF")
    border        = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Заголовок (рядок 1): зливаємо по ширині всіх стовпців
    total_cols = 1 + len(all_labels)
    end_col_letter = get_col_letter(total_cols)
    ws.merge_cells(f"A1:{end_col_letter}1")
    ws["A1"] = f"Таблиця охоплення щепленнями — {st.session_state.get('report_label', '')}"
    ws["A1"].fill = header_fill
    ws["A1"].font = Font(color="FFFFFF", bold=True, size=12)
    ws["A1"].alignment = center

    # Рядок заголовків (рядок 2): A2 = "Заклад", далі — показники
    ws["A2"] = "Заклад"
    ws["A2"].fill = subhead_fill
    ws["A2"].font = subhead_font
    ws["A2"].alignment = center
    ws["A2"].border = border
    ws.column_dimensions["A"].width = 35

    for col_idx, label in enumerate(all_labels, start=2):
        cell = ws.cell(row=2, column=col_idx, value=label)
        cell.fill = subhead_fill
        cell.font = subhead_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
        ws.column_dimensions[get_col_letter(col_idx)].width = 16

    ws.row_dimensions[2].height = 60

    # Дані: кожен рядок = один ЗОЗ
    for row_idx, fd in enumerate(facilities, start=3):
        cell_a = ws.cell(row=row_idx, column=1, value=fd["name"])
        cell_a.alignment = Alignment(vertical="center", wrap_text=True)
        cell_a.border = border

        for col_idx, label in enumerate(all_labels, start=2):
            pct = None
            for item in fd["coverage"]:
                if item["label"] == label:
                    pct = item["pct"]
                    break

            cell = ws.cell(row=row_idx, column=col_idx)
            if pct is not None:
                cell.value = pct / 100
                cell.number_format = "0.0%"
                if pct >= 95:
                    cell.fill = green_fill
                elif pct >= 85:
                    cell.fill = yellow_fill
                else:
                    cell.fill = red_fill
            else:
                cell.value = "—"
            cell.alignment = center
            cell.border = border

    # Легенда
    legend_row = len(facilities) + 4
    ws.cell(row=legend_row,   column=1, value="Легенда:").font = Font(bold=True)
    ws.cell(row=legend_row+1, column=1, value="≥ 95% — виконання плану").fill  = green_fill
    ws.cell(row=legend_row+2, column=1, value="85–95% — потребує уваги").fill  = yellow_fill
    ws.cell(row=legend_row+3, column=1, value="< 85% — критично").fill         = red_fill

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def get_col_letter(col_idx):
    """Converts column index to Excel letter (1=A, 27=AA, etc.)"""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


# ─────────────────────────────────────────────
# ІНТЕРФЕЙС
# ─────────────────────────────────────────────

# ── КРОК 1: Реквізити ────────────────────────────────────────────────
st.header("🏥 Крок 1 — Реквізити організації що формує зведений звіт")

with st.form("org_form"):
    col1, col2 = st.columns([3, 1])
    with col1:
        org_name = st.text_input("Назва закладу", placeholder="Наприклад: Тернопільський обласний ЦКПХ")
    with col2:
        org_edrpou = st.text_input("Код ЄДРПОУ", placeholder="12345678")

    col3, col4, col5 = st.columns(3)
    with col3:
        report_year = st.selectbox("Звітний рік", options=list(range(2024, 2031)), index=2)
    with col4:
        month_names = {1:"Січень",2:"Лютий",3:"Березень",4:"Квітень",5:"Травень",
                       6:"Червень",7:"Липень",8:"Серпень",9:"Вересень",10:"Жовтень",
                       11:"Листопад",12:"Грудень"}
        report_month = st.selectbox("Звітний місяць", options=list(month_names.keys()),
                                    format_func=lambda x: month_names[x], index=1)
    with col5:
        expected_count = st.number_input("Кількість ЗОЗ що мають подати звіт",
                                         min_value=1, max_value=500, value=None,
                                         step=1, placeholder="Введіть кількість...")

    submitted = st.form_submit_button("✅ Підтвердити реквізити", type="primary", use_container_width=True)

if submitted:
    form_errors = []
    if not org_name.strip():
        form_errors.append("Введіть назву закладу")
    if not org_edrpou.strip():
        form_errors.append("Введіть код ЄДРПОУ")
    elif not org_edrpou.strip().isdigit():
        form_errors.append("Код ЄДРПОУ повинен містити тільки цифри")
    if expected_count is None:
        form_errors.append("Введіть кількість ЗОЗ що мають подати звіт")
    if form_errors:
        for e in form_errors:
            st.error(f"❌ {e}")
    else:
        st.session_state["org_name"]       = org_name.strip()
        st.session_state["org_edrpou"]     = org_edrpou.strip()
        st.session_state["report_period"]  = datetime(report_year, report_month, 1)
        st.session_state["report_label"]   = f"{month_names[report_month]} {report_year}"
        st.session_state["expected_count"] = int(expected_count)

if "org_name" in st.session_state:
    st.success(
        f"✅ **{st.session_state['org_name']}**  |  "
        f"ЄДРПОУ: **{st.session_state['org_edrpou']}**  |  "
        f"Звітний період: **{st.session_state['report_label']}**  |  "
        f"Очікується ЗОЗ: **{st.session_state['expected_count']}**"
    )
    st.divider()

    # ── КРОК 2: Завантаження файлів ──────────────────────────────────
    st.header("📂 Крок 2 — Завантажте файли ЗОЗ")
    uploaded_files = st.file_uploader("Оберіть файли Excel (.xlsx)", type=["xlsx"],
                                      accept_multiple_files=True)

    if uploaded_files:
        expected        = st.session_state["expected_count"]
        submitted_count = len(uploaded_files)
        missing_count   = max(0, expected - submitted_count)
        percent         = min(100, round(submitted_count / expected * 100))

        st.subheader("📊 Стан подання звітів")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📋 Очікується", expected)
        m2.metric("📥 Подано",     submitted_count)
        m3.metric("⏳ Не подали",  missing_count, delta_color="inverse",
                  delta=f"-{missing_count}" if missing_count > 0 else None)
        m4.metric("✅ Виконання",  f"{percent}%")
        st.progress(percent / 100, text=f"Подано {submitted_count} з {expected} закладів ({percent}%)")
        if missing_count == 0:
            st.success("🎉 Усі заклади подали звіти!")
        else:
            st.warning(f"⚠️ Ще не подали звіт: **{missing_count}** заклад(ів).")

        st.divider()
        st.header("📋 Крок 3 — Список завантажених файлів")
        file_data = [{"Файл": f.name, "Розмір": f"{round(f.size/1024,1)} КБ",
                      "Статус": "⏳ Очікує перевірки"} for f in uploaded_files]
        st.dataframe(pd.DataFrame(file_data), use_container_width=True, hide_index=True)

        st.header("🔍 Крок 4 — Перевірка файлів")
        if st.button("▶️ Запустити перевірку", type="primary", use_container_width=True):
            files_bytes = [(f.name, f.read()) for f in uploaded_files]
            results = []
            pb = st.progress(0, text="Перевірка виконується...")
            for i, (fname, fbytes) in enumerate(files_bytes):
                r = validate_file(fbytes, fname)
                r["_bytes"] = fbytes
                results.append(r)
                pb.progress((i+1)/len(files_bytes), text=f"Перевірено {i+1} з {len(files_bytes)}: {fname}")
            st.session_state["results"] = results
            st.rerun()


# ─── Результати перевірки ─────────────────────────────────────────────
if "results" in st.session_state and "org_name" in st.session_state:
    results  = st.session_state["results"]
    expected = st.session_state["expected_count"]
    ok   = [r for r in results if r["status"] == "🟢 OK"]
    warn = [r for r in results if r["status"] == "🟡 Попередження"]
    bad  = [r for r in results if r["status"] == "🔴 Помилки"]
    good = [r for r in results if r["status"] in ("🟢 OK", "🟡 Попередження")]
    missing_count = max(0, expected - len(results))

    st.subheader("📊 Підсумок після перевірки")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("📋 Очікується",       expected)
    c2.metric("📥 Подано",           len(results))
    c3.metric("🟢 Без помилок",      len(ok))
    c4.metric("🟡 З попередженнями", len(warn))
    c5.metric("🔴 З помилками",      len(bad))

    if missing_count > 0:
        st.warning(f"⚠️ Ще не подали звіт: **{missing_count}** заклад(ів) з {expected}")

    periods    = list({r["period"] for r in results if r["period"] != "—"})
    if len(periods) > 1:
        st.warning(f"⚠️ Різні звітні періоди у файлах: {', '.join(periods)}")
    edrpous    = [r["edrpou"] for r in results if r["edrpou"] != "—"]
    duplicates = {e for e in edrpous if edrpous.count(e) > 1}
    if duplicates:
        st.error(f"❌ Дублікати ЄДРПОУ: {', '.join(duplicates)}")

    st.subheader("Детальні результати по файлах")
    for r in results:
        with st.expander(f"{r['status']}  |  {r['file']}  |  {r['name']}  |  ЄДРПОУ: {r['edrpou']}"):
            if r["errors"]:
                st.error("**Помилки:**")
                for e in r["errors"]:
                    st.write(f"• {e}")
            if r["warnings"]:
                st.warning("**Попередження:**")
                for w in r["warnings"]:
                    st.write(f"• {w}")
            if not r["errors"] and not r["warnings"]:
                st.success("Файл пройшов усі перевірки!")

    # ── КРОК 5: Зведення + Таблиця охоплення + Дашборди ─────────────
    if good:
        st.divider()
        st.header("📊 Крок 5 — Результати")

        if missing_count > 0:
            st.warning(f"⚠️ Увага: **{missing_count}** заклад(ів) ще не подали звіт. Зведення буде неповним.")
        if bad:
            with st.expander("🔴 Виключені файли (критичні помилки)"):
                for r in bad:
                    st.write(f"• {r['file']} — {r['name']}")

        # Витягуємо дані для дашбордів з усіх валідних файлів
        facility_data_list = []
        for r in good:
            try:
                fd = extract_facility_data(r["_bytes"], r["name"], r["edrpou"])
                facility_data_list.append(fd)
            except Exception:
                pass

        # ── ТРИ ВКЛАДКИ ──────────────────────────────────────────────
        tab_main, tab_coverage, tab_dash = st.tabs([
            "📥 Зведений файл",
            "📊 Таблиця охоплення",
            "📈 Дашборди"
        ])

        # ────────────────────────────────────────────────────────────
        # ВКЛАДКА 1 — ЗВЕДЕНИЙ ФАЙЛ
        # ────────────────────────────────────────────────────────────
        with tab_main:
            st.markdown(f"До зведення включено **{len(good)}** файл(ів) з **{expected}** очікуваних.")
            if st.button("⚙️ Створити зведений файл", type="primary", use_container_width=True):
                with st.spinner("⏳ Зведення виконується..."):
                    try:
                        file_bytes_list = [(r["file"], r["_bytes"]) for r in good]
                        result_bytes = aggregate_files(
                            file_bytes_list,
                            org_name      = st.session_state["org_name"],
                            org_edrpou    = st.session_state["org_edrpou"],
                            report_period = st.session_state["report_period"]
                        )
                        period_str = st.session_state["report_label"].replace(" ", "_")
                        st.success(f"✅ Успішно зведено {len(good)} файлів!")
                        st.download_button(
                            label="⬇️ Завантажити зведений файл",
                            data=result_bytes,
                            file_name=f"Зведений_звіт_{period_str}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True, type="primary"
                        )
                    except Exception as e:
                        st.error(f"❌ Помилка: {e}")

        # ────────────────────────────────────────────────────────────
        # ВКЛАДКА 2 — ТАБЛИЦЯ ОХОПЛЕННЯ
        # ────────────────────────────────────────────────────────────
        with tab_coverage:
            st.markdown("Оберіть параметри таблиці охоплення:")

            all_facility_names = [fd["name"] for fd in facility_data_list]
            all_labels = []
            seen = set()
            for fd in facility_data_list:
                for item in fd["coverage"]:
                    if item["label"] not in seen:
                        all_labels.append(item["label"])
                        seen.add(item["label"])

            col_f, col_i = st.columns(2)
            with col_f:
                st.markdown("**Заклади:**")
                select_all_f = st.checkbox("Обрати всі заклади", value=True, key="all_fac")
                if select_all_f:
                    selected_facilities = all_facility_names
                else:
                    selected_facilities = st.multiselect(
                        "Оберіть заклади:", all_facility_names, default=all_facility_names
                    )

            with col_i:
                st.markdown("**Показники (вакцини):**")
                select_all_i = st.checkbox("Обрати всі показники", value=True, key="all_ind")
                if select_all_i:
                    selected_indicators = all_labels
                else:
                    selected_indicators = st.multiselect(
                        "Оберіть показники:", all_labels, default=all_labels
                    )

            if selected_facilities and selected_indicators:
                if st.button("⚙️ Сформувати таблицю охоплення", type="primary", use_container_width=True):
                    with st.spinner("⏳ Формування таблиці..."):
                        try:
                            coverage_bytes = build_coverage_excel(
                                facility_data_list, selected_facilities, selected_indicators
                            )
                            period_str = st.session_state["report_label"].replace(" ", "_")
                            st.success("✅ Таблиця охоплення сформована!")
                            st.download_button(
                                label="⬇️ Завантажити таблицю охоплення",
                                data=coverage_bytes,
                                file_name=f"Таблиця_охоплення_{period_str}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True, type="primary"
                            )
                        except Exception as e:
                            st.error(f"❌ Помилка: {e}")
            else:
                st.info("Оберіть хоча б один заклад і один показник.")

        # ────────────────────────────────────────────────────────────
        # ВКЛАДКА 3 — ДАШБОРДИ
        # ────────────────────────────────────────────────────────────
        with tab_dash:
            d1, d2, d3, d4 = st.tabs([
                "💉 Охоплення",
                "🏆 Рейтинг ЗОЗ",
                "🏥 Залишки вакцин",
                "❌ Відмови та протипокази"
            ])

            # ── Дашборд 1: Охоплення ─────────────────────────────────
            with d1:
                st.subheader("💉 Охоплення щепленнями (план vs факт)")

                # Збираємо зведені дані по всіх ЗОЗ
                agg = {}
                for fd in facility_data_list:
                    for item in fd["coverage"]:
                        lbl = item["label"]
                        if lbl not in agg:
                            agg[lbl] = {"plan": 0, "executed": 0}
                        agg[lbl]["plan"]     += item["plan"]
                        agg[lbl]["executed"] += item["executed"]

                rows_agg = []
                for lbl, vals in agg.items():
                    pct = round(vals["executed"] / vals["plan"] * 100, 1) if vals["plan"] > 0 else 0
                    rows_agg.append({"Вакцина / Вік": lbl, "План": vals["plan"],
                                     "Виконано": vals["executed"], "%": pct})
                df_agg = pd.DataFrame(rows_agg).sort_values("%", ascending=True)

                # Фільтр
                all_vacc_labels = df_agg["Вакцина / Вік"].tolist()
                selected_vacc = st.multiselect(
                    "Оберіть вакцини для відображення:",
                    all_vacc_labels,
                    default=all_vacc_labels[:15] if len(all_vacc_labels) > 15 else all_vacc_labels,
                    key="dash_vacc_filter"
                )

                df_filtered = df_agg[df_agg["Вакцина / Вік"].isin(selected_vacc)]

                if not df_filtered.empty:
                    # Колір по % виконання
                    colors = []
                    for pct in df_filtered["%"]:
                        if pct >= 95:
                            colors.append("#2ECC71")
                        elif pct >= 85:
                            colors.append("#F39C12")
                        else:
                            colors.append("#E74C3C")

                    fig = go.Figure(go.Bar(
                        x=df_filtered["%"],
                        y=df_filtered["Вакцина / Вік"],
                        orientation="h",
                        marker_color=colors,
                        text=[f"{p}%" for p in df_filtered["%"]],
                        textposition="outside",
                        hovertemplate="<b>%{y}</b><br>Виконано: %{x}%<extra></extra>"
                    ))
                    fig.add_vline(x=95, line_dash="dash", line_color="green",
                                  annotation_text="95%", annotation_position="top")
                    fig.add_vline(x=85, line_dash="dash", line_color="orange",
                                  annotation_text="85%", annotation_position="top")
                    fig.update_layout(
                        xaxis_title="% виконання плану",
                        yaxis_title="",
                        height=max(400, len(df_filtered) * 30),
                        xaxis=dict(range=[0, 120]),
                        margin=dict(l=10, r=80, t=20, b=40),
                        plot_bgcolor="white"
                    )
                    st.plotly_chart(fig, use_container_width=True)

                    # Легенда
                    lc1, lc2, lc3 = st.columns(3)
                    lc1.success("🟢 ≥ 95% — виконання плану")
                    lc2.warning("🟡 85–95% — потребує уваги")
                    lc3.error("🔴 < 85% — критично")

            # ── Дашборд 2: Рейтинг ЗОЗ ───────────────────────────────
            with d2:
                st.subheader("🏆 Рейтинг закладів")

                # Збираємо всі унікальні вакцини
                all_vacc_for_rating = list({item["label"]
                                            for fd in facility_data_list
                                            for item in fd["coverage"]})
                all_vacc_for_rating.sort()

                rating_mode = st.radio(
                    "Показник для рейтингу:",
                    ["Середній % по всіх вакцинах", "По обраній вакцині"],
                    horizontal=True
                )

                if rating_mode == "По обраній вакцині":
                    selected_vacc_rating = st.selectbox(
                        "Оберіть вакцину:", all_vacc_for_rating, key="rating_vacc"
                    )

                rating_rows = []
                for fd in facility_data_list:
                    if rating_mode == "Середній % по всіх вакцинах":
                        pcts = [item["pct"] for item in fd["coverage"] if item["plan"] > 0]
                        avg_pct = round(sum(pcts) / len(pcts), 1) if pcts else 0
                        rating_rows.append({"Заклад": fd["name"], "%": avg_pct})
                    else:
                        pct = 0
                        for item in fd["coverage"]:
                            if item["label"] == selected_vacc_rating:
                                pct = item["pct"]
                                break
                        rating_rows.append({"Заклад": fd["name"], "%": pct})

                df_rating = pd.DataFrame(rating_rows).sort_values("%", ascending=False).reset_index(drop=True)
                df_rating.index += 1  # нумерація з 1

                colors_r = []
                for pct in df_rating["%"]:
                    if pct >= 95:
                        colors_r.append("#2ECC71")
                    elif pct >= 85:
                        colors_r.append("#F39C12")
                    else:
                        colors_r.append("#E74C3C")

                fig_r = go.Figure(go.Bar(
                    x=df_rating["%"],
                    y=df_rating["Заклад"],
                    orientation="h",
                    marker_color=colors_r,
                    text=[f"{p}%" for p in df_rating["%"]],
                    textposition="outside",
                    hovertemplate="<b>%{y}</b><br>%{x}%<extra></extra>"
                ))
                fig_r.add_vline(x=95, line_dash="dash", line_color="green")
                fig_r.add_vline(x=85, line_dash="dash", line_color="orange")
                fig_r.update_layout(
                    xaxis_title="% виконання",
                    yaxis_title="",
                    height=max(300, len(df_rating) * 40),
                    xaxis=dict(range=[0, 120]),
                    yaxis=dict(autorange="reversed"),
                    margin=dict(l=10, r=80, t=20, b=40),
                    plot_bgcolor="white"
                )
                st.plotly_chart(fig_r, use_container_width=True)

                # Таблиця рейтингу
                df_display = df_rating.copy()
                df_display["Місце"] = range(1, len(df_display)+1)
                df_display["Статус"] = df_display["%"].apply(
                    lambda x: "🟢 OK" if x >= 95 else ("🟡 Увага" if x >= 85 else "🔴 Критично")
                )
                st.dataframe(df_display[["Місце","Заклад","%","Статус"]],
                             use_container_width=True, hide_index=True)

            # ── Дашборд 3: Залишки вакцин ─────────────────────────────
            with d3:
                st.subheader("🏥 Залишки вакцин")

                critical_threshold = st.number_input(
                    "⚙️ Поріг критично малого залишку (доз):",
                    min_value=0, max_value=10000, value=50, step=10,
                    help="Залишки нижче цього значення будуть підсвічені червоним"
                )

                # Агрегуємо залишки по всіх ЗОЗ
                stock_agg = {}
                for fd in facility_data_list:
                    for s in fd["stocks"]:
                        vac = s["vaccine"]
                        if vac not in stock_agg:
                            stock_agg[vac] = {"closing": 0, "used": 0}
                        stock_agg[vac]["closing"] += s["closing"]
                        stock_agg[vac]["used"]    += s["used"]

                stock_rows = [{"Вакцина": vac, "Залишок (доз)": vals["closing"],
                               "Витрачено (доз)": vals["used"]}
                              for vac, vals in stock_agg.items() if vals["closing"] > 0 or vals["used"] > 0]
                df_stock = pd.DataFrame(stock_rows).sort_values("Залишок (доз)", ascending=True)

                if not df_stock.empty:
                    colors_s = ["#E74C3C" if v <= critical_threshold else "#3498DB"
                                for v in df_stock["Залишок (доз)"]]

                    fig_s = go.Figure(go.Bar(
                        x=df_stock["Залишок (доз)"],
                        y=df_stock["Вакцина"],
                        orientation="h",
                        marker_color=colors_s,
                        text=df_stock["Залишок (доз)"].astype(int),
                        textposition="outside",
                        hovertemplate="<b>%{y}</b><br>Залишок: %{x} доз<extra></extra>"
                    ))
                    fig_s.add_vline(x=critical_threshold, line_dash="dash", line_color="red",
                                   annotation_text=f"Критично ({critical_threshold})",
                                   annotation_position="top right")
                    fig_s.update_layout(
                        xaxis_title="Кількість доз",
                        yaxis_title="",
                        height=max(400, len(df_stock) * 30),
                        margin=dict(l=10, r=80, t=20, b=40),
                        plot_bgcolor="white"
                    )
                    st.plotly_chart(fig_s, use_container_width=True)

                    # Таблиця з виділенням критичних
                    critical_vaccines = df_stock[df_stock["Залишок (доз)"] <= critical_threshold]
                    if not critical_vaccines.empty:
                        st.error(f"🚨 Критично мало залишків ({critical_threshold} доз і менше):")
                        st.dataframe(critical_vaccines, use_container_width=True, hide_index=True)

            # ── Дашборд 4: Відмови та протипокази ────────────────────
            with d4:
                st.subheader("❌ Відмови та протипокази")

                col_ref, col_contra = st.columns(2)

                with col_ref:
                    st.markdown("**Відмови від щеплень (по нозологіях)**")
                    ref_agg = {}
                    for fd in facility_data_list:
                        for r in fd["refusals"]:
                            d = r["disease"]
                            ref_agg[d] = ref_agg.get(d, 0) + r["count"]

                    df_ref = pd.DataFrame([{"Нозологія": k, "Відмов": v}
                                           for k, v in ref_agg.items() if v > 0])

                    if not df_ref.empty:
                        fig_ref = px.bar(df_ref, x="Відмов", y="Нозологія",
                                         orientation="h", color="Відмов",
                                         color_continuous_scale=["#FFF3CD", "#E74C3C"],
                                         text="Відмов")
                        fig_ref.update_traces(textposition="outside")
                        fig_ref.update_layout(height=300, showlegend=False,
                                              plot_bgcolor="white",
                                              margin=dict(l=10, r=60, t=10, b=20))
                        st.plotly_chart(fig_ref, use_container_width=True)
                        st.dataframe(df_ref.sort_values("Відмов", ascending=False),
                                     use_container_width=True, hide_index=True)
                    else:
                        st.info("Відмов не зареєстровано")

                with col_contra:
                    st.markdown("**Протипокази до вакцинації**")
                    total_temp = sum(fd["temp_contraindications"] for fd in facility_data_list)
                    total_perm = sum(fd["perm_contraindications"] for fd in facility_data_list)
                    total_all  = total_temp + total_perm

                    if total_all > 0:
                        fig_pie = px.pie(
                            values=[total_temp, total_perm],
                            names=["Тимчасові", "Постійні"],
                            color_discrete_sequence=["#F39C12", "#E74C3C"],
                            hole=0.4
                        )
                        fig_pie.update_traces(textposition="inside",
                                              textinfo="percent+label+value")
                        fig_pie.update_layout(height=300, showlegend=True,
                                              margin=dict(l=10, r=10, t=10, b=10))
                        st.plotly_chart(fig_pie, use_container_width=True)

                        p1, p2, p3 = st.columns(3)
                        p1.metric("Тимчасові", int(total_temp))
                        p2.metric("Постійні",  int(total_perm))
                        p3.metric("Всього",    int(total_all))

                        # По ЗОЗ
                        st.markdown("**По закладах:**")
                        contra_rows = [{"Заклад": fd["name"],
                                        "Тимчасові": int(fd["temp_contraindications"]),
                                        "Постійні":  int(fd["perm_contraindications"]),
                                        "Всього":    int(fd["temp_contraindications"] + fd["perm_contraindications"])}
                                       for fd in facility_data_list]
                        st.dataframe(pd.DataFrame(contra_rows).sort_values("Всього", ascending=False),
                                     use_container_width=True, hide_index=True)
                    else:
                        st.info("Протипоказів не зареєстровано")

    else:
        st.error("❌ Немає файлів придатних для зведення — виправте помилки і спробуйте знову.")
