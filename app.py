import streamlit as st
import pandas as pd
from openpyxl import load_workbook
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


def validate_file(file_bytes, filename):
    errors = []
    warnings = []
    name = "—"
    edrpou = "—"
    period = "—"
    wb = None

    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)

        # 1. Обов'язкові аркуші
        required = ["План", "Виконання", "Залишки", "Зведений звіт", "Аркуш1"]
        missing = [s for s in required if s not in wb.sheetnames]
        if missing:
            errors.append(f"Відсутні аркуші: {', '.join(missing)}")
            return dict(file=filename, name=name, edrpou=edrpou, period=period,
                        status="🔴 Помилки", errors=errors, warnings=warnings, wb=wb)

        ws_plan = wb["План"]
        ws_exec = wb["Виконання"]
        ws_rem  = wb["Залишки"]

        # 2. Назва закладу
        raw_name = ws_plan["D8"].value
        name = str(raw_name).strip() if raw_name else "—"
        if not raw_name or name == "":
            errors.append("Порожня назва закладу (План!D8)")

        # 3. Код ЄДРПОУ
        raw_edrpou = ws_plan["E8"].value
        edrpou = str(raw_edrpou).strip().lstrip("'") if raw_edrpou else "—"
        if not raw_edrpou:
            errors.append("Порожній код ЄДРПОУ (План!E8)")
        elif not edrpou.isdigit() or len(edrpou) not in (7, 8, 9, 10):
            warnings.append(f"Код ЄДРПОУ '{edrpou}' має нестандартну довжину")

        # 4. Звітний період
        raw_period = ws_exec["F6"].value
        period = str(raw_period) if raw_period else "—"
        if not raw_period:
            errors.append("Відсутній звітний період (Виконання!F6)")

        # 5. Узгодженість назви між аркушами
        name_exec = ws_exec["C4"].value
        if name_exec and raw_name and str(name_exec).strip() != name:
            warnings.append("Назва закладу різниться між аркушами «План» і «Виконання»")

        # 6. Від'ємні значення у Виконанні
        for row in range(8, 105):
            val = ws_exec.cell(row=row, column=5).value
            if isinstance(val, (int, float)) and val < 0:
                vac = ws_exec.cell(row=row, column=3).value
                age = ws_exec.cell(row=row, column=4).value
                errors.append(f"Від'ємна кількість щеплень: {vac} / {age} = {val}")

        # 7. Балансова формула залишків
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
                    errors.append(
                        f"Залишки — помилка балансу для «{str(vaccine).strip()}»"
                        f" (є {d}, має бути {expected})"
                    )
                if d < 0:
                    errors.append(f"Залишки — від'ємний залишок для «{str(vaccine).strip()}»")

        # 8. Своєчасність КДП-3
        for row in range(8, 11):
            l_val = safe_num(ws_exec.cell(row=row, column=12).value)
            m_val = safe_num(ws_exec.cell(row=row, column=13).value)
            if m_val > l_val > 0:
                errors.append(
                    f"Рядок {row}: «Отримали КДП-3» ({m_val}) > «Народилося за 7 міс.» ({l_val})"
                )

        # 9. Протипокази: ВСЬОГО = Тимчасові + Постійні
        for row in range(8, 11):
            p = safe_num(ws_exec.cell(row=row, column=16).value)
            q = safe_num(ws_exec.cell(row=row, column=17).value)
            r = ws_exec.cell(row=row, column=18).value
            if isinstance(r, (int, float)) and abs(r - (p + q)) > 0.5:
                errors.append(
                    f"Рядок {row}: протипокази ВСЬОГО ({r}) ≠ Тимчасові+Постійні ({p + q})"
                )

    except Exception as e:
        errors.append(f"Не вдалось прочитати файл: {e}")

    status = "🔴 Помилки" if errors else ("🟡 Попередження" if warnings else "🟢 OK")
    return dict(file=filename, name=name, edrpou=edrpou, period=period,
                status=status, errors=errors, warnings=warnings, wb=wb)


def aggregate_files(file_bytes_list, org_name, org_edrpou, report_period):
    workbooks = [load_workbook(io.BytesIO(fbytes), data_only=True)
                 for _, fbytes in file_bytes_list]
    template_wb = load_workbook(io.BytesIO(file_bytes_list[0][1]))

    # ── 1. ВИКОНАННЯ ─────────────────────────────────────────────────
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

    group_sums = {
        11: list(range(8, 12)),
        23: list(range(12, 24)),
        35: list(range(24, 36)),
        42: list(range(36, 43)),
        48: list(range(43, 49)),
        61: list(range(49, 62)),
    }
    for sum_row, rows in group_sums.items():
        ws_out.cell(row=sum_row, column=7).value = sum(
            safe_num(ws_out.cell(row=r, column=5).value) for r in rows
        )
    ws_out.cell(row=99,  column=7).value = sum(safe_num(ws_out.cell(row=r, column=5).value) for r in [62,64,66,68,70,72,74,76,78])
    ws_out.cell(row=100, column=7).value = sum(safe_num(ws_out.cell(row=r, column=5).value) for r in [63,65,67,69,71,73,75,77,79,81])
    ws_out.cell(row=101, column=7).value = safe_num(ws_out.cell(row=80, column=5).value)
    ws_out.cell(row=102, column=7).value = sum(safe_num(ws_out.cell(row=r, column=5).value) for r in [82,84,86,88,90,92,94,96,98,100,102])
    ws_out.cell(row=103, column=7).value = sum(safe_num(ws_out.cell(row=r, column=5).value) for r in [83,85,87,89,91,93,95,97,99,101])
    ws_out.cell(row=104, column=7).value = sum(safe_num(ws_out.cell(row=r, column=5).value) for r in [103, 104])

    # ── 2. ЗАЛИШКИ ───────────────────────────────────────────────────
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

    # ── 3. ПЛАН ──────────────────────────────────────────────────────
    ws_plan_out = template_wb["План"]
    for row in range(11, 47):
        ws_plan_out.cell(row=row, column=6).value = 0
    for wb in workbooks:
        ws = wb["План"]
        for row in range(11, 47):
            cur = ws_plan_out.cell(row=row, column=6).value or 0
            ws_plan_out.cell(row=row, column=6).value = cur + safe_num(ws.cell(row=row, column=6).value)

    # ── 4. ЗВЕДЕНИЙ ЗВІТ ─────────────────────────────────────────────
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

    # ── 5. РЕКВІЗИТИ ─────────────────────────────────────────────────
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


# ─────────────────────────────────────────────
# ІНТЕРФЕЙС
# ─────────────────────────────────────────────

# ── КРОК 1: Реквізити ────────────────────────────────────────────────
st.header("🏥 Крок 1 — Реквізити організації що формує зведений звіт")

with st.form("org_form"):
    col1, col2 = st.columns([3, 1])
    with col1:
        org_name = st.text_input(
            "Назва закладу",
            placeholder="Наприклад: Тернопільський обласний ЦКПХ",
        )
    with col2:
        org_edrpou = st.text_input(
            "Код ЄДРПОУ",
            placeholder="12345678",
        )

    col3, col4, col5 = st.columns(3)
    with col3:
        report_year = st.selectbox(
            "Звітний рік",
            options=list(range(2024, 2031)),
            index=2
        )
    with col4:
        month_names = {
            1: "Січень", 2: "Лютий", 3: "Березень", 4: "Квітень",
            5: "Травень", 6: "Червень", 7: "Липень", 8: "Серпень",
            9: "Вересень", 10: "Жовтень", 11: "Листопад", 12: "Грудень"
        }
        report_month = st.selectbox(
            "Звітний місяць",
            options=list(month_names.keys()),
            format_func=lambda x: month_names[x],
            index=1
        )
    with col5:
        expected_count = st.number_input(
            "Кількість ЗОЗ що мають подати звіт",
            min_value=1,
            max_value=500,
            value=10,
            step=1,
            help="Загальна кількість закладів, від яких очікується звіт за цей місяць"
        )

    submitted = st.form_submit_button(
        "✅ Підтвердити реквізити", type="primary", use_container_width=True
    )

if submitted:
    form_errors = []
    if not org_name.strip():
        form_errors.append("Введіть назву закладу")
    if not org_edrpou.strip():
        form_errors.append("Введіть код ЄДРПОУ")
    elif not org_edrpou.strip().isdigit():
        form_errors.append("Код ЄДРПОУ повинен містити тільки цифри")

    if form_errors:
        for e in form_errors:
            st.error(f"❌ {e}")
    else:
        st.session_state["org_name"]       = org_name.strip()
        st.session_state["org_edrpou"]     = org_edrpou.strip()
        st.session_state["report_period"]  = datetime(report_year, report_month, 1)
        st.session_state["report_label"]   = f"{month_names[report_month]} {report_year}"
        st.session_state["expected_count"] = int(expected_count)

# Показуємо підтверджені реквізити
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

    uploaded_files = st.file_uploader(
        "Оберіть файли Excel (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
    )

    if uploaded_files:

        # ── Лічильник прогресу подання ────────────────────────────────
        expected  = st.session_state["expected_count"]
        submitted_count = len(uploaded_files)
        missing_count   = max(0, expected - submitted_count)
        percent         = min(100, round(submitted_count / expected * 100))

        st.subheader("📊 Стан подання звітів")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📋 Очікується",  expected)
        m2.metric("📥 Подано",      submitted_count,
                  delta=f"+{submitted_count}" if submitted_count > 0 else None)
        m3.metric("⏳ Не подали",   missing_count,
                  delta=f"-{missing_count}" if missing_count > 0 else None,
                  delta_color="inverse")
        m4.metric("✅ Виконання",   f"{percent}%")

        # Прогрес-бар
        st.progress(percent / 100,
                    text=f"Подано {submitted_count} з {expected} закладів ({percent}%)")

        if missing_count == 0:
            st.success("🎉 Усі заклади подали звіти!")
        elif missing_count > 0:
            st.warning(
                f"⚠️ Ще не подали звіт: **{missing_count}** заклад(ів).  "
                f"Можна звести вже зараз або дочекатися решти."
            )

        st.divider()

        # ── КРОК 3: Список файлів ─────────────────────────────────────
        st.header("📋 Крок 3 — Список завантажених файлів")
        file_data = [{"Файл": f.name,
                      "Розмір": f"{round(f.size / 1024, 1)} КБ",
                      "Статус": "⏳ Очікує перевірки"} for f in uploaded_files]
        st.dataframe(pd.DataFrame(file_data), use_container_width=True, hide_index=True)

        # ── КРОК 4: Перевірка ─────────────────────────────────────────
        st.header("🔍 Крок 4 — Перевірка файлів")

        if st.button("▶️ Запустити перевірку", type="primary", use_container_width=True):
            files_bytes = [(f.name, f.read()) for f in uploaded_files]
            results = []
            progress_bar = st.progress(0, text="Перевірка виконується...")
            for i, (fname, fbytes) in enumerate(files_bytes):
                r = validate_file(fbytes, fname)
                r["_bytes"] = fbytes
                results.append(r)
                progress_bar.progress(
                    (i + 1) / len(files_bytes),
                    text=f"Перевірено {i+1} з {len(files_bytes)}: {fname}"
                )
            st.session_state["results"] = results
            st.rerun()


# ─── Результати перевірки ─────────────────────────────────────────────
if "results" in st.session_state and "org_name" in st.session_state:
    results      = st.session_state["results"]
    expected     = st.session_state["expected_count"]

    ok   = [r for r in results if r["status"] == "🟢 OK"]
    warn = [r for r in results if r["status"] == "🟡 Попередження"]
    bad  = [r for r in results if r["status"] == "🔴 Помилки"]
    good = [r for r in results if r["status"] in ("🟢 OK", "🟡 Попередження")]

    # Підсумкові метрики після перевірки
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

    # Однорідність звітного періоду
    periods = list({r["period"] for r in results if r["period"] != "—"})
    if len(periods) > 1:
        st.warning(f"⚠️ Різні звітні періоди у файлах: {', '.join(periods)}")

    # Унікальність ЄДРПОУ
    edrpous    = [r["edrpou"] for r in results if r["edrpou"] != "—"]
    duplicates = {e for e in edrpous if edrpous.count(e) > 1}
    if duplicates:
        st.error(f"❌ Дублікати ЄДРПОУ: {', '.join(duplicates)}")

    # Детальні результати
    st.subheader("Детальні результати по файлах")
    for r in results:
        with st.expander(
            f"{r['status']}  |  {r['file']}  |  {r['name']}  |  ЄДРПОУ: {r['edrpou']}"
        ):
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

    # ── КРОК 5: Зведення ─────────────────────────────────────────────
    if good:
        st.divider()
        st.header("📊 Крок 5 — Зведення")

        st.info(
            f"До зведення будуть включені **{len(good)}** файл(ів)  "
            f"(з {expected} очікуваних).  \n"
            f"Виключено через критичні помилки: **{len(bad)}** файл(ів)."
        )

        if missing_count > 0:
            st.warning(
                f"⚠️ Увага: **{missing_count}** заклад(ів) ще не подали звіт. "
                f"Зведення буде неповним."
            )

        if bad:
            with st.expander("🔴 Виключені файли (критичні помилки)"):
                for r in bad:
                    st.write(f"• {r['file']} — {r['name']}")

        if st.button("📥 Створити зведений файл", type="primary", use_container_width=True):
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
                    filename   = f"Зведений_звіт_{period_str}.xlsx"

                    st.success(
                        f"✅ Зведено **{len(good)}** файлів "
                        f"з **{expected}** очікуваних!"
                    )

                    st.download_button(
                        label="⬇️ Завантажити зведений файл",
                        data=result_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )
                except Exception as e:
                    st.error(f"❌ Помилка під час зведення: {e}")
    else:
        st.error("❌ Немає файлів придатних для зведення — виправте помилки і спробуйте знову.")
