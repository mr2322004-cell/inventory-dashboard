
import io
import re
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(
    page_title="Inventory Intelligence",
    page_icon="📦",
    layout="wide",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
html, body, [class*="css"] {font-family: 'Cairo', sans-serif;}
.stApp {background: #0b1120; color: #e5e7eb;}
[data-testid="stSidebar"] {background: #111827;}
[data-testid="stMetric"] {
    background: #151f32;
    padding: 16px;
    border: 1px solid #29364b;
    border-radius: 14px;
}
</style>
""", unsafe_allow_html=True)


# =====================================
# التعرف على الأعمدة
# =====================================

ALIASES = {
    "code": [
        "كود الصنف", "كود المادة", "كود", "رمز الصنف",
        "item code", "item_id", "sku", "code", "barcode"
    ],
    "name": [
        "اسم الصنف", "اسم المادة", "الصنف",
        "item name", "product name", "description", "name"
    ],
    "expected": [
        "الكمية الدفترية", "كمية النظام", "الرصيد الدفتري",
        "الكمية المفترضة", "رصيد النظام",
        "system quantity", "expected quantity",
        "book quantity", "system stock"
    ],
    "actual": [
        "الكمية الفعلية", "الجرد الفعلي", "كمية الجرد",
        "الكمية المعدودة", "الرصيد الفعلي",
        "actual quantity", "counted quantity",
        "physical quantity", "counted stock"
    ],
    "difference": [
        "الفرق", "فرق الكمية", "الفروقات", "الفارق",
        "difference", "variance", "qty difference"
    ],
    "month": [
        "الشهر", "شهر الجرد", "شهر", "month", "period"
    ],
}


def normalize(value):
    value = str(value).strip().lower()
    value = re.sub(r"\s+", " ", value)
    return value


def detect_columns(df):
    result = {}
    normalized = {normalize(c): c for c in df.columns}

    for key, aliases in ALIASES.items():
        for alias in aliases:
            if normalize(alias) in normalized:
                result[key] = normalized[normalize(alias)]
                break

    return result


def read_file(file):
    filename = file.name.lower()

    if filename.endswith(".csv"):
        try:
            return pd.read_csv(file, encoding="utf-8-sig")
        except UnicodeDecodeError:
            file.seek(0)
            return pd.read_csv(file, encoding="cp1256")

    if filename.endswith((".xlsx", ".xls", ".xlsm")):
        return pd.read_excel(file)

    raise ValueError("صيغة الملف غير مدعومة.")


def to_numeric(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("٫", ".", regex=False)
        .str.strip(),
        errors="coerce",
    )


def identify_month(value):
    """تحويل قيم الشهر الشائعة إلى سبتمبر أو أكتوبر."""
    if pd.isna(value):
        return None

    text = normalize(value)

    # أسماء الأشهر العربية والإنجليزية
    if any(x in text for x in [
        "سبتمبر", "شهر 9", "الشهر 9", "sep", "september"
    ]):
        return "سبتمبر"

    if any(x in text for x in [
        "أكتوبر", "اكتوبر", "شهر 10", "الشهر 10",
        "oct", "october"
    ]):
        return "أكتوبر"

    # قيم رقمية أو تواريخ
    try:
        number = float(text)
        if number == 9:
            return "سبتمبر"
        if number == 10:
            return "أكتوبر"
    except (ValueError, TypeError):
        pass

    # محاولة قراءة تاريخ فعلي
    try:
        dt = pd.to_datetime(value, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            if dt.month == 9:
                return "سبتمبر"
            if dt.month == 10:
                return "أكتوبر"
    except (ValueError, TypeError):
        pass

    return None


# =====================================
# تجهيز ملف الشهرين
# =====================================

def prepare_data(raw):
    df = raw.copy()
    cols = detect_columns(df)

    if "code" not in cols:
        raise ValueError("لم أجد عمود كود الصنف.")

    if "month" not in cols:
        raise ValueError(
            "لم أجد عمود الشهر. يجب أن يحتوي الملف على عمود "
            "يوضح هل الصف يخص سبتمبر أم أكتوبر."
        )

    if "difference" not in cols and not (
        "expected" in cols and "actual" in cols
    ):
        raise ValueError(
            "يلزم وجود عمود الفرق أو الكمية الدفترية والفعلية."
        )

    df["كود الصنف"] = df[cols["code"]].astype(str).str.strip()

    if "name" in cols:
        df["اسم الصنف"] = df[cols["name"]].astype(str).str.strip()
    else:
        df["اسم الصنف"] = df["كود الصنف"]

    df["الشهر"] = df[cols["month"]].apply(identify_month)

    if "expected" in cols:
        df["الكمية الدفترية"] = to_numeric(df[cols["expected"]])

    if "actual" in cols:
        df["الكمية الفعلية"] = to_numeric(df[cols["actual"]])

    if "difference" in cols:
        # حسب نظامك: الموجب عجز والسالب زيادة
        df["الفرق"] = to_numeric(df[cols["difference"]])
    else:
        # الدفترية - الفعلية = موجب عند العجز
        df["الفرق"] = (
            df["الكمية الدفترية"] - df["الكمية الفعلية"]
        )

    if "الكمية الدفترية" not in df.columns:
        df["الكمية الدفترية"] = float("nan")

    if "الكمية الفعلية" not in df.columns:
        df["الكمية الفعلية"] = float("nan")

    df = df[
        df["كود الصنف"].notna()
        & (df["كود الصنف"].str.strip() != "")
        & (df["كود الصنف"].str.lower() != "nan")
        & df["الشهر"].isin(["سبتمبر", "أكتوبر"])
    ].copy()

    if df.empty:
        raise ValueError(
            "لم أجد صفوفًا لشهر سبتمبر أو أكتوبر. "
            "راجع قيم عمود الشهر."
        )

    # التجميع حسب الشهر وكود الصنف، وليس الاسم
    grouped = df.groupby(
        ["الشهر", "كود الصنف"],
        as_index=False,
        sort=False,
    ).agg({
        "اسم الصنف": "first",
        "الكمية الدفترية": lambda x: x.sum(min_count=1),
        "الكمية الفعلية": lambda x: x.sum(min_count=1),
        "الفرق": lambda x: x.sum(min_count=1),
    })

    # الموجب = عجز، السالب = زيادة
    grouped["كمية العجز"] = grouped["الفرق"].clip(lower=0)
    grouped["كمية الزيادة"] = (-grouped["الفرق"]).clip(lower=0)

    grouped["حالة الجرد"] = grouped["الفرق"].apply(
        lambda x: (
            "عجز" if x > 0
            else "زيادة" if x < 0
            else "مطابق"
        )
    )

    return grouped


# =====================================
# حساب المؤشرات
# =====================================

def calculate_metrics(df):
    total_codes = df["كود الصنف"].nunique()

    shortage_codes = int((df["الفرق"] > 0).sum())
    surplus_codes = int((df["الفرق"] < 0).sum())
    mismatch_codes = int((df["الفرق"] != 0).sum())

    shortage = df["كمية العجز"].sum()
    surplus = df["كمية الزيادة"].sum()

    reference = df["الكمية الدفترية"].sum(min_count=1)

    if pd.notna(reference) and reference > 0:
        loss_rate = shortage / reference * 100
    else:
        loss_rate = None

    mismatch_rate = (
        mismatch_codes / total_codes * 100
        if total_codes else 0
    )

    return {
        "codes": total_codes,
        "shortage_codes": shortage_codes,
        "surplus_codes": surplus_codes,
        "mismatch_codes": mismatch_codes,
        "shortage": shortage,
        "surplus": surplus,
        "reference": reference,
        "loss_rate": loss_rate,
        "mismatch_rate": mismatch_rate,
    }


def fmt_pct(value):
    if value is None or pd.isna(value):
        return "غير متاح"
    return f"{value:.2f}%"


def export_excel(sheets):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, data in sheets.items():
            data.to_excel(
                writer,
                sheet_name=name[:31],
                index=False,
            )

    return output.getvalue()


# =====================================
# واجهة البرنامج
# =====================================

st.title("📦 Inventory Intelligence")
st.subheader("تحليل الفقد ودقة الجرد")
st.caption(
    "ملف واحد — شهر سبتمبر وشهر أكتوبر — بدون أسعار"
)

st.sidebar.header("إعدادات التقرير")
st.sidebar.info(
    "الموجب = عجز | السالب = زيادة | الصفر = مطابق"
)

uploaded = st.file_uploader(
    "📁 ارفع ملف Excel أو CSV الذي يحتوي على الشهرين",
    type=["xlsx", "xls", "xlsm", "csv"],
)

if uploaded is None:
    st.info(
        "ارفع الملف الذي يحتوي على بيانات سبتمبر وأكتوبر. "
        "يجب أن يكون لكل صف كود صنف وقيمة للشهر."
    )
    st.stop()

try:
    raw = read_file(uploaded)
    data = prepare_data(raw)
except Exception as exc:
    st.error(f"تعذر تجهيز البيانات: {exc}")
    st.stop()

september = data[data["الشهر"] == "سبتمبر"].copy()
october = data[data["الشهر"] == "أكتوبر"].copy()

if september.empty or october.empty:
    st.warning(
        "الملف لم يتضمن بيانات للشهرين معًا. "
        "راجع عمود الشهر قبل الاعتماد على المقارنة."
    )

m9 = calculate_metrics(september)
m10 = calculate_metrics(october)

st.divider()
st.header("📊 مؤشرات الأداء")

a, b, c, d = st.columns(4)

a.metric("عجز سبتمبر", f'{m9["shortage"]:,.2f}')
b.metric(
    "عجز أكتوبر",
    f'{m10["shortage"]:,.2f}',
    delta=f'{m10["shortage"] - m9["shortage"]:,.2f}',
    delta_color="inverse",
)
c.metric("نسبة الفقد — سبتمبر", fmt_pct(m9["loss_rate"]))
d.metric("نسبة الفقد — أكتوبر", fmt_pct(m10["loss_rate"]))

a, b, c, d = st.columns(4)

a.metric("أكواد سبتمبر", f'{m9["codes"]:,}')
b.metric("أكواد أكتوبر", f'{m10["codes"]:,}')
c.metric("الجرد غير المضبوط — سبتمبر", fmt_pct(m9["mismatch_rate"]))
d.metric("الجرد غير المضبوط — أكتوبر", fmt_pct(m10["mismatch_rate"]))

st.caption(
    "نسبة الفقد = إجمالي كمية العجز ÷ إجمالي الكمية الدفترية × 100. "
    "نسبة الجرد غير المضبوط = الأكواد ذات الفرق ÷ إجمالي الأكواد المجردة × 100."
)

# =====================================
# التقارير التفصيلية
# =====================================

tab1, tab2, tab3, tab4 = st.tabs([
    "الأصناف الفاقدة",
    "الأصناف الزائدة",
    "تفاصيل الجرد",
    "مقارنة الشهرين",
])

with tab1:
    st.subheader("الأصناف التي بها عجز")

    shortages = data[data["كمية العجز"] > 0].copy()
    shortages = shortages.sort_values("كمية العجز", ascending=False)

    if shortages.empty:
        st.success("لا توجد أصناف بها عجز.")
    else:
        st.dataframe(
            shortages[[
                "الشهر", "كود الصنف", "اسم الصنف",
                "الكمية الدفترية", "الكمية الفعلية",
                "الفرق", "كمية العجز"
            ]],
            use_container_width=True,
            hide_index=True,
        )

        st.plotly_chart(
            px.bar(
                shortages.head(15),
                x="كمية العجز",
                y="اسم الصنف",
                color="الشهر",
                orientation="h",
                title="أكثر 15 صنفًا بها عجز",
                template="plotly_dark",
            ),
            use_container_width=True,
        )

with tab2:
    st.subheader("الأصناف التي بها زيادة")

    surpluses = data[data["كمية الزيادة"] > 0].copy()
    surpluses = surpluses.sort_values("كمية الزيادة", ascending=False)

    if surpluses.empty:
        st.info("لا توجد أصناف بها زيادة.")
    else:
        st.dataframe(
            surpluses[[
                "الشهر", "كود الصنف", "اسم الصنف",
                "الكمية الدفترية", "الكمية الفعلية",
                "الفرق", "كمية الزيادة"
            ]],
            use_container_width=True,
            hide_index=True,
        )

with tab3:
    st.subheader("تفاصيل الأصناف")

    search = st.text_input("ابحث باسم الصنف أو كوده")
    filtered = data.copy()

    if search:
        filtered = filtered[
            filtered["كود الصنف"].str.contains(
                search, case=False, na=False
            )
            | filtered["اسم الصنف"].str.contains(
                search, case=False, na=False
            )
        ]

    month_filter = st.selectbox(
        "اختر الشهر",
        ["الكل", "سبتمبر", "أكتوبر"],
    )

    if month_filter != "الكل":
        filtered = filtered[filtered["الشهر"] == month_filter]

    st.dataframe(
        filtered.sort_values(["الشهر", "كود الصنف"]),
        use_container_width=True,
        hide_index=True,
    )

with tab4:
    st.subheader("المقارنة بين الشهرين")

    comparison = pd.DataFrame([
        {
            "الشهر": "سبتمبر",
            "إجمالي الأكواد": m9["codes"],
            "أكواد العجز": m9["shortage_codes"],
            "أكواد الزيادة": m9["surplus_codes"],
            "أكواد غير مطابقة": m9["mismatch_codes"],
            "كمية العجز": m9["shortage"],
            "كمية الزيادة": m9["surplus"],
            "نسبة الفقد %": m9["loss_rate"],
            "نسبة الجرد غير المضبوط %": m9["mismatch_rate"],
        },
        {
            "الشهر": "أكتوبر",
            "إجمالي الأكواد": m10["codes"],
            "أكواد العجز": m10["shortage_codes"],
            "أكواد الزيادة": m10["surplus_codes"],
            "أكواد غير مطابقة": m10["mismatch_codes"],
            "كمية العجز": m10["shortage"],
            "كمية الزيادة": m10["surplus"],
            "نسبة الفقد %": m10["loss_rate"],
            "نسبة الجرد غير المضبوط %": m10["mismatch_rate"],
        },
    ])

    st.dataframe(
        comparison,
        use_container_width=True,
        hide_index=True,
    )

    chart_data = comparison.melt(
        id_vars="الشهر",
        value_vars=["كمية العجز", "كمية الزيادة"],
        var_name="نوع الفرق",
        value_name="الكمية",
    )

    st.plotly_chart(
        px.bar(
            chart_data,
            x="الشهر",
            y="الكمية",
            color="نوع الفرق",
            barmode="group",
            title="كميات العجز والزيادة",
            template="plotly_dark",
        ),
        use_container_width=True,
    )

# =====================================
# تنزيل تقرير Excel
# =====================================

st.divider()
st.header("📥 تنزيل التقرير")

summary = pd.DataFrame([
    {
        "الشهر": "سبتمبر",
        "إجمالي الأكواد": m9["codes"],
        "إجمالي الكمية المرجعية": m9["reference"],
        "كمية العجز": m9["shortage"],
        "كمية الزيادة": m9["surplus"],
        "نسبة الفقد %": m9["loss_rate"],
        "نسبة الجرد غير المضبوط %": m9["mismatch_rate"],
    },
    {
        "الشهر": "أكتوبر",
        "إجمالي الأكواد": m10["codes"],
        "إجمالي الكمية المرجعية": m10["reference"],
        "كمية العجز": m10["shortage"],
        "كمية الزيادة": m10["surplus"],
        "نسبة الفقد %": m10["loss_rate"],
        "نسبة الجرد غير المضبوط %": m10["mismatch_rate"],
    },
])

try:
    file_bytes = export_excel({
        "ملخص المقارنة": summary,
        "تفاصيل الجرد": data,
        "الأصناف الفاقدة": data[data["كمية العجز"] > 0],
        "الأصناف الزائدة": data[data["كمية الزيادة"] > 0],
    })

    st.download_button(
        "تحميل تقرير Excel الكامل",
        data=file_bytes,
        file_name="Inventory_Report_September_October.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

except Exception as exc:
    st.error(f"تعذر إنشاء ملف Excel: {exc}")

st.caption(
    "ملاحظة: يجب أن يكون عمود الشهر واضحًا، مثل سبتمبر وأكتوبر "
    "أو 9 و10. نسبة الفقد لا تُحسب دون كمية مرجعية موثوقة."
)

