
import io
import re
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(
    page_title="Inventory Intelligence",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
html, body, [class*="css"] {
    font-family: 'Cairo', sans-serif;
}
.stApp {
    background: #0b1120;
    color: #e5e7eb;
}
[data-testid="stSidebar"] {
    background: #111827;
}
[data-testid="stMetric"] {
    background: #151f32;
    padding: 18px;
    border: 1px solid #29364b;
    border-radius: 14px;
}
h1, h2, h3 { color: #f8fafc; }
</style>
""", unsafe_allow_html=True)


# ---------------------------
# Helpers
# ---------------------------

def normalize(value):
    value = str(value).strip().lower()
    value = re.sub(r"\s+", " ", value)
    return value


ALIASES = {
    "code": [
        "كود الصنف", "كود", "كود المادة", "رمز الصنف",
        "item code", "item_id", "sku", "code", "barcode"
    ],
    "name": [
        "اسم الصنف", "الصنف", "اسم المادة",
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
    "month": ["الشهر", "month", "period"],
}


def detect_columns(df):
    result = {}
    columns = {normalize(c): c for c in df.columns}

    for key, aliases in ALIASES.items():
        for alias in aliases:
            if normalize(alias) in columns:
                result[key] = columns[normalize(alias)]
                break

    return result


def read_file(file):
    name = file.name.lower()

    if name.endswith(".csv"):
        try:
            return pd.read_csv(file, encoding="utf-8-sig")
        except UnicodeDecodeError:
            file.seek(0)
            return pd.read_csv(file, encoding="cp1256")

    if name.endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(file)

    raise ValueError("صيغة الملف غير مدعومة")


def numeric(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("٫", ".", regex=False)
        .str.strip(),
        errors="coerce",
    )


def prepare_data(df, month_label):
    df = df.copy()
    cols = detect_columns(df)

    if "code" not in cols:
        raise ValueError("لم أجد عمود كود الصنف.")

    if "expected" not in cols and "difference" not in cols:
        raise ValueError(
            "يلزم وجود عمود الكمية الدفترية أو عمود الفرق."
        )

    df["كود الصنف"] = df[cols["code"]].astype(str).str.strip()

    if "name" in cols:
        df["اسم الصنف"] = df[cols["name"]].astype(str).str.strip()
    else:
        df["اسم الصنف"] = df["كود الصنف"]

    if "expected" in cols:
        df["الكمية الدفترية"] = numeric(df[cols["expected"]])

    if "actual" in cols:
        df["الكمية الفعلية"] = numeric(df[cols["actual"]])

    if "difference" in cols:
        df["الفرق"] = numeric(df[cols["difference"]])
    elif "actual" in cols and "expected" in cols:
        df["الفرق"] = (
            df["الكمية الفعلية"] - df["الكمية الدفترية"]
        )
    else:
        raise ValueError(
            "لا يمكن حساب الفرق بدون الكمية الفعلية."
        )

    if "الكمية الدفترية" not in df.columns:
        df["الكمية الدفترية"] = float("nan")

    if "الكمية الفعلية" not in df.columns:
        df["الكمية الفعلية"] = float("nan")

    df["الفرق"] = df["الفرق"].fillna(0)
    df["الشهر"] = month_label

    # تجميع كل كود في كل شهر، دون تكرار
    grouped = df.groupby("كود الصنف", as_index=False).agg({
        "اسم الصنف": "first",
        "الكمية الدفترية": "sum",
        "الكمية الفعلية": "sum",
        "الفرق": "sum",
    })

    grouped["الشهر"] = month_label
    grouped["كمية العجز"] = (-grouped["الفرق"]).clip(lower=0)
    grouped["كمية الزيادة"] = grouped["الفرق"].clip(lower=0)
    grouped["حالة الجرد"] = grouped["الفرق"].apply(
        lambda x: "عجز" if x < 0 else ("زيادة" if x > 0 else "مطابق")
    )

    return grouped


def calculate_metrics(df):
    total_codes = df["كود الصنف"].nunique()
    shortage_codes = (df["الفرق"] < 0).sum()
    surplus_codes = (df["الفرق"] > 0).sum()
    mismatch_codes = (df["الفرق"] != 0).sum()

    total_shortage = df["كمية العجز"].sum()
    total_surplus = df["كمية الزيادة"].sum()

    reference = df["الكمية الدفترية"].sum(min_count=1)

    if pd.notna(reference) and reference > 0:
        loss_rate = total_shortage / reference * 100
    else:
        loss_rate = None

    accuracy_error_rate = (
        mismatch_codes / total_codes * 100
        if total_codes else 0
    )

    return {
        "codes": total_codes,
        "shortage_codes": int(shortage_codes),
        "surplus_codes": int(surplus_codes),
        "mismatch_codes": int(mismatch_codes),
        "shortage": total_shortage,
        "surplus": total_surplus,
        "loss_rate": loss_rate,
        "mismatch_rate": accuracy_error_rate,
    }


def fmt_percent(value):
    return "غير متاح" if value is None else f"{value:.2f}%"


def excel_bytes(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, data in sheets.items():
            data.to_excel(writer, sheet_name=name[:31], index=False)
    return output.getvalue()


# ---------------------------
# Header
# ---------------------------

st.title("📦 Inventory Intelligence")
st.caption("نظام تحليل الفقد ودقة الجرد — سبتمبر وأكتوبر")

st.sidebar.header("⚙️ إعدادات التقرير")
st.sidebar.write("ارفع ملف كل شهر بصورة مستقلة.")
st.sidebar.info(
    "يدعم Excel وCSV. لا يتم استخدام أسعار أو قيم مالية."
)

col1, col2 = st.columns(2)

with col1:
    file_sep = st.file_uploader(
        "📁 ملف جرد سبتمبر",
        type=["xlsx", "xls", "xlsm", "csv"],
        key="september",
    )

with col2:
    file_oct = st.file_uploader(
        "📁 ملف جرد أكتوبر",
        type=["xlsx", "xls", "xlsm", "csv"],
        key="october",
    )

if not file_sep or not file_oct:
    st.info("ارفع ملف سبتمبر وملف أكتوبر لعرض التقرير.")
    st.stop()

try:
    september = prepare_data(read_file(file_sep), "سبتمبر")
    october = prepare_data(read_file(file_oct), "أكتوبر")
except Exception as exc:
    st.error(f"تعذر تجهيز التقرير: {exc}")
    st.stop()

# ---------------------------
# Metrics
# ---------------------------

m9 = calculate_metrics(september)
m10 = calculate_metrics(october)

st.divider()
st.subheader("📊 مقارنة المؤشرات الرئيسية")

c1, c2, c3, c4 = st.columns(4)

c1.metric(
    "عجز سبتمبر",
    f'{m9["shortage"]:,.2f}',
    help="مجموع كميات العجز في سبتمبر",
)
c2.metric(
    "عجز أكتوبر",
    f'{m10["shortage"]:,.2f}',
    delta=f'{m10["shortage"] - m9["shortage"]:,.2f}',
    delta_color="inverse",
)
c3.metric(
    "نسبة فقد سبتمبر",
    fmt_percent(m9["loss_rate"]),
)
c4.metric(
    "نسبة فقد أكتوبر",
    fmt_percent(m10["loss_rate"]),
)

c5, c6, c7, c8 = st.columns(4)

c5.metric("أكواد سبتمبر", f'{m9["codes"]:,}')
c6.metric("أكواد أكتوبر", f'{m10["codes"]:,}')
c7.metric(
    "الجرد غير المضبوط — سبتمبر",
    fmt_percent(m9["mismatch_rate"]),
)
c8.metric(
    "الجرد غير المضبوط — أكتوبر",
    fmt_percent(m10["mismatch_rate"]),
)

st.caption(
    "نسبة الفقد = كمية العجز ÷ إجمالي الكمية الدفترية × 100. "
    "نسبة الجرد غير المضبوط = عدد الأكواد ذات الفرق ÷ "
    "إجمالي الأكواد المجردة × 100."
)

# ---------------------------
# Monthly detail
# ---------------------------

tab1, tab2, tab3, tab4 = st.tabs([
    "📉 الأصناف الفاقدة",
    "📈 الزيادة",
    "🔎 تفاصيل الجرد",
    "📊 المقارنة",
])

with tab1:
    st.subheader("الأصناف التي بها عجز")
    shortages = pd.concat([september, october], ignore_index=True)
    shortages = shortages[shortages["كمية العجز"] > 0].copy()
    shortages = shortages.sort_values("كمية العجز", ascending=False)

    if shortages.empty:
        st.success("لا توجد أصناف بها عجز في الملفات المرفوعة.")
    else:
        st.dataframe(shortages[[
            "الشهر", "كود الصنف", "اسم الصنف",
            "كمية العجز", "الكمية الدفترية",
            "الكمية الفعلية", "الفرق"
        ]], use_container_width=True, hide_index=True)

with tab2:
    st.subheader("الأصناف التي بها زيادة")
    surpluses = pd.concat([september, october], ignore_index=True)
    surpluses = surpluses[surpluses["كمية الزيادة"] > 0].copy()
    surpluses = surpluses.sort_values("كمية الزيادة", ascending=False)

    if surpluses.empty:
        st.info("لا توجد أصناف بها زيادة.")
    else:
        st.dataframe(surpluses[[
            "الشهر", "كود الصنف", "اسم الصنف",
            "كمية الزيادة", "الكمية الدفترية",
            "الكمية الفعلية", "الفرق"
        ]], use_container_width=True, hide_index=True)

with tab3:
    st.subheader("تفاصيل الجرد")
    all_data = pd.concat([september, october], ignore_index=True)
    search = st.text_input("ابحث بكود الصنف أو الاسم")

    if search:
        mask = (
            all_data["كود الصنف"].str.contains(search, case=False, na=False)
            | all_data["اسم الصنف"].str.contains(search, case=False, na=False)
        )
        all_data = all_data[mask]

    st.dataframe(
        all_data.sort_values(["الشهر", "كود الصنف"]),
        use_container_width=True,
        hide_index=True,
    )

with tab4:
    st.subheader("مقارنة العجز والزيادة")
    comparison = pd.DataFrame([
        {
            "الشهر": "سبتمبر",
            "كمية العجز": m9["shortage"],
            "كمية الزيادة": m9["surplus"],
            "أكواد بها عجز": m9["shortage_codes"],
            "أكواد بها زيادة": m9["surplus_codes"],
            "أكواد غير مطابقة": m9["mismatch_codes"],
            "نسبة الجرد غير المضبوط": m9["mismatch_rate"],
        },
        {
            "الشهر": "أكتوبر",
            "كمية العجز": m10["shortage"],
            "كمية الزيادة": m10["surplus"],
            "أكواد بها عجز": m10["shortage_codes"],
            "أكواد بها زيادة": m10["surplus_codes"],
            "أكواد غير مطابقة": m10["mismatch_codes"],
            "نسبة الجرد غير المضبوط": m10["mismatch_rate"],
        },
    ])

    st.dataframe(comparison, use_container_width=True, hide_index=True)

    chart = comparison.melt(
        id_vars="الشهر",
        value_vars=["كمية العجز", "كمية الزيادة"],
        var_name="نوع الفرق",
        value_name="الكمية",
    )

    fig = px.bar(
        chart,
        x="الشهر",
        y="الكمية",
        color="نوع الفرق",
        barmode="group",
        title="العجز والزيادة حسب الشهر",
        template="plotly_dark",
    )
    st.plotly_chart(fig, use_container_width=True)

# ---------------------------
# Export
# ---------------------------

st.divider()
st.subheader("⬇️ تحميل التقرير")

combined = pd.concat([september, october], ignore_index=True)

report = pd.DataFrame([
    {
        "الشهر": "سبتمبر",
        "إجمالي الأكواد": m9["codes"],
        "كمية العجز": m9["shortage"],
        "كمية الزيادة": m9["surplus"],
        "نسبة الفقد %": m9["loss_rate"],
        "نسبة الجرد غير المضبوط %": m9["mismatch_rate"],
    },
    {
        "الشهر": "أكتوبر",
        "إجمالي الأكواد": m10["codes"],
        "كمية العجز": m10["shortage"],
        "كمية الزيادة": m10["surplus"],
        "نسبة الفقد %": m10["loss_rate"],
        "نسبة الجرد غير المضبوط %": m10["mismatch_rate"],
    },
])

try:
    data = excel_bytes({
        "ملخص المقارنة": report,
        "تفاصيل الجرد": combined,
        "الأصناف الفاقدة": combined[combined["كمية العجز"] > 0],
        "الأصناف الزائدة": combined[combined["كمية الزيادة"] > 0],
    })

    st.download_button(
        "📥 تنزيل تقرير Excel الكامل",
        data=data,
        file_name="Inventory_September_October.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
except Exception as exc:
    st.error(f"تعذر إنشاء ملف Excel: {exc}")

st.caption(
    "تنبيه: تأكد من أن الفرق محسوب على أساس "
    "الكمية الفعلية ناقص الكمية الدفترية. "
    "إذا كانت بيانات الكمية الدفترية غير متاحة، "
    "فلن يمكن حساب نسبة الفقد بدقة."
)

