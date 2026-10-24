import io
import re
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(
    page_title="متابعة محمد قطان - تقارير الجرد والفروقات",
    page_icon="📦",
    layout="wide",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
html, body, [class*="css"] {font-family: 'Cairo', sans-serif;}
.stApp {background: #0b1120; color: #e5e7eb;}
[data-testid="stMetric"] {
    background: #151f32;
    padding: 16px;
    border: 1px solid #29364b;
    border-radius: 14px;
}
</style>
""", unsafe_allow_html=True)

st.title("📦 متابعة محمد قطان - تقارير الجرد والفروقات والتحليل المتقدم")
st.caption("تقارير دقيقة ومفصلة لشهر سبتمبر وأكتوبر (الموجب عجز / السالب زيادة)")

# محاولة تحميل الملف تلقائياً من المستودع أو السماح برفعه يدويًا كخيار احتياطي
@st.cache_data
def load_default_data():
    try:
        # ضع هنا اسم ملف الإكسيل الموجود في جيت هاب بدقة (مثلاً report.xlsx أو تقرير شهرين.xlsx)
        return pd.read_excel("تقرير شهرين.xlsx")
    except:
        return None

df = load_default_data()

uploaded_file = st.file_uploader("📁 أو ارفع ملف الإكسيل يدوياً من هنا (اختياري)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, encoding="utf-8-sig")
    else:
        df = pd.read_excel(uploaded_file)

if df is not None:
    try:
        # استخراج الشهر بدقة من تواريخ الملف
        date_col = None
        for col in df.columns:
            if "date" in str(col).lower() or "تاريخ" in str(col):
                date_col = col
                break
            elif "source" in str(col).lower():
                date_col = col
                break

        if date_col:
            dt_series = pd.to_datetime(df[date_col], errors="coerce")
            if dt_series.isna().all() and "source" in str(date_col).lower():
                extracted_dates = df[date_col].astype(str).str.extract(r'(\d{4}-\d{2}-\d{2})')[0]
                dt_series = pd.to_datetime(extracted_dates, errors="coerce")

            df["الشهر_رقم"] = dt_series.dt.month
            df["الشهر"] = df["الشهر_رقم"].map({9: "سبتمبر", 10: "أكتوبر"})
        
        if "الشهر" not in df.columns or df["الشهر"].isna().all():
            full_text = df.astype(str).agg(' '.join, axis=1).str.lower()
            df["الشهر"] = "سبتمبر"
            df.loc[full_text.str.contains("10|oct|أكتوبر|اكتوبر", na=False), "الشهر"] = "أكتوبر"
            df.loc[full_text.str.contains("9|sep|سبتمبر", na=False), "الشهر"] = "سبتمبر"

        if "Difference" not in df.columns and "SystemQty" in df.columns and "ActualQty" in df.columns:
            df["Difference"] = df["SystemQty"] - df["ActualQty"]

        code_col = "Code" if "Code" in df.columns else df.columns[1]
        name_col = "Detail" if "Detail" in df.columns else df.columns[3]

        df["كود الصنف"] = df[code_col].astype(str).str.strip()
        df["اسم الصنف"] = df[name_col].astype(str).str.strip()
        df["الكمية الدفترية"] = pd.to_numeric(df["SystemQty"], errors="coerce").fillna(0)
        df["الكمية الفعلية"] = pd.to_numeric(df["ActualQty"], errors="coerce").fillna(0)
        df["الفرق"] = pd.to_numeric(df["Difference"], errors="coerce").fillna(0)

        # تجميع البيانات لمنع التكرار لكل شهر
        agg_df = df.groupby(["الشهر", "كود الصنف", "اسم الصنف"], as_index=False).agg({
            "الكمية الدفترية": "sum",
            "الكمية الفعلية": "sum",
            "الفرق": "sum"
        })

        # حسب طلبك: الموجب عجز (نقص) والسالب زيادة
        agg_df["كمية العجز"] = agg_df["الفرق"].apply(lambda x: x if x > 0 else 0)
        agg_df["كمية الزيادة"] = agg_df["الفرق"].apply(lambda x: -x if x < 0 else 0)

        def get_month_metrics(m_name):
            m_data = agg_df[agg_df["الشهر"] == m_name]
            total_codes = m_data["كود الصنف"].nunique()
            mismatch_codes = m_data[m_data["الفرق"] != 0]["كود الصنف"].nunique()
            total_system = m_data["الكمية الدفترية"].sum()
            total_shortage = m_data["كمية العجز"].sum()
            total_surplus = m_data["كمية الزيادة"].sum()
            
            loss_rate = (total_shortage / total_system * 100) if total_system > 0 else 0
            mismatch_rate = (mismatch_codes / total_codes * 100) if total_codes > 0 else 0
            
            return {
                "codes": total_codes,
                "mismatch_codes": mismatch_codes,
                "system": total_system,
                "shortage": total_shortage,
                "surplus": total_surplus,
                "loss_rate": loss_rate,
                "mismatch_rate": mismatch_rate,
                "data": m_data
            }

        m9 = get_month_metrics("سبتمبر")
        m10 = get_month_metrics("أكتوبر")

        st.divider()
        st.subheader("📊 مؤشرات الأداء ونسب الفقد والجرد لشهر 9 وشهر 10")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📉 نسبة الفقد — سبتمبر", f"{m9['loss_rate']:.2f}%", f"عجز: {m9['shortage']:,.2f}")
        col2.metric("📉 نسبة الفقد — أكتوبر", f"{m10['loss_rate']:.2f}%", f"عجز: {m10['shortage']:,.2f}", delta_color="inverse")
        col3.metric("⚠️ جرد غير مضبوط — سبتمبر", f"{m9['mismatch_rate']:.2f}%", f"{m9['mismatch_codes']:,} صنف")
        col4.metric("⚠️ جرد غير مضبوط — أكتوبر", f"{m10['mismatch_rate']:.2f}%", f"{m10['mismatch_codes']:,} صنف")

        tab1, tab2, tab3, tab4 = st.tabs([
            "🔥 أعلى 30 صنف معجز (عجز)",
            "📋 الملخص والمقارنة",
            "➕ الأصناف الزائدة",
            "🔍 بحث وفلترة متقدمة",
        ])

        with tab1:
            st.subheader("🔥 أعلى 30 صنف بها عجز ونقص (أكثر الأصناف استنزافاً)")
            selected_month_top = st.selectbox("اختر الشهر لعرض أعلى 30 صنف:", ["الكل", "سبتمبر", "أكتوبر"])
            top_query_df = agg_df if selected_month_top == "الكل" else agg_df[agg_df["الشهر"] == selected_month_top]
            
            top_30_shortages = top_query_df.sort_values("كمية العجز", ascending=False).head(30)
            
            st.dataframe(
                top_30_shortages[["الشهر", "كود الصنف", "اسم الصنف", "الكمية الدفترية", "الكمية الفعلية", "الفرق", "كمية العجز"]],
                use_container_width=True,
                hide_index=True
            )

            st.plotly_chart(
                px.bar(
                    top_30_shortages.head(15),
                    x="كمية العجز",
                    y="اسم الصنف",
                    color="الشهر",
                    orientation="h",
                    title="أكثر الأصناف التي بها عجز بالمخزن",
                    template="plotly_dark"
                ),
                use_container_width=True
            )

        with tab2:
            st.subheader("مقارنة شاملة بين شهر سبتمبر وأكتوبر")
            summary_table = pd.DataFrame([
                {
                    "الشهر": "سبتمبر",
                    "إجمالي الأصناف": m9["codes"],
                    "أصناف بها فروقات": m9["mismatch_codes"],
                    "إجمالي الكمية الدفترية": round(m9["system"], 2),
                    "إجمالي العجز": round(m9["shortage"], 2),
                    "إجمالي الزيادة": round(m9["surplus"], 2),
                    "نسبة الفقد %": f"{m9['loss_rate']:.2f}%",
                    "نسبة الجرد غير المضبوط %": f"{m9['mismatch_rate']:.2f}%",
                },
                {
                    "الشهر": "أكتوبر",
                    "إجمالي الأصناف": m10["codes"],
                    "أصناف بها فروقات": m10["mismatch_codes"],
                    "إجمالي الكمية الدفترية": round(m10["system"], 2),
                    "إجمالي العجز": round(m10["shortage"], 2),
                    "إجمالي الزيادة": round(m10["surplus"], 2),
                    "نسبة الفقد %": f"{m10['loss_rate']:.2f}%",
                    "نسبة الجرد غير المضبوط %": f"{m10['mismatch_rate']:.2f}%",
                }
            ])
            st.dataframe(summary_table, use_container_width=True)

        with tab3:
            st.subheader("➕ الأصناف التي بها زيادة في المخزون (السالب)")
            surpluses_all = agg_df[agg_df["كمية الزيادة"] > 0].sort_values("كمية الزيادة", ascending=False)
            st.dataframe(surpluses_all[["الشهر", "كود الصنف", "اسم الصنف", "الكمية الدفترية", "الكمية الفعلية", "الفرق", "كمية الزيادة"]], use_container_width=True, hide_index=True)

        with tab4:
            st.subheader("🔍 بحث وفلترة متقدمة للأصناف بدون تكرار")
            search_query = st.text_input("ابحث باسم الصنف أو الكود:")
            month_choice = st.selectbox("اختر الشهر للبحث:", ["الكل", "سبتمبر", "أكتوبر"])
            
            filtered = agg_df.copy()
            if search_query:
                filtered = filtered[filtered["كود الصنف"].str.contains(search_query, case=False, na=False) | filtered["اسم الصنف"].str.contains(search_query, case=False, na=False)]
            if month_choice != "الكل":
                filtered = filtered[filtered["الشهر"] == month_choice]
                
            st.dataframe(filtered, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"حدث خطأ أثناء معالجة البيانات: {e}")
else:
  st.info("👈 يرجى رفع ملف الإكسيل في مستودع جيت هاب أو رفعه يدوياً لعرض المؤشرات الفورية.")
