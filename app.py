import io
import re
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(
    page_title="متابعة محمد قطان - تقارير الجرد",
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

st.title("📦 متابعة محمد قطان - تقارير الجرد والفروقات")
st.caption("لوحة تحكم ذكية ومفصلة لتحليل أرصدة سبتمبر وأكتوبر")

uploaded_file = st.file_uploader("📁 ارفع ملف الإكسيل المجمع هنا", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, encoding="utf-8-sig")
        else:
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_name = excel_file.sheet_names[0]
            for s in excel_file.sheet_names:
                if "9" in s or "10" in s:
                    sheet_name = s
                    break
            df = pd.read_excel(excel_file, sheet_name=sheet_name)

        # استخراج الشهر تلقائياً من عمود Date أو Source.Name
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

        if "Difference" in df.columns:
            total_items = len(df)
            total_diff = df["Difference"].sum()
            
            # العجز والزيادة بمنطق المخازن السليم
            shortage_df = df[df["Difference"] > 0] # الدفتر أكبر من الفعلي = عجز
            excess_df = df[df["Difference"] < 0]   # الفعلي أكبر من الدفتر = زيادة
            exact_df = df[df["Difference"] == 0]

            # مربعات إحصائية فخمة
            st.markdown("---")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("📦 إجمالي السجلات", f"{total_items:,}")
            col2.metric("📉 صافي الفروقات", f"{total_diff:,.2f}")
            col3.metric("⚠️ سجلات بها عجز", f"{len(shortage_df):,}")
            col4.metric("➕ سجلات بها زيادة", f"{len(excess_df):,}")
            st.markdown("---")

            # تبويبات مفصلة بالكامل
            tab1, tab2, tab3, tab4 = st.tabs([
                "📋 كل البيانات",
                "📉 تفاصيل العجز (النقص)",
                "➕ تفاصيل الزيادة",
                "🔍 بحث وفلترة متقدمة",
            ])

            with tab1:
                st.subheader("عرض شامل لكافة سجلات الجرد")
                st.dataframe(df, use_container_width=True)

            with tab2:
                st.subheader("⚠️ السجلات التي بها عجز (نقص في المخزون)")
                st.metric("عدد سجلات العجز", len(shortage_df))
                st.dataframe(shortage_df, use_container_width=True)
                st.download_button(
                    "📥 تحميل تقرير العجز فقط (CSV)",
                    shortage_df.to_csv(index=False).encode("utf-8-sig"),
                    file_name="shortage_report.csv",
                    mime="text/csv",
                )

            with tab3:
                st.subheader("➕ السجلات التي بها زيادة في المخزون")
                st.metric("عدد سجلات الزيادة", len(excess_df))
                st.dataframe(excess_df, use_container_width=True)
                st.download_button(
                    "📥 تحميل تقرير الزيادة فقط (CSV)",
                    excess_df.to_csv(index=False).encode("utf-8-sig"),
                    file_name="excess_report.csv",
                    mime="text/csv",
                )

            with tab4:
                st.subheader("🔍 البحث المخصص حسب الصنف أو الكود")
                search_query = st.text_input("أدخل اسم الصنف أو الكود أو الباركود:")
                if search_query:
                    filtered_search = df[
                        df.astype(str)
                        .apply(lambda r: r.str.contains(search_query, case=False).any(), axis=1)
                    ]
                    st.write(f"نتائج البحث المطابقة: {len(filtered_search)} سجل")
                    st.dataframe(filtered_search, use_container_width=True)
                else:
                    st.info("اكتب أي كلمة للبحث الفوري في كافة الأعمدة.")

        else:
            st.error("لم يتم العثور على عمود الفروقات أو الكميات بشكل صحيح في الملف.")

    except Exception as e:
        st.error(f"حدث خطأ أثناء قراءة ومعالجة الملف: {e}")
else:
  st.info("👈 يرجى رفع ملف الإكسيل المجمع لعرض التحليل المفصل بالكامل.")
