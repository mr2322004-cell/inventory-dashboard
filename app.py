import io
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
st.caption("نظام التحليل المالي والمخزني المتقدم لشهر سبتمبر وأكتوبر + أداة دمج وتصدير الجرد اليومي")

# إنشاء القائمة الرئيسية للتحكم
main_mode = st.sidebar.radio(
    "اختر القسم المطلوبة:",
    ["📊 التقرير التنفيذي (شهر 9 وشهر 10)", "📂 دمج ملفات الجرد اليومي وتصدير الإكسيل"]
)

# ---------------------------------------------------------
# القسم الأول: التقرير الرئيسي القديم (سبتمبر وأكتوبر)
# ---------------------------------------------------------
if main_mode == "📊 التقرير التنفيذي (شهر 9 وشهر 10)":
    st.subheader("📊 لوحة تحكم ومؤشرات أداء الجرد (سبتمبر وأكتوبر)")
    
    @st.cache_data
    def load_default_data():
        try:
            return pd.read_excel("تقرير شهرين.xlsx")
        except:
            return None

    df = load_default_data()
    uploaded_file = st.file_uploader("📁 ارفع ملف الإكسيل المجمع هنا (اختياري):", type=["xlsx", "xls", "csv"])

    if uploaded_file is not None:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, encoding="utf-8-sig")
        else:
            df = pd.read_excel(uploaded_file)

    if df is not None:
        try:
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

            # تجميع البيانات لمنع التكرار
            agg_df = df.groupby(["الشهر", "كود الصنف", "اسم الصنف"], as_index=False).agg({
                "الكمية الدفترية": "sum",
                "الكمية الفعلية": "sum",
                "الفرق": "sum"
            })

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

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("📉 نسبة الفقد — سبتمبر", f"{m9['loss_rate']:.2f}%", f"عجز: {m9['shortage']:,.2f}")
            col2.metric("📉 نسبة الفقد — أكتوبر", f"{m10['loss_rate']:.2f}%", f"عجز: {m10['shortage']:,.2f}", delta_color="inverse")
            col3.metric("⚠️ جرد غير مضبوط — سبتمبر", f"{m9['mismatch_rate']:.2f}%", f"{m9['mismatch_codes']:,} صنف")
            col4.metric("⚠️ جرد غير مضبوط — أكتوبر", f"{m10['mismatch_rate']:.2f}%", f"{m10['mismatch_codes']:,} صنف")

            st.divider()

            tab1, tab2, tab3 = st.tabs(["🔥 أعلى 30 صنف معجز", "📋 الملخص والمقارنة", "🔍 بحث وتصفية"])

            with tab1:
                st.subheader("🔥 أعلى 30 صنف بها عجز ونقص (أكثر الأصناف استنزافاً)")
                top_30 = agg_df.sort_values("كمية العجز", ascending=False).head(30)
                st.dataframe(top_30[["الشهر", "كود الصنف", "اسم الصنف", "الكمية الدفترية", "الكمية الفعلية", "كمية العجز"]], use_container_width=True, hide_index=True)

            with tab2:
                summary_table = pd.DataFrame([
                    {"الشهر": "سبتمبر", "إجمالي الأصناف": m9["codes"], "أصناف بها فروقات": m9["mismatch_codes"], "نسبة الفقد %": f"{m9['loss_rate']:.2f}%", "نسبة الجرد غير المضبوط %": f"{m9['mismatch_rate']:.2f}%"},
                    {"الشهر": "أكتوبر", "إجمالي الأصناف": m10["codes"], "أصناف بها فروقات": m10["mismatch_codes"], "نسبة الفقد %": f"{m10['loss_rate']:.2f}%", "نسبة الجرد غير المضبوط %": f"{m10['mismatch_rate']:.2f}%"}
                ])
                st.dataframe(summary_table, use_container_width=True)

            with tab3:
                search_q = st.text_input("ابحث بالكود أو الاسم:")
                if search_q:
                    st.dataframe(agg_df[agg_df["كود الصنف"].str.contains(search_q, case=False) | agg_df["اسم الصنف"].str.contains(search_q, case=False)], use_container_width=True, hide_index=True)
                else:
                    st.dataframe(agg_df, use_container_width=True, hide_index=True)

        except Exception as e:
            st.error(f"حدث خطأ في معالجة الملف: {e}")
    else:
        st.info("👈 يرجى رفع ملف الإكسيل المجمع لعرض التقارير والمؤشرات.")

# ---------------------------------------------------------
# القسم الثاني: الأداة الجديدة لدمج تقارير الجرد اليومية والتصدير
# ---------------------------------------------------------
else:
    st.subheader("📂 مركز دمج ملفات الجرد اليومي واستخراج التقارير")
    st.caption("ارمي هنا كل تقارير الجرد اليومية (مثل ملفات الكاشيرات والمجموعات)، وسيتم تجميعهم وتوليد ملفات إكسيل للعجز والزيادة والرموز.")

    daily_files = st.file_uploader("ارفع تقارير الجرد اليومية (يمكنك تحديد ملفات متعددة CSV/Excel):", type=["csv", "xlsx", "xls"], accept_multiple_files=True)

    if daily_files:
        dfs = []
        for f in daily_files:
            try:
                if f.name.endswith(".csv"):
                    dfs.append(pd.read_csv(f, encoding="utf-8-sig"))
                else:
                    dfs.append(pd.read_excel(f))
            except Exception as e:
                st.error(f"خطأ في قراءة {f.name}: {e}")

        if dfs:
            m_df = pd.concat(dfs, ignore_index=True)

            c_code = "Code" if "Code" in m_df.columns else m_df.columns[0]
            c_name = "Detail" if "Detail" in m_df.columns else m_df.columns[2]
            c_sys = "SystemQty" if "SystemQty" in m_df.columns else m_df.columns[3]
            c_act = "ActualQty" if "ActualQty" in m_df.columns else m_df.columns[4]
            c_diff = "Difference" if "Difference" in m_df.columns else m_df.columns[5]

            m_df["رمز المادة"] = m_df[c_code].astype(str).str.strip()
            m_df["اسم الصنف"] = m_df[c_name].astype(str).str.strip()
            m_df["الكمية الدفترية"] = pd.to_numeric(m_df[c_sys], errors="coerce").fillna(0)
            m_df["الكمية الفعلية"] = pd.to_numeric(m_df[c_act], errors="coerce").fillna(0)
            m_df["الفرق"] = pd.to_numeric(m_df[c_diff], errors="coerce").fillna(0)

            # تجميع الأصناف المتشابهة لمنع التكرار
            daily_agg = m_df.groupby(["رمز المادة", "اسم الصنف"], as_index=False).agg({
                "الكمية الدفترية": "sum",
                "الكمية الفعلية": "sum",
                "الفرق": "sum"
            })

            daily_agg["كمية العجز"] = daily_agg["الفرق"].apply(lambda x: x if x > 0 else 0)
            daily_agg["كمية الزيادة"] = daily_agg["الفرق"].apply(lambda x: -x if x < 0 else 0)

            shortage_excel = daily_agg[daily_agg["كمية العجز"] > 0][["رمز المادة", "اسم الصنف", "الكمية الدفترية", "الكمية الفعلية", "كمية العجز"]].sort_values("كمية العجز", ascending=False)
            surplus_excel = daily_agg[daily_agg["كمية الزيادة"] > 0][["رمز المادة", "اسم الصنف", "الكمية الدفترية", "الكمية الفعلية", "كمية الزيادة"]].sort_values("كمية الزيادة", ascending=False)
            reference_excel = daily_agg[["رمز المادة", "اسم الصنف"]].drop_duplicates()

            st.success(f"تم دمج {len(daily_files)} ملف بنجاح وتجميع الأصناف المتشابهة!")

            d_tab1, d_tab2, d_tab3 = st.tabs(["🔥 شيت العجز (Excel)", "➕ شيت الزيادة (Excel)", "📚 مرجع الرموز"])

            with d_tab1:
                st.subheader("الأصناف التي بها عجز (نقص)")
                st.dataframe(shortage_excel, use_container_width=True, hide_index=True)
                
                out1 = io.BytesIO()
                with pd.ExcelWriter(out1, engine='xlsxwriter') as writer:
                    shortage_excel.to_excel(writer, sheet_name='العجز', index=False)
                st.download_button("📥 تحميل شيت العجز (Excel)", out1.getvalue(), "تقارير_العجز_اليومي.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with d_tab2:
                st.subheader("الأصناف التي بها زيادة (فائض)")
                st.dataframe(surplus_excel, use_container_width=True, hide_index=True)
                
                out2 = io.BytesIO()
                with pd.ExcelWriter(out2, engine='xlsxwriter') as writer:
                    surplus_excel.to_excel(writer, sheet_name='الزيادة', index=False)
                st.download_button("📥 تحميل شيت الزيادة (Excel)", out2.getvalue(), "تقارير_الزيادة_اليومية.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with d_tab3:
                st.subheader("مرجع رموز المواد والأسماء")
                st.dataframe(reference_excel, use_container_width=True, hide_index=True)
                
                out3 = io.BytesIO()
                with pd.ExcelWriter(out3, engine='xlsxwriter') as writer:
                    reference_excel.to_excel(writer, sheet_name='مرجع_الرموز', index=False)
                st.download_button("📥 تحميل شيت المرجع (Excel)", out3.getvalue(), "مرجع_الرموز_والاسماء.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
