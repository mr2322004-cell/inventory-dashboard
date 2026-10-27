import io
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(
    page_title="متابعة محمد قطان - نظام إدارة المخازن والجرد الذكي",
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

st.title("📦 متابعة محمد قطان - نظام إدارة المخازن والجرد الذكي")
st.caption("النظام المتكامل: تقارير الشهور + دمج اليومي + التقرير المستقل + جرد الباركود والموبايل")

main_mode = st.sidebar.radio(
    "📂 اختر القسم المطلوب:",
    [
        "📊 التقرير التنفيذي (شهر 9 وشهر 10)", 
        "📂 دمج ملفات الجرد اليومي وتصدير الإكسيل", 
        "🔍 تقرير الفروقات المستقل (شامل وفوري)",
        "📱 جرد الباركود والموبايل (سريع وفوري)"
    ]
)

# ---------------------------------------------------------
# القسم الأول: التقرير الرئيسي (سبتمبر وأكتوبر)
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
                st.subheader("🔥 أعلى 30 صنف بها عجز ونقص")
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
# القسم الثاني: دمج تقارير الجرد اليومي وتصدير الإكسيل
# ---------------------------------------------------------
elif main_mode == "📂 دمج ملفات الجرد اليومي وتصدير الإكسيل":
    st.subheader("📂 مركز دمج ملفات الجرد اليومي واستخراج التقارير")
    st.caption("ارمي هنا كل تقارير الجرد اليومية، وسيتم تجميعهم وتوليد ملفات CSV نظيفة وسريعة للتحميل الفوري.")

    daily_files = st.file_uploader("ارفع تقارير الجرد اليومية (حدد ملفات متعددة CSV/Excel):", type=["csv", "xlsx", "xls"], accept_multiple_files=True)

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

            daily_agg = m_df.groupby(["رمز المادة", "اسم الصنف"], as_index=False).agg({
                "الكمية الدفترية": "sum",
                "الكمية الفعلية": "sum",
                "الفرق": "sum"
            })

            daily_agg["كمية العجز"] = daily_agg["الفرق"].apply(lambda x: x if x > 0 else 0)
            daily_agg["كمية الزيادة"] = daily_agg["الفرق"].apply(lambda x: -x if x < 0 else 0)

            shortage_df = daily_agg[daily_agg["كمية العجز"] > 0][["رمز المادة", "اسم الصنف", "الكمية الدفترية", "الكمية الفعلية", "كمية العجز"]].sort_values("كمية العجز", ascending=False)
            surplus_df = daily_agg[daily_agg["كمية الزيادة"] > 0][["رمز المادة", "اسم الصنف", "الكمية الدفترية", "الكمية الفعلية", "كمية الزيادة"]].sort_values("كمية الزيادة", ascending=False)
            reference_df = daily_agg[["رمز المادة", "اسم الصنف"]].drop_duplicates()

            st.success(f"تم دمج {len(daily_files)} ملف بنجاح وتجميع الأصناف المتشابهة!")

            d_tab1, d_tab2, d_tab3 = st.tabs(["🔥 شيت العجز", "➕ شيت الزيادة", "📚 مرجع الرموز"])

            with d_tab1:
                st.subheader("الأصناف التي بها عجز (نقص)")
                st.dataframe(shortage_df, use_container_width=True, hide_index=True)
                csv_shortage = shortage_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 تحميل ملف العجز", csv_shortage, "تقارير_العجز_اليومي.csv", "text/csv")

            with d_tab2:
                st.subheader("الأصناف التي بها زيادة (فائض)")
                st.dataframe(surplus_df, use_container_width=True, hide_index=True)
                csv_surplus = surplus_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 تحميل ملف الزيادة", csv_surplus, "تقارير_الزيادة_اليومية.csv", "text/csv")

            with d_tab3:
                st.subheader("مرجع رموز المواد والأسماء")
                st.dataframe(reference_df, use_container_width=True, hide_index=True)
                csv_ref = reference_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 تحميل ملف المرجع", csv_ref, "مرجع_الرموز_والاسماء.csv", "text/csv")

# ---------------------------------------------------------
# القسم الثالث: تقرير الفروقات المستقل
# ---------------------------------------------------------
elif main_mode == "🔍 تقرير الفروقات المستقل (شامل وفوري)":
    st.subheader("🔍 تقرير الفروقات المستقل والتحليل الفوري")
    st.caption("قسم خاص ومستقل لتحليل الفاقد والعجز والزيادة بدقة ومرونة عالية لأي ملفات جرد ترفعها الآن.")

    ind_files = st.file_uploader("ارفع ملفات الجرد للتحليل المستقل (CSV أو Excel):", type=["csv", "xlsx", "xls"], accept_multiple_files=True)

    if ind_files:
        ind_dfs = []
        for f in ind_files:
            try:
                if f.name.endswith(".csv"):
                    ind_dfs.append(pd.read_csv(f, encoding="utf-8-sig"))
                else:
                    ind_dfs.append(pd.read_excel(f))
            except Exception as e:
                st.error(f"خطأ في قراءة الملف {f.name}: {e}")

        if ind_dfs:
            master_ind = pd.concat(ind_dfs, ignore_index=True)

            c_code = "Code" if "Code" in master_ind.columns else master_ind.columns[0]
            c_name = "Detail" if "Detail" in master_ind.columns else master_ind.columns[2]
            c_sys = "SystemQty" if "SystemQty" in master_ind.columns else master_ind.columns[3]
            c_act = "ActualQty" if "ActualQty" in master_ind.columns else master_ind.columns[4]
            c_diff = "Difference" if "Difference" in master_ind.columns else master_ind.columns[5]

            master_ind["رمز المادة"] = master_ind[c_code].astype(str).str.strip()
            master_ind["اسم الصنف"] = master_ind[c_name].astype(str).str.strip()
            master_ind["الكمية الدفترية"] = pd.to_numeric(master_ind[c_sys], errors="coerce").fillna(0)
            master_ind["الكمية الفعلية"] = pd.to_numeric(master_ind[c_act], errors="coerce").fillna(0)
            master_ind["الفرق"] = pd.to_numeric(master_ind[c_diff], errors="coerce").fillna(0)

            final_ind = master_ind.groupby(["رمز المادة", "اسم الصنف"], as_index=False).agg({
                "الكمية الدفترية": "sum",
                "الكمية الفعلية": "sum",
                "الفرق": "sum"
            })

            final_ind["كمية العجز (الفاقد)"] = final_ind["الفرق"].apply(lambda x: x if x > 0 else 0)
            final_ind["كمية الزيادة"] = final_ind["الفرق"].apply(lambda x: -x if x < 0 else 0)
            final_ind["حالة الصنف"] = final_ind["الفرق"].apply(lambda x: "عجز (نقص)" if x > 0 else ("زيادة (فائض)" if x < 0 else "مطابق تماماً"))

            total_items_count = len(final_ind)
            shortage_count = len(final_ind[final_ind["كمية العجز (الفاقد)"] > 0])
            surplus_count = len(final_ind[final_ind["كمية الزيادة"] > 0])
            matched_count = len(final_ind[final_ind["الفرق"] == 0])

            col_a, col_b, col_c, col_d = st.columns(4)
            col_a.metric("📦 إجمالي الأصناف", f"{total_items_count:,}")
            col_b.metric("🔥 أصناف عجز", f"{shortage_count:,}")
            col_c.metric("➕ أصناف زيادة", f"{surplus_count:,}")
            col_d.metric("✅ أصناف مطابقة", f"{matched_count:,}")

            st.divider()

            t_tab1, t_tab2, t_tab3 = st.tabs(["🔥 تقرير الفاقد والعجز", "➕ تقرير الزيادة", "📋 كافة الفروقات"])

            with t_tab1:
                st.subheader("الأصناف التي سجلت عجزاً وفاقداً")
                ind_shortage = final_ind[final_ind["كمية العجز (الفاقد)"] > 0][["رمز المادة", "اسم الصنف", "الكمية الدفترية", "الكمية الفعلية", "كمية العجز (الفاقد)"]].sort_values("كمية العجز (الفاقد)", ascending=False)
                st.dataframe(ind_shortage, use_container_width=True, hide_index=True)
                csv_ind_s = ind_shortage.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 تحميل تقرير الفاقد المستقل (CSV)", csv_ind_s, "تقرير_الفاقد_المستقل.csv", "text/csv")

            with t_tab2:
                st.subheader("الأصناف التي سجلت زيادة")
                ind_surplus = final_ind[final_ind["كمية الزيادة"] > 0][["رمز المادة", "اسم الصنف", "الكمية الدفترية", "الكمية الفعلية", "كمية الزيادة"]].sort_values("كمية الزيادة", ascending=False)
                st.dataframe(ind_surplus, use_container_width=True, hide_index=True)
                csv_ind_su = ind_surplus.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 تحميل تقرير الزيادة المستقل (CSV)", csv_ind_su, "تقرير_الزيادة_المستقل.csv", "text/csv")

            with t_tab3:
                st.subheader("الشامل الكامل لكل الأصناف")
                st.dataframe(final_ind[["رمز المادة", "اسم الصنف", "الكمية الدفترية", "الكمية الفعلية", "الفرق", "حالة الصنف"]], use_container_width=True, hide_index=True)
                csv_all_ind = final_ind.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 تحميل التقرير الشامل المستقل (CSV)", csv_all_ind, "التقرير_الشامل_المستقل.csv", "text/csv")
    else:
        st.info("👈 يرجى رفع ملفات الجرد في هذا القسم لعرض التقرير المستقل وتحليل الفاقد والزيادة فوراً.")

# ---------------------------------------------------------
# القسم الرابع: جرد الباركود والموبايل (بحث بالباركود، الكمية النظامية، والغير مجرودة)
# ---------------------------------------------------------
else:
    st.subheader("📱 جرد الباركود والموبايل (سريع وفوري)")
    st.caption("ارفع ملف المخزن الأساسي، وامسح أو اكتب الباركود/اسم الصنف لتظهر الكمية النظامية وتدخل الفعلية فوراً، مع متابعة الأصناف غير المجرودة.")

    base_file = st.file_uploader("📁 ارفع ملف المخزن الأساسي (يحتوي على: المادة، الكمية الحالية، إلخ):", type=["xlsx", "xls", "csv"])

    if base_file is not None:
        try:
            if base_file.name.endswith(".csv"):
                base_df = pd.read_csv(base_file, encoding="utf-8-sig")
            else:
                base_df = pd.read_excel(base_file)

            c_material = "المادة" if "المادة" in base_df.columns else base_df.columns[1]
            c_qty = "الكمية الحالية" if "الكمية الحالية" in base_df.columns else base_df.columns[2]

            base_df["الصنف والمادة"] = base_df[c_material].astype(str).str.strip()
            base_df["الكمية النظامية"] = pd.to_numeric(base_df[c_qty], errors="coerce").fillna(0)

            if 'mobile_inventory_v3' not in st.session_state:
                st.session_state['mobile_inventory_v3'] = {
                    row["الصنف والمادة"]: {"نظامي": row["الكمية النظامية"], "فعلي": None}
                    for _, row in base_df.iterrows()
                }

            st.success("تم تحميل شيت المخزن بنجاح وجاهز للجرد عبر الموبايل!")

            scan_tab1, scan_tab2, scan_tab3 = st.tabs(["📷 ماسح الباركود والبحث الفوري", "⚠️ الأصناف غير المجرودة (مع كمياتها)", "📥 تصدير تقرير الجرد النهائي"])

            with scan_tab1:
                st.subheader("🔍 مسح الباركود أو البحث السريع")
                
                # استخدام حقل إدخال يدعم الباركود عند استخدام ماسح الكاميرا الخارجي أو اللاسلكي
                scan_input = st.text_input("مسح الباركود بالكاميرا أو كتابة اسم/كود الصنف للبحث الفوري:", placeholder="امسح الباركود أو اكتب هنا...")

                target_items = base_df.copy()
                if scan_input:
                    target_items = target_items[
                        target_items["الصنف والمادة"].str.contains(scan_input, case=False, na=False)
                    ]

                if not target_items.empty:
                    st.write(f"الأصناف المطابقة (إجمالي: {len(target_items)} صنف):")
                    for idx, row in target_items.head(10).iterrows():
                        item_name = row["الصنف والمادة"]
                        sys_q = row["الكمية النظامية"]
                        current_actual = st.session_state['mobile_inventory_v3'][item_name]["فعلي"]

                        st.markdown(f"---")
                        st.markdown(f"📦 **{item_name}**")
                        st.markdown(f"📌 **الكمية النظامية (الدفترية):** `{sys_q:,.2f}`")

                        col_in, col_btn = st.columns([3, 1])
                        new_actual = col_in.number_input(
                            "أدخل الكمية الفعلية:", 
                            min_value=0.0, 
                            value=float(current_actual) if current_actual is not None else float(sys_q), 
                            key=f"act_v3_{idx}"
                        )
                        
                        if col_btn.button("حفظ الفعلي ✓", key=f"btn_v3_{idx}"):
                            st.session_state['mobile_inventory_v3'][item_name]["فعلي"] = new_actual
                            st.success(f"تم حفظ الكمية الفعلية بنجاح!")
                else:
                    st.warning("لم يتم العثور على صنف مطابق. تأكد من الباركود أو الكود المدخل.")

            with scan_tab2:
                st.subheader("⚠️ الأصناف التي لم يتم جردها بعد (مع كمياتها النظامية)")
                ungraced_list = [
                    {"الصنف والمادة": k, "الكمية النظامية": v["نظامي"]}
                    for k, v in st.session_state['mobile_inventory_v3'].items() if v["فعلي"] is None
                ]
                df_ungraced = pd.DataFrame(ungraced_list)
                st.info(f"إجمالي الأصناف المتبقية ولم تُجرَد بعد: {len(df_ungraced)} صنف")
                
                if not df_ungraced.empty:
                    st.dataframe(df_ungraced, use_container_width=True, hide_index=True)

            with scan_tab3:
                st.subheader("📥 تقرير نتائج الجرد النهائي وتصدير الإكسيل")
                graced_list = []
                for k, v in st.session_state['mobile_inventory_v3'].items():
                    if v["فعلي"] is not None:
                        diff = v["نظامي"] - v["فعلي"]  # موجب عجز، سالب زيادة
                        status = "عجز (نقص)" if diff > 0 else ("زيادة (فائض)" if diff < 0 else "مطابق تماماً")
                        graced_list.append({
                            "الصنف والمادة": k,
                            "الكمية النظامية": v["نظامي"],
                            "الكمية الفعلية": v["فعلي"],
                            "الفرق": diff,
                            "حالة الصنف": status
                        })

                if graced_list:
                    df_final_result = pd.DataFrame(graced_list)
                    st.dataframe(df_final_result, use_container_width=True, hide_index=True)

                    csv_mob = df_final_result.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        "📥 تحميل تقرير الجرد النهائي (CSV جاهز للإكسيل)",
                        csv_mob,
                        "تقرير_جرد_الموبايل_النهائي.csv",
                        "text/csv"
                    )
                else:
                    st.warning("لم يتم جرد أي صنف حتى الآن لطباعة أو تصدير التقرير النهائي.")

        except Exception as e:
            st.error(f"حدث خطأ أثناء معالجة ملف المخزن: {e}")
    else:
        st.info("👈 يرجى رفع ملف المخزن الأساسي في هذا القسم لبدء الجرد السريع عبر الموبايل.")
