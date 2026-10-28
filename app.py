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
st.caption("النظام المتكامل: تقارير الشهور + تقرير الأسبوع + دمج اليومي + التقرير المستقل + الجرد السريع + مقارنة الجرد الفعلي")

main_mode = st.sidebar.radio(
    "📂 اختر القسم المطلوب:",
    [
        "📊 التقرير التنفيذي (شهر 9 وشهر 10)", 
        "📅 تقرير الأسبوع (متابعة محمد قطان وتحليل الفاقد)",
        "📂 دمج ملفات الجرد اليومي وتصدير الإكسيل", 
        "🔍 تقرير الفروقات المستقل (شامل وفوري)",
        "📱 الجرد السريع والبحث الذكي بالموبايل",
        "⚖️ مقارنة الجرد الخارجي/الفعلي مع النظام والفروقات"
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
    uploaded_file = st.file_uploader("📁 ارفع ملف الإكسيل هنا (يقبل تقارير الشهور أو ملف المخزن العام):", type=["xlsx", "xls", "csv"])

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file, encoding="utf-8-sig")
            else:
                df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"خطأ في قراءة الملف: {e}")
            df = None

    if df is not None:
        try:
            is_warehouse_file = any("الكمية الحالية" in str(c) for c in df.columns) or any("المادة" in str(c) for c in df.columns)

            if is_warehouse_file:
                st.info("💡 لوحظ أنك رفعت ملف المخزن الأساسي هنا. إليك عرض مباشر لأصناف المخزن ورصيدها:")
                c_mat = "المادة" if "المادة" in df.columns else df.columns[1]
                c_q = "الكمية الحالية" if "الكمية الحالية" in df.columns else df.columns[2]
                for col in df.columns:
                    if "الكمية الحالية" in str(col):
                        c_q = col
                        break
                
                temp_view = df[[c_mat, c_q]].copy()
                temp_view.columns = ["اسم الصنف / المادة", "الكمية الحالية"]
                st.dataframe(temp_view, use_container_width=True, hide_index=True)
            else:
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
                    
                    loss_rate = (total_shortage / total_system * 100) if total_system > 0 else 0
                    mismatch_rate = (mismatch_codes / total_codes * 100) if total_codes > 0 else 0
                    
                    return {
                        "codes": total_codes,
                        "mismatch_codes": mismatch_codes,
                        "system": total_system,
                        "shortage": total_shortage,
                        "loss_rate": loss_rate,
                        "mismatch_rate": mismatch_rate,
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
                    top_30 = agg_df.sort_values("كمية العجز", ascending=False).head(30)
                    st.dataframe(top_30[["الشهر", "كود الصنف", "اسم الصنف", "الكمية الدفترية", "الكمية الفعلية", "كمية العجز"]], use_container_width=True, hide_index=True)
                with tab2:
                    summary_table = pd.DataFrame([
                        {"الشهر": "سبتمبر", "إجمالي الأصناف": m9["codes"], "أصناف بها فروقات": m9["mismatch_codes"], "نسبة الفقد %": f"{m9['loss_rate']:.2f}%"},
                        {"الشهر": "أكتوبر", "إجمالي الأصناف": m10["codes"], "أصناف بها فروقات": m10["mismatch_codes"], "نسبة الفقد %": f"{m10['loss_rate']:.2f}%"}
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
        st.info("👈 يرجى رفع ملف الإكسيل لعرض التقارير.")

# ---------------------------------------------------------
# القسم الثاني: تقرير الأسبوع (متابعة محمد قطان وتحليل الفاقد) - جديد وجبار
# ---------------------------------------------------------
elif main_mode == "📅 تقرير الأسبوع (متابعة محمد قطان وتحليل الفاقد)":
    st.subheader("📅 تقرير الأسبوع — متابعة محمد قطان وتحليل الفاقد والزيادة")
    st.caption("ارفع هنا جميع ملفات جرد أيام الأسبوع، وسيتم تجميعها وتحليلها لتحديد الأصناف التي بها عجز كبير، وزيادة متكررة، والمواد التي تحتاج إلى ضبط عاجل.")

    weekly_files = st.file_uploader("📁 ارفع تقارير الجرد لأيام الأسبوع (حدد عدة ملفات Excel أو CSV):", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

    if weekly_files:
        try:
            week_dfs = []
            for wf in weekly_files:
                if wf.name.endswith(".csv"):
                    week_dfs.append(pd.read_csv(wf, encoding="utf-8-sig"))
                else:
                    week_dfs.append(pd.read_excel(wf))

            combined_week = pd.concat(week_dfs, ignore_index=True)

            # استخراج وتوحيد الأعمدة بدقة
            w_code = combined_week.columns[0]
            w_name = combined_week.columns[1] if len(combined_week.columns) > 1 else combined_week.columns[0]
            w_sys = combined_week.columns[2] if len(combined_week.columns) > 2 else combined_week.columns[1]
            w_act = combined_week.columns[3] if len(combined_week.columns) > 3 else combined_week.columns[2]
            w_diff = combined_week.columns[4] if len(combined_week.columns) > 4 else combined_week.columns[3]

            for col in combined_week.columns:
                cs = str(col).lower()
                if any(k in cs for k in ["code", "كود", "رمز"]): w_code = col
                if any(k in cs for k in ["detail", "name", "صنف", "مادة", "المادة"]): w_name = col
                if any(k in cs for k in ["system", "دفتر", "نظام", "current"]): w_sys = col
                if any(k in cs for k in ["actual", "فعلي"]): w_act = col
                if any(k in cs for k in ["diff", "فرق"]): w_diff = col

            combined_week["الكود"] = combined_week[w_code].astype(str).str.strip()
            combined_week["اسم الصنف"] = combined_week[w_name].astype(str).str.strip()
            combined_week["الكمية الدفترية"] = pd.to_numeric(combined_week[w_sys], errors="coerce").fillna(0)
            combined_week["الكمية الفعلية"] = pd.to_numeric(combined_week[w_act], errors="coerce").fillna(0)
            combined_week["الفرق الأسبوعي"] = combined_week["الكمية الفعلية"] - combined_week["الكمية الدفترية"]

            # تجميع البيانات على مستوى الصنف للأسبوع كله
            weekly_agg = combined_week.groupby(["الكود", "اسم الصنف"], as_index=False).agg({
                "الكمية الدفترية": "sum",
                "الكمية الفعلية": "sum",
                "الفرق الأسبوعي": "sum"
            })

            weekly_agg["كمية العجز (الفاقد)"] = weekly_agg["الفرق الأسبوعي"].apply(lambda x: -x if x < 0 else 0)
            weekly_agg["كمية الزيادة"] = weekly_agg["الفرق الأسبوعي"].apply(lambda x: x if x > 0 else 0)
            weekly_agg["حالة الصنف"] = weekly_agg["الفرق الأسبوعي"].apply(lambda x: "عجز (نقص خطر)" if x < 0 else ("زيادة (فائض)" if x > 0 else "مطابق تماماً"))

            tot_w_items = len(weekly_agg)
            w_shortage_items = len(weekly_agg[weekly_agg["الفرق الأسبوعي"] < 0])
            w_surplus_items = len(weekly_agg[weekly_agg["الفرق الأسبوعي"] > 0])
            w_total_def = weekly_agg["كمية العجز (الفاقد)"].sum()

            st.success(f"تم تجميع وتحليل عدد {len(weekly_files)} ملفات أسبوعية بنجاح!")

            w_m1, w_m2, w_m3, w_m4 = st.columns(4)
            w_m1.metric("📦 إجمالي أصناف الأسبوع", f"{tot_w_items:,}")
            w_m2.metric("🔥 أصناف بها عجز وفاقد", f"{w_shortage_items:,}")
            w_m3.metric("📉 إجمالي كمية العجز أسبوعياً", f"{w_total_def:,.2f}")
            w_m4.metric("➕ أصناف بها زيادة", f"{w_surplus_items:,}")

            st.divider()

            w_tab1, w_tab2, w_tab3, w_tab4 = st.tabs([
                "🚨 المواد التي تعجز بكثرة وتحتاج ضبط", 
                "🔥 تقرير العجز الأسبوعي", 
                "➕ تقرير الزيادة الأسبوعية", 
                "📋 التقرير الأسبوعي الشامل"
            ])

            with w_tab1:
                st.subheader("🚨 المواد الأكثر عجزاً وفاقداً خلال الأسبوع (تتطلب ضبط عاجل)")
                st.caption("هذه القائمة ترتب الأصناف التي سجلت أعلى كميات عجز ونقص خلال أيام الأسبوع لمتابعتها وفحصها.")
                top_deficits = weekly_agg[weekly_agg["كمية العجز (الفاقد)"] > 0].sort_values("كمية العجز (الفاقد)", ascending=False).head(30)
                st.dataframe(top_deficits, use_container_width=True, hide_index=True)
                
                if not top_deficits.empty:
                    csv_top_def = top_deficits.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        "📥 تحميل تقرير المواد المعجزة بكثرة (CSV)",
                        csv_top_def,
                        "المواد_التي_تحتاج_ضبط_عاجل.csv",
                        "text/csv"
                    )

            with w_tab2:
                st.subheader("🔥 تقرير العجز الأسبوعي الكامل")
                df_w_short = weekly_agg[weekly_agg["الفرق الأسبوعي"] < 0].sort_values("الفرق الأسبوعي", ascending=True)
                st.dataframe(df_w_short, use_container_width=True, hide_index=True)
                if not df_w_short.empty:
                    csv_w_short = df_w_short.to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📥 تحميل تقرير العجز الأسبوعي (CSV)", csv_w_short, "تقرير_العجز_الأسبوعي.csv", "text/csv")

            with w_tab3:
                st.subheader("➕ تقرير الزيادة أسبوعياً")
                df_w_surp = weekly_agg[weekly_agg["الفرق الأسبوعي"] > 0].sort_values("الفرق الأسبوعي", ascending=False)
                st.dataframe(df_w_surp, use_container_width=True, hide_index=True)
                if not df_w_surp.empty:
                    csv_w_surp = df_w_surp.to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📥 تحميل تقرير الزيادة الأسبوعي (CSV)", csv_w_surp, "تقرير_الزيادة_الأسبوعي.csv", "text/csv")

            with w_tab4:
                st.subheader("📋 التقرير الأسبوعي الشامل لكافة الأصناف")
                st.dataframe(weekly_agg, use_container_width=True, hide_index=True)
                csv_w_all = weekly_agg.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 تحميل التقرير الأسبوعي الشامل (CSV)", csv_w_all, "التقرير_الأسبوعي_الشامل.csv", "text/csv")

        except Exception as e:
            st.error(f"حدث خطأ أثناء معالجة ملفات الأسبوع: {e}")
    else:
        st.info("👈 يرجى رفع تقارير الجرد الخاصة بأيام الأسبوع في هذا القسم لتوليد تقرير الأسبوع الشامل وتحليل المواد المعجزة.")

# ---------------------------------------------------------
# القسم الثالث: دمج تقارير الجرد اليومي وتصدير الإكسيل
# ---------------------------------------------------------
elif main_mode == "📂 دمج ملفات الجرد اليومي وتصدير الإكسيل":
    st.subheader("📂 مركز دمج ملفات الجرد اليومي واستخراج التقارير")
    daily_files = st.file_uploader("ارفع تقارير الجرد اليومية:", type=["csv", "xlsx", "xls"], accept_multiple_files=True)
    if daily_files:
        dfs = [pd.read_csv(f, encoding="utf-8-sig") if f.name.endswith(".csv") else pd.read_excel(f) for f in daily_files]
        m_df = pd.concat(dfs, ignore_index=True)
        st.success("تم الدمج بنجاح!")
        st.dataframe(m_df.head(20), use_container_width=True)

# ---------------------------------------------------------
# القسم الرابع: تقرير الفروقات المستقل
# ---------------------------------------------------------
elif main_mode == "🔍 تقرير الفروقات المستقل (شامل وفوري)":
    st.subheader("🔍 تقرير الفروقات المستقل والتحليل الفوري")
    ind_files = st.file_uploader("ارفع ملفات الجرد للتحليل:", type=["csv", "xlsx", "xls"], accept_multiple_files=True)
    if ind_files:
        ind_dfs = [pd.read_csv(f, encoding="utf-8-sig") if f.name.endswith(".csv") else pd.read_excel(f) for f in ind_files]
        master_ind = pd.concat(ind_dfs, ignore_index=True)
        st.success("تم التحليل بنجاح!")
        st.dataframe(master_ind.head(20), use_container_width=True)

# ---------------------------------------------------------
# القسم الخامس: الجرد السريع والبحث الذكي بالموبايل
# ---------------------------------------------------------
elif main_mode == "📱 الجرد السريع والبحث الذكي بالموبايل":
    st.subheader("📱 الجرد السريع والبحث الذكي بالموبايل")
    st.caption("ارفع هنا ملف المخزن الأساسي (جرد المخزن 22)، وابحث بالاسم أو الكود لتعديل الكمية الفعلية ومتابعة الفروقات.")

    base_file = st.file_uploader("📁 ارفع ملف المخزن الأساسي (جرد المخزن 22):", type=["xlsx", "xls", "csv"])

    if base_file is not None:
        try:
            if base_file.name.endswith(".csv"):
                base_df = pd.read_csv(base_file, encoding="utf-8-sig")
            else:
                base_df = pd.read_excel(base_file)

            c_material = "المادة" if "المادة" in base_df.columns else base_df.columns[1]
            c_qty = None
            for col in base_df.columns:
                if "الكمية الحالية" in str(col).strip():
                    c_qty = col
                    break
            if not c_qty:
                c_qty = base_df.columns[2]

            base_df["الصنف والمادة"] = base_df[c_material].astype(str).str.strip()
            base_df["الكمية النظامية"] = pd.to_numeric(base_df[c_qty], errors="coerce").fillna(0)

            if 'mobile_inventory_v11' not in st.session_state:
                st.session_state['mobile_inventory_v11'] = {
                    row["الصنف والمادة"]: {"نظامي": row["الكمية النظامية"], "فعلي": row["الكمية النظامية"]}
                    for _, row in base_df.iterrows()
                }

            st.success("تم تحميل ملف المخزن وقراءة الكميات النظامية بنجاح تام!")

            scan_tab1, scan_tab2, scan_tab3 = st.tabs(["🔍 البحث السريع وتعديل الكميات", "⚠️ الأصناف غير المجرودة", "📥 تقرير الجرد النهائي وتصدير الإكسيل"])

            with scan_tab1:
                st.subheader("📝 جدول الجرد السريع والبحث الذكي")
                search_term = st.text_input("ابحث باسم الصنف أو الكود (بحث تقريبي سريع):", placeholder="اكتب اسم الصنف أو الكود هنا...")

                display_df = base_df.copy()
                if search_term:
                    display_df = display_df[
                        display_df["الصنف والمادة"].str.contains(search_term, case=False, na=False)
                    ]

                st.write(f"عرض عدد {len(display_df)} صنف (يمكنك تعديل الكمية الفعلية مباشرة):")

                for idx, row in display_df.head(30).iterrows():
                    item_name = row["الصنف والمادة"]
                    sys_q = row["الكمية النظامية"]
                    
                    if item_name not in st.session_state['mobile_inventory_v11']:
                        st.session_state['mobile_inventory_v11'][item_name] = {"نظامي": sys_q, "فعلي": sys_q}

                    current_actual = st.session_state['mobile_inventory_v11'][item_name]["فعلي"]

                    col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
                    col1.markdown(f"**{item_name}**")
                    col2.markdown(f"النظام: **{sys_q}**")
                    
                    new_act = col3.number_input(
                        "الفعلي", 
                        value=float(current_actual), 
                        key=f"act_direct_v11_{idx}",
                        label_visibility="collapsed"
                    )
                    
                    st.session_state['mobile_inventory_v11'][item_name]["فعلي"] = new_act
                    
                    diff = new_act - sys_q
                    
                    if diff > 0:
                        col4.markdown(f"زيادة: **+{diff}**")
                    elif diff < 0:
                        col4.markdown(f"عجز: **{diff}**")
                    else:
                        col4.markdown(f"الفرق: **0 (مطابق)**")
                    st.markdown("---")

            with scan_tab2:
                st.subheader("⚠️ الأصناف التي لم يتم تعديلها أو جردها بعد (مطابقة للنظام)")
                ungraced_list = [
                    {"الصنف والمادة": k, "الكمية النظامية": v["نظامي"], "الكمية الفعلية": v["فعلي"]}
                    for k, v in st.session_state['mobile_inventory_v11'].items() if v["نظامي"] == v["فعلي"]
                ]
                df_ungraced = pd.DataFrame(ungraced_list)
                st.info(f"إجمالي الأصناف التي بنفس الكمية النظامية: {len(df_ungraced)} صنف")
                if not df_ungraced.empty:
                    st.dataframe(df_ungraced, use_container_width=True, hide_index=True)

            with scan_tab3:
                st.subheader("📥 تقرير نتائج الجرد النهائي وتصدير الإكسيل")
                final_list = []
                for k, v in st.session_state['mobile_inventory_v11'].items():
                    diff = v["فعلي"] - v["نظامي"]
                    if diff < 0:
                        status = "عجز (نقص)"
                    elif diff > 0:
                        status = "زيادة (فائض)"
                    else:
                        status = "مطابق تماماً"

                    final_list.append({
                        "الصنف والمادة": k,
                        "الكمية النظامية": v["نظامي"],
                        "الكمية الفعلية": v["فعلي"],
                        "الفرق": diff,
                        "حالة الصنف": status
                    })

                df_final = pd.DataFrame(final_list)
                st.dataframe(df_final, use_container_width=True, hide_index=True)

                csv_mob = df_final.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    "📥 تحميل تقرير الجرد النهائي (CSV جاهز للإكسيل)",
                    csv_mob,
                    "تقرير_جرد_الموبايل_النهائي.csv",
                    "text/csv"
                )

        except Exception as e:
            st.error(f"حدث خطأ أثناء معالجة ملف المخزن: {e}")
    else:
        st.info("👈 يرجى رفع ملف المخزن الأساسي (جرد المخزن 22) في **القسم الرابع** لبدء الجرد السريع وتعديل الكميات.")

# ---------------------------------------------------------
# القسم السادس: مقارنة الجرد الفعلي الخارجي مع النظام
# ---------------------------------------------------------
else:
    st.subheader("⚖️ قسم مقارنة الجرد الفعلي الخارجي مع النظام واستخراج الفروقات")
    st.caption("ارفع ملف النظام الأساسي وملف الجرد الخارجي/اليدوي لمقارنتهما أوتوماتيكياً واستخراج العجز والزيادة.")

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        sys_file = st.file_uploader("📁 1. ارفع ملف النظام الأساسي (المادة، الكمية الحالية):", type=["xlsx", "xls", "csv"], key="sys_file_ext")
    with col_up2:
        ext_file = st.file_uploader("📁 2. ارفع ملف الجرد الفعلي الخارجي/اليدوي:", type=["xlsx", "xls", "csv"], key="ext_file_ext")

    if sys_file is not None and ext_file is not None:
        try:
            sys_df = pd.read_csv(sys_file, encoding="utf-8-sig") if sys_file.name.endswith(".csv") else pd.read_excel(sys_file)
            ext_df = pd.read_csv(ext_file, encoding="utf-8-sig") if ext_file.name.endswith(".csv") else pd.read_excel(ext_file)

            c_sys_mat = "المادة" if "المادة" in sys_df.columns else sys_df.columns[1]
            c_sys_qty = "الكمية الحالية"
            for col in sys_df.columns:
                if "الكمية الحالية" in str(col).strip():
                    c_sys_qty = col
                    break

            sys_df["الصنف"] = sys_df[c_sys_mat].astype(str).str.strip()
            sys_df["الكمية النظامية"] = pd.to_numeric(sys_df[c_sys_qty], errors="coerce").fillna(0)

            c_ext_mat = ext_df.columns[0]
            c_ext_qty = ext_df.columns[1] if len(ext_df.columns) > 1 else ext_df.columns[0]
            for col in ext_df.columns:
                col_s = str(col).strip()
                if any(k in col_s for k in ["مادة", "صنف", "اسم", "Detail", "Code"]): c_ext_mat = col
                if any(k in col_s for k in ["فعلي", "الفعلي", "الكمية", "Actual", "Qty"]): c_ext_qty = col

            ext_df["الصنف"] = ext_df[c_ext_mat].astype(str).str.strip()
            ext_df["الكمية الفعلية"] = pd.to_numeric(ext_df[c_ext_qty], errors="coerce").fillna(0)

            merged_comp = pd.merge(
                sys_df[["الصنف", "الكمية النظامية"]],
                ext_df[["الصنف", "الكمية الفعلية"]],
                on="الصنف",
                how="outer"
            ).fillna({"الكمية النظامية": 0, "الكمية الفعلية": 0})

            merged_comp["الفرق"] = merged_comp["الكمية الفعلية"] - merged_comp["الكمية النظامية"]
            merged_comp["حالة الصنف"] = merged_comp["الفرق"].apply(lambda x: "عجز (نقص)" if x < 0 else ("زيادة (فائض)" if x > 0 else "مطابق تماماً"))

            total_items = len(merged_comp)
            shortage_items = len(merged_comp[merged_comp["الفرق"] < 0])
            surplus_items = len(merged_comp[merged_comp["الفرق"] > 0])
            matched_items = len(merged_comp[merged_comp["الفرق"] == 0])

            st.success("تمت مطارحة ومقارنة الجرد الخارجي مع النظام بنجاح تام!")

            m_col1, m_col2, m_col3, m_col4 = st.columns(4)
            m_col1.metric("📦 إجمالي الأصناف بالمقارنة", f"{total_items:,}")
            m_col2.metric("🔥 أصناف بها عجز", f"{shortage_items:,}")
            m_col3.metric("➕ أصناف بها زيادة", f"{surplus_items:,}")
            m_col4.metric("✅ أصناف مطابقة", f"{matched_items:,}")

            st.divider()

            comp_tab1, comp_tab2, comp_tab3 = st.tabs(["📋 جدول المقارنة الشامل", "🔥 تقرير العجز فقط", "➕ تقرير الزيادة فقط"])

            with comp_tab1:
                st.subheader("جدول المقارنة الكامل بين النظام والجرد الخارجي")
                st.dataframe(merged_comp, use_container_width=True, hide_index=True)
                csv_all_comp = merged_comp.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 تحميل تقرير المقارنة الشامل (CSV)", csv_all_comp, "تقرير_مقارنة_الجرد_الشامل.csv", "text/csv")

            with comp_tab2:
                st.subheader("الأصناف التي سجلت عجزاً ونقصاً")
                df_short = merged_comp[merged_comp["الفرق"] < 0].sort_values("الفرق", ascending=True)
                st.dataframe(df_short, use_container_width=True, hide_index=True)
                if not df_short.empty:
                    csv_short = df_short.to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📥 تحميل تقرير العجز فقط (CSV)", csv_short, "تقرير_عجز_المخزن_الخارجي.csv", "text/csv")

            with comp_tab3:
                st.subheader("الأصناف التي سجلت زيادة وفائضاً")
                df_surp = merged_comp[merged_comp["الفرق"] > 0].sort_values("الفرق", ascending=False)
                st.dataframe(df_surp, use_container_width=True, hide_index=True)
                if not df_surp.empty:
                    csv_surp = df_surp.to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📥 تحميل تقرير الزيادة فقط (CSV)", csv_surp, "تقرير_زيادة_المخزن_الخارجي.csv", "text/csv")

        except Exception as e:
            st.error(f"حدث خطأ أثناء معالجة ومقارنة الملفين: {e}")
    else:
        st.info("👈 يرجى رفع ملف النظام الأساسي وملف الجرد الخارجي في هذا القسم لمقارنتهما واستخراج الفروقات الفورية.")
