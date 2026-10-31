import io
import pandas as pd
import streamlit as st
import datetime

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
st.caption("النظام المتكامل: معادلة الجرد اليومي المدمج (المصححة بدقة) + تقارير الأسبوع + الجرد السريع")

main_mode = st.sidebar.radio(
    "📂 اختر القسم المطلوب:",
    [
        "🧮 معادلة الجرد اليومي المدمج (الفصل والتصدير الفوري)",
        "✅ مهام ومجموعات جرد الأسبوع (حسب جدول الأيام)",
        "📊 التقرير التنفيذي (شهر 9 وشهر 10)", 
        "📅 تقرير الأسبوع (متابعة محمد قطان - تحليل الفاقد وأكثر 20 صنف عجز)",
        "📂 دمج ملفات الجرد اليومي وتصدير الإكسيل", 
        "🔍 تقرير الفروقات المستقل (شامل وفوري)",
        "📱 الجرد السريع والبحث الذكي بالموبايل"
    ]
)

def format_code_as_text(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    try:
        if "e" in val_str.lower():
            val_str = f"{float(val_str):.0f}"
    except:
        pass
    return val_str

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# ---------------------------------------------------------
# القسم الأول: معادلة الجرد اليومي المدمج (المصححة بدقة تامة)
# ---------------------------------------------------------
if main_mode == "🧮 معادلة الجرد اليومي المدمج (الفصل والتصدير الفوري)":
    st.subheader("🧮 معادلة الجرد اليومي المدمج (تصحيح حساب الفروقات والكميات الدفترية)")
    st.caption("ارفع جميع ملفات الجرد اليومي هنا. سيقوم النظام بدمجها بدقة تامة، تجميع الأصناف بدون تكرار، وحساب الزيادة والفقد بدقة مطلقة.")

    eq_files = st.file_uploader(
        "📁 ارفع ملفات الجرد اليومي دفعة واحدة:", 
        type=["xlsx", "xls", "csv"], 
        accept_multiple_files=True
    )

    if eq_files:
        try:
            eq_dfs = []
            for ef in eq_files:
                if ef.name.endswith(".csv"):
                    df_t = pd.read_csv(ef, encoding="utf-8-sig", dtype=str)
                else:
                    df_t = pd.read_excel(ef, dtype=str)
                eq_dfs.append(df_t)

            combined_eq = pd.concat(eq_dfs, ignore_index=True)

            # استخراج دقيق للأعمدة الأساسية من الملفات المرفوعة
            c_code = "Code" if "Code" in combined_eq.columns else combined_eq.columns[0]
            c_name = "Detail" if "Detail" in combined_eq.columns else combined_eq.columns[2]
            c_sys = "SystemQty" if "SystemQty" in combined_eq.columns else combined_eq.columns[3]
            c_act = "ActualQty" if "ActualQty" in combined_eq.columns else combined_eq.columns[4]

            for col in combined_eq.columns:
                cs = str(col).lower()
                if any(k in cs for k in ["code", "كود"]): c_code = col
                if any(k in cs for k in ["detail", "name", "صنف", "مادة"]): c_name = col
                if any(k in cs for k in ["systemqty", "دفترية", "نظام"]): c_sys = col
                if any(k in cs for k in ["actualqty", "فعلي"]): c_act = col

            combined_eq["كود الصنف"] = combined_eq[c_code].apply(format_code_as_text)
            combined_eq["اسم الصنف"] = combined_eq[c_name].astype(str).str.strip()
            combined_eq["الكمية الدفترية"] = pd.to_numeric(combined_eq[c_sys], errors="coerce").fillna(0)
            combined_eq["الكمية الفعلية"] = pd.to_numeric(combined_eq[c_act], errors="coerce").fillna(0)

            # تجميع البيانات لكل صنف بدون تكرار (جمع الدفترية والفعلية معا)
            eq_agg = combined_eq.groupby(["كود الصنف", "اسم الصنف"], as_index=False).agg({
                "الكمية الدفترية": "sum",
                "الكمية الفعلية": "sum"
            })

            # حساب الفرق الدقيق = الفعلية - الدفترية
            eq_agg["الفرق"] = eq_agg["الكمية الفعلية"] - eq_agg["الكمية الدفترية"]

            # فصل الزيادة والعجز بناء على الفرق
            eq_agg["كمية الزيادة"] = eq_agg["الفرق"].apply(lambda x: x if x > 0 else 0)
            eq_agg["كمية الفقد (العجز)"] = eq_agg["الفرق"].apply(lambda x: -x if x < 0 else 0)

            st.success(f"تم دمج وتحليل عدد {len(eq_files)} ملفات بنجاح تام وبدون أي أخطاء في الحسابات!")

            eq_tab1, eq_tab2, eq_tab3, eq_tab4, eq_tab5 = st.tabs([
                "📋 إجمالي الجرد المدمج", 
                "➕ أصناف الزيادة", 
                "📉 أصناف الفقد (العجز)", 
                "📌 رمز المادة + كمية الزيادة فقط", 
                "📌 رمز المادة + كمية الفقد فقط"
            ])

            with eq_tab1:
                st.subheader("📋 إجمالي الجرد المدمج (بدون تكرار)")
                st.dataframe(eq_agg, use_container_width=True, hide_index=True)
                excel_data = convert_df_to_excel(eq_agg)
                st.download_button("📥 تحميل إجمالي الجرد المدمج (Excel)", excel_data, "اجمالي_الجرد_المدمج.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with eq_tab2:
                st.subheader("➕ الأصناف التي بها زيادة")
                df_surplus = eq_agg[eq_agg["كمية الزيادة"] > 0].copy()
                st.dataframe(df_surplus, use_container_width=True, hide_index=True)
                if not df_surplus.empty:
                    excel_data = convert_df_to_excel(df_surplus)
                    st.download_button("📥 تحميل تقرير الزيادة (Excel)", excel_data, "اصناف_الزيادة.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with eq_tab3:
                st.subheader("📉 الأصناف التي بها فقد أو عجز")
                df_short = eq_agg[eq_agg["كمية الفقد (العجز)"] > 0].copy()
                st.dataframe(df_short, use_container_width=True, hide_index=True)
                if not df_short.empty:
                    excel_data = convert_df_to_excel(df_short)
                    st.download_button("📥 تحميل تقرير الفقد (Excel)", excel_data, "اصناف_الفقد.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with eq_tab4:
                st.subheader("📌 رمز المادة وكمية الزيادة فقط")
                df_code_surplus = eq_agg[eq_agg["كمية الزيادة"] > 0][["كود الصنف", "كمية الزيادة"]].copy()
                df_code_surplus.columns = ["رمز المادة", "كمية الزيادة"]
                st.dataframe(df_code_surplus, use_container_width=True, hide_index=True)
                if not df_code_surplus.empty:
                    excel_data = convert_df_to_excel(df_code_surplus)
                    st.download_button("📥 تحميل رمز المادة وكمية الزيادة (Excel)", excel_data, "رمز_المادة_وكمية_الزيادة.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with eq_tab5:
                st.subheader("📌 رمز المادة وكمية الفقد فقط")
                df_code_short = eq_agg[eq_agg["كمية الفقد (العجز)"] > 0][["كود الصنف", "كمية الفقد (العجز)"]].copy()
                df_code_short.columns = ["رمز المادة", "كمية الفقد"]
                st.dataframe(df_code_short, use_container_width=True, hide_index=True)
                if not df_code_short.empty:
                    excel_data = convert_df_to_excel(df_code_short)
                    st.download_button("📥 تحميل رمز المادة وكمية الفقد (Excel)", excel_data, "رمز_المادة_وكمية_الفقد.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"حدث خطأ أثناء معالجة معادلة الجرد: {e}")
    else:
        st.info("👈 ارفع ملفات الجرد اليومي في زرار الرفع أعلاه لتطبيق معادلة الجرد المدمج بدقة.")

# ---------------------------------------------------------
# باقي الأقسام
# ---------------------------------------------------------
elif main_mode == "✅ مهام ومجموعات جرد الأسبوع (حسب جدول الأيام)":
    st.subheader("✅ جدول ومجموعات جرد الأسبوع (متابعة محمد قطان)")
    today = datetime.date.today()
    current_week_number = today.isocalendar()[1]
    tasks_schedule = {
        "🧊 السبت: جرد الثلاجة (الجبن، اللانشون، الألبان، الخضار، الصوصات، زبدة، سمنة)": False,
        "🗓️ الأحد والأربعاء (الدفعة الأولى): مناديل، حفاضات، مواد التنظيف، غسيل الملابس والصحون، عناية شخصية": False,
        "🗓️ الإثنين: شيبس و سناكس، بسكويت وعلويات، نودلز، بقوليات سابية، إيس كريم، تمور، كاندي": False,
        "📦 الثلاثاء: جرد المخزن الشامل (بهارات، زيت وزيتون، دقيق وسكر، كورن فليكس، بن، حبوب)": False,
        "🗓️ الأربعاء والسبت (الدفعة الثانية): أخري استخدام مرة واحدة، بطاريات، منعمات الأقمشة، معطر جو، مبيدات حشرية": False,
        "🗓️ الخميس: المتبقيات والأقسام الإضافية (مشروم، ورق عنب، مخلل، صلصة، ضيافة العيد)": False,
    }
    if 'todo_week_num_v3' not in st.session_state or st.session_state['todo_week_num_v3'] != current_week_number:
        st.session_state['todo_week_num_v3'] = current_week_number
        st.session_state['tasks_state_v3'] = tasks_schedule

    completed_count = sum(1 for v in st.session_state['tasks_state_v3'].values() if v)
    total_tasks = len(st.session_state['tasks_state_v3'])
    for task_name in list(st.session_state['tasks_state_v3'].keys()):
        is_checked = st.checkbox(task_name, value=st.session_state['tasks_state_v3'][task_name], key=f"task_v3_{task_name}")
        st.session_state['tasks_state_v3'][task_name] = is_checked

    progress_pct = (completed_count / total_tasks) * 100 if total_tasks > 0 else 0
    st.progress(progress_pct / 100)
    st.write(f"نسبة الإنجاز: **{progress_pct:.1f}%**")

elif main_mode == "📊 التقرير التنفيذي (شهر 9 وشهر 10)":
    st.subheader("📊 لوحة تحكم ومؤشرات أداء الجرد (سبتمبر وأكتوبر)")
    st.info("قسم التقارير التنفيذية")

elif main_mode == "📅 تقرير الأسبوع (متابعة محمد قطان - تحليل الفاقد وأكثر 20 صنف عجز)":
    st.subheader("📅 تقرير الأسبوع — متابعة محمد قطان")
    st.info("قسم تقارير الأسبوع وتحليل الفاقد")

elif main_mode == "📂 دمج ملفات الجرد اليومي وتصدير الإكسيل":
    st.subheader("📂 مركز دمج ملفات الجرد اليومي")
    st.info("قسم دمج الملفات")

elif main_mode == "🔍 تقرير الفروقات المستقل (شامل وفوري)":
    st.subheader("🔍 تقرير الفروقات المستقل")
    st.info("قسم التقرير المستقل")

else:
    st.subheader("📱 الجرد السريع والبحث الذكي بالموبايل")
    st.info("قسم الجرد بالموبايل")
