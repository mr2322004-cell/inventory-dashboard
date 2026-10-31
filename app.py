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
st.caption("النظام المتكامل: تحويل الأكواد لنص كامل 100% + معادلة الجرد اليومي المدمج + تقارير الشهور + تحليل الفاقد + الجرد السريع")

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

# دالة ذكية لتحويل الأكواد لنص صريح دون تحويلها لرموز علمية E+12
def format_code_as_text(val):
    if pd.isna(val):
        return ""
    # تحويل لسطر نصي مع إزالة البوينت صفر لو موجودة من الإكسيل
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    # لو كان رقم علمي بصيغة e/E تحويله لرقم صحيح نصي
    try:
        if "e" in val_str.lower():
            val_str = f"{float(val_str):.0f}"
    except:
        pass
    return val_str

# دالة لتصدير الإكسيل مع الحفاظ على الأكواد كنصوص صريحة 100%
def to_excel_with_text_codes(df, code_cols=["كود الصنف", "رمز المادة"]):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        text_format = workbook.add_format({'num_format': '@'})
        
        for col_idx, col_name in enumerate(df.columns):
            if col_name in code_cols:
                worksheet.set_column(col_idx, col_idx, 20, text_format)
                for row_idx, value in enumerate(df[col_name]):
                    worksheet.write_string(row_idx + 1, col_idx, str(value), text_format)
    return output.getvalue()

# ---------------------------------------------------------
# القسم الأول: معادلة الجرد اليومي المدمج (معالجة الأكواد كنص)
# ---------------------------------------------------------
if main_mode == "🧮 معادلة الجرد اليومي المدمج (الفصل والتصدير الفوري)":
    st.subheader("🧮 معادلة الجرد اليومي المدمج (تصدير كود المادة كـ نص كامل)")
    st.caption("ارفع جميع ملفات الجرد اليومي. يتم دمجها وتجميع الكميات، وتصدير الأكواد كنص صريح يفتح في الإكسيل بدون Scientific Notation (E+12).")

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
                    df_temp = pd.read_csv(ef, encoding="utf-8-sig", dtype=str)
                else:
                    df_temp = pd.read_excel(ef, dtype=str)
                eq_dfs.append(df_temp)

            combined_eq = pd.concat(eq_dfs, ignore_index=True)

            c_code = "Code" if "Code" in combined_eq.columns else combined_eq.columns[0]
            c_name = "Detail" if "Detail" in combined_eq.columns else combined_eq.columns[2]
            c_sys = "SystemQty" if "SystemQty" in combined_eq.columns else combined_eq.columns[3]
            c_act = "ActualQty" if "ActualQty" in combined_eq.columns else combined_eq.columns[4]
            c_diff = "Difference" if "Difference" in combined_eq.columns else combined_eq.columns[5]

            for col in combined_eq.columns:
                cs = str(col).lower()
                if any(k in cs for k in ["code", "كود"]): c_code = col
                if any(k in cs for k in ["detail", "name", "صنف", "مادة"]): c_name = col
                if any(k in cs for k in ["systemqty", "نظام"]): c_sys = col
                if any(k in cs for k in ["actualqty", "فعلي"]): c_act = col
                if any(k in cs for k in ["difference", "فرق"]): c_diff = col

            combined_eq["كود الصنف"] = combined_eq[c_code].apply(format_code_as_text)
            combined_eq["اسم الصنف"] = combined_eq[c_name].astype(str).str.strip()
            combined_eq["الكمية الدفترية"] = pd.to_numeric(combined_eq[c_sys], errors="coerce").fillna(0)
            combined_eq["الكمية الفعلية"] = pd.to_numeric(combined_eq[c_act], errors="coerce").fillna(0)
            combined_eq["الفرق"] = pd.to_numeric(combined_eq[c_diff], errors="coerce").fillna(0)

            eq_agg = combined_eq.groupby(["كود الصنف", "اسم الصنف"], as_index=False).agg({
                "الكمية الدفترية": "sum",
                "الكمية الفعلية": "sum",
                "الفرق": "sum"
            })

            eq_agg["كمية الزيادة"] = eq_agg.apply(lambda r: (r["الكمية الفعلية"] - r["الكمية الدفترية"]) if r["الكمية الفعلية"] > r["الكمية الدفترية"] else 0, axis=1)
            eq_agg["كمية الفقد (العجز)"] = eq_agg.apply(lambda r: (r["الكمية الدفترية"] - r["الكمية الفعلية"]) if r["الكمية الدفترية"] > r["الكمية الفعلية"] else 0, axis=1)

            st.success(f"تم دمج عدد {len(eq_files)} ملفات بنجاح وحفظ رموز المواد كنص كامل!")

            eq_tab1, eq_tab2, eq_tab3, eq_tab4, eq_tab5 = st.tabs([
                "📋 إجمالي الجرد المدمج", 
                "➕ أصناف الزيادة", 
                "📉 أصناف الفقد (العجز)", 
                "📌 رمز المادة + كمية الزيادة فقط", 
                "📌 رمز المادة + كمية الفقد فقط"
            ])

            with eq_tab1:
                st.subheader("📋 إجمالي الجرد المدمج")
                st.dataframe(eq_agg, use_container_width=True, hide_index=True)
                excel_bytes = to_excel_with_text_codes(eq_agg, code_cols=["كود الصنف"])
                st.download_button("📥 تحميل إجمالي الجرد المدمج (Excel نصي)", excel_bytes, "اجمالي_الجرد_المدمج.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with eq_tab2:
                st.subheader("➕ الأصناف التي بها زيادة")
                df_surplus = eq_agg[eq_agg["كمية الزيادة"] > 0].copy()
                st.dataframe(df_surplus, use_container_width=True, hide_index=True)
                if not df_surplus.empty:
                    excel_surp = to_excel_with_text_codes(df_surplus, code_cols=["كود الصنف"])
                    st.download_button("📥 تحميل تقرير الزيادة (Excel نصي)", excel_surp, "اصناف_الزيادة.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with eq_tab3:
                st.subheader("📉 الأصناف التي بها فقد أو عجز")
                df_short = eq_agg[eq_agg["كمية الفقد (العجز)"] > 0].copy()
                st.dataframe(df_short, use_container_width=True, hide_index=True)
                if not df_short.empty:
                    excel_short = to_excel_with_text_codes(df_short, code_cols=["كود الصنف"])
                    st.download_button("📥 تحميل تقرير الفقد (Excel نصي)", excel_short, "اصناف_الفقد.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with eq_tab4:
                st.subheader("📌 رمز المادة وكمية الزيادة فقط")
                df_code_surplus = eq_agg[eq_agg["كمية الزيادة"] > 0][["كود الصنف", "كمية الزيادة"]].copy()
                df_code_surplus.columns = ["رمز المادة", "كمية الزيادة"]
                st.dataframe(df_code_surplus, use_container_width=True, hide_index=True)
                if not df_code_surplus.empty:
                    excel_cs = to_excel_with_text_codes(df_code_surplus, code_cols=["رمز المادة"])
                    st.download_button("📥 تحميل رمز المادة وكمية الزيادة (Excel نصي)", excel_cs, "رمز_المادة_وكمية_الزيادة.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with eq_tab5:
                st.subheader("📌 رمز المادة وكمية الفقد فقط")
                df_code_short = eq_agg[eq_agg["كمية الفقد (العجز)"] > 0][["كود الصنف", "كمية الفقد (العجز)"]].copy()
                df_code_short.columns = ["رمز المادة", "كمية الفقد"]
                st.dataframe(df_code_short, use_container_width=True, hide_index=True)
                if not df_code_short.empty:
                    excel_csh = to_excel_with_text_codes(df_code_short, code_cols=["رمز المادة"])
                    st.download_button("📥 تحميل رمز المادة وكمية الفقد (Excel نصي)", excel_csh, "رمز_المادة_وكمية_الفقد.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"حدث خطأ أثناء معالجة معادلة الجرد: {e}")
    else:
        st.info("👈 ارفع ملفات الجرد اليومي في زرار الرفع أعلاه لتطبيق معادلة الجرد المدمج وفصل النتائج.")

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
