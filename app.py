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
.product-card {
    background: #151f32;
    padding: 15px;
    border-radius: 12px;
    border: 1px solid #29364b;
    margin-bottom: 15px;
}
</style>
""", unsafe_allow_html=True)

st.title("📦 متابعة محمد قطان - نظام إدارة المخازن والجرد الذكي")
st.caption("النظام المبسط: جرد أصناف المخزن مباشرة + معادلة الجرد اليومي + تقرير الأسبوع")

main_mode = st.sidebar.radio(
    "📂 الأقسام الرئيسية:",
    [
        "🛒 جرد الأصناف والكميات",
        "🧮 معادلة الجرد اليومي المدمج",
        "✅ تقرير ومجموعات جرد الأسبوع"
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
    except Exception:
        pass
    return val_str

def convert_df_to_excel(df_to_export):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_export.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# ---------------------------------------------------------
# 1. قسم جرد الأصناف والكميات (مباشرة من ملف كميات.xlsx بدون PDF)
# ---------------------------------------------------------
if main_mode == "🛒 جرد الأصناف والكميات":
    st.subheader("🛒 جرد أصناف المخزن (عرض كميات النظام وإدخال الكمية الفعلية)")
    st.caption("ارفع ملف `كميات.xlsx` لعرض الأصناف وأكوادها وكميات النظام مباشرة مع خانة لإدخال الكمية الفعلية.")

    qty_file = st.file_uploader("📁 ارفع ملف كميات المخزن (كميات.xlsx):", type=["xlsx", "xls"])

    df_qty = None
    if qty_file is not None:
        try:
            df_qty = pd.read_excel(qty_file)
            df_qty.columns = [str(c).strip() for c in df_qty.columns]
            st.success("تم قراءة ملف الكميات بنجاح!")
        except Exception as e:
            st.error(f"خطأ في قراءة الملف: {e}")

    if df_qty is not None:
        # اكتشاف أعمدة الكود، الصنف، والكمية تلقائياً
        code_col = None
        item_col = None
        qty_col = None

        for col in df_qty.columns:
            c_lower = col.lower()
            if any(k in c_lower for k in ["code", "كود"]):
                code_col = col
            if any(k in c_lower for k in ["صنف", "اسم", "مادة", "detail", "name", "item"]):
                if item_col is None:
                    item_col = col
            if any(k in c_lower for k in ["كمية", "كميه", "نظام", "حالية", "متاحة", "qty", "quantity", "stock"]):
                if qty_col is None:
                    qty_col = col

        # بدائل احتياطية إذا لم تتطابق الأسماء
        if item_col is None:
            item_col = df_qty.columns[1] if len(df_qty.columns) > 1 else df_qty.columns[0]
        if qty_col is None:
            qty_col = df_qty.columns[-1]

        df_qty = df_qty.dropna(subset=[item_col]).copy()
        
        st.markdown("---")
        inventory_data = []

        for i, row in df_qty.iterrows():
            item_name = str(row[item_col]).strip() if pd.notna(row[item_col]) else ""
            item_code = str(row[code_col]).strip() if (code_col and pd.notna(row[code_col])) else f"PRD-{i+1}"
            system_qty = float(row[qty_col]) if pd.notna(row[qty_col]) else 0.0

            with st.container():
                st.markdown(f"""
                <div style="background: #151f32; padding: 12px; border-radius: 10px; border: 1px solid #29364b; margin-bottom: 10px;">
                    <h4 style="margin: 0; color: #38bdf8;">📦 الكود: {item_code} | {item_name}</h4>
                </div>
                """, unsafe_allow_html=True)

                col_sys, col_act = st.columns([1, 1])
                with col_sys:
                    st.metric("كمية النظام", f"{system_qty}")
                with col_act:
                    actual_qty = st.number_input("الكمية الفعلية", min_value=0.0, step=1.0, key=f"act_{i}")

            inventory_data.append({
                "م": i + 1,
                "كود الصنف": item_code,
                "اسم الصنف": item_name,
                "كمية النظام": system_qty,
                "الكمية الفعلية": actual_qty
            })
            st.markdown("---")

        if st.button("💾 حفظ وتصدير تقرير الجرد (Excel)"):
            df_result = pd.DataFrame(inventory_data)
            excel_bytes = convert_df_to_excel(df_result)
            st.download_button(
                "📥 تحميل تقرير الجرد (Excel)",
                excel_bytes,
                "تقرير_جرد_المخزن.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("👈 يرجى رفع ملف `كميات.xlsx` أعلاه لعرض أصناف المخزن مباشرة.")

# ---------------------------------------------------------
# 2. معادلة الجرد اليومي المدمج
# ---------------------------------------------------------
elif main_mode == "🧮 معادلة الجرد اليومي المدمج":
    st.subheader("🧮 معادلة الجرد اليومي المدمج (رمز الصنف، اسم الصنف، والكمية)")
    st.caption("ارفع أولاً ملف (رمز المادة الأساسي)، ثم ارفع ملفات الجرد اليومي لاستخراج التقارير بالصيغة المطلوبة.")

    master_file = st.file_uploader("📁 1. ارفع الملف الأساسي (رمز المادة.xlsx):", type=["xlsx", "xls"])

    eq_files = st.file_uploader(
        "📁 2. ارفع ملفات الجرد اليومي دفعة واحدة:",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True
    )

    master_name_to_code = {}
    if master_file is not None:
        try:
            df_m = pd.read_excel(master_file)
            for val in df_m.iloc[:, 0].dropna():
                val_str = str(val).strip()
                if '-' in val_str:
                    p = val_str.split('-', 1)
                    code_val = format_code_as_text(p[0])
                    name_val = p[1].strip().lower()
                    master_name_to_code[name_val] = code_val
            st.success("تم قراءة الملف الأساسي لرموز المواد بنجاح!")
        except Exception as e:
            st.error(f"خطأ في قراءة الملف الأساسي: {e}")

    if eq_files and master_file:
        try:
            eq_dfs = []
            for ef in eq_files:
                if ef.name.endswith(".csv"):
                    df_t = pd.read_csv(ef, encoding="utf-8-sig", dtype=str)
                else:
                    df_t = pd.read_excel(ef, dtype=str)
                eq_dfs.append(df_t)

            combined_eq = pd.concat(eq_dfs, ignore_index=True)

            c_code = "Code" if "Code" in combined_eq.columns else combined_eq.columns[0]
            c_name = "Detail" if "Detail" in combined_eq.columns else combined_eq.columns[2]
            c_sys = "SystemQty" if "SystemQty" in combined_eq.columns else combined_eq.columns[3]
            c_act = "ActualQty" if "ActualQty" in combined_eq.columns else combined_eq.columns[4]

            for col in combined_eq.columns:
                cs = str(col).lower()
                if any(k in cs for k in ["code", "كود"]):
                    c_code = col
                if any(k in cs for k in ["detail", "name", "صنف", "مادة"]):
                    c_name = col
                if any(k in cs for k in ["systemqty", "دفترية", "نظام"]):
                    c_sys = col
                if any(k in cs for k in ["actualqty", "فعلي"]):
                    c_act = col

            def get_official_code(row):
                item_name = str(row[c_name]).strip().lower()
                for m_name, m_code in master_name_to_code.items():
                    if m_name in item_name or item_name in m_name:
                        return m_code
                return format_code_as_text(row[c_code])

            combined_eq["رمز المادة الرسمي"] = combined_eq.apply(get_official_code, axis=1)
            combined_eq["اسم الصنف"] = combined_eq[c_name].astype(str).str.strip()
            combined_eq["الكمية الدفترية"] = pd.to_numeric(combined_eq[c_sys], errors="coerce").fillna(0)
            combined_eq["الكمية الفعلية"] = pd.to_numeric(combined_eq[c_act], errors="coerce").fillna(0)

            eq_agg = combined_eq.groupby(["رمز المادة الرسمي", "اسم الصنف"], as_index=False).agg({
                "الكمية الدفترية": "first",
                "الكمية الفعلية": "sum"
            })

            eq_agg["الفرق"] = eq_agg["الكمية الفعلية"] - eq_agg["الكمية الدفترية"]
            eq_agg["كمية الزيادة"] = eq_agg["الفرق"].apply(lambda x: x if x > 0 else 0)
            eq_agg["كمية الفقد (العجز)"] = eq_agg["الفرق"].apply(lambda x: -x if x < 0 else 0)

            st.success(f"تم دمج وتحليل عدد {len(eq_files)} ملفات بنجاح واستخراج رموز المواد الرسمية!")

            eq_tab1, eq_tab2, eq_tab3, eq_tab4, eq_tab5 = st.tabs([
                "📋 إجمالي الجرد المدمج",
                "➕ أصناف الزيادة",
                "📉 أصناف الفقد (العجز)",
                "📌 رمز الصنف + اسم الصنف + كمية الزيادة",
                "📌 رمز الصنف + اسم الصنف + كمية الفقد"
            ])

            with eq_tab1:
                st.subheader("📋 إجمالي الجرد المدمج")
                st.dataframe(eq_agg, use_container_width=True, hide_index=True)
                excel_data = convert_df_to_excel(eq_agg)
                st.download_button("📥 تحميل إجمالي الجرد (Excel)", excel_data, "اجمالي_الجرد.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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
                st.subheader("📌 رمز الصنف + اسم الصنف + كمية الزيادة")
                df_code_surplus = eq_agg[eq_agg["كمية الزيادة"] > 0][["رمز المادة الرسمي", "اسم الصنف", "كمية الزيادة"]].copy()
                st.dataframe(df_code_surplus, use_container_width=True, hide_index=True)
                if not df_code_surplus.empty:
                    excel_data = convert_df_to_excel(df_code_surplus)
                    st.download_button("📥 تحميل تقرير الزيادة (رمز واسم وكمية)", excel_data, "رمز_واسم_وكمية_الزيادة.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with eq_tab5:
                st.subheader("📌 رمز الصنف + اسم الصنف + كمية الفقد (العجز)")
                df_code_short = eq_agg[eq_agg["كمية الفقد (العجز)"] > 0][["رمز المادة الرسمي", "اسم الصنف", "كمية الفقد (العجز)"]].copy()
                st.dataframe(df_code_short, use_container_width=True, hide_index=True)
                if not df_code_short.empty:
                    excel_data = convert_df_to_excel(df_code_short)
                    st.download_button("📥 تحميل تقرير الفقد (رمز واسم وكمية)", excel_data, "رمز_واسم_وكمية_الفقد.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"حدث خطأ أثناء معالجة معادلة الجرد: {e}")
    else:
        st.info("👈 يرجى رفع الملف الأساسي (رمز المادة) أولاً، ثم رفع ملفات الجرد اليومي لتظهر النتائج فوراً.")

# ---------------------------------------------------------
# 3. تقرير ومجموعات جرد الأسبوع
# ---------------------------------------------------------
else:
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
