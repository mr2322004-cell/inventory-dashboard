import io
import pandas as pd
import streamlit as st
import datetime
import difflib

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
st.caption("النظام المبسط: جرد أصناف الطلب (بالصور، كمية النظام، وخانة الكمية الفعلية) + معادلة الجرد اليومي + تقرير الأسبوع")

main_mode = st.sidebar.radio(
    "📂 الأقسام الرئيسية:",
    [
        "🛒 جرد الطلبات بالصور والكميات",
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
    except:
        pass
    return val_str

def convert_df_to_excel(df_to_export):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_export.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# قائمة أصناف الطلب المستخرجة من ملفات الـ PDF
pdf_order_items = [
    "تيفاني بريك 2 صباع 12.5 جم", "شيبسي كباب عالفحم ب 5", "شيبسي خل وملح ب 10", "شيبسي جبنة ب 10",
    "شيبسي ملح ب 10", "شيبسي جبنة البارميزان 10 ج", "شيبسي شطة وليمون ب 10", "شيبسي بطعم جبنة متبلة ب 10",
    "فن داي كيك بنكهة الفانيليا وقطع الفواكة", "لارش كيك 10 ج", "رونیسكا بسكويت بالشوكولا ب 3 ج",
    "تيفاني بريك داي بريك شوكولاه بالبندق 16 جم", "ريتش بيك توست بالحليب غ500", "ريتش بيك توست سادة 500غ",
    "تيفاني بريك 4 اصابع 25 جم", "تيفاني بريكيستا 13 جم", "شيبسي كرانشي فراخ ب 10 ج",
    "شيبسي ويفي فراخ بالصوص الحار ب 15 ج", "بسكويت سادة جامبو", "مندولين ويفر بالشوكولا جم 33 ب 10 ج",
    "جيلي عيون ب 5 ج", "لارش كيك بالكراميل 10 ج", "كيك لارش مولتن كيك", "ديمة سوفتي بالتمر 21 غ",
    "شوكولا ليرات ع10", "من و سلوی رش", "بسكو مصر ويفر ب 5 ج", "زيرو بروتين بار ۵۰جم بطعم القهوة",
    "بيج شيبسي فلفل حار مدخن ب 20 ج", "بیج شيبسي بالكريمة المتبلة ب 10 ج", "تيفاني بريك داي بريك شوكولاه 25 جم",
    "ماجیکال مارشميلو 5 ج", "بيبيتو جيلي ب 5 ج", "اصابع هريسة اللوز", "ابوالوفا كرواسون", "بسكوت روكستا 2 ق",
    "شيبسي وافي رانش جبنة مع لسعة ليمون 10 ج", "تيفاني شنكوز كوكيز شوكولا", "دومتي باتيه ب 15 ج بطعم جبنة كريمي سادة",
    "دومتي باتيه ب 15 ج بطعم فيتا بالزيتون", "دومتي باتيه ب 15 ج بطعم كريمي رومي", "شيبسي جبنة متبلة ب 7 ج",
    "بريدواي باتيه بالشوكولا", "بريدواي باتيه جبنة هلابينو", "شيبسي فراخ مع الصوص الحار ب 10 ج",
    "توست توستيلا سادة 500 جم", "بيج شيبسي بالليمون المولع ب 15 ج", "دومتي كرواسون ب 15 ج بطعم بالفلمنك",
    "دومتي باتيه ب ١٥ ج بطعم مربى التين"
]

# ---------------------------------------------------------
# 1. قسم جرد الطلبات بالصور والكميات
# ---------------------------------------------------------
if main_mode == "🛒 جرد الطلبات بالصور والكميات":
    st.subheader("🛒 جرد أصناف الطلب (عرض الصور، كمية النظام، وإدخال الكمية الفعلية)")
    st.caption("ارفع ملف `كميات.xlsx` ليسحب كمية النظام لكل صنف تلقائياً، مع عرض صورة المنتج وخانة فارغة للكمية الفعلية.")

    qty_file = st.file_uploader("📁 ارفع ملف كميات المخزن (كميات.xlsx):", type=["xlsx", "xls"])

    df_qty = None
    if qty_file is not None:
        try:
            df_qty = pd.read_excel(qty_file)
            st.success("تم قراءة ملف الكميات بنجاح!")
        except Exception as e:
            st.error(f"خطأ في قراءة الملف: {e}")

    if df_qty is not None:
        st.markdown("---")
        inventory_data = []
        
        for i, item in enumerate(pdf_order_items):
            system_qty = 0.0
            matched_name = "غير مطابق"
            for _, row in df_qty.iterrows():
                excel_item = str(row['الصنف'])
                if item in excel_item or excel_item in item:
                    system_qty = float(row['الكمية الحالية'])
                    matched_name = excel_item
                    break
            if matched_name == "غير مطابق":
                matches = difflib.get_close_matches(item, df_qty['الصنف'].astype(str).tolist(), n=1, cutoff=0.3)
                if matches:
                    matched_row = df_qty[df_qty['الصنف'].astype(str) == matches[0]]
                    if not matched_row.empty:
                        system_qty = float(matched_row.iloc[0]['الكمية الحالية'])
                        matched_name = matches[0]

            # تصميم بطاقة المنتج مع صورة توضيحية واسم الصنف
            with st.container():
                st.markdown(f"""
                <div style="background: #151f32; padding: 12px; border-radius: 10px; border: 1px solid #29364b; margin-bottom: 10px;">
                    <h4 style="margin: 0; color: #38bdf8;">📦 {i+1}. {item}</h4>
                    <p style="margin: 2px 0 0 0; font-size: 13px; color: #94a3b8;">مطابق في النظام: {matched_name}</p>
                </div>
                """, unsafe_allow_html=True)

                col_img, col_sys, col_act, col_diff = st.columns([1, 2, 2, 2])
                with col_img:
                    # عرض صورة مصغرة توضيحية لكل صنف
                    st.image("https://images.unsplash.com/photo-1586816001966-79b736744398?w=100&auto=format&fit=crop&q=60", width=70)
                with col_sys:
                    st.metric("كمية النظام", f"{system_qty}")
                with col_act:
                    actual_qty = st.number_input("الكمية الفعلية", min_value=0.0, step=1.0, key=f"act_{i}", label_visibility="collapsed")
                with col_diff:
                    diff = actual_qty - system_qty
                    st.metric("الفرق", f"{diff}", delta=f"{diff}")

            inventory_data.append({
                "م": i+1,
                "اسم الصنف": item,
                "الصنف المطابق بالملف": matched_name,
                "كمية النظام": system_qty,
                "الكمية الفعلية": actual_qty,
                "الفرق": diff
            })
            st.markdown("---")

        if st.button("💾 حفظ وتصدير تقرير الجرد بالصور (Excel)"):
            df_result = pd.DataFrame(inventory_data)
            excel_bytes = convert_df_to_excel(df_result)
            st.download_button(
                "📥 تحميل تقرير الجرد (Excel)",
                excel_bytes,
                "تقرير_جرد_الطلبات_بالصور.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("👈 يرجى رفع ملف `كميات.xlsx` أعلاه لعرض المنتجات بالصور وربط كميات النظام تلقائياً.")

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
                if any(k in cs for k in ["code", "كود"]): c_code = col
                if any(k in cs for k in ["detail", "name", "صنف", "مادة"]): c_name = col
                if any(k in cs for k in ["systemqty", "دفترية", "نظام"]): c_sys = col
                if any(k in cs for k in ["actualqty", "فعلي"]): c_act = col

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
