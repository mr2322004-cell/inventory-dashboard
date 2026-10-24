import streamlit as st
import pandas as pd
import plotly.express as px

# ضبط إعدادات الصفحة
st.set_page_config(
    page_title="لوحة تحكم تقارير الجرد والتالف",
    page_icon="📊",
    layout="wide"
)

# تخصيص واجهة المستخدم باللغة العربية (RTL)
st.markdown("""
    <style>
    .main { text-align: right; direction: rtl; }
    h1, h2, h3, h4 { text-align: right; }
    div[data-testid="stMetricValue"] { text-align: right; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 لوحة تحكم تقارير الجرد والفروقات (شهر 9 وشهر 10)")
st.write("مرحباً بك! قم برفع ملف الإكسيل المجمع لاستعراض الفروقات والأصناف المفقودة فوراً.")

# رفع الملف
uploaded_file = st.file_uploader("📁 ارفع ملف الإكسيل المجمع هنا", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    # قراءة البيانات
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    # تنظيف البيانات وتقسيم الشهور
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df['Month'] = df['Date'].dt.month
    
    # قائمة الفلترة بالشهور في الشريط الجانبي
    st.sidebar.header("🎯 فلاتر العرض")
    selected_month = st.sidebar.selectbox("اختر الشهر:", ["الكل", "شهر 9", "شهر 10"])
    
    if selected_month == "شهر 9":
        filtered_df = df[df['Month'] == 9]
    elif selected_month == "شهر 10":
        filtered_df = df[df['Month'] == 10]
    else:
        filtered_df = df.copy()

    # --- حساب مؤشرات الأداء الرئيسية (KPIs) ---
    total_system = filtered_df['SystemQty'].sum()
    total_actual = filtered_df['ActualQty'].sum()
    total_diff = filtered_df['Difference'].sum()
    
    # العجز (الفقد بالسالب)
    total_loss = filtered_df[filtered_df['Difference'] < 0]['Difference'].sum()
    
    # نسب الأداء
    loss_percentage = (abs(total_loss) / total_system * 100) if total_system > 0 else 0
    incorrect_counts = len(filtered_df[filtered_df['Difference'] != 0])
    total_counts = len(filtered_df)
    error_rate = (incorrect_counts / total_counts * 100) if total_counts > 0 else 0

    # عرض المؤشرات العليا
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("إجمالي الدفتري", f"{total_system:,.0f}")
    col2.metric("إجمالي الفعلي", f"{total_actual:,.0f}")
    col3.metric("نسبة الفقد الإجمالية", f"{loss_percentage:.2f}%")
    col4.metric("نسبة الجرد غير المضبوط", f"{error_rate:.1f}%")

    st.markdown("---")

    # --- أزرار الفلترة السريعة بنقرة واحدة ---
    st.subheader("🔘 الفلترة السريعة حسب حالة الجرد")
    filter_status = st.radio(
        "اختر العرض المطلوب:",
        ["🔴 الأصناف التي بها عجز (مفقودة)", "🟢 الأصناف المطابقة", "🔵 الأصناف التي بها زيادة", "📋 كل البيانات"],
        horizontal=True
    )

    if "عجز" in filter_status:
        display_df = filtered_df[filtered_df['Difference'] < 0]
    elif "المطابقة" in filter_status:
        display_df = filtered_df[filtered_df['Difference'] == 0]
    elif "زيادة" in filter_status:
        display_df = filtered_df[filtered_df['Difference'] > 0]
    else:
        display_df = filtered_df.copy()

    # --- رسم بياني لأعلى الأصناف فقداً ---
    if not display_df.empty and "عجز" in filter_status:
        st.subheader("📌 أعلى 10 أصناف بها عجز (تستوجب إعادة الطلب)")
        top_losses = display_df.groupby('Detail')['Difference'].sum().reset_index()
        top_losses['Difference'] = top_losses['Difference'].abs()
        top_losses = top_losses.sort_values(by='Difference', ascending=False).head(10)
        
        fig = px.bar(top_losses, x='Detail', y='Difference', text='Difference',
                     title="أعلى 10 أصناف مفقودة (بالكمية)",
                     labels={'Detail': 'اسم الصنف', 'Difference': 'كمية العجز'},
                     color_discrete_sequence=['#E74C3C'])
        st.plotly_chart(fig, use_container_width=True)

    # --- جدول البيانات التفصيلي ---
    st.subheader("📑 جدول التفاصيل")
    st.dataframe(
        display_df[['Code', 'Barcode', 'Detail', 'SystemQty', 'ActualQty', 'Difference', 'Date']],
        use_container_width=True
    )

    # --- زر تحميل التقرير لطلبيات الشراء ---
    csv_data = display_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="📥 تحميل التقرير الحالي كملف Excel/CSV لطلبات الشراء",
        data=csv_data,
        file_name="Inventory_Discrepancies_Report.csv",
        mime="text/csv"
    )
else:
    st.info("👈 يرجى رفع ملف الإكسيل المجمع من الزر بالأعلى للبدء في تحليل البيانات.")
