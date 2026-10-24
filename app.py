import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

# إعداد الصفحة لتكون واسعة وبتنسيق احترافي
st.set_page_config(
    page_title="لوحة تحكم تقارير الجرد المفصلة", page_icon="🚀", layout="wide"
)

# تصميم العنوان بشكل فخم
st.markdown(
    """
    <div style='background-color: #1E1E2F; padding: 20px; border-radius: 10px; text-align: center;'>
        <h1 style='color: #00FFA3; margin: 0;'>🚀 لوحة تحكم تقارير الجرد والفروقات (التحليل المفصل)</h1>
        <p style='color: #FFFFFF; margin-top: 10px;'>تحليل شامل ودقيق لأرصدة شهر 9 وشهر 10</p>
    </div>
""",
    unsafe_allow_html=True,
)

st.write("")

# رفع الملف
uploaded_file = st.file_uploader(
    "📁 ارفع ملف الإكسيل المجمع هنا", type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:
  try:
    if uploaded_file.name.endswith(".csv"):
      df = pd.read_csv(uploaded_file)
    else:
      excel_file = pd.ExcelFile(uploaded_file)
      df = pd.read_excel(excel_file, sheet_name=excel_file.sheet_names[0])

    # معالجة التاريخ إن وجد
    for col in df.columns:
      if "date" in str(col).lower() or "تاريخ" in str(col):
        df["Date"] = pd.to_datetime(df[col], errors="coerce")
        break

    if "Difference" in df.columns:
      total_items = len(df)
      total_diff = df["Difference"].sum()
      shortage_df = df[df["Difference"] < 0]
      excess_df = df[df["Difference"] > 0]
      exact_df = df[df["Difference"] == 0]

      # مربعات إحصائية فخمة
      st.markdown("---")
      col1, col2, col3, col4 = st.columns(4)
      col1.metric("📦 إجمالي الأصناف", f"{total_items:,}")
      col2.metric("📉 إجمالي الفروقات", f"{total_diff:,.2f}")
      col3.metric("⚠️ أصناف بها عجز", f"{len(shortage_df):,}")
      col4.metric("➕ أصناف بها زيادة", f"{len(excess_df):,}")
      st.markdown("---")

      # تبويبات مفصلة بالكامل
      tab1, tab2, tab3, tab4 = st.tabs([
          "📋 كل البيانات",
          "📉 تفاصيل العجز (الواضح)",
          "➕ تفاصيل الزيادة",
          "🔍 بحث وفلترة متقدمة",
      ])

      with tab1:
        st.subheader("عرض شامل لكافة سجلات الجرد")
        st.dataframe(df, use_container_width=True)

      with tab2:
        st.subheader("⚠️ الأصناف التي بها عجز (نقص في المخزون)")
        st.metric("عدد أصناف العجز", len(shortage_df))
        st.dataframe(shortage_df, use_container_width=True)
        # زر تحميل العجز
        st.download_button(
            "📥 تحميل تقرير العجز فقط (CSV)",
            shortage_df.to_csv(index=False).encode("utf-8"),
            file_name="shortage_report.csv",
            mime="text/csv",
        )

      with tab3:
        st.subheader("➕ الأصناف التي بها زيادة في المخزون")
        st.metric("عدد أصناف الزيادة", len(excess_df))
        st.dataframe(excess_df, use_container_width=True)
        # زر تحميل الزيادة
        st.download_button(
            "📥 تحميل تقرير الزيادة فقط (CSV)",
            excess_df.to_csv(index=False).encode("utf-8"),
            file_name="excess_report.csv",
            mime="text/csv",
        )

      with tab4:
        st.subheader("🔍 البحث المخصص حسب الصنف أو الكود")
        search_query = st.text_input("أدخل اسم الصنف أو الباركود أو الكود:")
        if search_query:
          filtered_search = df[
              df.astype(str)
              .apply(
                  lambda r: r.str.contains(search_query, case=False).any(),
                  axis=1,
              )
          ]
          st.write(f"نتائج البحث المطابقة: {len(filtered_search)} صنف")
          st.dataframe(filtered_search, use_container_width=True)
        else:
          st.info("اكتب أي كلمة للبحث الفوري في كافة الأعمدة.")

    else:
      st.error("عمود 'Difference' غير موجود في الملف، تأكد من مطابقة أسماء الأعمدة.")

  except Exception as e:
    st.error(fحدث خطأ أثناء المعالجة: {e})
else:
  st.info("👈 يرجى رفع ملف الإكسيل المجمع لعرض التحليل المفصل بالكامل.")
