import matplotlib.pyplot as plt
import pandas as pd
import seaborn as stt_sns
import streamlit as st

# إعداد الصفحة لتكون واسعة وبتنسيق احترافي
st.set_page_config(
    page_title="لوحة تحكم تقارير الجرد والفروقات الاحترافية",
    page_icon="🚀",
    layout="wide",
)

# تصميم العنوان بشكل فخم
st.markdown(
    """
    <div style='background-color: #1E1E2F; padding: 20px; border-radius: 10px; text-align: center;'>
        <h1 style='color: #00FFA3; margin: 0;'>🚀 لوحة تحكم تقارير الجرد والفروقات الاحترافية</h1>
        <p style='color: #FFFFFF; margin-top: 10px;'>تحليل شامل ودقيق للفروقات والأصناف (شهر 9 وشهر 10)</p>
    </div>
""",
    unsafe_allow_html=True,
)

st.write("")

# رفع الملف
uploaded_file = st.file_uploader(
    "📁 ارفع ملف الإكسيل المجمع هنا لاستعراض التحليل الواو",
    type=["xlsx", "xls", "csv"],
)

if uploaded_file is not None:
  try:
    # قراءة الملف بذكاء
    if uploaded_file.name.endswith(".csv"):
      df = pd.read_csv(uploaded_file)
    else:
      excel_file = pd.ExcelFile(uploaded_file)
      sheet_name = excel_file.sheet_names[0]
      df = pd.read_excel(excel_file, sheet_name=sheet_name)

    # معالجة عمود التاريخ وتوحيد اسم عمود الفرق إن وجد
    for col in df.columns:
      if "date" in str(col).lower() or "تاريخ" in str(col):
        df["Date"] = pd.to_datetime(df[col], errors="coerce")
        break

    # التأكد من الأعمدة الأساسية
    if "Difference" in df.columns:
      # حساب الإحصائيات الكبرى
      total_items = len(df)
      total_diff = df["Difference"].sum()
      shortage_items = df[df["Difference"] < 0]
      excess_items = df[df["Difference"] > 0]

      # عرض مربعات إحصائية فخمة (Metrics)
      st.markdown("---")
      col1, col2, col3, col4 = st.columns(4)
      col1.metric("📦 إجمالي الأصناف", f"{total_items:,}")
      col2.metric(
          "📉 إجمالي الفروقات",
          f"{total_diff:,.2f}",
          delta_color="inverse",
      )
      col3.metric("⚠️ أصناف بها عجز", f"{len(shortage_items):,}")
      col4.metric("➕ أصناف بها زيادة", f"{len(excess_items):,}")
      st.markdown("---")

      # قسم الرسوم البيانية الاحترافية
      st.subheader("📊 الرسوم البيانية والتحليلات البصرية")
      fig, ax = plt.subplots(figsize=(10, 4))
      top_diffs = df.nlargest(10, "Difference") if "Detail" in df.columns else df
      if "Detail" in df.columns:
        ax.bar(
            top_diffs["Detail"].astype(str),
            top_diffs["Difference"],
            color="#00FFA3",
        )
        plt.xticks(rotation=45, ha="right")
        ax.set_title("أكبر الفروقات المسجلة للأصناف", color="white")
        fig.patch.set_facecolor("#0E1117")
        ax.set_facecolor("#1E1E2F")
        ax.tick_params(colors="white")
        st.pyplot(fig)

      # جدول البيانات مع خيارات البحث والفلترة
      st.subheader("📋 جدول البيانات التفصيلي والفلاتر")
      search_term = st.text_input(
          "🔍 ابحث عن صنف أو كود معين:", ""
      )
      if search_term:
        df_filtered = df[
            df.astype(str)
            .apply(lambda row: row.str.contains(search_term, case=False).any(), axis=1)
        ]
      else:
        df_filtered = df

      st.dataframe(df_filtered, use_container_width=True)

      # زر تحميل النتائج المفدلة
      csv_data = df_filtered.to_csv(index=False).encode("utf-8")
      st.download_button(
          label="📥 تحميل النتائج المفلترة كملف CSV",
          data=csv_data,
          file_name="filtered_inventory_report.csv",
          mime="text/csv",
      )

    else:
      st.success("تم رفع الملف بنجاح ولكن يرجى التأكد من تطابق اسم عمود الفروقات.")
      st.dataframe(df.head(10))

  except Exception as e:
    st.error(f"حدث خطأ أثناء معالجة الملف: {e}")
else:
  st.info(
      "👈 يرجى رفع ملف الإكسيل المجمع من الزر بالأعلى لنبدأ عرض لوحة التحكم"
      " الاحترافية."
  )
