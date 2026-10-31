import io
import pandas as pd
import streamlit as st
import datetime

st.set_page_config(
    page_title="متابعة محمد قطان - نظام إدارة المخازن",
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
st.caption("النظام المبسط: معادلة الجرد اليومي المدمج + تقرير الأسبوع فقط")

# القائمة الجانبية مختصرة ومقتصرة على المطلوب فقط
main_mode = st.sidebar.radio(
    "📂 الأقسام الرئيسية:",
    [
        "🧮 معادلة الجرد اليومي المدمج (تصدير .xls)",
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

def convert_df_to_xls(df_to_export):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_export.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
    return output.getvalue()

@st.cache_data
def load_master_codes():
    try:
        df_m = pd.read_excel("رمز المادة.xlsx")
        name_to_code = {}
        for val in df_m.iloc[:, 0].dropna():
            val_str = str(val).strip()
            if '-' in val_str:
                p = val_str.split('-', 1)
                code_val = format_code_as_text(p[0])
                name_val = p[1].strip().lower()
                name_to
