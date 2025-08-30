
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import arabic_reshaper
from bidi.algorithm import get_display

st.set_page_config(page_title="لوحة المؤشرات", layout="wide")

def ar_label(s: str) -> str:
    """Reshape + apply bidi so Arabic renders correctly in charts and UI."""
    try:
        return get_display(arabic_reshaper.reshape(str(s)))
    except Exception:
        return str(s)

# ---------- Sidebar: data input ----------
st.sidebar.header(ar_label("البيانات"))
st.sidebar.write(ar_label("ارفع ملفات CSV أو استخدم البيانات التجريبية"))

emp_file = st.sidebar.file_uploader(ar_label("ملف الموظفين (يحتوي على عمود Nationality)"), type=["csv"])
contract_file = st.sidebar.file_uploader(ar_label("ملف العقود (يحتوي على عمود Status)"), type=["csv"])

# Demo fallback
if emp_file is not None:
    emp_df = pd.read_csv(emp_file)
else:
    emp_df = pd.DataFrame({
        "Nationality": ["Saudi Arabia"] * 260 + ["Egyptian"] * 8 + ["Pakistani"] * 6 + ["Jordanian"] * 5 +
                        ["Lebanese"] * 3 + ["Indian"] * 3 + ["British"] * 2 + ["Portuguese"] * 2 +
                        ["Syrian"] * 1 + ["Ukraine"] * 1 + ["New Zealand"] * 1 + ["Kingdom of Morocco"] * 1
    })

if contract_file is not None:
    contract_df = pd.read_csv(contract_file)
else:
    contract_df = pd.DataFrame({
        "Status": (["نشط"] * 210) + (["منتهي"] * 80)
    })

# ---------- Arabic mapping for nationalities ----------
NATIONALITY_AR_MAP = {
    'Saudi Arabia': 'السعودية',
    'Syrian': 'سوري',
    'Jordanian': 'أردني',
    'Pakistani': 'باكستاني',
    'Egyptian': 'مصري',
    'Ukraine': 'أوكراني',
    'Portuguese': 'برتغالي',
    'Lebanese': 'لبناني',
    'New Zealand': 'نيوزيلندي',
    'Kingdom of Morocco': 'المملكة المغربية',
    'Indian': 'هندي',
    'British': 'بريطاني',
}

# ---------- Charts ----------
st.title(ar_label("لوحة المؤشرات"))

col1, col2 = st.columns([2, 3])

with col1:
    # Pie: contract status distribution
    status_counts = contract_df["Status"].value_counts().reset_index()
    status_counts.columns = ["status", "count"]
    status_counts["status_label"] = status_counts["status"].apply(ar_label)

    fig_pie = px.pie(
        status_counts,
        names="status_label",
        values="count",
        hole=0.0
    )
    fig_pie.update_traces(textinfo="percent+label", hovertemplate="%{label}: %{value}")
    fig_pie.update_layout(
        title_text=ar_label("توزيع حالة العقد"),
        legend_title_text="",
        font=dict(family="Cairo, Noto Kufi Arabic, Amiri, DejaVu Sans, Arial", size=14),
    )
    st.plotly_chart(fig_pie, use_container_width=True)


with col2:
    # Bar: nationality distribution
    nat_counts_df = emp_df["Nationality"].value_counts().reset_index()
    nat_counts_df.columns = ["nationality", "count"]
    nat_counts_df["nat_ar"] = nat_counts_df["nationality"].map(NATIONALITY_AR_MAP).fillna(nat_counts_df["nationality"])
    nat_counts_df["nat_label"] = nat_counts_df["nat_ar"].apply(ar_label)

    fig_bar = px.bar(
        nat_counts_df.sort_values("count", ascending=False),
        x="nat_label",
        y="count",
        labels={"nat_label": ar_label("الجنسية"), "count": ar_label("عدد الموظفين")},
    )
    fig_bar.update_traces(hovertemplate="%{x}: %{y}")
    fig_bar.update_layout(
        title_text=ar_label("توزيع الجنسيات"),
        xaxis_title=ar_label("الجنسية"),
        yaxis_title=ar_label("عدد الموظفين"),
        font=dict(family="Cairo, Noto Kufi Arabic, Amiri, DejaVu Sans, Arial", size=14),
        margin=dict(l=20, r=20, t=60, b=20)
    )
    st.plotly_chart(fig_bar, use_container_width=True)

st.caption(ar_label("ملاحظة: تتم معالجة النص العربي باستخدام arabic-reshaper و python-bidi لضمان عرض صحيح."))
