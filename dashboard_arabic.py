"""
Arabic Dashboard for Workforce Insights
--------------------------------------

This Streamlit application allows users to upload two Excel spreadsheets and
an optional PDF report to generate key workforce metrics for their
organisation. It supports Arabic labels throughout the interface and
calculates the following indicators based on the uploaded data:

* عدد العقود محددة المدة وعدد العقود غير محددة المدة
* عدد الموظفين الأجانب وعدد الموظفين الأجانب الذين لديهم أكثر من أربعة تابعين
* معدل توثيق العقود ومعدل التوطين – يتم استخلاصهما تلقائياً من تقرير PDF إن وجد

The app expects the following structure in the Excel files:

1. **Employee file (قوى.xlsx)** – must contain at least the columns:
   - ``Id number`` – a unique identifier for each employee (may be the national ID or iqama number).
   - ``Contract Status`` – values such as ``محدد`` (fixed term) or ``غير محدد`` (indefinite).
   - ``Nationality`` – the employee’s nationality written in English (e.g., ``Saudi Arabia``).

2. **Dependents file (المقيمين النشطين و تابعيهم.xlsx)** – must contain at least:
   - ``رقم إقامة رب الأسرة`` – the iqama number of the head of household (employee).
   - one row per dependent.

3. **Monthly report PDF** – optional. If provided, the app uses ``pdfplumber`` to extract the
   contract documentation rate (معدل توثيق العقود) and the Saudisation rate (معدل التوطين).
   The PDF must contain these Arabic phrases followed by the percentage values. If the text
   extraction fails, the app will ask the user to enter the values manually.

To run the app in Google Colab you can execute the following in a cell:

```python
!pip install streamlit pdfplumber
!streamlit run dashboard_arabic.py --server.port 8501 --server.address 0.0.0.0
```

Streamlit will provide a public URL when run inside Colab. Open that URL to view the dashboard.

"""

import io
import re
from typing import Optional, Tuple

import pandas as pd
import streamlit as st
# Matplotlib is no longer required for charting (we use Plotly instead), but we
# import it conditionally in case other parts of the code or future expansions
# rely on it.  If it is not installed, charts will still function via Plotly.
try:
    import matplotlib.pyplot as plt  # type: ignore
except ImportError:
    plt = None  # type: ignore
# Plotly is used for interactive charts with hover support.  It provides a richer
# user experience compared to Matplotlib and handles right‑to‑left text better.
import plotly.express as px

# Attempt to import pdfplumber for PDF parsing; fall back gracefully if unavailable
try:
    import pdfplumber  # type: ignore
except ImportError:
    pdfplumber = None  # pdfplumber will be None if not installed

# Arabic text rendering helpers
# Matplotlib does not support proper Arabic shaping and right‑to‑left rendering out of the box.  To
# display Arabic labels correctly in charts, we utilise the arabic_reshaper and python‑bidi
# libraries.  If they are not installed, we fall back to using the raw text, which may appear
# reversed or broken.  These imports are optional to keep the app functional even without the
# dependencies.
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
except ImportError:
    arabic_reshaper = None  # type: ignore
    # Define a no‑op fallback so code below can call get_display safely
    def get_display(text: str) -> str:  # type: ignore
        return text



def extract_pdf_metrics(file: io.BytesIO) -> Tuple[Optional[str], Optional[str]]:
    """Extract Saudisation and contract documentation rates from a PDF.

    This function attempts to locate the phrases ``معدل التوطين`` and ``معدل توثيق العقود``
    in the PDF text and then extract the first percentage value that appears nearby.
    It normalizes whitespace and searches within a fixed window after the phrase.

    Parameters
    ----------
    file : io.BytesIO
        A PDF file-like object.

    Returns
    -------
    tuple
        ``(saudisation_rate, contract_doc_rate)`` where each value is a string
        including the percent sign (e.g., ``"75%"``) or ``None`` if not found.
    """
    if pdfplumber is None:
        # PDF extraction library not installed
        return None, None
    try:
        saudisation_rate: Optional[str] = None
        contract_doc_rate: Optional[str] = None
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                # Normalize whitespace to single spaces
                text = re.sub(r"\s+", " ", text)
                # Search for Saudisation phrase and extract first percentage afterwards
                idx_saud = text.find("معدل التوطين")
                if idx_saud != -1 and saudisation_rate is None:
                    window = text[idx_saud:idx_saud + 60]  # look ahead 60 characters
                    m = re.search(r"([0-9٠-٩]+)\s*%", window)
                    if m:
                        # Convert Eastern Arabic numerals to Western if present
                        num = m.group(1)
                        # Replace Arabic-Indic digits with Western digits
                        arabic_to_western = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
                        num = num.translate(arabic_to_western)
                        saudisation_rate = f"{num}%"
                # Search for contract documentation phrase
                idx_doc = text.find("معدل توثيق العقود")
                if idx_doc != -1 and contract_doc_rate is None:
                    window = text[idx_doc:idx_doc + 60]
                    m = re.search(r"([0-9٠-٩]+)\s*%", window)
                    if m:
                        num = m.group(1)
                        arabic_to_western = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
                        num = num.translate(arabic_to_western)
                        contract_doc_rate = f"{num}%"
        return saudisation_rate, contract_doc_rate
    except Exception:
        return None, None


def main() -> None:
    """Run the Streamlit dashboard application."""
    st.set_page_config(page_title="لوحة مؤشرات الموارد البشرية", layout="wide")
    st.title("لوحة مؤشرات الموارد البشرية")
    st.markdown("""
    **مرحبًا!** يمكنكم هنا رفع ملفات البيانات لتحليل معلومات الموظفين بشكل تلقائي.
    الرجاء رفع ملف الموظفين، ملف التابعين، ويمكن رفع تقرير PDF لاستخراج نسب التوطين وتوثيق العقود.
    """)

    # Uploaders
    emp_file = st.file_uploader("❶ – تحميل ملف الموظفين (Excel)", type=["xlsx"], key="emp")
    dep_file = st.file_uploader("❷ – تحميل ملف التابعين (Excel)", type=["xlsx"], key="dep")
    pdf_file = st.file_uploader("❸ – تحميل التقرير الشهري (PDF) – اختياري", type=["pdf"], key="pdf")

    # Initialize placeholders for metrics
    saudisation_rate: Optional[str] = None
    contract_doc_rate: Optional[str] = None
    # New metric: wage protection compliance rate. This will always be entered manually
    wage_protection_rate: Optional[str] = None

    if pdf_file is not None:
        saudisation_rate, contract_doc_rate = extract_pdf_metrics(pdf_file)

    # If PDF extraction failed, allow manual input. We also include a field for the wage protection compliance rate.
    # The expander is open if any of the three rates are missing.
    with st.expander("تعديل نسب التوطين وتوثيق العقود يدويًا", expanded=(saudisation_rate is None or contract_doc_rate is None or wage_protection_rate is None)):
        saudisation_rate = st.text_input(
            "نسبة التوطين (%)",
            value=saudisation_rate if saudisation_rate is not None else "",
            help="أدخل الرقم فقط إذا لم يتم استخلاصه من التقرير"
        )
        contract_doc_rate = st.text_input(
            "نسبة توثيق العقود (%)",
            value=contract_doc_rate if contract_doc_rate is not None else "",
            help="أدخل الرقم فقط إذا لم يتم استخلاصه من التقرير"
        )
        wage_protection_rate = st.text_input(
            "نسبة التزام حماية الأجور (%)",
            value=wage_protection_rate if wage_protection_rate is not None else "",
            help="أدخل الرقم فقط، سيتم عرضه مع مؤشرات الامتثال"
        )

    if emp_file and dep_file:
        try:
            # Read dependents data from the first sheet
            dep_xls = pd.ExcelFile(dep_file)
            dep_df = pd.read_excel(dep_xls, sheet_name=dep_xls.sheet_names[0])

            # Read employee file and automatically detect relevant sheets
            emp_xls = pd.ExcelFile(emp_file)
            contract_sheet = None
            master_sheet = None
            # Identify sheet containing 'Contract Status' and 'Nationality'
            for sheet in emp_xls.sheet_names:
                try:
                    tmp_df = pd.read_excel(emp_xls, sheet_name=sheet, nrows=1)
                except Exception:
                    continue
                cols = tmp_df.columns
                if 'Contract Status' in cols and contract_sheet is None:
                    contract_sheet = sheet
                if 'Nationality' in cols and 'Id number' in cols and master_sheet is None:
                    master_sheet = sheet
            if contract_sheet is None:
                st.error("لم يتم العثور على ورقة تحتوي على عمود 'Contract Status'. يرجى التحقق من ملف الموظفين.")
                return
            if master_sheet is None:
                st.error("لم يتم العثور على ورقة تحتوي على أعمدة 'Nationality' و 'Id number'. يرجى التحقق من ملف الموظفين.")
                return

            # Load sheets fully
            contract_df = pd.read_excel(emp_xls, sheet_name=contract_sheet)
            master_df = pd.read_excel(emp_xls, sheet_name=master_sheet)

            # Ensure the contract status column exists
            if 'Contract Status' not in contract_df.columns:
                st.error("العمود 'Contract Status' غير موجود في ورقة العقد.")
                return
            # Normalize contract status values: remove surrounding spaces and drop missing
            contract_df['Contract Status'] = contract_df['Contract Status'].astype(str).str.strip()
            contract_counts = contract_df['Contract Status'][
                ~contract_df['Contract Status'].isin(['', 'nan', 'NaN'])
            ].value_counts()

            # Prepare for merging
            # Ensure 'Id number' exists in both
            if 'Id number' not in master_df.columns or 'Id number' not in contract_df.columns:
                st.error("العمود 'Id number' غير موجود في إحدى أوراق الموظفين.")
                return
            master_df['Id number'] = master_df['Id number'].astype(str)
            contract_df['Id number'] = contract_df['Id number'].astype(str)

            # Merge nationality and contract status
            emp_df = pd.merge(master_df[['Id number', 'Nationality']], contract_df[['Id number', 'Contract Status']], on='Id number', how='inner')

            # Count foreign employees
            foreign_emps = emp_df[emp_df['Nationality'].str.lower() != 'saudi arabia']
            num_foreign = len(foreign_emps)

            # Process dependents: convert to numeric for merge
            dep_df['رقم إقامة رب الأسرة'] = pd.to_numeric(dep_df['رقم إقامة رب الأسرة'], errors='coerce').astype('Int64')
            emp_df['Id number_int'] = pd.to_numeric(emp_df['Id number'], errors='coerce').astype('Int64')

            # Group dependents count by head id
            dep_counts = dep_df.groupby('رقم إقامة رب الأسرة').size().reset_index(name='dependents')
            emp_merged = pd.merge(emp_df, dep_counts, left_on='Id number_int', right_on='رقم إقامة رب الأسرة', how='left')
            emp_merged['dependents'] = emp_merged['dependents'].fillna(0)

            foreign_with_many_dep = emp_merged[(emp_merged['Nationality'].str.lower() != 'saudi arabia') & (emp_merged['dependents'] > 4)]

            # Display metrics
            st.subheader("النتائج الرئيسية")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("العقود محددة المدة", int(contract_counts.get('محدد', 0)))
            col2.metric("العقود غير محددة المدة", int(contract_counts.get('غير محدد', 0)))
            col3.metric("عدد الموظفين الأجانب", num_foreign)
            col4.metric("عدد الأجانب مع أكثر من أربعة تابعين", len(foreign_with_many_dep))

            # Display compliance metrics, including wage protection compliance if provided
            with st.container():
                st.subheader("مؤشرات الامتثال")
                metric_cols = st.columns(3)
                # Saudisation rate metric
                if saudisation_rate:
                    metric_cols[0].metric("نسبة التوطين", f"{saudisation_rate}")
                else:
                    metric_cols[0].write("لم يتم توفير نسبة التوطين")
                # Contract documentation rate metric
                if contract_doc_rate:
                    metric_cols[1].metric("نسبة توثيق العقود", f"{contract_doc_rate}")
                else:
                    metric_cols[1].write("لم يتم توفير نسبة توثيق العقود")
                # Wage protection compliance rate metric
                if wage_protection_rate:
                    metric_cols[2].metric("نسبة التزام حماية الأجور", f"{wage_protection_rate}")
                else:
                    metric_cols[2].write("لم يتم توفير نسبة التزام حماية الأجور")

            # Additional statistics
            with st.expander("إحصائيات إضافية"):
                # Use interactive charts via Plotly to improve readability and provide hover
                # information.  Arabic labels are reshaped and displayed correctly if the
                # arabic_reshaper and bidi libraries are available.  Otherwise, the raw
                # labels are used.

                # Contract status distribution as a pie chart
                st.subheader("توزيع حالة العقد")
                # Create a DataFrame for plotly with status and counts
                contract_data = contract_counts.reset_index()
                contract_data.columns = ['status', 'count']
                # Generate Arabic‑aware labels
                if arabic_reshaper is not None:
                    contract_data['status_label'] = contract_data['status'].apply(
                        lambda lbl: get_display(arabic_reshaper.reshape(str(lbl)))
                    )
                else:
                    contract_data['status_label'] = contract_data['status'].astype(str)
                # Create a pie chart using Plotly Express
                fig1 = px.pie(
                    contract_data,
                    names='status_label',
                    values='count',
                    hole=0.0,
                )
                fig1.update_traces(textinfo='percent+label')
                # Helper to properly display Arabic text if the reshaper is available
                def arabic_display(text: str) -> str:
                    """Reshape and reorder Arabic text for right‑to‑left display."""
                    if arabic_reshaper is not None:
                        return get_display(arabic_reshaper.reshape(text))
                    return text

                fig1.update_layout(
                    title_text=arabic_display("توزيع حالة العقد"),
                    showlegend=True,
                    legend_title_text='',
                    # Ensure a font that supports Arabic is used
                    font=dict(family="DejaVu Sans", size=14),
                )
                st.plotly_chart(fig1, use_container_width=True)

                # Nationality distribution as a bar chart
                st.subheader("توزيع الجنسيات")
                nat_counts_df = emp_df['Nationality'].value_counts().reset_index()
                nat_counts_df.columns = ['nationality', 'count']
                # Map English nationality names to Arabic.  If a nationality is not in the
                # mapping, fall back to the original name.  This ensures the bar chart
                # displays nationalities in Arabic for known values.
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
                nat_counts_df['nat_ar'] = nat_counts_df['nationality'].map(NATIONALITY_AR_MAP).fillna(nat_counts_df['nationality'])
                # Prepare Arabic labels using reshaper and bidi for proper display
                if arabic_reshaper is not None:
                    nat_counts_df['nat_label'] = nat_counts_df['nat_ar'].apply(
                        lambda lbl: get_display(arabic_reshaper.reshape(str(lbl)))
                    )
                else:
                    nat_counts_df['nat_label'] = nat_counts_df['nat_ar'].astype(str)
                # Create a bar chart with Plotly
                fig2 = px.bar(
                    nat_counts_df,
                    x='nat_label',
                    y='count',
                    labels={'nat_label': 'الجنسية', 'count': 'عدد الموظفين'},
                )
                fig2.update_layout(
                    title_text=arabic_display("توزيع الجنسيات"),
                    xaxis_title=arabic_display("الجنسية"),
                    yaxis_title=arabic_display("عدد الموظفين"),
                    font=dict(family="DejaVu Sans", size=14),
                )
                # Show numeric values on hover
                fig2.update_traces(hovertemplate='%{x}: %{y}')
                st.plotly_chart(fig2, use_container_width=True)

        except Exception as exc:
            st.error(f"حدث خطأ أثناء معالجة البيانات: {exc}")


if __name__ == "__main__":
    main()
