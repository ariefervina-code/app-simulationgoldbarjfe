
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime

# =========================================================
# 1. PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="Employee Gold Award Dashboard",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# 2. CUSTOM CSS
# =========================================================

st.markdown("""
<style>

.stApp {
    background-color: #F5F7FB;
}

.block-container {
    padding-top: 1.5rem;
    padding-bottom: 2rem;
}

.dashboard-title {
    font-size: 32px;
    font-weight: 800;
    color: #172554;
}

.dashboard-subtitle {
    font-size: 14px;
    color: #64748B;
}

[data-testid="stMetric"] {
    background-color: white;
    padding: 18px;
    border-radius: 14px;
    border: 1px solid #E2E8F0;
    box-shadow: 0px 3px 10px rgba(0,0,0,0.03);
}

[data-testid="stMetricValue"] {
    color: #172554;
    font-weight: 800;
}

[data-testid="stMetricLabel"] {
    color: #64748B;
    font-weight: 600;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# 3. HELPER FUNCTIONS
# =========================================================

def format_rupiah(value):

    return "Rp " + f"{value:,.0f}".replace(",", ".")


def parse_date(data):

    return pd.to_datetime(
        data,
        errors="coerce",
        dayfirst=True
    )


def clean_id(value):

    if pd.isna(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()


def find_column(columns, keywords):

    # Exact match first
    for keyword in keywords:

        for col in columns:

            if str(col).strip().lower() == keyword.lower():
                return col

    # Partial match
    for keyword in keywords:

        for col in columns:

            if keyword.lower() in str(col).strip().lower():
                return col

    return None


def export_excel(summary, detail, department):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        summary.to_excel(
            writer,
            sheet_name="Budget Projection",
            index=False
        )

        detail.to_excel(
            writer,
            sheet_name="Employee Award",
            index=False
        )

        department.to_excel(
            writer,
            sheet_name="Department Analysis",
            index=False
        )

    output.seek(0)

    return output.getvalue()


# =========================================================
# 4. HEADER
# =========================================================

st.markdown(
    '<div class="dashboard-title">'
    '🏆 Employee Gold Award Dashboard'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="dashboard-subtitle">'
    'Long Service Award | Employee Recognition | Budget Planning'
    '</div>',
    unsafe_allow_html=True
)

st.write("")


# =========================================================
# 5. SIDEBAR SETTINGS
# =========================================================

FILE_PATH = (
    r"C:\Users\vb222426\Documents"
    r"\SIMULATIONGOLDBAR\Simulation Gold Bar.xlsx"
)

with st.sidebar:

    st.header("⚙️ Dashboard Settings")

    st.subheader("Gold Award Policy")

    gold_5 = st.number_input(
        "5 Years Award (Gram)",
        min_value=0.0,
        value=5.0,
        step=0.5
    )

    gold_10 = st.number_input(
        "10 Years Award (Gram)",
        min_value=0.0,
        value=10.0,
        step=0.5
    )

    st.divider()

    st.subheader("Gold Price Simulation")

    gold_price = st.number_input(
        "Gold Price per Gram (Rp)",
        min_value=0,
        value=0,
        step=25000,
        help="Masukkan asumsi harga emas per gram."
    )

    st.caption(
        "Harga emas dapat diubah sesuai kebutuhan simulasi HR."
    )


# =========================================================
# 6. READ EXCEL AUTOMATICALLY
# =========================================================

try:

    excel = pd.ExcelFile(FILE_PATH)

    with st.sidebar:

        st.divider()

        st.subheader("Excel Source")

        sheet_name = st.selectbox(
            "Select Worksheet",
            excel.sheet_names
        )

    df = pd.read_excel(
        FILE_PATH,
        sheet_name=sheet_name
    )

except FileNotFoundError:

    st.error(
        "File Excel tidak ditemukan. "
        "Periksa kembali lokasi file."
    )

    st.code(FILE_PATH)

    st.stop()

except Exception as e:

    st.error(f"Gagal membaca Excel: {e}")

    st.stop()


if df.empty:

    st.warning("Data Excel kosong.")

    st.stop()


df.columns = [
    str(col).strip()
    for col in df.columns
]

columns = list(df.columns)


# =========================================================
# 7. COLUMN MAPPING
# =========================================================

with st.sidebar:

    st.divider()

    st.subheader("📋 Column Mapping")

    def select_column(label, keywords, optional=False):

        guess = find_column(
            columns,
            keywords
        )

        options = (
            ["-- Not Available --"] + columns
            if optional
            else columns
        )

        if optional:

            index = (
                columns.index(guess) + 1
                if guess in columns
                else 0
            )

        else:

            index = (
                columns.index(guess)
                if guess in columns
                else 0
            )

        return st.selectbox(
            label,
            options,
            index=index
        )


    nik_col = select_column(
        "NIK",
        ["NIK", "Employee ID"]
    )

    name_col = select_column(
        "Employee Name",
        ["Name", "Employee Name"]
    )

    dept_col = select_column(
        "Department",
        ["Dept", "Department"]
    )

    section_col = select_column(
        "Section",
        ["Section"],
        optional=True
    )

    join_col = select_column(
        "Join Date",
        ["Join Date", "Date of Joining"],
        optional=True
    )

    five_col = select_column(
        "5 Years Award Date",
        ["5 Years", "5 Year"]
    )

    ten_col = select_column(
        "10 Years Award Date",
        ["10 Years", "10 Year"]
    )

    retired_col = select_column(
        "Retirement Date",
        ["Retired", "Retirement Date"],
        optional=True
    )


# =========================================================
# 8. PREPARE EMPLOYEE DATA
# =========================================================

employee = pd.DataFrame()

employee["NIK"] = (
    df[nik_col].apply(clean_id)
)

employee["Employee Name"] = (
    df[name_col]
    .fillna("")
    .astype(str)
    .str.strip()
)

employee["Department"] = (
    df[dept_col]
    .fillna("Not Available")
    .astype(str)
    .str.strip()
)


if section_col != "-- Not Available --":

    employee["Section"] = (
        df[section_col]
        .fillna("Not Available")
        .astype(str)
        .str.strip()
    )

else:

    employee["Section"] = "Not Available"


if join_col != "-- Not Available --":

    employee["Join Date"] = parse_date(
        df[join_col]
    )

else:

    employee["Join Date"] = pd.NaT


employee["5 Years Date"] = parse_date(
    df[five_col]
)

employee["10 Years Date"] = parse_date(
    df[ten_col]
)


if retired_col != "-- Not Available --":

    employee["Retired Date"] = parse_date(
        df[retired_col]
    )

else:

    employee["Retired Date"] = pd.NaT


# Remove blank employee names

employee = employee[
    employee["Employee Name"].ne("")
].copy()


# =========================================================
# 9. CREATE AWARD DATA
# =========================================================

five_award = employee[
    employee["5 Years Date"].notna()
].copy()

five_award["Award"] = "5 Years"

five_award["Award Date"] = (
    five_award["5 Years Date"]
)

five_award["Gold Gram"] = gold_5


ten_award = employee[
    employee["10 Years Date"].notna()
].copy()

ten_award["Award"] = "10 Years"

ten_award["Award Date"] = (
    ten_award["10 Years Date"]
)

ten_award["Gold Gram"] = gold_10


award_data = pd.concat(
    [five_award, ten_award],
    ignore_index=True
)


if award_data.empty:

    st.warning(
        "Tidak ditemukan tanggal award "
        "pada kolom 5 Years atau 10 Years."
    )

    st.stop()


# =========================================================
# 10. REMOVE DUPLICATE AWARD RECORDS
# =========================================================

award_data["_Employee Key"] = award_data.apply(
    lambda row: (
        "NIK:" + row["NIK"]
        if row["NIK"]
        else "NAME:" + row["Employee Name"].strip().lower()
    ),
    axis=1
)

# Satu karyawan tidak dihitung dua kali
# untuk jenis award yang sama.

award_data = award_data.drop_duplicates(
    subset=["_Employee Key", "Award"],
    keep="first"
).copy()


# =========================================================
# 11. RETIREMENT FILTER
# =========================================================

with st.sidebar:

    st.divider()

    st.subheader("Employee Eligibility")

    exclude_retired = st.checkbox(
        "Exclude employees retired before award date",
        value=True,
        disabled=(retired_col == "-- Not Available --")
    )


if exclude_retired:

    award_data = award_data[
        award_data["Retired Date"].isna()
        |
        (
            award_data["Retired Date"]
            >= award_data["Award Date"]
        )
    ].copy()


if award_data.empty:

    st.warning(
        "Tidak ada penerima award setelah filter eligibility."
    )

    st.stop()


# =========================================================
# 12. AWARD CALCULATION
# =========================================================

award_data["Award Year"] = (
    award_data["Award Date"].dt.year
)

award_data["Gold Budget"] = (
    award_data["Gold Gram"] * gold_price
)


award_data = award_data.sort_values(
    ["Award Year", "Award", "Employee Name"]
)


# =========================================================
# 13. SIDEBAR FILTER
# =========================================================

st.sidebar.divider()

st.sidebar.subheader("🔎 Data Filters")

min_year = int(
    award_data["Award Year"].min()
)

max_year = int(
    award_data["Award Year"].max()
)

available_years = list(
    range(min_year, max_year + 1)
)

selected_years = st.sidebar.multiselect(
    "Award Year",
    available_years,
    default=available_years
)

selected_awards = st.sidebar.multiselect(
    "Award Type",
    ["5 Years", "10 Years"],
    default=["5 Years", "10 Years"]
)

departments = sorted(
    award_data["Department"].dropna().unique().tolist()
)

selected_departments = st.sidebar.multiselect(
    "Department",
    departments,
    default=departments
)


# =========================================================
# 14. APPLY FILTER
# =========================================================

filtered = award_data[
    award_data["Award Year"].isin(selected_years)
    &
    award_data["Award"].isin(selected_awards)
    &
    award_data["Department"].isin(selected_departments)
].copy()


# =========================================================
# 15. EXECUTIVE SUMMARY
# =========================================================

st.divider()

st.subheader("📊 Executive Summary")

total_employee = employee["_Employee Key"] if "_Employee Key" in employee.columns else employee["NIK"]

total_employee_count = (
    employee["NIK"].replace("", pd.NA).dropna().nunique()
)

count_5 = len(
    filtered[filtered["Award"] == "5 Years"]
)

count_10 = len(
    filtered[filtered["Award"] == "10 Years"]
)

total_recipients = len(filtered)

total_gold = filtered["Gold Gram"].sum()

total_budget = filtered["Gold Budget"].sum()


c1, c2, c3, c4 = st.columns(4)

c1.metric(
    "Total Employees",
    f"{total_employee_count:,}"
)

c2.metric(
    "5 Years Recipients",
    f"{count_5:,}"
)

c3.metric(
    "10 Years Recipients",
    f"{count_10:,}"
)

c4.metric(
    "Estimated Budget",
    format_rupiah(total_budget)
)


c5, c6 = st.columns(2)

c5.metric(
    "Total Award Recipients",
    f"{total_recipients:,}"
)

c6.metric(
    "Total Gold Requirement",
    f"{total_gold:,.1f} Gram"
)


if gold_price == 0:

    st.warning(
        "Masukkan harga emas per gram pada sidebar "
        "untuk menghitung estimasi anggaran."
    )


# =========================================================
# 16. YEARLY PROJECTION
# =========================================================

st.divider()

st.subheader("📅 Yearly Award Projection")


summary_rows = []

for year in selected_years:

    year_data = filtered[
        filtered["Award Year"] == year
    ]

    n5 = len(
        year_data[year_data["Award"] == "5 Years"]
    )

    n10 = len(
        year_data[year_data["Award"] == "10 Years"]
    )

    gram5 = n5 * gold_5

    gram10 = n10 * gold_10

    total_gram = gram5 + gram10

    budget = total_gram * gold_price

    summary_rows.append({

        "Year": year,

        "5 Years Recipients": n5,

        "10 Years Recipients": n10,

        "Total Recipients": n5 + n10,

        "Gold 5 Gram": gram5,

        "Gold 10 Gram": gram10,

        "Total Gold (Gram)": total_gram,

        "Estimated Budget (Rp)": budget

    })


summary = pd.DataFrame(summary_rows)


if summary.empty:

    st.info(
        "Pilih minimal satu tahun pada sidebar "
        "untuk menampilkan proyeksi."
    )

else:

    # =====================================================
    # 17. RECIPIENT CHART
    # =====================================================

    col1, col2 = st.columns(2)

    with col1:

        chart_count = summary.melt(
            id_vars="Year",

            value_vars=[
                "5 Years Recipients",
                "10 Years Recipients"
            ],

            var_name="Award Type",

            value_name="Employees"
        )

        fig_count = px.bar(

            chart_count,

            x="Year",

            y="Employees",

            color="Award Type",

            barmode="group",

            text_auto=True,

            title="Award Recipients by Year",

            color_discrete_sequence=[
                "#D6A84F",
                "#64748B"
            ]

        )

        fig_count.update_layout(
            plot_bgcolor="white",
            paper_bgcolor="white",
            legend_title_text="",
            xaxis=dict(dtick=1)
        )

        st.plotly_chart(
            fig_count,
            use_container_width=True
        )


    # =====================================================
    # 18. BUDGET CHART
    # =====================================================

    with col2:

        fig_budget = px.bar(

            summary,

            x="Year",

            y="Estimated Budget (Rp)",

            text_auto=".2s",

            title="Estimated Gold Award Budget",

            color_discrete_sequence=[
                "#2563EB"
            ]

        )

        fig_budget.update_layout(
            plot_bgcolor="white",
            paper_bgcolor="white",
            xaxis=dict(dtick=1),
            yaxis_title="Budget (Rp)"
        )

        st.plotly_chart(
            fig_budget,
            use_container_width=True
        )


    # =====================================================
    # 19. GOLD REQUIREMENT CHART
    # =====================================================

    st.divider()

    st.subheader("🪙 Gold Requirement Projection")

    gold_chart = summary.melt(
        id_vars="Year",

        value_vars=[
            "Gold 5 Gram",
            "Gold 10 Gram"
        ],

        var_name="Gold Category",

        value_name="Gold Gram"
    )

    fig_gold = px.bar(

        gold_chart,

        x="Year",

        y="Gold Gram",

        color="Gold Category",

        barmode="stack",

        text_auto=True,

        title="Annual Gold Requirement",

        color_discrete_sequence=[
            "#D6A84F",
            "#64748B"
        ]

    )

    fig_gold.update_layout(
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis=dict(dtick=1),
        yaxis_title="Gold (Gram)"
    )

    st.plotly_chart(
        fig_gold,
        use_container_width=True
    )


    # =====================================================
    # 20. ANNUAL BUDGET TABLE
    # =====================================================

    st.divider()

    st.subheader("📋 Annual Budget Planning")

    budget_table = summary.copy()

    budget_table["Estimated Budget"] = (
        budget_table["Estimated Budget (Rp)"]
        .apply(format_rupiah)
    )

    budget_table = budget_table.drop(
        columns=["Estimated Budget (Rp)"]
    )

    st.dataframe(
        budget_table,
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# 21. EMPLOYEE AWARD DETAILS
# =========================================================

st.divider()

st.subheader("👥 Employee Award Details")

search = st.text_input(
    "Search Employee by NIK or Name"
)

detail = filtered.copy()

if search.strip():

    detail = detail[
        detail["NIK"].str.contains(
            search,
            case=False,
            na=False
        )
        |
        detail["Employee Name"].str.contains(
            search,
            case=False,
            na=False
        )
    ]


display_detail = detail.copy()

display_detail["Award Date"] = (
    display_detail["Award Date"].dt.strftime("%d-%b-%Y")
)

display_detail["Join Date"] = (
    display_detail["Join Date"].dt.strftime("%d-%b-%Y")
)

display_detail["Retired Date"] = (
    display_detail["Retired Date"].dt.strftime("%d-%b-%Y")
)

display_detail["Gold Budget"] = (
    display_detail["Gold Budget"].apply(format_rupiah)
)

st.dataframe(

    display_detail[
        [
            "Award Year",
            "NIK",
            "Employee Name",
            "Department",
            "Section",
            "Join Date",
            "Award",
            "Award Date",
            "Gold Gram",
            "Gold Budget"
        ]
    ],

    use_container_width=True,

    hide_index=True

)


# =========================================================
# 22. DEPARTMENT ANALYSIS
# =========================================================

st.divider()

st.subheader("🏢 Department Analysis")


dept_summary = (

    filtered.groupby(
        ["Department", "Award"]
    )

    .agg(

        Recipients=("NIK", "count"),

        Gold_Gram=("Gold Gram", "sum"),

        Budget=("Gold Budget", "sum")

    )

    .reset_index()

)


if not dept_summary.empty:

    fig_dept = px.bar(

        dept_summary,

        x="Department",

        y="Recipients",

        color="Award",

        barmode="group",

        text_auto=True,

        title="Award Recipients by Department",

        color_discrete_sequence=[
            "#D6A84F",
            "#64748B"
        ]

    )

    fig_dept.update_layout(
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis_tickangle=-35,
        legend_title_text=""
    )

    st.plotly_chart(
        fig_dept,
        use_container_width=True
    )

    dept_display = dept_summary.copy()

    dept_display["Budget"] = (
        dept_display["Budget"].apply(format_rupiah)
    )

    st.dataframe(
        dept_display,
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# 23. GOLD PRICE SCENARIO SIMULATION
# =========================================================

st.divider()

st.subheader("💰 Gold Price Scenario Simulation")

st.caption(
    "Simulasikan perubahan harga emas "
    "dan dampaknya terhadap anggaran perusahaan."
)

price_change = st.slider(

    "Gold Price Change (%)",

    min_value=-30,

    max_value=100,

    value=0,

    step=5

)

scenario_price = (
    gold_price * (1 + price_change / 100)
)

scenario_budget = (
    total_gold * scenario_price
)


s1, s2, s3 = st.columns(3)

s1.metric(
    "Current Gold Price",
    format_rupiah(gold_price)
)

s2.metric(
    "Scenario Gold Price",
    format_rupiah(scenario_price)
)

s3.metric(
    "Scenario Total Budget",
    format_rupiah(scenario_budget)
)


# =========================================================
# 24. EXPORT EXCEL
# =========================================================

st.divider()

st.subheader("📥 Download Report")

export_detail = filtered.drop(
    columns=["_Employee Key"],
    errors="ignore"
).copy()

export_summary = (
    summary.copy()
    if not summary.empty
    else pd.DataFrame()
)

export_department = dept_summary.copy()


excel_bytes = export_excel(
    export_summary,
    export_detail,
    export_department
)


st.download_button(

    label="📊 Download Excel Report",

    data=excel_bytes,

    file_name="Employee_Gold_Award_Report.xlsx",

    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),

    use_container_width=True

)


# =========================================================
# 25. FOOTER
# =========================================================

st.divider()

st.caption(
    "Employee Gold Award Dashboard | HR & GA | "
    "Long Service Award & Budget Projection"
)

