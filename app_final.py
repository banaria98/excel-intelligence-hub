import io
import re
import warnings

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

warnings.filterwarnings("ignore")


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Excel Intelligence Hub",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>

    /* ---------------------------------------------------------
       GLOBAL
    --------------------------------------------------------- */

    .stApp {
        background: #f4f7fb;
        color: #172033;
    }

    .block-container {
        max-width: 1550px;
        padding-top: 1.2rem;
        padding-bottom: 3rem;
    }


    /* ---------------------------------------------------------
       SIDEBAR
    --------------------------------------------------------- */

    [data-testid="stSidebar"] {
        background: #172033;
    }

    [data-testid="stSidebar"] * {
        color: #ffffff !important;
    }

    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p {
        color: #ffffff !important;
    }

    /* Sidebar select boxes */

    [data-testid="stSidebar"] div[data-baseweb="select"] {
        background: #ffffff !important;
    }

    [data-testid="stSidebar"] div[data-baseweb="select"] * {
        color: #172033 !important;
    }

    [data-testid="stSidebar"] input {
        color: #172033 !important;
        background: #ffffff !important;
    }

    /* Multiselect */

    [data-testid="stSidebar"] div[data-baseweb="tag"] {
        background: #e8edf5 !important;
    }

    [data-testid="stSidebar"] div[data-baseweb="tag"] span {
        color: #172033 !important;
    }

    /* Dropdown menu */

    div[data-baseweb="popover"] {
        background: #ffffff !important;
    }

    div[data-baseweb="popover"] * {
        color: #172033 !important;
    }

    li[role="option"] {
        color: #172033 !important;
        background: #ffffff !important;
    }

    li[role="option"]:hover {
        background: #edf2f8 !important;
        color: #172033 !important;
    }


    /* ---------------------------------------------------------
       HERO
    --------------------------------------------------------- */

    .hero {
        background: linear-gradient(
            135deg,
            #172033 0%,
            #304b75 100%
        );

        padding: 32px 38px;
        border-radius: 20px;
        color: white;
        margin-bottom: 22px;

        box-shadow:
            0 12px 35px rgba(23, 32, 51, 0.14);
    }

    .hero h1 {
        font-size: 38px;
        margin: 0 0 7px 0;
        font-weight: 800;
        color: white;
    }

    .hero p {
        margin: 0;
        font-size: 15px;
        color: #dce5f1;
    }


    /* ---------------------------------------------------------
       SECTION HEADINGS
    --------------------------------------------------------- */

    .section-title {
        font-size: 25px;
        font-weight: 800;
        color: #172033;
        margin-top: 32px;
        margin-bottom: 5px;
    }

    .section-subtitle {
        color: #65738a;
        font-size: 14px;
        margin-bottom: 15px;
    }


    /* ---------------------------------------------------------
       CARDS
    --------------------------------------------------------- */

    .card {
        background: #ffffff;
        border: 1px solid #e1e7ef;
        border-radius: 16px;
        padding: 20px;

        box-shadow:
            0 5px 18px rgba(23, 32, 51, 0.05);

        margin-bottom: 14px;
    }

    .card h3 {
        color: #172033;
        margin-top: 0;
    }


    /* ---------------------------------------------------------
       KPI CARDS
    --------------------------------------------------------- */

    .kpi-card {
        background: #ffffff;
        border: 1px solid #e1e7ef;
        border-radius: 16px;
        padding: 20px;
        min-height: 125px;

        box-shadow:
            0 5px 18px rgba(23, 32, 51, 0.05);
    }

    .kpi-label {
        color: #697890;
        font-size: 11px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: .06em;
    }

    .kpi-value {
        color: #172033;
        font-size: 29px;
        font-weight: 800;
        margin-top: 8px;
    }

    .kpi-description {
        color: #7a879a;
        font-size: 12px;
        margin-top: 5px;
    }


    /* ---------------------------------------------------------
       CHART DESCRIPTION
    --------------------------------------------------------- */

    .chart-description {
        background: #f7f9fc;
        border-left: 4px solid #6c7f9d;

        padding: 11px 14px;
        margin-bottom: 12px;

        border-radius: 0 8px 8px 0;

        color: #536176;
        font-size: 13px;
        line-height: 1.45;
    }

    .chart-description b {
        color: #172033;
    }


    /* ---------------------------------------------------------
       DATA QUALITY
    --------------------------------------------------------- */

    .status-good {
        background: #edf8f1;
        border: 1px solid #cdebd7;
        color: #21613a;

        padding: 12px 15px;
        border-radius: 10px;
    }

    .status-warning {
        background: #fff8e8;
        border: 1px solid #f0dfad;
        color: #765d18;

        padding: 12px 15px;
        border-radius: 10px;
    }

    .status-danger {
        background: #fff0f0;
        border: 1px solid #efcccc;
        color: #8a3030;

        padding: 12px 15px;
        border-radius: 10px;
    }


    /* ---------------------------------------------------------
       TABLE
    --------------------------------------------------------- */

    [data-testid="stDataFrame"] {
        background: white;
    }


    /* ---------------------------------------------------------
       FOOTER
    --------------------------------------------------------- */

    .footer {
        text-align: center;
        color: #7a879a;
        padding: 30px;
        font-size: 12px;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# CLEANING
# ============================================================

def clean_column_name(name):
    name = str(name).strip()
    name = re.sub(r"\s+", " ", name)
    return name


def clean_dataframe(df):
    df = df.copy()

    df.columns = [
        clean_column_name(c)
        for c in df.columns
    ]

    # Remove completely empty rows/columns
    df = df.dropna(
        axis=0,
        how="all"
    )

    df = df.dropna(
        axis=1,
        how="all"
    )

    # Convert numeric-looking text
    for col in df.columns:

        if df[col].dtype == "object":

            text = (
                df[col]
                .astype(str)
                .str.strip()
            )

            numeric = pd.to_numeric(
                text
                .str.replace(",", "", regex=False)
                .str.replace("$", "", regex=False)
                .str.replace("%", "", regex=False),
                errors="coerce",
            )

            if numeric.notna().mean() >= 0.90:
                df[col] = numeric

    # Convert date-looking columns
    for col in df.columns:

        if df[col].dtype == "object":

            name = col.lower()

            if any(
                word in name
                for word in [
                    "date",
                    "timestamp",
                    "datetime",
                    "month",
                ]
            ):

                converted = pd.to_datetime(
                    df[col],
                    errors="coerce"
                )

                if converted.notna().mean() >= 0.70:
                    df[col] = converted

    return df.reset_index(drop=True)


# ============================================================
# BUSINESS SEMANTICS
# ============================================================

KEYWORDS = {

    "revenue": [
        "revenue",
        "sales revenue",
        "net sales",
        "gross sales",
        "sales amount",
    ],

    "profit": [
        "profit",
        "net profit",
        "gross profit",
        "operating profit",
    ],

    "cost": [
        "cost",
        "expense",
        "spend",
        "expenditure",
    ],

    "price": [
        "price",
        "unit price",
        "selling price",
        "sale price",
    ],

    "quantity": [
        "quantity",
        "qty",
        "units",
        "volume",
        "count",
        "number of units",
    ],

    "sold": [
        "sold",
        "units sold",
        "number sold",
        "sales volume",
    ],

    "inventory": [
        "inventory",
        "stock",
        "hand-in-stock",
        "on hand",
        "stock on hand",
        "ending stock",
        "closing stock",
    ],

    "opening_inventory": [
        "opening stock",
        "opening inventory",
        "beginning inventory",
        "beginning stock",
    ],

    "purchase": [
        "purchase",
        "purchases",
        "stock in",
        "received",
        "receipts",
        "inbound",
    ],

    "product": [
        "product",
        "item",
        "sku",
        "material",
        "part",
        "article",
    ],

    "customer": [
        "customer",
        "client",
        "account",
        "buyer",
    ],

    "region": [
        "region",
        "country",
        "state",
        "city",
        "location",
        "territory",
    ],

    "status": [
        "status",
        "state",
        "result",
        "outcome",
        "condition",
    ],

    "date": [
        "date",
        "timestamp",
        "datetime",
        "month",
        "time",
    ],
}


def semantic_score(column, role):

    name = column.lower()
    score = 0

    for keyword in KEYWORDS.get(role, []):

        if keyword == name:
            score += 12

        elif keyword in name:
            score += 6

    return score


def detect_columns(df):

    numeric = (
        df
        .select_dtypes(
            include=np.number
        )
        .columns
        .tolist()
    )

    dates = (
        df
        .select_dtypes(
            include=["datetime64[ns]"]
        )
        .columns
        .tolist()
    )

    categorical = [
        c
        for c in df.columns
        if c not in numeric
        and c not in dates
    ]

    identifier = []

    for col in df.columns:

        unique_ratio = (
            df[col].nunique(
                dropna=True
            )
            / max(len(df), 1)
        )

        if (
            unique_ratio > 0.95
            and col not in numeric
        ):
            identifier.append(col)

    roles = {}

    for role in KEYWORDS:

        matches = []

        for col in df.columns:

            score = semantic_score(
                col,
                role
            )

            if score > 0:

                matches.append(
                    (
                        col,
                        score
                    )
                )

        matches.sort(
            key=lambda x: x[1],
            reverse=True
        )

        roles[role] = [
            x[0]
            for x in matches
        ]

    roles["numeric"] = numeric
    roles["dates"] = dates
    roles["categorical"] = categorical
    roles["identifier"] = identifier

    return roles


# ============================================================
# PRIMARY MEASURE
# ============================================================

def choose_primary_measure(df, roles):

    priority = [
        "revenue",
        "profit",
        "sold",
        "quantity",
        "inventory",
        "purchase",
        "cost",
        "price",
    ]

    for role in priority:

        for col in roles.get(role, []):

            if col in roles["numeric"]:

                if col not in roles["identifier"]:
                    return col

    candidates = [
        c
        for c in roles["numeric"]
        if c not in roles["identifier"]
    ]

    if candidates:
        return candidates[0]

    return None


def choose_primary_category(df, roles):

    priority = [
        "product",
        "customer",
        "region",
        "status",
    ]

    for role in priority:

        for col in roles.get(role, []):

            if col in roles["categorical"]:

                unique = (
                    df[col]
                    .nunique(
                        dropna=True
                    )
                )

                if 2 <= unique <= 100:
                    return col

    for col in roles["categorical"]:

        if col in roles["identifier"]:
            continue

        unique = (
            df[col]
            .nunique(
                dropna=True
            )
        )

        if 2 <= unique <= 100:
            return col

    return None


# ============================================================
# QUALITY
# ============================================================

def calculate_quality(df):

    cells = max(
        df.shape[0] * df.shape[1],
        1
    )

    missing = int(
        df.isna()
        .sum()
        .sum()
    )

    duplicates = int(
        df.duplicated()
        .sum()
    )

    missing_pct = (
        missing
        / cells
        * 100
    )

    duplicate_pct = (
        duplicates
        / max(len(df), 1)
        * 100
    )

    score = 100

    score -= min(
        missing_pct * 0.75,
        40
    )

    score -= min(
        duplicate_pct * 0.50,
        25
    )

    return {
        "score": max(
            0,
            round(score, 1)
        ),
        "missing": missing,
        "missing_pct": missing_pct,
        "duplicates": duplicates,
        "duplicate_pct": duplicate_pct,
    }


# ============================================================
# NUMBER FORMATTING
# ============================================================

def format_number(value):

    if value is None:
        return "—"

    try:

        if pd.isna(value):
            return "—"

    except Exception:
        pass

    try:
        value = float(value)

    except Exception:
        return str(value)

    if abs(value) >= 1_000_000_000:

        return (
            f"{value / 1_000_000_000:.1f}B"
        )

    if abs(value) >= 1_000_000:

        return (
            f"{value / 1_000_000:.1f}M"
        )

    if abs(value) >= 1_000:

        return (
            f"{value / 1_000:.1f}K"
        )

    if value.is_integer():

        return f"{int(value):,}"

    return f"{value:,.2f}"


def is_money_column(column):

    name = column.lower()

    words = [
        "revenue",
        "sales",
        "profit",
        "cost",
        "expense",
        "price",
        "value",
        "amount",
        "income",
        "budget",
    ]

    return any(
        word in name
        for word in words
    )


def format_metric(value, column):

    if pd.isna(value):
        return "—"

    if is_money_column(column):

        return (
            "$"
            + format_number(value)
        )

    return format_number(value)


# ============================================================
# OUTLIERS
# ============================================================

def detect_outliers(df, numeric_columns):

    results = []

    for col in numeric_columns:

        series = pd.to_numeric(
            df[col],
            errors="coerce"
        ).dropna()

        if len(series) < 5:
            continue

        q1 = series.quantile(0.25)
        q3 = series.quantile(0.75)

        iqr = q3 - q1

        if iqr == 0:
            continue

        lower = q1 - 1.5 * iqr
        upper = q3 + 1.5 * iqr

        mask = (
            (df[col] < lower)
            |
            (df[col] > upper)
        )

        if mask.any():

            temp = df.loc[
                mask
            ].copy()

            temp["_Outlier Field"] = col
            temp["_Outlier Value"] = temp[col]

            results.append(temp)

    if results:

        return pd.concat(
            results,
            ignore_index=True
        )

    return pd.DataFrame()


# ============================================================
# BUSINESS FINDINGS
# ============================================================

def build_findings(
    df,
    roles,
    primary_measure,
    primary_category,
    quality
):

    findings = []
    warnings_list = []

    # --------------------------------------------------------
    # Quality
    # --------------------------------------------------------

    if quality["missing"] > 0:

        findings.append(
            f"{quality['missing']:,} missing cells "
            f"were found ({quality['missing_pct']:.1f}% of all cells)."
        )

    if quality["duplicates"] > 0:

        findings.append(
            f"{quality['duplicates']:,} duplicate rows "
            "were detected."
        )

    # --------------------------------------------------------
    # Primary measure
    # --------------------------------------------------------

    if primary_measure:

        series = pd.to_numeric(
            df[primary_measure],
            errors="coerce"
        ).dropna()

        if len(series):

            total = series.sum()
            average = series.mean()
            minimum = series.min()
            maximum = series.max()

            findings.append(
                f"{primary_measure} totals "
                f"{format_metric(total, primary_measure)}, "
                f"with an average of "
                f"{format_metric(average, primary_measure)} "
                f"per record."
            )

            if maximum > average * 3:

                warnings_list.append(
                    f"{primary_measure} contains unusually "
                    "large observations compared with its average."
                )

    # --------------------------------------------------------
    # Category
    # --------------------------------------------------------

    if (
        primary_category
        and primary_measure
    ):

        temp = df[
            [
                primary_category,
                primary_measure
            ]
        ].copy()

        temp[primary_measure] = pd.to_numeric(
            temp[primary_measure],
            errors="coerce"
        )

        grouped = (
            temp
            .groupby(
                primary_category,
                dropna=False
            )[primary_measure]
            .sum()
            .sort_values(
                ascending=False
            )
        )

        if len(grouped) >= 2:

            top_name = grouped.index[0]
            bottom_name = grouped.index[-1]

            total = grouped.sum()

            if total != 0:

                share = (
                    grouped.iloc[0]
                    / total
                    * 100
                )

                findings.append(
                    f"{top_name} is the largest "
                    f"{primary_category} contributor, "
                    f"representing approximately "
                    f"{share:.1f}% of total {primary_measure}."
                )

                findings.append(
                    f"{bottom_name} has the lowest total "
                    f"{primary_measure} among the analyzed "
                    f"{primary_category} values."
                )

    # --------------------------------------------------------
    # Inventory
    # --------------------------------------------------------

    if (
        roles["sold"]
        and roles["inventory"]
    ):

        sold_col = roles["sold"][0]
        inventory_col = roles["inventory"][0]

        sold = pd.to_numeric(
            df[sold_col],
            errors="coerce"
        ).sum()

        inventory = pd.to_numeric(
            df[inventory_col],
            errors="coerce"
        ).sum()

        if inventory > 0:

            ratio = (
                sold
                / inventory
            )

            findings.append(
                f"Total units sold are "
                f"{ratio:.2f} times the current "
                "inventory level."
            )

    # --------------------------------------------------------
    # Correlation
    # --------------------------------------------------------

    numeric = roles["numeric"]

    if len(numeric) >= 2:

        corr = df[
            numeric
        ].corr(
            numeric_only=True
        )

        strongest = None

        for i in range(len(numeric)):

            for j in range(
                i + 1,
                len(numeric)
            ):

                a = numeric[i]
                b = numeric[j]

                value = corr.loc[
                    a,
                    b
                ]

                if pd.notna(value):

                    candidate = (
                        abs(value),
                        value,
                        a,
                        b
                    )

                    if (
                        strongest is None
                        or candidate[0] > strongest[0]
                    ):

                        strongest = candidate

        if strongest:

            _, value, a, b = strongest

            if abs(value) >= 0.70:

                direction = (
                    "positive"
                    if value > 0
                    else "negative"
                )

                findings.append(
                    f"{a} and {b} show a strong "
                    f"{direction} relationship "
                    f"(correlation {value:.2f})."
                )

    return (
        findings,
        warnings_list
    )


# ============================================================
# GROUP DATA
# ============================================================

def aggregate_data(
    df,
    dimension,
    measure,
    aggregation
):

    temp = df.copy()

    temp[measure] = pd.to_numeric(
        temp[measure],
        errors="coerce"
    )

    if aggregation == "Average":

        grouped = (
            temp
            .groupby(
                dimension,
                dropna=False
            )[measure]
            .mean()
        )

    elif aggregation == "Count":

        grouped = (
            temp
            .groupby(
                dimension,
                dropna=False
            )[measure]
            .count()
        )

    elif aggregation == "Minimum":

        grouped = (
            temp
            .groupby(
                dimension,
                dropna=False
            )[measure]
            .min()
        )

    elif aggregation == "Maximum":

        grouped = (
            temp
            .groupby(
                dimension,
                dropna=False
            )[measure]
            .max()
        )

    else:

        grouped = (
            temp
            .groupby(
                dimension,
                dropna=False
            )[measure]
            .sum()
        )

    return (
        grouped
        .sort_values(
            ascending=False
        )
        .head(30)
        .reset_index()
    )


# ============================================================
# CHART STYLE
# ============================================================

def style_chart(fig):

    fig.update_layout(
        template="plotly_white",
        paper_bgcolor="white",
        plot_bgcolor="white",

        font=dict(
            family="Arial",
            color="#172033"
        ),

        margin=dict(
            l=45,
            r=45,
            t=55,
            b=55
        ),

        hovermode="closest",

        legend=dict(
            bgcolor="rgba(255,255,255,0.85)"
        )
    )

    return fig


# ============================================================
# CHART FUNCTIONS
# ============================================================

def make_bar_chart(
    df,
    category,
    measure,
    horizontal=False
):

    data = aggregate_data(
        df,
        category,
        measure,
        "Sum"
    )

    if horizontal:

        fig = px.bar(
            data,
            x=measure,
            y=category,
            orientation="h",
            text_auto=".2s"
        )

    else:

        fig = px.bar(
            data,
            x=category,
            y=measure,
            text_auto=".2s"
        )

    return style_chart(fig)


def make_pie_chart(
    df,
    category,
    measure,
    donut=False
):

    data = aggregate_data(
        df,
        category,
        measure,
        "Sum"
    ).head(12)

    fig = px.pie(
        data,
        names=category,
        values=measure,
        hole=0.55 if donut else 0
    )

    return style_chart(fig)


def make_histogram(
    df,
    measure
):

    fig = px.histogram(
        df,
        x=measure,
        nbins=30
    )

    return style_chart(fig)


def make_box_plot(
    df,
    measure,
    category=None
):

    if category:

        top = (
            df[category]
            .value_counts()
            .head(10)
            .index
        )

        temp = df[
            df[category].isin(top)
        ]

        fig = px.box(
            temp,
            x=category,
            y=measure,
            points="outliers"
        )

    else:

        fig = px.box(
            df,
            y=measure,
            points="outliers"
        )

    return style_chart(fig)


def make_heatmap(
    df,
    numeric
):

    corr = df[
        numeric
    ].corr(
        numeric_only=True
    )

    fig = px.imshow(
        corr,
        text_auto=".2f",
        aspect="auto"
    )

    return style_chart(fig)


def make_scatter(
    df,
    x,
    y,
    category=None
):

    fig = px.scatter(
        df,
        x=x,
        y=y,
        color=(
            category
            if category
            else None
        )
    )

    return style_chart(fig)


def make_trend(
    df,
    date_col,
    measure,
    aggregation="Sum"
):

    temp = df.copy()

    temp[date_col] = pd.to_datetime(
        temp[date_col],
        errors="coerce"
    )

    temp = temp.dropna(
        subset=[date_col]
    )

    if temp.empty:
        return None

    temp["_Month"] = (
        temp[date_col]
        .dt.to_period("M")
        .dt.to_timestamp()
    )

    temp[measure] = pd.to_numeric(
        temp[measure],
        errors="coerce"
    )

    if aggregation == "Average":

        grouped = (
            temp
            .groupby("_Month")[measure]
            .mean()
            .reset_index()
        )

    elif aggregation == "Count":

        grouped = (
            temp
            .groupby("_Month")[measure]
            .count()
            .reset_index()
        )

    elif aggregation == "Minimum":

        grouped = (
            temp
            .groupby("_Month")[measure]
            .min()
            .reset_index()
        )

    elif aggregation == "Maximum":

        grouped = (
            temp
            .groupby("_Month")[measure]
            .max()
            .reset_index()
        )

    else:

        grouped = (
            temp
            .groupby("_Month")[measure]
            .sum()
            .reset_index()
        )

    fig = px.line(
        grouped,
        x="_Month",
        y=measure,
        markers=True
    )

    return style_chart(fig)


def make_pareto(
    df,
    category,
    measure
):

    data = aggregate_data(
        df,
        category,
        measure,
        "Sum"
    ).head(15)

    total = data[
        measure
    ].sum()

    if total == 0:
        return None

    data["Cumulative %"] = (
        data[measure]
        .cumsum()
        / total
        * 100
    )

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=data[category],
            y=data[measure],
            name=measure
        )
    )

    fig.add_trace(
        go.Scatter(
            x=data[category],
            y=data["Cumulative %"],
            mode="lines+markers",
            name="Cumulative %",
            yaxis="y2"
        )
    )

    fig.update_layout(
        yaxis2=dict(
            title="Cumulative %",
            overlaying="y",
            side="right",
            range=[0, 110]
        )
    )

    return style_chart(fig)


def make_treemap(
    df,
    category,
    measure
):

    data = aggregate_data(
        df,
        category,
        measure,
        "Sum"
    ).head(25)

    fig = px.treemap(
        data,
        path=[category],
        values=measure
    )

    return style_chart(fig)


# ============================================================
# PLOTLY DISPLAY CONFIGURATION
# ============================================================

PLOTLY_CONFIG = {
    "displayModeBar": True,
    "displaylogo": False,
}


def display_plotly(fig):

    if fig is not None:

        st.plotly_chart(
            fig,
            width="stretch",
            config=PLOTLY_CONFIG
        )


# ============================================================
# HERO
# ============================================================

st.markdown(
    """
    <div class="hero">

    <h1>📊 Excel Intelligence Hub</h1>

    <p>
    Interactive business intelligence dashboard for exploring,
    understanding and visualizing Excel data.
    </p>

    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "📁 Upload an Excel workbook",
    type=["xlsx", "xls"]
)


if uploaded_file is None:

    st.markdown(
        """
        <div class="card">

        <h3>Start your analysis</h3>

        <p>
        Upload an Excel workbook to automatically analyze its
        structure, data quality, business measures, categories,
        trends, distributions, relationships and anomalies.
        </p>

        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="footer">

        Excel Intelligence Hub • Business Data Visualization

        </div>
        """,
        unsafe_allow_html=True
    )

    st.stop()


# ============================================================
# LOAD WORKBOOK
# ============================================================

try:

    workbook = pd.ExcelFile(
        uploaded_file
    )

    sheets = workbook.sheet_names

except Exception as e:

    st.error(
        f"Unable to read the Excel workbook: {e}"
    )

    st.stop()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        "## ⚙️ Dashboard Filters"
    )

    selected_sheet = st.selectbox(
        "Worksheet",
        sheets,
        key="sheet_selector"
    )

    st.markdown("---")


# ============================================================
# LOAD SELECTED SHEET
# ============================================================

try:

    raw_df = pd.read_excel(
        uploaded_file,
        sheet_name=selected_sheet
    )

except Exception as e:

    st.error(
        f"Unable to read '{selected_sheet}': {e}"
    )

    st.stop()


df = clean_dataframe(
    raw_df
)


if df.empty:

    st.error(
        "The selected worksheet does not contain usable data."
    )

    st.stop()


# ============================================================
# DETECT STRUCTURE
# ============================================================

roles = detect_columns(
    df
)

primary_measure = (
    choose_primary_measure(
        df,
        roles
    )
)

primary_category = (
    choose_primary_category(
        df,
        roles
    )
)


# ============================================================
# FILTERS
# ============================================================

with st.sidebar:

    filter_columns = [
        c
        for c in roles["categorical"]
        if c not in roles["identifier"]
        and df[c].nunique(
            dropna=True
        ) <= 100
    ]

    if filter_columns:

        for col in filter_columns[:8]:

            values = sorted(
                df[col]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            st.multiselect(
                col,
                values,
                default=[],
                key=f"filter_{col}"
            )


# ============================================================
# APPLY FILTERS
# ============================================================

filtered_df = df.copy()

for col in filter_columns[:8]:

    selected_values = st.session_state.get(
        f"filter_{col}",
        []
    )

    if selected_values:

        filtered_df = filtered_df[
            filtered_df[col]
            .astype(str)
            .isin(
                selected_values
            )
        ]


filtered_df = (
    filtered_df
    .reset_index(drop=True)
)


# Re-detect after filters

roles = detect_columns(
    filtered_df
)

primary_measure = (
    choose_primary_measure(
        filtered_df,
        roles
    )
)

primary_category = (
    choose_primary_category(
        filtered_df,
        roles
    )
)

quality = calculate_quality(
    filtered_df
)

outliers = detect_outliers(
    filtered_df,
    roles["numeric"]
)

findings, warnings_list = (
    build_findings(
        filtered_df,
        roles,
        primary_measure,
        primary_category,
        quality
    )
)


# ============================================================
# OVERVIEW
# ============================================================

st.markdown(
    '<div class="section-title">📌 Dashboard Overview</div>',
    unsafe_allow_html=True
)

st.markdown(
    f"""
    <div class="section-subtitle">

    <b>{uploaded_file.name}</b>

    &nbsp; • &nbsp;

    <b>{selected_sheet}</b>

    &nbsp; • &nbsp;

    {len(filtered_df):,} records after filtering

    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# KPI CARDS
# ============================================================

kpis = [

    (
        "Records",
        format_number(
            len(filtered_df)
        ),
        "Rows currently included"
    ),

    (
        "Columns",
        format_number(
            len(filtered_df.columns)
        ),
        "Fields in the worksheet"
    ),

    (
        "Data Quality",
        f"{quality['score']:.0f}/100",
        "Overall structural score"
    ),

    (
        "Missing Cells",
        format_number(
            quality["missing"]
        ),
        f"{quality['missing_pct']:.1f}% of all cells"
    ),

]

cols = st.columns(4)

for col, item in zip(
    cols,
    kpis
):

    label, value, description = item

    with col:

        st.markdown(
            f"""
            <div class="kpi-card">

            <div class="kpi-label">
            {label}
            </div>

            <div class="kpi-value">
            {value}
            </div>

            <div class="kpi-description">
            {description}
            </div>

            </div>
            """,
            unsafe_allow_html=True
        )


# ============================================================
# DATA STRUCTURE
# ============================================================

st.markdown(
    '<div class="section-title">🧠 Detected Data Structure</div>',
    unsafe_allow_html=True
)

structure = [

    (
        "Primary Measure",
        primary_measure
        or "Not identified"
    ),

    (
        "Primary Dimension",
        primary_category
        or "Not identified"
    ),

    (
        "Date Field",
        roles["dates"][0]
        if roles["dates"]
        else "Not identified"
    ),

    (
        "Numeric Fields",
        str(
            len(
                roles["numeric"]
            )
        )
    ),

]

cols = st.columns(4)

for col, item in zip(
    cols,
    structure
):

    label, value = item

    with col:

        st.markdown(
            f"""
            <div class="card">

            <b>{label}</b>

            <br><br>

            <span style="
            font-size:17px;
            font-weight:700;
            color:#172033;
            ">

            {value}

            </span>

            </div>
            """,
            unsafe_allow_html=True
        )


# ============================================================
# KEY FINDINGS
# ============================================================

st.markdown(
    '<div class="section-title">💡 Key Statistics & Findings</div>',
    unsafe_allow_html=True
)

if findings:

    for finding in findings[:6]:

        st.markdown(
            f"""
            <div class="card">

            <b>•</b> {finding}

            </div>
            """,
            unsafe_allow_html=True
        )


# ============================================================
# TOP CATEGORY ANALYSIS
# ============================================================

if (
    primary_category
    and primary_measure
):

    st.markdown(
        '<div class="section-title">📊 Category Performance</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            f"""
            <div class="card">

            <h3>Top {primary_category} Values</h3>

            <div class="chart-description">

            <b>What this shows:</b>

            The chart ranks the highest-performing
            {primary_category} values according to total
            <b>{primary_measure}</b>. This helps identify
            which categories contribute the most to the dataset.

            </div>
            """,
            unsafe_allow_html=True
        )

        fig = make_bar_chart(
            filtered_df,
            primary_category,
            primary_measure,
            horizontal=True
        )

        display_plotly(fig)

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            f"""
            <div class="card">

            <h3>{primary_measure} Composition</h3>

            <div class="chart-description">

            <b>What this shows:</b>

            This chart shows how the total
            <b>{primary_measure}</b> is distributed across
            the major {primary_category} values. Larger sections
            represent a larger share of the total.

            </div>
            """,
            unsafe_allow_html=True
        )

        fig = make_pie_chart(
            filtered_df,
            primary_category,
            primary_measure,
            donut=True
        )

        display_plotly(fig)

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )


# ============================================================
# PARETO + TREEMAP
# ============================================================

if (
    primary_category
    and primary_measure
):

    st.markdown(
        '<div class="section-title">🔥 Contribution & Concentration</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            f"""
            <div class="card">

            <h3>Pareto Analysis</h3>

            <div class="chart-description">

            <b>What this shows:</b>

            The bars rank categories from largest to smallest,
            while the line shows cumulative contribution.
            This helps determine whether a small number of
            categories account for most of the total
            <b>{primary_measure}</b>.

            </div>
            """,
            unsafe_allow_html=True
        )

        fig = make_pareto(
            filtered_df,
            primary_category,
            primary_measure
        )

        if fig:

            display_plotly(fig)

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            f"""
            <div class="card">

            <h3>Contribution Treemap</h3>

            <div class="chart-description">

            <b>What this shows:</b>

            Each rectangle represents a category and its
            relative contribution to <b>{primary_measure}</b>.
            Larger rectangles indicate greater contribution.

            </div>
            """,
            unsafe_allow_html=True
        )

        fig = make_treemap(
            filtered_df,
            primary_category,
            primary_measure
        )

        display_plotly(fig)

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )


# ============================================================
# DISTRIBUTION
# ============================================================

if primary_measure:

    st.markdown(
        '<div class="section-title">📦 Distribution Analysis</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            f"""
            <div class="card">

            <h3>{primary_measure} Distribution</h3>

            <div class="chart-description">

            <b>What this shows:</b>

            This histogram shows how frequently different
            values of <b>{primary_measure}</b> occur. It helps
            identify the typical range, concentration and
            unusually large or small observations.

            </div>
            """,
            unsafe_allow_html=True
        )

        fig = make_histogram(
            filtered_df,
            primary_measure
        )

        display_plotly(fig)

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            f"""
            <div class="card">

            <h3>Variation & Outliers</h3>

            <div class="chart-description">

            <b>What this shows:</b>

            The box plot summarizes the typical range of
            <b>{primary_measure}</b> and highlights observations
            that fall unusually far from the normal range.

            </div>
            """,
            unsafe_allow_html=True
        )

        fig = make_box_plot(
            filtered_df,
            primary_measure,
            primary_category
        )

        display_plotly(fig)

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )


# ============================================================
# TREND
# ============================================================

if (
    roles["dates"]
    and primary_measure
):

    date_col = roles["dates"][0]

    st.markdown(
        '<div class="section-title">📈 Trend Analysis</div>',
        unsafe_allow_html=True
    )

    aggregation = st.selectbox(
        "Trend statistic",
        [
            "Sum",
            "Average",
            "Count",
            "Minimum",
            "Maximum"
        ],
        key="trend_statistic"
    )

    st.markdown(
        f"""
        <div class="card">

        <h3>{primary_measure} Over Time</h3>

        <div class="chart-description">

        <b>What this shows:</b>

        This line chart tracks the monthly
        <b>{aggregation.lower()}</b> of
        <b>{primary_measure}</b> over time using
        <b>{date_col}</b>. Peaks and declines can reveal
        changes in business activity or performance.

        </div>
        """,
        unsafe_allow_html=True
    )

    fig = make_trend(
        filtered_df,
        date_col,
        primary_measure,
        aggregation
    )

    if fig:

        display_plotly(fig)

    else:

        st.info(
            "There is not enough valid date information to create a trend."
        )

    st.markdown(
        "</div>",
        unsafe_allow_html=True
    )


# ============================================================
# CORRELATION + SCATTER
# ============================================================

if len(roles["numeric"]) >= 2:

    st.markdown(
        '<div class="section-title">🔗 Relationships Between Measures</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        numeric_for_heatmap = (
            roles["numeric"][:15]
        )

        st.markdown(
            """
            <div class="card">

            <h3>Correlation Matrix</h3>

            <div class="chart-description">

            <b>What this shows:</b>

            The matrix measures how strongly numeric fields
            move together. Values closer to +1 indicate a
            strong positive relationship, while values closer
            to -1 indicate a strong negative relationship.

            </div>
            """,
            unsafe_allow_html=True
        )

        fig = make_heatmap(
            filtered_df,
            numeric_for_heatmap
        )

        display_plotly(fig)

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    with col2:

        x_col = st.selectbox(
            "X-axis measure",
            roles["numeric"],
            key="relationship_x"
        )

        y_options = [
            c
            for c in roles["numeric"]
            if c != x_col
        ]

        if y_options:

            y_col = st.selectbox(
                "Y-axis measure",
                y_options,
                key="relationship_y"
            )

            st.markdown(
                f"""
                <div class="card">

                <h3>{x_col} vs {y_col}</h3>

                <div class="chart-description">

                <b>What this shows:</b>

                Each point represents one record. The chart helps
                identify whether higher values of <b>{x_col}</b>
                tend to occur with higher or lower values of
                <b>{y_col}</b>.

                </div>
                """,
                unsafe_allow_html=True
            )

            fig = make_scatter(
                filtered_df,
                x_col,
                y_col,
                primary_category
            )

            display_plotly(fig)

            st.markdown(
                "</div>",
                unsafe_allow_html=True
            )


# ============================================================
# ANOMALIES
# ============================================================

st.markdown(
    '<div class="section-title">🚨 Anomaly Detection</div>',
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="card">

    <div class="chart-description">

    <b>What this shows:</b>

    Potential outliers are observations that fall outside
    the normal statistical range of a numeric field using
    the interquartile range method. These observations
    may deserve additional business review.

    </div>
    """,
    unsafe_allow_html=True
)

if not outliers.empty:

    st.warning(
        f"{len(outliers):,} potential outlier observations detected."
    )

    display_columns = [
        c
        for c in outliers.columns
        if not c.startswith("_")
    ]

    display_columns += [
        "_Outlier Field",
        "_Outlier Value"
    ]

    st.dataframe(
        outliers[
            display_columns
        ].head(150),
        width="stretch"
    )

else:

    st.success(
        "No significant statistical outliers were detected."
    )

st.markdown(
    "</div>",
    unsafe_allow_html=True
)


# ============================================================
# DATA QUALITY
# ============================================================

st.markdown(
    '<div class="section-title">🧹 Data Quality</div>',
    unsafe_allow_html=True
)

if quality["score"] >= 90:

    st.markdown(
        f"""
        <div class="status-good">

        <b>Good data quality:</b>

        The dataset received a quality score of
        {quality['score']:.0f}/100.

        </div>
        """,
        unsafe_allow_html=True
    )

elif quality["score"] >= 75:

    st.markdown(
        f"""
        <div class="status-warning">

        <b>Moderate data quality:</b>

        The dataset received a quality score of
        {quality['score']:.0f}/100.

        Some issues should be reviewed.

        </div>
        """,
        unsafe_allow_html=True
    )

else:

    st.markdown(
        f"""
        <div class="status-danger">

        <b>Data quality concern:</b>

        The dataset received a quality score of
        {quality['score']:.0f}/100.

        Cleaning should be considered before using the
        data for important decisions.

        </div>
        """,
        unsafe_allow_html=True
    )


quality_table = pd.DataFrame(
    {
        "Data Quality Metric": [
            "Rows",
            "Columns",
            "Missing Cells",
            "Missing Percentage",
            "Duplicate Rows",
            "Duplicate Percentage",
            "Quality Score"
        ],

        "Value": [
            f"{len(filtered_df):,}",
            f"{len(filtered_df.columns):,}",
            f"{quality['missing']:,}",
            f"{quality['missing_pct']:.2f}%",
            f"{quality['duplicates']:,}",
            f"{quality['duplicate_pct']:.2f}%",
            f"{quality['score']:.0f}/100"
        ]
    }
)

st.dataframe(
    quality_table,
    width="stretch",
    hide_index=True
)


# ============================================================
# INTERACTIVE ANALYSIS BUILDER
# ============================================================

st.markdown(
    '<div class="section-title">🧩 Analysis Builder</div>',
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="section-subtitle">

    Build additional visualizations by selecting the fields and
    statistical method you want to examine.

    </div>
    """,
    unsafe_allow_html=True
)

dimensions = (
    roles["categorical"]
    + roles["dates"]
)

measures = roles["numeric"]


if dimensions and measures:

    col1, col2, col3 = st.columns(3)

    with col1:

        builder_dimension = st.selectbox(
            "Dimension",
            dimensions,
            key="builder_dimension"
        )

    with col2:

        builder_measure = st.selectbox(
            "Measure",
            measures,
            key="builder_measure"
        )

    with col3:

        builder_aggregation = st.selectbox(
            "Statistic",
            [
                "Sum",
                "Average",
                "Count",
                "Minimum",
                "Maximum"
            ],
            key="builder_aggregation"
        )

    chart_type = st.selectbox(
        "Chart type",
        [
            "Bar",
            "Horizontal Bar",
            "Line",
            "Area",
            "Pie",
            "Donut",
            "Histogram",
            "Box Plot",
            "Scatter",
            "Pareto",
            "Treemap"
        ],
        key="builder_chart_type"
    )

    st.markdown(
        """
        <div class="card">
        """,
        unsafe_allow_html=True
    )

    fig = None

    # --------------------------------------------------------
    # Histogram
    # --------------------------------------------------------

    if chart_type == "Histogram":

        fig = px.histogram(
            filtered_df,
            x=builder_measure,
            nbins=30
        )

    # --------------------------------------------------------
    # Box plot
    # --------------------------------------------------------

    elif chart_type == "Box Plot":

        category = (
            builder_dimension
            if builder_dimension in roles["categorical"]
            else None
        )

        fig = make_box_plot(
            filtered_df,
            builder_measure,
            category
        )

    # --------------------------------------------------------
    # Scatter
    # --------------------------------------------------------

    elif chart_type == "Scatter":

        if len(measures) >= 2:

            scatter_x = st.selectbox(
                "Scatter X-axis",
                measures,
                key="builder_scatter_x"
            )

            scatter_y_options = [
                c
                for c in measures
                if c != scatter_x
            ]

            scatter_y = st.selectbox(
                "Scatter Y-axis",
                scatter_y_options,
                key="builder_scatter_y"
            )

            fig = make_scatter(
                filtered_df,
                scatter_x,
                scatter_y,
                (
                    builder_dimension
                    if builder_dimension in roles["categorical"]
                    else None
                )
            )

        else:

            st.warning(
                "At least two numeric fields are required."
            )

            fig = None

    # --------------------------------------------------------
    # Line
    # --------------------------------------------------------

    elif chart_type == "Line":

        if builder_dimension in roles["dates"]:

            fig = make_trend(
                filtered_df,
                builder_dimension,
                builder_measure,
                builder_aggregation
            )

        else:

            data = aggregate_data(
                filtered_df,
                builder_dimension,
                builder_measure,
                builder_aggregation
            )

            fig = px.line(
                data,
                x=builder_dimension,
                y=builder_measure,
                markers=True
            )

    # --------------------------------------------------------
    # Area
    # --------------------------------------------------------

    elif chart_type == "Area":

        data = aggregate_data(
            filtered_df,
            builder_dimension,
            builder_measure,
            builder_aggregation
        )

        fig = px.area(
            data,
            x=builder_dimension,
            y=builder_measure
        )

    # --------------------------------------------------------
    # Pie / Donut
    # --------------------------------------------------------

    elif chart_type in [
        "Pie",
        "Donut"
    ]:

        data = aggregate_data(
            filtered_df,
            builder_dimension,
            builder_measure,
            builder_aggregation
        )

        fig = px.pie(
            data,
            names=builder_dimension,
            values=builder_measure,
            hole=(
                0.55
                if chart_type == "Donut"
                else 0
            )
        )

    # --------------------------------------------------------
    # Pareto
    # --------------------------------------------------------

    elif chart_type == "Pareto":

        fig = make_pareto(
            filtered_df,
            builder_dimension,
            builder_measure
        )

    # --------------------------------------------------------
    # Treemap
    # --------------------------------------------------------

    elif chart_type == "Treemap":

        fig = make_treemap(
            filtered_df,
            builder_dimension,
            builder_measure
        )

    # --------------------------------------------------------
    # Standard bars
    # --------------------------------------------------------

    else:

        data = aggregate_data(
            filtered_df,
            builder_dimension,
            builder_measure,
            builder_aggregation
        )

        if chart_type == "Horizontal Bar":

            fig = px.bar(
                data,
                x=builder_measure,
                y=builder_dimension,
                orientation="h",
                text_auto=".2s"
            )

        else:

            fig = px.bar(
                data,
                x=builder_dimension,
                y=builder_measure,
                text_auto=".2s"
            )

    if fig is not None:

        fig = style_chart(fig)

        display_plotly(fig)

    st.markdown(
        "</div>",
        unsafe_allow_html=True
    )


# ============================================================
# DATA EXPLORER
# ============================================================

st.markdown(
    '<div class="section-title">🔎 Data Explorer</div>',
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="section-subtitle">

    View the cleaned data currently being used by the dashboard.

    </div>
    """,
    unsafe_allow_html=True
)

st.dataframe(
    filtered_df,
    width="stretch",
    height=450
)


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div class="footer">
        <strong>Excel Intelligence Hub</strong><br>
        Interactive business intelligence and data visualization<br>
        <span>Built by Alina Banari</span>
    </div>
    """,
    unsafe_allow_html=True
)