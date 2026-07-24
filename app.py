from __future__ import annotations

from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from streamlit_autorefresh import st_autorefresh


st.set_page_config(
    page_title="TATA Claims Dashboard",
    page_icon="📊",
    layout="wide",
)

TATA_COLUMN = "TATA AIG PLAN"
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
PREFERRED_MASTER_FILE = DATA_DIR / "TATA_Master_Data.xlsx"
GARANTIE_SHEET = "TATA MASTER DATA"
TATA_RAW_SHEET = "Tata Raw Status"


def clean_column_name(value: object) -> str:
    return " ".join(str(value).replace("\n", " ").split()).strip()


def find_column(columns: pd.Index, *names: str) -> str | None:
    lookup = {clean_column_name(col).casefold(): col for col in columns}
    for name in names:
        match = lookup.get(clean_column_name(name).casefold())
        if match is not None:
            return match
    return None


def normalize_sheet_name(value: object) -> str:
    return " ".join(str(value).split()).strip().casefold()


def resolve_sheet_name(
    excel_file: pd.ExcelFile,
    expected_name: str,
    keywords: tuple[str, ...],
) -> str:
    normalized = {
        normalize_sheet_name(sheet): sheet for sheet in excel_file.sheet_names
    }
    expected = normalize_sheet_name(expected_name)

    if expected in normalized:
        return normalized[expected]

    for normalized_name, original_name in normalized.items():
        if all(keyword.casefold() in normalized_name for keyword in keywords):
            return original_name

    available = ", ".join(excel_file.sheet_names)
    raise ValueError(
        f"Could not find the required sheet '{expected_name}'. "
        f"Available sheets: {available}"
    )


def find_master_workbook() -> Path:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    candidates = [
        path
        for path in DATA_DIR.glob("*.xlsx")
        if not path.name.startswith("~$")
    ]
    if PREFERRED_MASTER_FILE in candidates:
        candidates.remove(PREFERRED_MASTER_FILE)
        candidates.insert(0, PREFERRED_MASTER_FILE)

    if not candidates:
        raise FileNotFoundError(
            f"No .xlsx workbook was found inside: {DATA_DIR}"
        )

    checked_files: list[str] = []
    for workbook in candidates:
        try:
            excel_file = pd.ExcelFile(workbook)
            resolve_sheet_name(
                excel_file,
                GARANTIE_SHEET,
                ("tata", "master", "data"),
            )
            resolve_sheet_name(
                excel_file,
                TATA_RAW_SHEET,
                ("tata", "raw", "status"),
            )
            return workbook
        except Exception as exc:
            checked_files.append(f"{workbook.name}: {exc}")

    details = "\n".join(checked_files)
    raise ValueError(
        "No workbook containing both required sheets was found.\n\n"
        f"Required sheets:\n- {GARANTIE_SHEET}\n- {TATA_RAW_SHEET}\n\n"
        f"Checked workbooks:\n{details}"
    )


def read_report(
    source,
    sheet_name: str | None = None,
    sheet_keywords: tuple[str, ...] = (),
) -> pd.DataFrame:
    name = getattr(source, "name", str(source)).lower()
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        excel_file = pd.ExcelFile(source)
        selected_sheet = 0
        if sheet_name:
            selected_sheet = resolve_sheet_name(
                excel_file,
                sheet_name,
                sheet_keywords,
            )
        return pd.read_excel(excel_file, sheet_name=selected_sheet, dtype=str)
    if name.endswith(".csv"):
        return pd.read_csv(source, dtype=str, low_memory=False)
    return pd.read_csv(source, sep="\t", dtype=str, low_memory=False)


def text_series(df: pd.DataFrame, *names: str) -> pd.Series:
    column = find_column(df.columns, *names)
    if column is None:
        return pd.Series("", index=df.index, dtype="string")
    return df[column].fillna("").astype(str).str.strip()


def number_series(df: pd.DataFrame, *names: str) -> pd.Series:
    values = text_series(df, *names)
    values = values.str.replace(",", "", regex=False).str.replace("₹", "", regex=False)
    return pd.to_numeric(values, errors="coerce").fillna(0)


def date_series(df: pd.DataFrame, *names: str) -> pd.Series:
    values = text_series(df, *names).replace(
        {"": pd.NA, "00-01-1900": pd.NA, "01-01-1900": pd.NA}
    )
    return pd.to_datetime(values, errors="coerce", dayfirst=True)


def yes_mask(values: pd.Series) -> pd.Series:
    return values.str.upper().isin({"YES", "Y", "TRUE", "1"})


@st.cache_data(show_spinner=False)
def prepare_report(source, modified_at: float | None = None) -> pd.DataFrame:
    del modified_at
    report = read_report(
        source,
        sheet_name=GARANTIE_SHEET,
        sheet_keywords=("tata", "master", "data"),
    )
    report.columns = [clean_column_name(col) for col in report.columns]

    tata_col = find_column(report.columns, TATA_COLUMN)
    if tata_col is None:
        raise ValueError(
            f"Required column '{TATA_COLUMN}' was not found in the report."
        )

    # Keep the complete master sheet. Some management-report rules are direct
    # Excel COUNTIF/SUMIF equivalents over the whole My Status (CL) column.
    # The TATA-plan filter is applied separately where the Excel rule uses BZ.
    report["_is_tata_plan"] = yes_mask(
        report[tata_col].fillna("").astype(str).str.strip()
    )
    report["_claim_number"] = text_series(report, "Claim Number", "SRN")
    report["_status"] = text_series(report, "My Status", "Actual Status")
    report["_sub_status"] = text_series(report, "Substatus", "Sub Status")
    report["_state"] = text_series(report, "State").replace("", "Not Available")
    report["_service_centre"] = text_series(
        report, "Service Centre Name", "Preferred Service Centre"
    ).replace("", "Not Available")
    report["_claim_date"] = date_series(report, "Claim Received Date")
    report["_payment"] = number_series(report, "Payment Received")
    report["_receivable"] = number_series(
        report, "Amount Receivable - As per TATA Agreement"
    )

    rejected = yes_mask(text_series(report, "Any Rejection"))
    rejected |= report["_status"].str.upper().str.contains("REJECT", na=False)
    rejected |= text_series(report, "Approval Type").str.upper().eq("CLAIM_REJECTED")

    approval_type = text_series(report, "Approval Type").str.upper()
    approved = approval_type.str.contains("APPROVED", na=False)
    approved |= report["_status"].str.upper().isin(
        {"APPROVED", "PARTIALLY APPROVED", "PAYMENT IN PROCESS", "PAID"}
    )
    approved &= ~rejected

    report["_rejected"] = rejected
    report["_approved"] = approved
    report["_payment_received"] = report["_payment"] > 0
    return report


@st.cache_data(show_spinner=False)
def prepare_tata_raw(
    source, modified_at: float | None = None
) -> pd.DataFrame:
    del modified_at
    report = read_report(
        source,
        sheet_name=TATA_RAW_SHEET,
        sheet_keywords=("tata", "raw", "status"),
    )
    report.columns = [clean_column_name(col) for col in report.columns]

    claim_id = find_column(report.columns, "Claim ID")
    garantie_id = find_column(report.columns, "Garantie Claim ID")
    if claim_id is None or garantie_id is None:
        raise ValueError(
            "The TATA raw report must contain 'Claim ID' and 'Garantie Claim ID'."
        )

    report["_tata_claim_id"] = text_series(report, "Claim ID")
    report["_garantie_claim_id"] = text_series(report, "Garantie Claim ID")
    report["_tata_status"] = text_series(report, "My Status")
    report["_task"] = text_series(report, "Task")
    report["_claim_status"] = text_series(report, "Claim Status")
    report["_tata_actual_status"] = text_series(report, "Actual Status")
    report["_tata_sub_status"] = text_series(report, "Sub-Status", "Substatus")
    report["_tata_state"] = text_series(report, "State").replace("", "Not Available")
    report["_tata_group"] = text_series(report, "Groups").replace("", "Not Available")
    report["_receivable_raw"] = number_series(report, "Receivable Amount")
    report["_rejection_date"] = date_series(report, "Rejection Date")

    status_text = (
        report["_tata_actual_status"] + " " + report["_tata_sub_status"]
    ).str.upper()
    report["_raw_rejected"] = status_text.str.contains("REJECT", na=False)
    report["_raw_rejected"] |= report["_rejection_date"].notna()
    report["_raw_approved"] = status_text.str.contains(
        r"APPROVED|PARTIALLY APPROVED|BER", regex=True, na=False
    )
    report["_raw_approved"] &= ~report["_raw_rejected"]
    return report


def money(value: float) -> str:
    return f"₹{value:,.0f}"


def indian_number(value: float) -> str:
    negative = value < 0
    whole, decimal = f"{abs(float(value)):.2f}".split(".")
    if len(whole) > 3:
        last_three = whole[-3:]
        remaining = whole[:-3]
        groups = []
        while remaining:
            groups.append(remaining[-2:])
            remaining = remaining[:-2]
        whole = ",".join(reversed(groups)) + "," + last_three
    result = f"{whole}.{decimal}"
    return f"-{result}" if negative else result


def summary_row(
    label: str,
    frame: pd.DataFrame,
    mask: pd.Series,
    amount_column: str,
) -> dict[str, object]:
    selected = frame.loc[mask]
    return {
        "Status": label,
        "No. of Claims": int(len(selected)),
        "Amount": float(selected[amount_column].fillna(0).sum()),
    }


def format_summary(rows: list[dict[str, object]]) -> pd.DataFrame:
    table = pd.DataFrame(rows)
    display = table.copy()
    display["No. of Claims"] = display["No. of Claims"].map(
        lambda value: f"{int(value):,}"
    )
    display["Amount"] = display["Amount"].map(indian_number)
    return display


def show_summary_table(
    title: str,
    rows: list[dict[str, object]],
    add_total: bool = True,
) -> None:
    if add_total:
        rows = [
            *rows,
            {
                "Status": "Total",
                "No. of Claims": sum(int(row["No. of Claims"]) for row in rows),
                "Amount": sum(float(row["Amount"]) for row in rows),
            },
        ]
    st.subheader(title)
    st.dataframe(format_summary(rows), use_container_width=True, hide_index=True)


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Filters")

    valid_dates = df["_claim_date"].dropna()
    selected_dates = None
    if not valid_dates.empty:
        min_date, max_date = valid_dates.min().date(), valid_dates.max().date()
        selected_dates = st.sidebar.date_input(
            "Claim received date",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
        )

    states = sorted(df["_state"].dropna().unique().tolist())
    selected_states = st.sidebar.multiselect("State", states)

    statuses = sorted(df["_status"].replace("", "Not Available").unique().tolist())
    selected_statuses = st.sidebar.multiselect("My Status", statuses)

    centres = sorted(df["_service_centre"].dropna().unique().tolist())
    selected_centres = st.sidebar.multiselect("Service Centre", centres)

    result = df.copy()
    if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
        start, end = pd.Timestamp(selected_dates[0]), pd.Timestamp(selected_dates[1])
        result = result[
            result["_claim_date"].isna()
            | result["_claim_date"].between(start, end + pd.Timedelta(days=1), inclusive="left")
        ]
    if selected_states:
        result = result[result["_state"].isin(selected_states)]
    if selected_statuses:
        normalized = result["_status"].replace("", "Not Available")
        result = result[normalized.isin(selected_statuses)]
    if selected_centres:
        result = result[result["_service_centre"].isin(selected_centres)]
    return result


st.title("TATA Claims Summary")
st.caption("Only records where TATA AIG PLAN = YES are included.")

try:
    master_file = find_master_workbook()
except Exception as exc:
    st.error(str(exc))
    st.info(
        "Place the master Excel workbook inside the project's data folder. "
        "The workbook filename can be anything, but it must contain both "
        "required worksheets."
    )
    st.stop()

with st.sidebar:
    st.header("Data Source")
    refresh_seconds = st.number_input(
        "Auto-refresh interval (seconds)",
        min_value=0,
        max_value=3600,
        value=60,
        step=30,
        help="Set to 0 to disable automatic refresh.",
    )
    st.caption("Local master workbook")
    st.code(str(master_file))
    if st.button("Refresh data now", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

if refresh_seconds:
    st_autorefresh(interval=int(refresh_seconds * 1000), key="tata_refresh")

source = master_file
modified_at = master_file.stat().st_mtime
tata_raw_file = master_file

try:
    master_data = prepare_report(source, modified_at)
except Exception as exc:
    st.error(f"Could not read the claims report: {exc}")
    st.stop()

data = master_data.loc[master_data["_is_tata_plan"]].copy()

if data.empty:
    st.warning("No rows have TATA AIG PLAN set to YES.")
    st.stop()

filtered = apply_filters(data)

claim_key = filtered["_claim_number"].replace("", pd.NA)
total_claims = int(claim_key.nunique()) if claim_key.notna().any() else len(filtered)
rejections = int(filtered.loc[filtered["_rejected"], "_claim_number"].nunique())
approvals = int(filtered.loc[filtered["_approved"], "_claim_number"].nunique())
payments = int(filtered.loc[filtered["_payment_received"], "_claim_number"].nunique())
payment_amount = filtered["_payment"].sum()
receivable_amount = filtered["_receivable"].sum()
outstanding = max(receivable_amount - payment_amount, 0)

k1, k2, k3, k4 = st.columns(4)
k1.metric("Number of Claims", f"{total_claims:,}")
k2.metric("Approvals", f"{approvals:,}")
k3.metric("Rejections", f"{rejections:,}")
k4.metric("Payment Received", f"{payments:,}", money(payment_amount))

p1, p2, p3 = st.columns(3)
p1.metric("Amount Receivable", money(receivable_amount))
p2.metric("Payment Amount Received", money(payment_amount))
p3.metric("Outstanding Amount", money(outstanding))

left, right = st.columns(2)
with left:
    status_counts = (
        filtered["_status"].replace("", "Not Available").value_counts().reset_index()
    )
    status_counts.columns = ["Status", "Claims"]
    fig = px.bar(
        status_counts,
        x="Status",
        y="Claims",
        title="Claims by Status",
        color="Claims",
        color_continuous_scale="Blues",
    )
    fig.update_layout(coloraxis_showscale=False)
    st.plotly_chart(fig, use_container_width=True)

with right:
    monthly = filtered.dropna(subset=["_claim_date"]).copy()
    monthly["Month"] = monthly["_claim_date"].dt.to_period("M").astype(str)
    monthly = monthly.groupby("Month", as_index=False).size().rename(columns={"size": "Claims"})
    fig = px.line(monthly, x="Month", y="Claims", markers=True, title="Monthly Claim Trend")
    st.plotly_chart(fig, use_container_width=True)

left, right = st.columns(2)
with left:
    state_counts = filtered["_state"].value_counts().head(15).reset_index()
    state_counts.columns = ["State", "Claims"]
    fig = px.bar(
        state_counts,
        x="Claims",
        y="State",
        orientation="h",
        title="Top States by Claims",
        color="Claims",
        color_continuous_scale="Teal",
    )
    fig.update_layout(yaxis={"categoryorder": "total ascending"}, coloraxis_showscale=False)
    st.plotly_chart(fig, use_container_width=True)

with right:
    outcome = pd.DataFrame(
        {
            "Outcome": ["Approved", "Rejected", "Other / Pending"],
            "Claims": [
                approvals,
                rejections,
                max(total_claims - approvals - rejections, 0),
            ],
        }
    )
    fig = px.pie(
        outcome,
        names="Outcome",
        values="Claims",
        hole=0.55,
        title="Claim Outcome",
        color="Outcome",
        color_discrete_map={
            "Approved": "#16a34a",
            "Rejected": "#dc2626",
            "Other / Pending": "#f59e0b",
        },
    )
    st.plotly_chart(fig, use_container_width=True)

st.subheader("Claim Details")
display_columns = [
    col
    for col in [
        find_column(filtered.columns, "Claim Number"),
        find_column(filtered.columns, "Claim Received Date"),
        find_column(filtered.columns, "Customer Name"),
        find_column(filtered.columns, "State"),
        find_column(filtered.columns, "Service Centre Name"),
        find_column(filtered.columns, "My Status"),
        find_column(filtered.columns, "Substatus", "Sub Status"),
        find_column(filtered.columns, "Approval Type"),
        find_column(filtered.columns, "Any Rejection"),
        find_column(filtered.columns, "Payment Received"),
        find_column(filtered.columns, "Amount Receivable - As per TATA Agreement"),
    ]
    if col is not None
]
st.dataframe(filtered[display_columns], use_container_width=True, hide_index=True)

csv_data = filtered[display_columns].to_csv(index=False).encode("utf-8")
st.download_button(
    "Download Filtered Claims",
    data=csv_data,
    file_name="tata_filtered_claims.csv",
    mime="text/csv",
)

if tata_raw_file is not None:
    st.divider()
    st.header("TATA Raw Dashboard Reconciliation")
    try:
        tata_raw = prepare_tata_raw(tata_raw_file, modified_at)
    except Exception as exc:
        st.error(f"Could not read the TATA raw report: {exc}")
    else:
        garantie_map = data.copy()
        garantie_map["_tata_claim_id"] = text_series(
            garantie_map, "TATA AIG CLAIM NO"
        )
        garantie_map["_garantie_claim_id"] = text_series(
            garantie_map, "Claim Number"
        )

        garantie_fields = [
            "_tata_claim_id",
            "_garantie_claim_id",
            "_payment",
            "_payment_received",
            "_status",
            "_sub_status",
        ]
        garantie_map = (
            garantie_map[garantie_fields]
            .sort_values("_payment", ascending=False)
            .drop_duplicates(
                subset=["_tata_claim_id", "_garantie_claim_id"], keep="first"
            )
        )

        reconciled = tata_raw.merge(
            garantie_map,
            on=["_tata_claim_id", "_garantie_claim_id"],
            how="left",
            indicator=True,
        )
        matched = reconciled["_merge"].eq("both")
        raw_total = int(tata_raw["_tata_claim_id"].replace("", pd.NA).nunique())
        raw_approvals = int(
            tata_raw.loc[tata_raw["_raw_approved"], "_tata_claim_id"].nunique()
        )
        raw_rejections = int(
            tata_raw.loc[tata_raw["_raw_rejected"], "_tata_claim_id"].nunique()
        )
        matched_count = int(
            reconciled.loc[matched, "_tata_claim_id"].replace("", pd.NA).nunique()
        )
        matched_payment_count = int(
            reconciled.loc[
                matched & reconciled["_payment_received"].fillna(False),
                "_tata_claim_id",
            ].nunique()
        )
        matched_payment_amount = reconciled.loc[matched, "_payment"].fillna(0).sum()

        r1, r2, r3, r4 = st.columns(4)
        r1.metric("TATA Raw Claims", f"{raw_total:,}")
        r2.metric("TATA Approved", f"{raw_approvals:,}")
        r3.metric("TATA Rejected", f"{raw_rejections:,}")
        r4.metric("Matched with Garantie", f"{matched_count:,}")

        q1, q2, q3 = st.columns(3)
        q1.metric("Matched Payments Received", f"{matched_payment_count:,}")
        q2.metric("Matched Payment Amount", money(matched_payment_amount))
        q3.metric("Unmatched TATA Claims", f"{raw_total - matched_count:,}")

        left, right = st.columns(2)
        with left:
            workflow = (
                tata_raw["_tata_status"]
                .replace("", "Not Available")
                .value_counts()
                .reset_index()
            )
            workflow.columns = ["TATA My Status", "Claims"]
            fig = px.bar(
                workflow,
                x="Claims",
                y="TATA My Status",
                orientation="h",
                title="TATA Workflow Status",
                color="Claims",
                color_continuous_scale="Purples",
            )
            fig.update_layout(
                yaxis={"categoryorder": "total ascending"},
                coloraxis_showscale=False,
            )
            st.plotly_chart(fig, use_container_width=True)

        with right:
            groups = tata_raw["_tata_group"].value_counts().reset_index()
            groups.columns = ["Ageing Group", "Claims"]
            fig = px.bar(
                groups,
                x="Ageing Group",
                y="Claims",
                title="TATA Claims by Ageing Group",
                color="Claims",
                color_continuous_scale="Oranges",
            )
            fig.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("Reconciliation Details")
        reconciliation_view = pd.DataFrame(
            {
                "TATA Claim ID": reconciled["_tata_claim_id"],
                "Garantie Claim ID": reconciled["_garantie_claim_id"],
                "TATA My Status": reconciled["_tata_status"],
                "TATA Actual Status": reconciled["_tata_actual_status"],
                "TATA Sub-Status": reconciled["_tata_sub_status"],
                "State": reconciled["_tata_state"],
                "Receivable Amount": reconciled["_receivable_raw"],
                "Garantie My Status": reconciled["_status"].fillna("Not Matched"),
                "Garantie Sub-Status": reconciled["_sub_status"].fillna("Not Matched"),
                "Payment Received": reconciled["_payment"].fillna(0),
                "Match Result": matched.map(
                    {True: "Matched", False: "Not found in Garantie report"}
                ),
            }
        )
        st.dataframe(reconciliation_view, use_container_width=True, hide_index=True)
        st.download_button(
            "Download Reconciliation",
            data=reconciliation_view.to_csv(index=False).encode("utf-8"),
            file_name="tata_garantie_reconciliation.csv",
            mime="text/csv",
        )

        st.divider()
        st.header("TATA Management Reports")

        task = tata_raw["_task"].str.strip().str.casefold()
        approval_pending = task.eq("awaiting approval")
        payment_pending = task.isin({"complete job order", "create payment"})
        work_in_progress = task.eq("work in progress")
        reopen_cwp = task.eq("send quote")
        reported_to_tata = (
            approval_pending | payment_pending | work_in_progress | reopen_cwp
        )

        current_rows = [
            summary_row(
                "Approval Pending TATA AIG",
                tata_raw,
                approval_pending,
                "_receivable_raw",
            ),
            summary_row(
                "Payments Pending - TATA AIG",
                tata_raw,
                payment_pending,
                "_receivable_raw",
            ),
            summary_row("WIP", tata_raw, work_in_progress, "_receivable_raw"),
            summary_row(
                "To be re-opened (CWP by TATA)",
                master_data,
                master_data["_status"]
                .str.strip()
                .str.casefold()
                .eq("to be re-opened"),
                "_receivable",
            ),
        ]
        show_summary_table("Current TATA Pending Summary", current_rows)

        # Excel-equivalent management calculations:
        # - Total TAGIC Claims uses COUNTIFS(BZ:BZ, "YES") -> `data`
        # - CL status rules use COUNTIF(CL:CL, ...) -> full `master_data`
        master_status = master_data["_status"].str.strip().str.casefold()
        all_tagic_mask = pd.Series(True, index=data.index)
        cwp_statuses = {
            "ok to cwp",
            "ok to cwp - back panel case",
            "ok to cwp - receivable is less than 100",
        }
        due_statuses = {
            "awaiting approval",
            "payment pending",
            "to be re-opened",
            "invoice pending from garantie",
        }

        cwp_mask = master_status.isin(cwp_statuses)
        total_tagic_count = int(len(data))
        total_tagic_amount = float(data["_receivable"].fillna(0).sum())
        cwp_count = int(cwp_mask.sum())
        cwp_amount = float(
            master_data.loc[cwp_mask, "_receivable"].fillna(0).sum()
        )
        paid_mask = master_status.eq("paid")
        paid_count = int(paid_mask.sum())
        paid_amount = float(
            master_data.loc[paid_mask, "_receivable"].fillna(0).sum()
        )
        valid_tagic_count = max(total_tagic_count - cwp_count, 0)
        valid_tagic_amount = max(total_tagic_amount - cwp_amount, 0)

        # Total Due From TATA is a balance, not a status-based count:
        # Total Valid TAGIC Claims - Already Paid / Payment Received From TATA.
        total_due_count = max(valid_tagic_count - paid_count, 0)
        total_due_amount = max(valid_tagic_amount - paid_amount, 0)

        overall_rows = [
            summary_row(
                "Total TAGIC Claims",
                data,
                all_tagic_mask,
                "_receivable",
            ),
            summary_row(
                "Ok to CWP (Cancelled/Duplicate/Rejected)",
                master_data,
                cwp_mask,
                "_receivable",
            ),
            {
                "Status": (
                    "Total Valid TAGIC Claims "
                    "(Excl. Cancelled/Duplicate/Rejected)"
                ),
                "No. of Claims": valid_tagic_count,
                "Amount": valid_tagic_amount,
            },
            {
                "Status": "Already Paid - Payment Received From TATA",
                "No. of Claims": paid_count,
                "Amount": paid_amount,
            },
            {
                "Status": "Total Due From TATA",
                "No. of Claims": total_due_count,
                "Amount": total_due_amount,
            },
        ]
        show_summary_table(
            "Overall TAGIC Claims Summary",
            overall_rows,
            add_total=False,
        )

        reported_to_tata_mask = master_status.isin(due_statuses)
        awaiting_final_documents_mask = master_status.eq("documents pending")
        estimate_pending_mask = master_status.eq("estimate pending")
        pending_upload_mask = master_status.eq("pending for upload")

        final_document_rows = [
            summary_row(
                "Final Documents Received - Reported to TATA AIG",
                master_data,
                reported_to_tata_mask,
                "_receivable",
            ),
            summary_row(
                "Claims Awaiting Final Documents - Approved",
                master_data,
                awaiting_final_documents_mask,
                "_receivable",
            ),
            summary_row(
                "Estimate Pending",
                master_data,
                estimate_pending_mask,
                "_receivable",
            ),
            summary_row(
                "Final Documents Received - Pending for Upload",
                master_data,
                pending_upload_mask,
                "_receivable",
            ),
        ]
        show_summary_table(
            "TATA Due and Final Document Summary",
            final_document_rows,
        )

with st.expander("Metric definitions"):
    st.markdown(
        """
- **Number of Claims:** Unique `Claim Number` records after applying filters.
- **Approvals:** `Approval Type` contains Approved, or `My Status` is an approved/paid state; rejected claims are excluded.
- **Rejections:** `Any Rejection = Yes`, `My Status` contains Rejected, or `Approval Type = CLAIM_REJECTED`.
- **Payment Received:** Claims where the numeric `Payment Received` amount is greater than zero.
- **Outstanding Amount:** Amount Receivable minus Payment Received, with a minimum of zero.
- **TATA reconciliation:** `Claim ID` matches `TATA AIG CLAIM NO` and `Garantie Claim ID` matches `Claim Number`.
- **Total TAGIC Claims:** Full-sheet equivalent of `COUNTIFS(BZ:BZ,"YES")`.
- **OK to CWP:** Full-sheet `My Status (CL)` match for `Ok to CWP`, `Ok to CWP - Back Panel Case`, or `Ok to CWP - Receivable is less than 100`.
- **Already Paid:** Full-sheet `My Status (CL) = Paid`.
- **Total Due From TATA:** `Total Valid TAGIC Claims − Already Paid / Payment Received From TATA`, calculated separately for claim count and amount.
- **Reported to TATA:** Full-sheet `My Status (CL)` match for `Awaiting Approval`, `Payment Pending`, `To be re-opened`, or `Invoice Pending from Garantie`.
- **Final-document categories:** Full-sheet exact `My Status (CL)` matches for `Documents Pending`, `Estimate Pending`, and `Pending for Upload`.
"""
    )
