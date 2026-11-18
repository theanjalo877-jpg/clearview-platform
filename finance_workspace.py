from collections import Counter
from datetime import date, datetime, timedelta
from html import escape
from io import BytesIO, StringIO
from zoneinfo import ZoneInfo
from zipfile import BadZipFile
import csv

import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.utils.exceptions import InvalidFileException

from finance import CLOSED, findings, load, loan_schedule, projection

today = datetime.now(ZoneInfo("Asia/Dubai")).date()



def demo():
    book = Workbook()
    sheet = book.active
    sheet.title = "Customer Invoice Tracker"
    sheet.append(["Client Name", "Brand Select", "Invoice Number", "Total Gross Amount (AED)", "Deposit Paid (AED)", "Due Date"])
    sheet.append(["Studio North", "Demo business", "INV-210", 24000, 6000, today - timedelta(days=17)])
    sheet.append(["Cedar Retail", "Demo business", "INV-211", 14500, 0, today + timedelta(days=20)])
    sheet = book.create_sheet("Supplier Payables")
    sheet.append(["Supplier Name", "Company", "Invoice Number", "Amount Payable Gross (AED)", "Amount Paid (AED)", "Due Date"])
    for _ in range(2):
        sheet.append(["Sample Freight", "Demo business", "SF-72", 3000, 0, today + timedelta(days=4)])
    sheet.append(["Sample Materials", "Demo business", "SM-12", 12000, 0, today + timedelta(days=7)])
    buffer = BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def downloads(headers, rows):
    output = StringIO()
    writer = csv.writer(output)
    writer.writerow(headers)
    for row in rows:
        writer.writerow(["'" + value if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")) else value for value in row])
    return output.getvalue()


def render():
    st.markdown("""
    <style>
    .block-container {max-width: 1260px; padding-top: 3.5rem;}
    .cv-hero {position:relative; overflow:hidden; border-radius:26px; min-height:238px;
     background:linear-gradient(116deg,#20133e,#46308b 75%,#5336ad); color:#fff;
     padding:34px 38px; margin-bottom:20px; display:flex; align-items:center; isolation:isolate;}
    .cv-copy {position:relative; z-index:2; max-width:550px}
    .cv-brand {font-size:.78rem; color:#dfd4ff; font-weight:750; letter-spacing:.16em}
    .cv-title {font-size:clamp(2rem,4vw,3.5rem); line-height:1.04; margin:.45em 0 .25em; font-weight:780}
    .cv-hero p {color:#e8e0ff; font-size:1rem; margin:0}
    .cv-art {position:absolute; right:1%; width:39%; min-width:210px; opacity:.9}
    .cv-line {stroke-dasharray:500; stroke-dashoffset:0; animation:draw 2s ease-out both}
    .cv-orbit {animation:pulse 3s ease-in-out infinite; transform-origin:265px 49px}
    @keyframes draw {from {stroke-dashoffset:500}}
    @keyframes pulse {50% {transform:scale(1.5);opacity:.45}}
    @media(max-width:680px){.cv-hero{padding:24px}.cv-art{opacity:.23;width:58%}}
    @media(prefers-reduced-motion:reduce){.cv-line,.cv-orbit{animation:none}}
    [data-testid="stMetric"] {background:#fff; border:1px solid #e8e4f4; border-radius:16px; padding:14px;}
    </style>
    <div class="cv-hero"><div class="cv-copy"><div class="cv-brand">◈ CLEARVIEW · BUSINESS INTELLIGENCE</div>
    <div class="cv-title">Know where your<br>money stands.</div>
    <p>Find overdue cash, check payments and see when you may need more money.</p></div>
    <svg class="cv-art" viewBox="0 0 320 185" fill="none" aria-hidden="true">
    <rect x="12" y="18" width="294" height="148" rx="22" fill="#ffffff" fill-opacity=".08" stroke="#cbb5ff" stroke-opacity=".35"/>
    <path d="M30 129H288 M30 95H288 M30 61H288" stroke="#e7d9ff" stroke-opacity=".22"/>
    <path d="M35 133L73 115L105 124L143 87L174 96L210 69L235 73L265 49L286 40V151H35Z" fill="#a58bff" fill-opacity=".26"/>
    <path class="cv-line" d="M35 133L73 115L105 124L143 87L174 96L210 69L235 73L265 49L286 40" stroke="#e0ffad" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>
    <circle class="cv-orbit" cx="265" cy="49" r="11" fill="#e0ffad" fill-opacity=".24"/>
    <circle cx="265" cy="49" r="4" fill="#e0ffad"/>
    </svg></div>
    """, unsafe_allow_html=True)

    with st.container(border=True):
        main, example = st.columns([3, 1])
        upload = main.file_uploader("Upload your Excel workbook", type=["xlsx", "xlsm"], help="Supported tables include invoices, supplier bills, expenses and cash movements.")
        sample = example.checkbox("Try sample data", value=False, disabled=upload is not None)
    with st.sidebar:
        st.header("Report settings")
        fallback = st.selectbox("Currency for unlabeled columns", ["AED", "USD", "EUR", "GBP", "SAR", "INR", "CAD", "AUD"])
        as_of = st.date_input("Report date", value=today)
        st.caption("Use one currency per table. Workbook data is processed in this app session.")
    entries, coverage, issues = [], [], []
    if upload is not None or sample:
        try:
            entries, coverage, issues = load(upload.getvalue() if upload is not None else demo(), fallback)
        except (ValueError, TypeError, OSError, BadZipFile, KeyError, InvalidFileException):
            st.error("Clearview could not read this workbook. Save an unprotected Excel copy and check the sheet size before retrying.")
    with st.sidebar:
        companies = sorted({e.company for e in entries})
        company = st.selectbox("Business", companies) if companies else None
        units = sorted({e.currency for e in entries if e.company == company})
        currency = st.selectbox("Report currency", units) if units else fallback
    def format_cash(value):
        return f"{currency} {value:,.2f}"

    scope = [e for e in entries if (e.company, e.currency) == (company, currency)]
    if sample and upload is None:
        st.caption("Sample data · Fictional transactions")
    leaks, forecast, loans, sources = st.tabs(["Cash leaks", "Cash forecast", "Loan planner", "Workbook details"])

    with leaks:
        st.header("What needs your attention?")
        if not scope:
            if coverage and any(kind != "No supported table" for _, kind, _, _ in coverage):
                st.info("Clearview recognized your template. Add invoice, bill, expense or cash movement rows to see findings. Check Workbook details for rows that need correction.")
            elif coverage:
                st.info("Clearview couldn’t identify a supported table. Check Workbook details for the column combinations it can read.")
            else:
                st.info("Upload a workbook or try the sample above.")
        else:
            overdue, results = findings(scope, as_of)
            left, mid, right = st.columns(3)
            left.metric("Overdue customer cash", format_cash(sum(e.amount for e in overdue)))
            mid.metric("Possible duplicate groups", sum(title.startswith("Possible duplicate") for title, *_ in results))
            right.metric("Missing invoice due dates", sum(title.startswith("Missing") for title, *_ in results))
            st.caption("Overdue cash is money to follow up on. Possible duplicates require a check against the original record. The figures cover the selected business and currency.")
            if not results:
                if any(e.kind in {"invoice", "payable", "actual"} for e in scope):
                    st.info("No findings from the available invoice, bill and payment checks.")
                else:
                    st.info("Add customer invoices, supplier bills or actual payments to check for cash leaks.")
            if not any(e.kind == "invoice" for e in scope):
                st.caption("Customer invoice data is needed to measure overdue balances or underbilling.")
            if overdue:
                ageing = [0, 0, 0, 0]
                for entry in overdue:
                    days_late = (as_of - entry.day).days
                    ageing[min((days_late - 1) // 30, 3)] += entry.amount
                chart = go.Figure(go.Bar(
                    x=["1–30 days", "31–60 days", "61–90 days", "91+ days"],
                    y=[float(value) for value in ageing],
                    marker_color=["#9a84ed", "#7553d8", "#5b39bc", "#40267e"],
                    hovertemplate=f"%{{x}}<br>{currency} %{{y:,.2f}}<extra></extra>",
                ))
                chart.update_layout(title="Overdue cash by age", template="plotly_white", height=250,
                                    margin=dict(l=16, r=16, t=43, b=16), yaxis_title=currency)
                st.plotly_chart(chart, config={"displayModeBar": False}, key="overdue-ageing")
            for title, amount, detail, action, source in results[:100]:
                with st.container(border=True):
                    st.subheader(f"{title} · {format_cash(amount)}")
                    st.text(detail)
                    st.markdown(f"**Next step**  {action}")
                    st.caption("Workbook rows · " + escape(source))
            if results:
                if len(results) > 100:
                    st.caption("Showing the first 100 findings. Download the full list below.")
                st.download_button("Download findings", downloads(["Finding", "Amount", "Details", "Next step", "Workbook rows"], results), "clearview-findings.csv", "text/csv")

    with loans:
        st.header("Plan a loan payment")
        a, b, c, d = st.columns(4)
        principal = a.number_input(f"Loan amount ({currency})", min_value=0.0, max_value=1e11, value=0.0, step=1000.0)
        rate = b.number_input("Annual interest rate (%)", min_value=0.0, max_value=100.0, value=8.0, step=0.25, format="%.4f")
        term = c.number_input("Term in months", min_value=1, max_value=480, value=36)
        first = d.date_input("First payment date", value=today + timedelta(days=30), max_value=date(2100, 1, 1))
        schedule = loan_schedule(principal, rate, term, first) if principal else []
        include = st.checkbox("Include this loan in the cash forecast") if schedule else False
        st.caption("Fixed-rate loan with monthly payments on a reducing balance. Enter a separate forecast outflow for any fees or insurance. Confirm actual dates and amounts with the lender.")
        if schedule:
            a, b, c = st.columns(3)
            a.metric("First monthly payment", format_cash(schedule[0][1]))
            b.metric("Total interest", format_cash(sum(r[3] for r in schedule)))
            c.metric("Total repayments", format_cash(sum(r[1] for r in schedule)))
            columns = ["Payment date", "Payment", "Principal", "Interest", "Balance"]
            st.dataframe([dict(zip(columns, r)) for r in schedule], hide_index=True)
            st.download_button("Download payment schedule", downloads(columns, schedule), "clearview-loan.csv", "text/csv")

    with forecast:
        st.header("How much extra cash might you need?")
        a, b, c, d = st.columns(4)
        opening = a.number_input(f"Cash available now ({currency})", value=None, step=1000.0, help="Enter the verified bank and cash balance at the start of the report date.")
        buffer = b.number_input(f"Minimum cash to keep ({currency})", min_value=0.0, value=0.0, step=500.0)
        days = c.selectbox("Forecast length", [30, 60, 90, 180], index=2)
        delay = d.slider("Receipt delay (days)", 0, 60, 14)
        using = st.radio("Use these workbook rows for the forecast", ["Planned cash movements", "Open invoices and supplier bills", "Loan only"], horizontal=True,
                         index=0 if any(e.kind == "forecast" and e.amount for e in scope) else 1)
        types = {"Planned cash movements": {"forecast"}, "Open invoices and supplier bills": {"invoice", "payable"}, "Loan only": set()}[using]
        selected = [e for e in scope if e.kind in types and e.status not in CLOSED and e.amount and (e.kind == "forecast" or e.amount > 0)]
        counts = Counter((e.kind, e.party.casefold(), e.ref.casefold()) for e in selected if e.party and e.ref)
        repeated = {key for key, count in counts.items() if count > 1 and key[0] != "forecast"}
        events = []
        missing = past = duplicate_rows = 0
        for entry in selected:
            if (entry.kind, entry.party.casefold(), entry.ref.casefold()) in repeated:
                duplicate_rows += 1
            elif entry.day is None:
                missing += 1
            elif entry.day < as_of and entry.kind != "payable":
                past += 1
            else:
                events.append((max(entry.day, as_of), -entry.amount if entry.kind == "payable" else entry.amount))
        if include:
            events.extend((due, -amount) for due, amount, *_ in schedule if due >= as_of)
            if types:
                st.caption("Loan payments are added separately. Use workbook rows that exclude this loan to avoid counting it twice.")
        if missing or past or duplicate_rows:
            st.warning(f"Excluded rows · {missing} missing a date · {past} past-due receipts or past planned movements · {duplicate_rows} possible duplicate invoice or bill rows. Correct duplicates and update the expected dates to include them in the next forecast.")
        if any(e.kind == "payable" and e.day and e.day < as_of for e in selected):
            st.caption("Overdue supplier bills with a usable date are assumed payable on the report date.")
        in_window = sum(as_of <= due <= as_of + timedelta(days=days) for due, _ in events)
        if opening is None:
            st.info("Enter your verified cash balance to calculate a funding gap.")
        elif not in_window:
            st.info("No usable cash movements or included loan payments fall within this period. Add dated entries or choose a different forecast period.")
        else:
            rows = projection(events, opening, as_of, days, delay, buffer)
            low_base = min(row[1] for row in rows)
            low_delayed = min(row[2] for row in rows)
            gap = max(0, buffer - min(low_base, opening))
            stress = max(0, buffer - min(low_delayed, opening))
            first_gap = as_of if opening < buffer else next((r[0] for r in rows if r[1] < buffer), None)
            a, b, c = st.columns(3)
            a.metric("Extra cash needed · expected", format_cash(gap))
            b.metric("Extra cash needed · receipts delayed", format_cash(stress))
            c.metric("First day below minimum", first_gap.strftime("%d %b %Y") if first_gap else "None in this period")
            chart = go.Figure()
            for index, label, color, dash in [(1, "Expected cash", "#6244da", "solid"), (2, "Receipts delayed", "#c75b29", "dash"), (3, "Cash minimum", "#677185", "dot")]:
                chart.add_scatter(x=[r[0] for r in rows], y=[r[index] for r in rows], name=label,
                                  mode="lines", line={"color": color, "dash": dash, "shape": "hv"})
            chart.update_layout(template="plotly_white", height=365, hovermode="x unified",
                                margin=dict(l=18, r=18, t=12, b=16), legend=dict(orientation="h", y=-.22),
                                yaxis_title=currency)
            st.plotly_chart(chart, config={"displayModeBar": False})
            if gap:
                st.markdown(f"**Next step**  By {first_gap:%d %b}, check which receipts can arrive sooner or which supplier deadlines can move. Plan for up to **{format_cash(gap)}** in extra cash if those dates hold.")
            st.caption(f"{in_window} cash movements in the period. The chart shows end-of-day balances. The delayed scenario moves receipts {delay} days later and keeps payment dates fixed. Expenses and actual bank activity are reflected through the opening balance or dated forecast entries.")
            st.download_button("Download forecast", downloads(["Date", "Expected cash", "Receipts delayed", "Cash minimum"], rows), "clearview-forecast.csv", "text/csv")

    with sources:
        actual = [e for e in scope if e.kind == "actual" and e.day <= as_of]
        if actual:
            receipts = sum(e.amount for e in actual if e.amount > 0)
            payments = -sum(e.amount for e in actual if e.amount < 0)
            st.header("Cleared cash movements")
            a, b, c = st.columns(3)
            a.metric("Cash received", format_cash(receipts))
            b.metric("Cash paid out", format_cash(payments))
            c.metric("Net cash movement", format_cash(receipts - payments))
            st.caption("Calculated from all dated actual cash rows through the report date. Compare with the bank before using these totals to make a payment decision.")
        st.header("What Clearview read")
        if coverage:
            st.dataframe([dict(zip(["Sheet", "Detected table", "Rows read", "Rows skipped"], row)) for row in coverage], hide_index=True)
        if issues:
            st.subheader("Rows to check")
            st.dataframe([{"Workbook row": source, "What to check": issue} for source, issue in issues], hide_index=True)
        st.markdown("""
    **Customer and supplier balances** need an invoice reference, a customer or supplier name or a clearly named sheet, and an amount. Add a due date and payment status for better results. For underbilling checks, add an **Agreed Gross Amount** column for the agreed value of that specific invoice, including tax on the same basis as the invoice amount.

    **Cash movements** need a date and either an amount with a cash direction, a signed amount, or separate debit and credit columns. Use a forecast sheet for planned movements. Add a reference and counterparty to check for repeated actual payments.

    **Company and currency** are taken from recognizable headers and sheet names when available. Transactions without a company appear under Unassigned. A workbook with totals only, uncached formula-only amounts or ambiguous columns may need a clearer export.
    """)
