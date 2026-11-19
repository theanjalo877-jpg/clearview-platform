from datetime import date, datetime, timedelta
from io import BytesIO, StringIO
from zoneinfo import ZoneInfo
from zipfile import BadZipFile
import csv

import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.utils.exceptions import InvalidFileException

from finance import CLOSED, cash_priorities, funding_gap, load, loan_schedule, matching_references, projection

def demo(today):
    book = Workbook()
    sheet = book.active
    sheet.title = "Customer Invoice Tracker"
    sheet.append(["Client Name", "Brand Select", "Invoice Number", "Total Gross Amount (AED)", "Deposit Paid (AED)", "Due Date"])
    sheet.append(["Studio North", "Demo business", "INV-210", 24000, 6000, today - timedelta(days=17)])
    sheet.append(["Cedar Retail", "Demo business", "INV-211", 14500, 0, today + timedelta(days=20)])
    sheet = book.create_sheet("Supplier Payables")
    sheet.append(["Supplier Name", "Company", "Invoice Number", "Amount Payable Gross (AED)", "Amount Paid (AED)", "Due Date"])
    sheet.append(["Sample Freight", "Demo business", "SF-72", 3000, 0, today + timedelta(days=4)])
    sheet.append(["Sample Materials", "Demo business", "SM-12", 12000, 0, today + timedelta(days=6)])
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


def clear_sample():
    st.session_state["finance_sample"] = False


def render():
    today = datetime.now(ZoneInfo("Asia/Dubai")).date()

    with st.container(border=True):
        main, example = st.columns([3, 1])
        upload = main.file_uploader("Upload your Excel workbook", type=["xlsx", "xlsm"],
                                    help="Supported tables include invoices, supplier bills, expenses and cash movements.",
                                    key="finance_upload", on_change=clear_sample)
        opted_in = example.checkbox("Try sample data", key="finance_sample", disabled=upload is not None)
        sample = opted_in and upload is None
    with st.sidebar:
        st.header("Report settings")
        fallback = st.selectbox("Currency for unlabeled columns", ["AED", "USD", "EUR", "GBP", "SAR", "INR", "CAD", "AUD"])
        as_of = st.date_input("Report date", value=today)
        st.caption("Use one currency per table. Workbook data is processed in this app session.")
    entries, coverage, issues = [], [], []
    if upload is not None or sample:
        try:
            entries, coverage, issues = load(upload.getvalue() if upload is not None else demo(as_of), fallback)
        except (ValueError, TypeError, OSError, BadZipFile, KeyError, InvalidFileException):
            st.error("ClearView could not read this workbook. Save an unprotected Excel copy and check the sheet size before retrying.")
    with st.sidebar:
        companies = sorted({e.company for e in entries})
        company = st.selectbox("Business", companies) if companies else None
        units = sorted({e.currency for e in entries if e.company == company})
        currency = st.selectbox("Report currency", units) if units else fallback
    def format_cash(value):
        return f"{currency} {value:,.2f}"

    scope = [e for e in entries if (e.company, e.currency) == (company, currency)]
    if sample:
        st.caption("Sample data · Fictional transactions")
    groups = matching_references(scope)
    ambiguous = {e.source for group in groups for e in group if e.kind in {"invoice", "payable"}}
    overdue, receipts, bills = cash_priorities(scope, as_of, ambiguous)
    priorities, forecast, loans, sources = st.tabs(["Cash priorities", "Cash forecast", "Loan planner", "Workbook details"])

    with priorities:
        st.header("What should I do this week?")
        if not scope:
            if coverage and any(kind != "No supported table" for _, kind, _, _ in coverage):
                st.info("ClearView recognized your template. Add populated invoice, bill or cash movement rows to see your cash priorities. See Workbook details for rows that need correction.")
            elif coverage:
                st.info("ClearView couldn’t identify a supported table. See Workbook details for the column combinations it can read.")
            else:
                st.info("Upload a workbook or try the sample above.")
        else:
            has_invoices = any(e.kind == "invoice" for e in scope)
            has_bills = any(e.kind == "payable" for e in scope)
            left, mid, right = st.columns(3)
            left.metric("Overdue customer balances", format_cash(sum(e.amount for e in overdue)) if has_invoices else "Unavailable")
            mid.metric("Customer payments due in 7 days", format_cash(sum(e.amount for e in receipts)) if has_invoices else "Unavailable")
            right.metric("Supplier bills due now or in 7 days", format_cash(sum(e.amount for e in bills)) if has_bills else "Unavailable")
            st.caption(f"Seven-day window · {as_of:%d %b} to {as_of + timedelta(days=6):%d %b %Y}. Supplier totals include overdue bills. Undated and ambiguous rows are listed in Workbook details.")
            if not has_invoices and not has_bills:
                st.info("This view needs invoice or supplier bill rows. Use Cash forecast for planned cash movements.")
            actions = [
                ("Follow up", e, "Check for a recent receipt, then ask the customer for a payment date.")
                for e in overdue
            ] + [
                ("Confirm receipt", e, "Confirm the customer’s scheduled payment date.")
                for e in receipts
            ] + [
                ("Review bill", e, "Check approval and bank records before scheduling payment.")
                for e in bills
            ]
            if actions:
                st.subheader("Your action list")
                headers = ["Action", "Customer or supplier", "Reference", "Due date", f"Open balance ({currency})", "Workbook row", "Next step"]
                action_rows = [(action, e.party or "—", e.ref or "—", e.day, float(e.amount), e.source, step)
                               for action, e, step in actions]
                st.dataframe([dict(zip(headers, row)) for row in action_rows[:150]], hide_index=True)
                if len(action_rows) > 150:
                    st.caption("Showing the first 150 items. Download the action list for all rows.")
                st.download_button("Download action list", downloads(headers, action_rows), "clearview-action-list.csv", "text/csv")
            elif has_invoices or has_bills:
                st.info("No dated items in this week’s action list. Check Workbook details for outstanding items that couldn’t be scheduled.")

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
        opening = a.number_input(f"Cash available now ({currency})", value=10000.0 if sample else None,
                                 key=f"opening-{company}-{currency}-{'sample' if sample else 'workbook'}",
                                 step=1000.0, help="Enter the verified bank and cash balance at the start of the report date.")
        buffer = b.number_input(f"Minimum cash to keep ({currency})", min_value=0.0, value=0.0, step=500.0)
        days = c.selectbox("Forecast length", [30, 60, 90, 180], index=2)
        delay = d.slider("Receipt delay (days)", 0, 60, 14)
        using = st.radio("Use these workbook rows for the forecast", ["Planned cash movements", "Open invoices and supplier bills", "Loan only"], horizontal=True,
                         index=0 if any(e.kind == "forecast" and e.amount for e in scope) else 1)
        types = {"Planned cash movements": {"forecast"}, "Open invoices and supplier bills": {"invoice", "payable"}, "Loan only": set()}[using]
        selected = [e for e in scope if e.kind in types and e.status not in CLOSED and e.amount and (e.kind == "forecast" or e.amount > 0)]
        scheduled = []
        missing = past = ambiguous_rows = 0
        for entry in selected:
            if entry.source in ambiguous:
                ambiguous_rows += 1
            elif entry.day is None:
                missing += 1
            elif entry.day < as_of and entry.kind != "payable":
                past += 1
            else:
                scheduled.append((entry, max(entry.day, as_of), -entry.amount if entry.kind == "payable" else entry.amount))
        events = [(due, amount) for _, due, amount in scheduled]
        if include:
            events.extend((due, -amount) for due, amount, *_ in schedule if due >= as_of)
            if types:
                st.caption("Loan payments are added separately. Use workbook rows that exclude this loan to avoid counting it twice.")
        if missing or past or ambiguous_rows:
            st.warning(f"Excluded rows · {missing} missing a date · {past} past-due receipts or past planned movements · {ambiguous_rows} rows sharing an invoice or bill reference. Review the rows in Workbook details and update unresolved dates or records before using this forecast.")
        if any(e.kind == "payable" and e.day and e.day < as_of for e in selected):
            st.caption("Overdue supplier bills with a usable date are assumed payable on the report date.")
        in_window = sum(as_of <= due <= as_of + timedelta(days=days) for due, _ in events)
        if opening is None:
            st.info("Enter your verified cash balance to calculate a funding gap.")
        elif not in_window and not (using == "Open invoices and supplier bills" and overdue):
            st.info("No usable cash movements or included loan payments fall within this period. Add dated entries or choose a different forecast period.")
        else:
            rows = projection(events, opening, as_of, days, delay, buffer)
            gap = funding_gap(rows, opening, buffer)
            stress = funding_gap(rows, opening, buffer, scenario=2)
            first_gap = as_of if opening < buffer else next((r[0] for r in rows if r[1] < buffer), None)
            a, b, c = st.columns(3)
            a.metric("Extra cash needed · expected", format_cash(gap))
            b.metric("Extra cash needed · receipts delayed", format_cash(stress))
            c.metric("First day below minimum", first_gap.strftime("%d %b %Y") if first_gap else "None in this period")
            scenario_rows = None
            end = as_of + timedelta(days=days)
            scheduled_choices = {
                i: (entry, due, amount) for i, (entry, due, amount) in enumerate(scheduled)
                if due <= end and (amount > 0 or due < end)
            }
            if using == "Open invoices and supplier bills":
                scheduled_choices.update({len(scheduled) + i: (e, None, e.amount) for i, e in enumerate(overdue)})
            if scheduled_choices:
                with st.expander("Test a cash decision", expanded=True):
                    st.caption("Try one change to a payment date and compare with the expected forecast. The workbook stays as uploaded.")
                    labels = {}
                    for i, (entry, due, amount) in scheduled_choices.items():
                        task = "Collect overdue" if due is None else "Collect sooner" if amount > 0 else "Pay later"
                        labels[i] = f"{task} · {entry.party or entry.category or 'Cash movement'} · {format_cash(abs(amount))} · {entry.source}"
                    choice = st.selectbox("Which cash movement could change?", list(scheduled_choices), index=None,
                                          format_func=labels.__getitem__, placeholder="Choose a receipt or payment")
                    if choice is not None:
                        _, due, amount = scheduled_choices[choice]
                        earliest = due + timedelta(days=1) if due and amount < 0 else as_of
                        latest = due if due and amount > 0 else end
                        new_date = st.date_input("Revised payment date", min_value=earliest, max_value=latest,
                                                 value=earliest, key=f"decision-{company}-{currency}-{scheduled_choices[choice][0].source}-{earliest}-{latest}")
                        adjusted = events.copy()
                        if due is None:
                            adjusted.append((new_date, amount))
                        else:
                            adjusted[choice] = (new_date, amount)
                        scenario_rows = projection(adjusted, opening, as_of, days, delay, buffer)
                        revised = funding_gap(scenario_rows, opening, buffer)
                        reduction = gap - revised
                        a, b = st.columns(2)
                        a.metric("Extra cash needed with this change", format_cash(revised))
                        b.metric("Reduction in funding gap", format_cash(reduction))
                        if reduction == 0 and gap > 0:
                            st.caption("This change does not reduce the largest shortfall in the period. Try a different date or cash movement.")
                        if due is None:
                            st.caption("This scenario assumes the full overdue balance is collected on the revised date.")
                        elif amount < 0:
                            st.caption("Use a revised supplier date only if it can be agreed.")
            chart = go.Figure()
            for index, label, color, dash in [(1, "Expected cash", "#087a72", "solid"), (2, "Receipts delayed", "#ba562b", "dash"), (3, "Cash minimum", "#687789", "dot")]:
                chart.add_scatter(x=[r[0] for r in rows], y=[r[index] for r in rows], name=label,
                                  mode="lines", line={"color": color, "dash": dash, "shape": "hv"})
            if scenario_rows:
                chart.add_scatter(x=[row[0] for row in scenario_rows], y=[row[1] for row in scenario_rows],
                                  name="With your change", mode="lines",
                                  line={"color": "#5042c2", "width": 3, "shape": "hv"})
            chart.update_layout(template="plotly_white", height=365, hovermode="x unified",
                                margin=dict(l=18, r=18, t=12, b=16), legend=dict(orientation="h", y=-.22),
                                yaxis_title=currency)
            st.plotly_chart(chart, config={"displayModeBar": False})
            if gap:
                st.markdown(f"**Next step**  By {first_gap:%d %b}, confirm which receipts can arrive sooner or whether a supplier can agree to a later date. The expected schedule needs up to **{format_cash(gap)}** in extra cash.")
            elif stress:
                delayed_gap_day = next(row[0] for row in rows if row[2] < buffer)
                st.markdown(f"**Next step**  Confirm customer payment dates. With a {delay}-day delay on receipts, cash first falls below your minimum on **{delayed_gap_day:%d %b}**.")
            st.caption(f"{in_window} cash movements in the period. The chart shows end-of-day balances. The delayed scenario moves receipts {delay} days later and keeps payment dates fixed. Expenses and actual bank activity are reflected through the opening balance or dated forecast entries.")
            headers = ["Date", "Expected cash", "Receipts delayed", "Cash minimum"]
            if scenario_rows:
                headers.append("With your change")
                rows = [(*row, scenario[1]) for row, scenario in zip(rows, scenario_rows)]
            st.download_button("Download forecast", downloads(headers, rows), "clearview-forecast.csv", "text/csv")

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
        st.header("What ClearView read")
        if coverage:
            st.dataframe([dict(zip(["Sheet", "Detected table", "Rows read", "Rows skipped"], row)) for row in coverage], hide_index=True)
        if issues:
            st.subheader("Rows to check")
            st.dataframe([{"Workbook row": source, "What to check": issue} for source, issue in issues], hide_index=True)
        pending = [e for e in scope if e.kind in {"invoice", "payable"} and e.amount > 0
                   and e.status not in CLOSED and e.day is None]
        if pending:
            st.subheader("Open balances missing a due date")
            st.dataframe([{"Customer or supplier": e.party, "Reference": e.ref, f"Open balance ({currency})": float(e.amount),
                           "Workbook row": e.source} for e in pending], hide_index=True)
        if groups:
            st.subheader("Matching references to review")
            st.caption("These rows share identifiers. Check whether they’re separate instalments, repeated entries, or already settled. Matching references alone don’t establish a loss.")
            st.dataframe([{"Record type": e.kind, "Counterparty": e.party, "Reference": e.ref, "Date": e.day,
                           f"Recorded amount ({currency})": float(abs(e.amount)), "Workbook row": e.source}
                          for group in groups for e in group], hide_index=True)
        underbilled = [e for e in scope if e.underbilled and e.source not in ambiguous]
        if underbilled:
            st.subheader("Invoices below their agreed amount")
            st.caption("Compare against the approved price, discounts and changes before raising an additional charge.")
            st.dataframe([{"Customer": e.party, "Reference": e.ref, f"Difference ({currency})": float(e.underbilled),
                           "Workbook row": e.source} for e in underbilled], hide_index=True)
        st.markdown("""
    **Customer and supplier balances** need an invoice reference, a customer or supplier name or a clearly named sheet, and an amount. Add a due date and payment status for better results. For underbilling checks, add an **Agreed Gross Amount** column for the agreed value of that specific invoice, including tax on the same basis as the invoice amount.

    **Cash movements** need a date and either an amount with a cash direction, a signed amount, or separate debit and credit columns. Use a forecast sheet for planned movements. Add a reference and counterparty to check for repeated actual payments.

    **Company and currency** are taken from recognizable headers and sheet names when available. Transactions without a company appear under Unassigned. A workbook with totals only, uncached formula-only amounts or ambiguous columns may need a clearer export.
    """)
