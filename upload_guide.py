import streamlit as st


def show_guide(workspace):
    with st.expander("How to do it · Prepare your file and read the results"):
        if workspace == "finance":
            st.markdown("""
1. **Choose an Excel workbook.** Upload an unprotected .xlsx or .xlsm file. Use one table per sheet, with column names in the first row and one transaction per row.
2. **Choose the columns for your goal.**
   - **Who owes you and what you owe.** Use clearly named sheets such as *Customer Invoice Tracker* and *Supplier Payables*. Include customer or supplier name, invoice number, gross amount, amount paid, and due date. Add payment status, company, and currency where available.
   - **What cash you’ll need.** Use a *Cash Forecast* sheet with a date and either a signed amount (positive for money in, negative for money out) or an amount plus direction. Use one currency per table. Enter your opening cash balance in Cash forecast.
   - **Possible underbilling.** Include **Agreed Gross Amount** alongside the gross invoice amount, with both amounts on the same tax basis.
   - **Loan payments.** Enter the loan amount, annual rate, term, and first payment date in Loan planner. You can include those payments in the forecast.
3. **Upload or try the sample.** Avoid merged headers, subtotal rows, and blank or uncached formula cells. Check the company, currency, and report date in Report settings.
4. **Check Workbook details first.** Review recognised sheets and excluded rows. Correct the original workbook and upload it again if amounts, dates, or references are missing. Open Cash priorities for next steps.
""")
        else:
            st.markdown("""
1. **Choose your file.** Use .csv, .xlsx, or .xls for a table. Use .docx for an actual Word table or supporting notes. Plain .txt and Word prose appear under Supporting evidence.
2. **Put column names in the first row.** Use one row per order, item, or employee. In Word use **Insert → Table**, with one header row and no merged cells. Use plain numbers and a consistent unit and currency. Leave totals and subtotals out.
3. **Include fields that answer your question.**
   - **Sales or margin:** Order ID, Date, Product, Category, Revenue, Cost, Discount.
   - **Delivery:** Order ID, Date, Delivery Status, Waiting Minutes.
   - **Stock:** SKU, Product, Stock, Reorder Level, Supplier.
   - **Staffing:** Employee ID, Department, Salary, Hours Worked.
4. **Upload or try the sample.** The charts below use recognised columns and numeric rows. A narrative-only document can explain context. Add a table when you want totals, trends, or comparisons.
5. **Check the source.** Review the results for each file and inspect its rows. If ClearView can’t use a field, give the column a clear name, check values and units, then re-upload.
""")


def chart_motion():
    st.html("""
<style>
[class*="st-key-cv-chart-"] [data-testid="stPlotlyChart"] {
  border:1px solid #cddcde; border-radius:16px; padding:8px; overflow:hidden;
  box-shadow:0 10px 22px #132e4216;
  animation:cv-chart-enter .75s cubic-bezier(.2,.74,.26,1) both;
}
@keyframes cv-chart-enter {
  from { opacity:0; transform:perspective(950px) translateY(12px) rotateX(7deg);
         clip-path:inset(0 0 10% 0 round 16px); }
  to { opacity:1; transform:none; clip-path:inset(0 round 16px); }
}
[class*="st-key-cv-chart-"] .barlayer {
  animation:cv-bars-rise .9s .12s cubic-bezier(.2,.74,.26,1) both;
}
[class*="st-key-cv-chart-"] .pielayer {
  animation:cv-pie-focus .9s .15s cubic-bezier(.2,.74,.26,1) both;
  transform-box:fill-box; transform-origin:center;
}
@keyframes cv-bars-rise {
  from {clip-path:inset(100% 0 0 0)} to {clip-path:inset(0)}
}
@keyframes cv-pie-focus {
  from {opacity:0; transform:scale(.87); filter:blur(3px)}
  to {opacity:1; transform:none; filter:none}
}
@media(prefers-reduced-motion:reduce) {
  [class*="st-key-cv-chart-"] [data-testid="stPlotlyChart"],
  [class*="st-key-cv-chart-"] .barlayer,
  [class*="st-key-cv-chart-"] .pielayer { animation:none; }
}
</style>
""")
