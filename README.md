Anjalo's# clearview-platform


## Run ClearView

```sh
python -m pip install -r requirements.txt
streamlit run app.py
```

Use the Finance workspace for customer and supplier trackers, actual cash movements,
planned cash movements, and loan estimates. Try the fictional sample from the upload panel.
Recognized empty templates are identified without counting formula placeholders as transactions.
Cash projections use an opening balance entered by the owner and one workbook source at a time.
Cash priorities lists overdue customer balances and bills or receipts due within the next seven days.
Use Test a cash decision to compare collecting an overdue balance, bringing a receipt forward,
or rescheduling a supplier payment with the expected funding gap. These scenarios do not edit the workbook.
Matching references are listed for review in Workbook details and carry no inferred loss amount.
Add an Agreed Gross Amount column per invoice to enable underbilling checks.
Use Business operations for sales, inventory, workforce and other general analysis.

Excel uploads for Finance are processed in the Streamlit session and are not stored by the application.
