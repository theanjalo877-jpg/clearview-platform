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
Add an Agreed Gross Amount column per invoice to enable underbilling checks.
Use Business operations for sales, inventory, workforce and other general analysis.

Excel uploads for Finance are processed in the Streamlit session and are not stored by the application.
