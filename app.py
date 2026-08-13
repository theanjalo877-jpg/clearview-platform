import streamlit as st
import io
import re
import textwrap
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from docx import Document

# ============================================================
# CLEARVIEW — SINGLE-PAGE BUSINESS INTELLIGENCE WORKSPACE
# ============================================================

st.set_page_config(
    page_title="ClearView — Business Intelligence",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ============================================================
# PREMIUM DESIGN — KEEP THIS VISUAL LANGUAGE CONSISTENT
# ============================================================

st.markdown(
    textwrap.dedent(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');

        :root{
            --ink:#0B1220;
            --muted:#667085;
            --line:#E7EAF0;
            --surface:#FFFFFF;
            --bg:#F5F7FB;
            --blue:#315CF6;
            --violet:#6C7CFF;
            --green:#12B76A;
            --amber:#F79009;
            --red:#F04438;
            --shadow:0 18px 55px rgba(15,23,42,.08);
        }
        *{box-sizing:border-box;}
        html,body,[class*="css"]{font-family:"DM Sans",sans-serif!important;}
        .stApp{
            background:
              radial-gradient(circle at 90% 0%,rgba(49,92,246,.09),transparent 26%),
              radial-gradient(circle at 0% 30%,rgba(108,124,255,.06),transparent 24%),
              var(--bg);
        }
        .block-container{max-width:1420px!important;padding-top:26px!important;padding-bottom:55px!important;}
        #MainMenu,footer,header{visibility:hidden;height:0;}

        .topbar{display:flex;align-items:center;justify-content:space-between;margin-bottom:22px;}
        .brand{display:flex;align-items:center;gap:11px;}
        .brand-mark{
            width:40px;height:40px;border-radius:13px;display:flex;align-items:center;justify-content:center;
            color:#fff;font-family:"Manrope",sans-serif;font-weight:800;
            background:linear-gradient(135deg,#315CF6,#7C6FF6);box-shadow:0 10px 26px rgba(49,92,246,.25);
        }
        .brand-name{font-family:"Manrope",sans-serif;font-size:18px;font-weight:800;color:var(--ink);letter-spacing:-.4px;}
        .brand-pill{padding:8px 13px;border:1px solid var(--line);border-radius:999px;background:rgba(255,255,255,.75);color:#475467;font-size:12px;font-weight:700;}

        .hero{
            position:relative;overflow:hidden;min-height:390px;border-radius:32px;padding:54px 58px;color:#fff;
            background:
              radial-gradient(circle at 83% 25%,rgba(124,111,246,.40),transparent 22%),
              radial-gradient(circle at 100% 100%,rgba(49,92,246,.38),transparent 32%),
              linear-gradient(135deg,#09111F 0%,#111C36 56%,#172A4C 100%);
            box-shadow:0 26px 75px rgba(11,18,32,.18);animation:fadeUp .6s ease both;
        }
        .hero:before{content:"";position:absolute;width:520px;height:520px;right:-180px;top:-190px;border:1px solid rgba(255,255,255,.10);border-radius:50%;box-shadow:0 0 0 70px rgba(255,255,255,.025),0 0 0 140px rgba(255,255,255,.018);}
        .hero-grid{position:relative;z-index:2;display:grid;grid-template-columns:1.35fr .65fr;gap:45px;align-items:center;}
        .eyebrow{display:inline-flex;align-items:center;gap:8px;color:#B9C3FF;font-size:11px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:18px;}
        .eyebrow-dot{width:7px;height:7px;border-radius:50%;background:#7C8CFF;box-shadow:0 0 0 6px rgba(124,140,255,.12);}
        .hero-title{font-family:"Manrope",sans-serif;font-size:clamp(42px,5vw,66px);line-height:1.01;letter-spacing:-3.2px;font-weight:800;margin:0;max-width:820px;}
        .hero-title span{background:linear-gradient(90deg,#A8B5FF,#7CD4FF);-webkit-background-clip:text;background-clip:text;color:transparent;}
        .hero-copy{max-width:760px;color:#C9D1E2;font-size:16px;line-height:1.75;margin-top:22px;}
        .signature{margin-top:25px;color:#AAB3C4;font-size:12px;font-family:"Manrope",sans-serif;letter-spacing:.3px;}
        .signature strong{color:#E9EDFA;font-style:italic;font-size:15px;}
        .hero-visual{min-height:260px;display:flex;align-items:center;justify-content:center;}
        .orbit{position:relative;width:245px;height:245px;border:1px solid rgba(255,255,255,.12);border-radius:50%;animation:spinSlow 18s linear infinite;}
        .orbit:before,.orbit:after{content:"";position:absolute;inset:27px;border:1px solid rgba(255,255,255,.09);border-radius:50%;}
        .orbit:after{inset:59px;}
        .orb-core{position:absolute;inset:82px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-family:"Manrope",sans-serif;font-size:25px;font-weight:800;color:#fff;background:linear-gradient(135deg,#315CF6,#7C6FF6);box-shadow:0 0 50px rgba(108,124,255,.48);animation:pulse 2.8s ease-in-out infinite;}
        .orb-dot{position:absolute;width:12px;height:12px;border-radius:50%;background:#fff;box-shadow:0 0 18px rgba(255,255,255,.75);}
        .d1{top:16px;left:114px}.d2{right:16px;top:112px}.d3{bottom:22px;left:70px}

        .section-head{margin-top:30px;margin-bottom:15px;}
        .section-kicker{color:var(--blue);font-size:11px;font-weight:800;letter-spacing:1.4px;text-transform:uppercase;}
        .section-title{font-family:"Manrope",sans-serif;color:var(--ink);font-size:28px;line-height:1.15;font-weight:800;letter-spacing:-1.1px;margin-top:4px;}
        .section-copy{color:var(--muted);font-size:14px;line-height:1.65;max-width:820px;margin-top:6px;}

        .workflow{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:18px 0 28px;}
        .workflow-card{background:rgba(255,255,255,.78);border:1px solid var(--line);border-radius:18px;padding:18px;transition:transform .22s ease,box-shadow .22s ease,border-color .22s ease;}
        .workflow-card:hover{transform:translateY(-4px);box-shadow:var(--shadow);border-color:#D5DAE5;}
        .workflow-no{font-family:"Manrope",sans-serif;font-size:11px;font-weight:800;color:#98A2B3;}
        .workflow-title{font-family:"Manrope",sans-serif;font-size:15px;font-weight:800;color:var(--ink);margin-top:12px;}
        .workflow-text{color:var(--muted);font-size:12px;line-height:1.55;margin-top:5px;}

        .upload-shell{background:rgba(255,255,255,.88);border:1px solid #DDE2EA;border-radius:25px;padding:25px;box-shadow:0 14px 42px rgba(15,23,42,.06);}
        .upload-heading{display:flex;align-items:center;justify-content:space-between;gap:15px;margin-bottom:15px;}
        .upload-title{font-family:"Manrope",sans-serif;font-size:20px;font-weight:800;color:var(--ink);}
        .upload-subtitle{color:var(--muted);font-size:13px;line-height:1.55;margin-top:4px;}
        [data-testid="stFileUploader"]{margin-top:8px;}
        [data-testid="stFileUploaderDropzone"]{background:linear-gradient(180deg,#FBFCFF,#F7F9FD)!important;border:1.5px dashed #AAB4C7!important;border-radius:18px!important;min-height:155px!important;transition:border-color .2s ease,background .2s ease,transform .2s ease;}
        [data-testid="stFileUploaderDropzone"]:hover{border-color:#6C7CFF!important;background:#F7F8FF!important;}
        .file-row{display:flex;align-items:center;gap:10px;padding:11px 13px;border:1px solid var(--line);background:#fff;border-radius:13px;margin-top:8px;}
        .file-icon{width:30px;height:30px;border-radius:9px;display:flex;align-items:center;justify-content:center;background:#EEF2FF;color:#315CF6;font-weight:800;font-size:12px;}
        .file-name{color:#344054;font-size:12px;font-weight:700;}
        .file-meta{color:#98A2B3;font-size:11px;margin-left:auto;}

        .empty-state{text-align:center;background:rgba(255,255,255,.72);border:1px solid var(--line);border-radius:24px;padding:48px 25px;margin-top:22px;animation:fadeUp .45s ease both;}
        .empty-icon{width:62px;height:62px;border-radius:20px;margin:0 auto 15px;display:flex;align-items:center;justify-content:center;color:#315CF6;font-size:25px;font-weight:800;background:linear-gradient(135deg,#EEF2FF,#F7F4FF);}
        .empty-title{font-family:"Manrope",sans-serif;font-size:20px;font-weight:800;color:var(--ink);}
        .empty-copy{color:var(--muted);font-size:13px;line-height:1.6;max-width:620px;margin:7px auto 0;}

        .metric-card{background:#fff;border:1px solid var(--line);border-radius:18px;padding:18px;min-height:104px;box-shadow:0 8px 25px rgba(15,23,42,.035);}
        .metric-label{color:#667085;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;}
        .metric-value{font-family:"Manrope",sans-serif;color:var(--ink);font-size:29px;font-weight:800;letter-spacing:-1px;margin-top:7px;}
        .metric-note{color:#98A2B3;font-size:11px;margin-top:2px;}

        .insight{background:#fff;border:1px solid var(--line);border-radius:20px;padding:21px;margin:10px 0;box-shadow:0 8px 25px rgba(15,23,42,.035);animation:fadeUp .45s ease both;}
        .insight.problem{border-left:4px solid var(--red)}.insight.solution{border-left:4px solid var(--green)}.insight.suggestion{border-left:4px solid var(--blue)}
        .insight-label{color:#98A2B3;font-size:10px;font-weight:800;letter-spacing:1.1px;text-transform:uppercase;}
        .insight-title{color:var(--ink);font-family:"Manrope",sans-serif;font-size:18px;font-weight:800;margin-top:6px;}
        .insight-body{color:#475467;font-size:13px;line-height:1.65;margin-top:7px;}
        .priority{display:inline-block;padding:5px 9px;border-radius:999px;font-size:10px;font-weight:800;margin-top:12px;background:#F2F4F7;color:#475467;}
        .priority.high{background:#FEF3F2;color:#B42318}.priority.medium{background:#FFFAEB;color:#B54708}.priority.low{background:#ECFDF3;color:#027A48}

        .chart-shell{background:#fff;border:1px solid var(--line);border-radius:22px;padding:18px 20px 10px;box-shadow:0 8px 25px rgba(15,23,42,.035);}
        .chart-title{font-family:"Manrope",sans-serif;font-size:17px;font-weight:800;color:var(--ink);}
        .chart-copy{color:var(--muted);font-size:12px;margin-top:3px;}
        .chart-caption{color:#667085;font-size:12px;margin:8px 0 4px;}

        .footer{margin-top:45px;padding-top:22px;border-top:1px solid var(--line);text-align:center;color:#98A2B3;font-size:11px;line-height:1.7;}
        .footer strong{color:#667085;}
        .stButton>button{border-radius:11px!important;font-family:"DM Sans",sans-serif!important;font-weight:700!important;min-height:40px!important;transition:all .18s ease!important;}
        .stButton>button:hover{transform:translateY(-1px);box-shadow:0 8px 18px rgba(15,23,42,.08);}
        div[data-testid="stMetric"]{display:none;}
        div[data-baseweb="select"]>div{border-radius:11px!important;border-color:#DDE2EA!important;}

        @keyframes fadeUp{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
        @keyframes pulse{0%,100%{transform:scale(1);box-shadow:0 0 50px rgba(108,124,255,.42)}50%{transform:scale(1.06);box-shadow:0 0 70px rgba(108,124,255,.60)}}
        @keyframes spinSlow{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
        @media(max-width:900px){.hero{padding:38px 28px}.hero-grid{grid-template-columns:1fr}.hero-visual{display:none}.workflow{grid-template-columns:1fr 1fr}}
        @media(max-width:600px){.workflow{grid-template-columns:1fr}.hero-title{font-size:42px;letter-spacing:-2px}}
        </style>
        """
    ),
    unsafe_allow_html=True,
)

# ============================================================
# HELPERS
# ============================================================

def html(markup: str):
    st.markdown(textwrap.dedent(markup), unsafe_allow_html=True)


def safe_text(value):
    return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def reset_workspace():
    st.session_state.files = {}
    st.session_state.uploader_key += 1


def file_signature(file):
    return f"{file.name}:{file.size}"


def extract_word_document(file):
    document = Document(file)
    paragraphs = [p.text.strip() for p in document.paragraphs if p.text.strip()]
    tables = []
    for table in document.tables:
        rows = []
        for row in table.rows:
            rows.append([cell.text.strip() for cell in row.cells])
        if rows:
            tables.append(rows)
    return {
        "type": "Document",
        "name": file.name,
        "text": "\n".join(paragraphs),
        "tables": tables,
    }


def read_file(file):
    name = file.name.lower()
    try:
        if name.endswith(".csv"):
            file.seek(0)
            df = pd.read_csv(file)
            return {"type": "Table", "name": file.name, "data": df}

        if name.endswith((".xlsx", ".xls")):
            file.seek(0)
            excel = pd.ExcelFile(file)
            frames = []
            for sheet in excel.sheet_names:
                sheet_df = pd.read_excel(excel, sheet_name=sheet)
                if not sheet_df.empty:
                    sheet_df["__source_sheet"] = sheet
                    frames.append(sheet_df)
            df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
            return {"type": "Table", "name": file.name, "data": df}

        if name.endswith(".docx"):
            file.seek(0)
            return extract_word_document(file)

        if name.endswith(".txt"):
            file.seek(0)
            raw = file.read()
            if isinstance(raw, bytes):
                raw = raw.decode("utf-8", errors="ignore")
            return {"type": "Document", "name": file.name, "text": str(raw), "tables": []}

        return {"type": "Unsupported", "name": file.name, "data": None}
    except Exception as exc:
        return {"type": "Error", "name": file.name, "data": str(exc)}


def clean_columns(df):
    clean = df.copy()
    clean.columns = [str(c).strip() for c in clean.columns]
    return clean


def numeric_columns(df):
    return list(df.select_dtypes(include=np.number).columns)


def normalised_name(name):
    return re.sub(r"[^a-z0-9]+", "_", str(name).lower()).strip("_")


def find_column(df, aliases):
    lookup = {normalised_name(c): c for c in df.columns}
    for alias in aliases:
        key = normalised_name(alias)
        if key in lookup:
            return lookup[key]
    for col in df.columns:
        key = normalised_name(col)
        if any(normalised_name(alias) in key for alias in aliases):
            return col
    return None


SEMANTIC_ALIASES = {
    "revenue": ["revenue", "sales", "sales_aed", "total_sales", "turnover", "income", "net_sales", "amount"],
    "cost": ["cost", "operating_cost", "expense", "expenses", "cogs", "purchase_cost", "total_cost"],
    "profit": ["profit", "net_profit", "gross_profit", "ebit", "ebitda"],
    "quantity": ["quantity", "qty", "units", "units_sold", "quantity_in_stock", "stock", "inventory", "order_quantity"],
    "price": ["price", "unit_price", "total_price", "selling_price", "amount"],
    "customer": ["customer", "customer_id", "client", "client_id", "buyer", "customer_name"],
    "order": ["order", "order_id", "transaction", "transaction_id", "invoice", "invoice_id"],
    "date": ["date", "order_date", "transaction_date", "invoice_date", "created_at", "timestamp"],
    "region": ["region", "area", "location", "territory", "branch", "store"],
    "product": ["product", "product_id", "sku", "item", "item_name", "category"],
    "status": ["status", "order_status", "payment_status", "delivery_status"],
    "wait": ["waiting_minutes", "wait_time", "processing_time", "lead_time", "cycle_time", "response_time"],
    "complaints": ["complaints", "customer_complaints", "returns", "refunds", "issues"],
    "delay": ["delayed_cases", "delayed_orders", "delay", "late_orders", "delivery_delay"],
}


def infer_semantics(df):
    return {key: find_column(df, aliases) for key, aliases in SEMANTIC_ALIASES.items()}


def is_financial_table(df, semantics):
    names = " ".join(normalised_name(c) for c in df.columns)
    financial_words = ["balance", "asset", "liability", "equity", "debit", "credit", "ledger", "account", "journal", "trial_balance", "income_statement", "profit_loss"]
    hits = sum(word in names for word in financial_words)
    return hits >= 2 or (semantics.get("revenue") and semantics.get("cost"))


def best_category(df):
    candidates = []
    for col in df.columns:
        if col == "__source_sheet":
            continue
        series = df[col].dropna()
        if series.empty or pd.api.types.is_numeric_dtype(series):
            continue
        unique = series.nunique()
        if 2 <= unique <= min(15, max(2, len(series) // 2)):
            candidates.append((unique, col))
    candidates.sort()
    return candidates[0][1] if candidates else None


def human_number(value):
    try:
        value = float(value)
        if abs(value) >= 1_000_000:
            return f"{value/1_000_000:.1f}M"
        if abs(value) >= 1_000:
            return f"{value/1_000:.1f}K"
        return f"{value:,.0f}"
    except Exception:
        return str(value)


def financial_findings(df, semantics):
    problems, solutions, suggestions = [], [], []
    revenue_col = semantics.get("revenue")
    cost_col = semantics.get("cost")
    profit_col = semantics.get("profit")

    if revenue_col and cost_col:
        revenue = pd.to_numeric(df[revenue_col], errors="coerce")
        cost = pd.to_numeric(df[cost_col], errors="coerce")
        valid = pd.concat([revenue, cost], axis=1).dropna()
        if not valid.empty:
            rev_total = valid.iloc[:, 0].sum()
            cost_total = valid.iloc[:, 1].sum()
            margin = ((rev_total - cost_total) / rev_total * 100) if rev_total else np.nan
            if pd.notna(margin) and margin < 15:
                problems.append(("Profit margin needs attention", f"Based on the uploaded revenue and cost fields, the calculated margin is approximately {margin:.1f}%.", "high"))
                solutions.append(("Review cost drivers", "Separate fixed, variable and exceptional costs, then identify the categories contributing most to the margin pressure."))
                suggestions.append(("Test margin by product, region or period", "Use the interactive analysis below to locate where the margin is strongest and weakest."))
            else:
                suggestions.append(("Margin appears comparatively healthy", f"The uploaded revenue and cost fields imply an estimated margin of {margin:.1f}%. Validate the business definitions before using it for formal reporting."))

    if profit_col:
        profit = pd.to_numeric(df[profit_col], errors="coerce").dropna()
        if not profit.empty and (profit < 0).any():
            negative_profit = int((profit < 0).sum())
            problems.append(("Loss-making records detected", f"{negative_profit:,} records contain a negative profit value. These may be genuine losses, adjustments or classification issues.", "high"))
            solutions.append(("Trace negative-profit records", "Review the underlying revenue and cost entries for those records and separate genuine losses from posting or classification errors."))

    return problems, solutions, suggestions


def analysis_for_table(df):
    df = clean_columns(df)
    if df.empty:
        return {"rows": 0, "columns": len(df.columns), "missing": 0, "duplicates": 0, "negative": 0, "numeric": [], "semantics": {}, "problems": [], "solutions": [], "suggestions": []}

    nums = numeric_columns(df)
    semantics = infer_semantics(df)
    missing = int(df.isna().sum().sum())
    duplicates = int(df.duplicated().sum())
    negative = sum(int((pd.to_numeric(df[c], errors="coerce") < 0).sum()) for c in nums)

    problems, solutions, suggestions = [], [], []

    if missing:
        pct = missing / max(1, df.size) * 100
        problems.append(("Incomplete information", f"{missing:,} empty cells were found ({pct:.1f}% of available cells). Missing values can distort totals, averages and comparisons.", "high" if pct >= 5 else "medium"))
        solutions.append(("Strengthen data completeness", "Identify which fields are essential, validate them at entry and document acceptable reasons for blanks."))

    if duplicates:
        pct = duplicates / max(1, len(df)) * 100
        problems.append(("Duplicate records", f"{duplicates:,} duplicate rows were detected ({pct:.1f}% of records). This can inflate business totals and counts.", "high" if pct >= 3 else "medium"))
        solutions.append(("Add a unique-record control", "Use a transaction, order, invoice, customer or case identifier and check it before records enter reporting."))

    if negative:
        problems.append(("Negative values require context", f"{negative:,} negative numeric values were found. Some may be valid refunds, credits or accounting adjustments.", "medium"))
        solutions.append(("Separate valid adjustments from errors", "Apply business rules to refunds, credits, returns and accounting adjustments instead of automatically treating every negative as an error."))

    p_fin, s_fin, sug_fin = financial_findings(df, semantics)
    problems.extend(p_fin); solutions.extend(s_fin); suggestions.extend(sug_fin)

    # Operational patterns.
    wait_col = semantics.get("wait")
    complaint_col = semantics.get("complaints")
    delay_col = semantics.get("delay")
    if wait_col:
        wait = pd.to_numeric(df[wait_col], errors="coerce").dropna()
        if not wait.empty:
            avg = wait.mean()
            high = int((wait > wait.quantile(.9)).sum())
            if high and avg > wait.median() * 1.35:
                problems.append(("Service or process time is uneven", f"The average {wait_col} is {avg:.1f}, while the upper 10% of records are materially higher. This suggests a smaller group of cases is creating disproportionate delay.", "medium"))
                solutions.append(("Investigate the slowest cases", f"Review the top 10% of {wait_col} records by customer, process, product, region or date to isolate the bottleneck."))

    if complaint_col and wait_col:
        complaint = pd.to_numeric(df[complaint_col], errors="coerce")
        wait = pd.to_numeric(df[wait_col], errors="coerce")
        valid = pd.concat([wait.rename("wait"), complaint.rename("complaints")], axis=1).dropna()
        if len(valid) >= 10 and valid["wait"].nunique() > 1 and valid["complaints"].nunique() > 1:
            corr = valid["wait"].corr(valid["complaints"])
            if pd.notna(corr) and corr >= .35:
                problems.append(("Longer waits are associated with more complaints", f"The uploaded data shows a positive relationship of approximately {corr:.2f} between {wait_col} and {complaint_col}.", "high"))
                solutions.append(("Reduce the high-delay journey", "Prioritise the process stages and cases with the longest waiting times, then monitor complaint levels after the change."))
                suggestions.append(("Track the relationship over time", "Use the interactive relationship view to see whether the pattern persists across periods or categories."))

    if delay_col and semantics.get("quantity"):
        delay = pd.to_numeric(df[delay_col], errors="coerce")
        if delay.mean(skipna=True) > 0:
            suggestions.append(("Prioritise delayed records", f"Use {delay_col} as a management measure and compare it against volume, product, region or customer segment."))

    # Outlier detection, but only for numeric columns with enough observations.
    for col in nums[:12]:
        s = pd.to_numeric(df[col], errors="coerce").dropna()
        if len(s) < 12 or s.std() == 0:
            continue
        high_out = int((s > s.mean() + 2 * s.std()).sum())
        low_out = int((s < s.mean() - 2 * s.std()).sum())
        if high_out + low_out >= max(3, int(len(s) * .01)):
            problems.append((f"Unusual variation in {col}", f"{high_out + low_out:,} values are more than two standard deviations from the average. These may be exceptional business events or data-quality issues.", "medium"))
            suggestions.append((f"Review exceptional {col} records", "Compare unusual records with dates, categories, customers, regions or transaction types before deciding whether action is required."))

    if not problems:
        suggestions.append(("No major issue detected by the current checks", "ClearView did not find a material problem in the checks it can support from this table. Use the interactive views to investigate business performance further."))

    suggestions.append(("Keep the decision loop consistent", "Review the same key measures regularly, investigate material exceptions and compare performance before and after operational changes."))

    return {
        "rows": len(df),
        "columns": len(df.columns),
        "missing": missing,
        "duplicates": duplicates,
        "negative": negative,
        "numeric": nums,
        "semantics": semantics,
        "problems": problems,
        "solutions": solutions,
        "suggestions": suggestions,
    }


def combine_findings(table_results):
    problems, solutions, suggestions = [], [], []
    for result in table_results:
        a = analysis_for_table(result["data"])
        for title, body, priority in a["problems"]:
            problems.append((title, body, priority, result["name"]))
        for title, body in a["solutions"]:
            solutions.append((title, body, result["name"]))
        for title, body in a["suggestions"]:
            suggestions.append((title, body, result["name"]))

    # Deduplicate while preserving the strongest file-specific evidence.
    def unique(items):
        seen, output = set(), []
        for item in items:
            key = tuple(item[:3])
            if key not in seen:
                seen.add(key); output.append(item)
        return output

    return unique(problems), unique(solutions), unique(suggestions)


def date_column(df, semantics=None):
    semantics = semantics or infer_semantics(df)
    col = semantics.get("date")
    if col:
        parsed = pd.to_datetime(df[col], errors="coerce")
        if parsed.notna().sum() >= 3:
            return col
    return None


def numeric_metric_options(df):
    semantics = infer_semantics(df)
    nums = numeric_columns(df)
    preferred = []
    for key in ["revenue", "cost", "profit", "quantity", "price", "wait", "complaints", "delay"]:
        col = semantics.get(key)
        if col and col in nums and col not in preferred:
            preferred.append(col)
    return preferred + [c for c in nums if c not in preferred]


def category_options(df):
    options = []
    semantics = infer_semantics(df)
    for key in ["product", "customer", "region", "status"]:
        col = semantics.get(key)
        if col and not pd.api.types.is_numeric_dtype(df[col]) and col not in options:
            options.append(col)
    auto = best_category(df)
    if auto and auto not in options:
        options.append(auto)
    return options


def make_chart(df, analysis_mode, metric=None, category=None):
    work = df.copy()
    semantics = infer_semantics(work)
    date_col = date_column(work, semantics)
    nums = numeric_metric_options(work)
    metric = metric if metric in nums else (nums[0] if nums else None)

    if analysis_mode == "Performance trend" and metric:
        if date_col:
            d = work[[date_col, metric]].copy()
            d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
            d[metric] = pd.to_numeric(d[metric], errors="coerce")
            d = d.dropna().sort_values(date_col)
            if not d.empty:
                d = d.groupby(date_col, as_index=False)[metric].sum()
                fig = px.area(d, x=date_col, y=metric, markers=True, template="plotly_white")
                fig.update_traces(line_color="#315CF6", fillcolor="rgba(49,92,246,.15)", hovertemplate="%{x|%d %b %Y}<br>" + metric + ": %{y:,.2f}<extra></extra>")
                return fig

        d = pd.to_numeric(work[metric], errors="coerce").dropna().reset_index(drop=True)
        if not d.empty:
            d = pd.DataFrame({"Record": np.arange(1, len(d) + 1), metric: d})
            fig = px.line(d, x="Record", y=metric, markers=True, template="plotly_white")
            fig.update_traces(line_color="#315CF6", marker_color="#6C7CFF", hovertemplate="Record %{x}<br>" + metric + ": %{y:,.2f}<extra></extra>")
            return fig

    if analysis_mode == "Distribution" and metric:
        clean = pd.to_numeric(work[metric], errors="coerce").dropna()
        if not clean.empty:
            fig = px.histogram(clean, x=metric, nbins=min(30, max(8, int(np.sqrt(len(clean))))), marginal="box", template="plotly_white")
            fig.update_traces(marker_color="#315CF6", hovertemplate=metric + ": %{x:,.2f}<br>Records: %{y}<extra></extra>")
            return fig

    if analysis_mode == "Category performance" and category and metric:
        d = work[[category, metric]].copy()
        d[category] = d[category].fillna("Missing").astype(str)
        d[metric] = pd.to_numeric(d[metric], errors="coerce")
        d = d.dropna().groupby(category, as_index=False)[metric].sum().sort_values(metric, ascending=False).head(12)
        if not d.empty:
            fig = px.bar(d, x=metric, y=category, orientation="h", text_auto=".2s", template="plotly_white")
            fig.update_traces(marker_color="#6C7CFF", hovertemplate="%{y}<br>" + metric + ": %{x:,.2f}<extra></extra>")
            fig.update_layout(yaxis=dict(autorange="reversed"))
            return fig

    if analysis_mode == "Share of total" and category:
        counts = work[category].fillna("Missing").astype(str).value_counts().head(10).reset_index()
        counts.columns = [category, "Records"]
        fig = px.pie(counts, names=category, values="Records", hole=.58, template="plotly_white")
        fig.update_traces(textposition="inside", textinfo="percent+label", hovertemplate="%{label}: %{value} records (%{percent})<extra></extra>")
        return fig

    if analysis_mode == "Relationship" and metric:
        nums = numeric_metric_options(work)
        if len(nums) >= 2:
            x, y = nums[0], nums[1] if nums[1] != metric else (nums[2] if len(nums) > 2 else nums[0])
            d = work[[x, y]].apply(pd.to_numeric, errors="coerce").dropna()
            if not d.empty:
                fig = px.scatter(d, x=x, y=y, trendline="ols" if len(d) >= 20 else None, template="plotly_white")
                fig.update_traces(marker=dict(size=8, color="#315CF6", opacity=.72), hovertemplate=x + ": %{x:,.2f}<br>" + y + ": %{y:,.2f}<extra></extra>")
                return fig

    # Fallback to a useful category chart or empty state.
    if category:
        counts = work[category].fillna("Missing").astype(str).value_counts().head(12).reset_index()
        counts.columns = [category, "Records"]
        fig = px.bar(counts, x="Records", y=category, orientation="h", template="plotly_white")
        fig.update_traces(marker_color="#315CF6")
        fig.update_layout(yaxis=dict(autorange="reversed"))
        return fig
    return None


def style_figure(fig):
    if fig is None:
        return None
    fig.update_layout(
        height=470,
        margin=dict(l=12, r=12, t=20, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="DM Sans", color="#344054"),
        hoverlabel=dict(bgcolor="#0B1220", font_color="#FFFFFF", font_family="DM Sans"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    return fig


def document_summary(doc):
    text = doc.get("text", "")
    words = re.findall(r"\b\w+\b", text)
    return len(words), len(doc.get("tables", []))

# ============================================================
# SESSION STATE
# ============================================================

if "files" not in st.session_state:
    st.session_state.files = {}
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# ============================================================
# BRAND + LANDING
# ============================================================

html(
    """
    <div class="topbar">
        <div class="brand">
            <div class="brand-mark">C</div>
            <div class="brand-name">ClearView</div>
        </div>
        <div class="brand-pill">Business Intelligence</div>
    </div>

    <div class="hero">
        <div class="hero-grid">
            <div>
                <div class="eyebrow"><span class="eyebrow-dot"></span> Business Operations Intelligence</div>
                <div class="hero-title">See what is happening.<br><span>Know what to do next.</span></div>
                <div class="hero-copy">
                    ClearView turns the business information you already have into one clear,
                    interactive decision view — showing what needs attention, why it matters,
                    what can be done, and where the next opportunity may be.
                </div>
                <div class="signature">ATW · independent business intelligence project</div>
            </div>
            <div class="hero-visual">
                <div class="orbit">
                    <div class="orb-core">CV</div>
                    <div class="orb-dot d1"></div><div class="orb-dot d2"></div><div class="orb-dot d3"></div>
                </div>
            </div>
        </div>
    </div>
    """
)

html(
    """
    <div class="section-head">
        <div class="section-kicker">How ClearView works</div>
        <div class="section-title">One workspace. From business files to decisions.</div>
        <div class="section-copy">
            Upload the information your business already has — spreadsheets, CSV exports and Word documents.
            ClearView reads the available structure, identifies measurable signals, tests practical business rules,
            and builds the analysis around the evidence actually found in your files.
        </div>
    </div>
    <div class="workflow">
        <div class="workflow-card"><div class="workflow-no">01</div><div class="workflow-title">Upload</div><div class="workflow-text">Add multiple files from the same business. Nothing is analysed before you upload.</div></div>
        <div class="workflow-card"><div class="workflow-no">02</div><div class="workflow-title">Understand</div><div class="workflow-text">ClearView identifies dates, money, quantities, customers, products, regions and other business signals.</div></div>
        <div class="workflow-card"><div class="workflow-no">03</div><div class="workflow-title">Find the issue</div><div class="workflow-text">The engine checks completeness, duplication, unusual values, financial relationships and operational patterns.</div></div>
        <div class="workflow-card"><div class="workflow-no">04</div><div class="workflow-title">Act</div><div class="workflow-text">Each important finding is paired with a practical solution and a management suggestion.</div></div>
    </div>
    """
)

# ============================================================
# UPLOAD WORKSPACE — MULTIPLE FILES, REMOVE / REPLACE
# ============================================================

html(
    """
    <div class="upload-shell">
        <div class="upload-heading">
            <div>
                <div class="upload-title">Upload your business information</div>
                <div class="upload-subtitle">CSV · XLSX · XLS · DOCX · TXT · multiple files supported</div>
            </div>
        </div>
    """
)

uploaded = st.file_uploader(
    "Choose one or more business files",
    type=["csv", "xlsx", "xls", "docx", "txt"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_key}",
    help="You can upload several files from the same business. ClearView analyses each file separately and also looks for useful relationships across files.",
)

if uploaded:
    for file in uploaded:
        st.session_state.files[file_signature(file)] = file

if st.session_state.files:
    st.markdown("**Files in this workspace**")
    file_keys = list(st.session_state.files.keys())
    for key in file_keys:
        file = st.session_state.files[key]
        c1, c2, c3 = st.columns([.08, .72, .20])
        suffix = Path(file.name).suffix.upper().replace(".", "") or "FILE"
        with c1:
            st.markdown(f"<div class='file-icon'>{safe_text(suffix[:4])}</div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"**{safe_text(file.name)}**<br><span style='color:#98A2B3;font-size:11px'>{file.size/1024:.0f} KB</span>", unsafe_allow_html=True)
        with c3:
            if st.button("Remove", key=f"remove_{key}", use_container_width=True):
                del st.session_state.files[key]
                st.rerun()
    if st.button("Clear all files", key="clear_all", type="secondary"):
        reset_workspace()
        st.rerun()

st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# EMPTY STATE — NOTHING BEFORE UPLOAD
# ============================================================

if not st.session_state.files:
    html(
        """
        <div class="empty-state">
            <div class="empty-icon">◈</div>
            <div class="empty-title">Your business workspace is ready.</div>
            <div class="empty-copy">
                No business data has been uploaded yet. Add your files above and ClearView will build the decision view from those files.
            </div>
        </div>
        <div class="footer">
            ClearView — Business Intelligence<br>
            <strong>ATW</strong> · an independent business analytics project<br>
            © 2026 Anjalo Theophine Wilson · All rights reserved
        </div>
        """
    )
    st.stop()

# ============================================================
# INGESTION
# ============================================================

results = [read_file(f) for f in st.session_state.files.values()]
table_results = [r for r in results if r["type"] == "Table"]
document_results = [r for r in results if r["type"] == "Document"]
errors = [r for r in results if r["type"] in ("Error", "Unsupported")]

# Clean tables once so every later chart/finding uses the same data.
for result in table_results:
    result["data"] = clean_columns(result["data"])

# ============================================================
# EXECUTIVE SNAPSHOT
# ============================================================

total_rows = sum(len(r["data"]) for r in table_results)
total_columns = sum(len(r["data"].columns) for r in table_results)
total_missing = sum(int(r["data"].isna().sum().sum()) for r in table_results)

html(
    """
    <div class="section-head" style="margin-top:30px">
        <div class="section-kicker">Your business snapshot</div>
        <div class="section-title">What ClearView found in the uploaded information.</div>
        <div class="section-copy">This view is generated from the current files. Replace or remove a file and the analysis updates.</div>
    </div>
    """
)

m1, m2, m3, m4 = st.columns(4)
metric_data = [
    ("Files", len(results), "in this workspace"),
    ("Records", f"{total_rows:,}", "rows across tables"),
    ("Fields", f"{total_columns:,}", "columns across tables"),
    ("Documents", len(document_results), "Word / text files"),
]
for col, (label, value, note) in zip((m1, m2, m3, m4), metric_data):
    with col:
        html(f"""
        <div class="metric-card">
            <div class="metric-label">{safe_text(label)}</div>
            <div class="metric-value">{safe_text(value)}</div>
            <div class="metric-note">{safe_text(note)}</div>
        </div>
        """)

# ============================================================
# PROBLEM → SOLUTION → SUGGESTION
# ============================================================

all_problems, all_solutions, all_suggestions = combine_findings(table_results)

html(
    """
    <div class="section-head" style="margin-top:30px">
        <div class="section-kicker">Decision view</div>
        <div class="section-title">Problem → Solution → Suggestion</div>
        <div class="section-copy">The goal is not to overwhelm the client with technical output. It is to turn the evidence into an understandable business decision.</div>
    </div>
    """
)

left, right = st.columns(2, gap="large")
with left:
    html('<div class="section-kicker" style="margin-top:5px;color:#F04438">Problems identified</div>')
    if all_problems:
        for title, body, priority, source in all_problems[:10]:
            html(f"""
            <div class="insight problem">
                <div class="insight-label">Problem · {safe_text(source)}</div>
                <div class="insight-title">{safe_text(title)}</div>
                <div class="insight-body">{safe_text(body)}</div>
                <span class="priority {priority}">{priority.upper()} PRIORITY</span>
            </div>
            """)
    else:
        html("""
        <div class="insight solution">
            <div class="insight-label">No major issue detected</div>
            <div class="insight-title">The available checks look clean.</div>
            <div class="insight-body">ClearView did not find a material issue in the supported checks. This does not mean the business has no problems; it means the uploaded data did not provide enough evidence for one of the rules to trigger.</div>
        </div>
        """)

with right:
    html('<div class="section-kicker" style="margin-top:5px;color:#12B76A">Solutions & suggestions</div>')
    if all_solutions or all_suggestions:
        for title, body, source in all_solutions[:7]:
            html(f"""
            <div class="insight solution">
                <div class="insight-label">Solution · {safe_text(source)}</div>
                <div class="insight-title">{safe_text(title)}</div>
                <div class="insight-body">{safe_text(body)}</div>
            </div>
            """)
        for title, body, source in all_suggestions[:5]:
            html(f"""
            <div class="insight suggestion">
                <div class="insight-label">Suggestion · {safe_text(source)}</div>
                <div class="insight-title">{safe_text(title)}</div>
                <div class="insight-body">{safe_text(body)}</div>
            </div>
            """)

# ============================================================
# DATA-AWARE INTERACTIVE ANALYTICS
# ============================================================

if table_results:
    html(
        """
        <div class="section-head" style="margin-top:34px">
            <div class="section-kicker">Interactive intelligence</div>
            <div class="section-title">Explore the business from different angles.</div>
            <div class="section-copy">The options below are generated from the uploaded data. ClearView does not force the same fixed chart onto every business.</div>
        </div>
        """
    )

    for idx, result in enumerate(table_results):
        df = result["data"]
        nums = numeric_metric_options(df)
        cats = category_options(df)
        default_metric = nums[0] if nums else None
        default_category = cats[0] if cats else None

        html(f"""
        <div class="chart-shell">
            <div class="chart-title">{safe_text(result['name'])}</div>
            <div class="chart-copy">{len(df):,} records · {len(df.columns):,} fields · analysis adapts to this file's structure</div>
        </div>
        """)

        c1, c2, c3 = st.columns([1.45, 1.45, 1.45])
        with c1:
            analysis_mode = st.selectbox(
                "Business view",
                ["Performance trend", "Category performance", "Distribution", "Share of total", "Relationship"],
                key=f"mode_{idx}_{result['name']}",
                help="Choose the question you want the chart to answer rather than choosing a raw database field.",
            )
        with c2:
            metric = st.selectbox(
                "Measure",
                nums if nums else ["No measurable numeric field detected"],
                index=0,
                key=f"metric_{idx}_{result['name']}",
                disabled=not bool(nums),
            )
        with c3:
            category = st.selectbox(
                "Break down by",
                cats if cats else ["No useful category detected"],
                index=0,
                key=f"category_{idx}_{result['name']}",
                disabled=not bool(cats),
            )

        fig = make_chart(
            df,
            analysis_mode,
            metric=metric if nums else None,
            category=category if cats else None,
        )
        fig = style_figure(fig)
        if fig is not None:
            st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False, "responsive": True, "scrollZoom": True})

        # Explain what the selected view means in plain business language.
        view_explanations = {
            "Performance trend": "Shows how the selected measure changes across records or time. Use it to spot growth, decline and sudden changes.",
            "Category performance": "Compares the selected measure across the strongest categories detected in the file.",
            "Distribution": "Shows where the values concentrate and highlights unusually high or low observations.",
            "Share of total": "Shows how the records are distributed across a meaningful category detected in the uploaded data.",
            "Relationship": "Compares two measurable fields to see whether they move together. A relationship is not proof of causation.",
        }
        st.markdown(f"<div class='chart-caption'><b>{safe_text(analysis_mode)}:</b> {safe_text(view_explanations[analysis_mode])}</div>", unsafe_allow_html=True)

        with st.expander(f"Inspect {result['name']} data", expanded=False):
            st.dataframe(df, use_container_width=True, height=330)

# ============================================================
# DOCUMENT EVIDENCE
# ============================================================

if document_results:
    html(
        """
        <div class="section-head" style="margin-top:34px">
            <div class="section-kicker">Supporting evidence</div>
            <div class="section-title">Documents are part of the business story.</div>
            <div class="section-copy">Word and text files are extracted and surfaced instead of being treated as irrelevant uploads.</div>
        </div>
        """
    )
    for doc in document_results:
        words, tables = document_summary(doc)
        html(f"""
        <div class="insight suggestion">
            <div class="insight-label">Document reviewed</div>
            <div class="insight-title">{safe_text(doc['name'])}</div>
            <div class="insight-body">ClearView extracted approximately <strong>{words:,}</strong> words and <strong>{tables:,}</strong> table(s). The extracted evidence can be inspected below.</div>
        </div>
        """)
        with st.expander(f"Read extracted content · {doc['name']}"):
            st.text_area("Extracted text", doc.get("text", "")[:30000], height=260, key=f"text_{doc['name']}")
            if doc.get("tables"):
                st.markdown("**Tables found in this document**")
                for table_index, table in enumerate(doc["tables"]):
                    st.dataframe(pd.DataFrame(table), use_container_width=True, height=220, key=f"doc_table_{doc['name']}_{table_index}")

# ============================================================
# FILE HANDLING NOTES
# ============================================================

if errors:
    html("""
    <div class="insight problem" style="margin-top:18px">
        <div class="insight-label">File handling</div>
        <div class="insight-title">Some files could not be analysed.</div>
        <div class="insight-body">Remove the affected file above and replace it with a readable CSV, Excel, Word or text file.</div>
    </div>
    """)
    for error in errors:
        st.caption(f"{error['name']}: {error.get('data', 'Unsupported file type')}")

# ============================================================
# FOOTER
# ============================================================

html(
    """
    <div class="footer">
        ClearView — Business Intelligence<br>
        <strong>ATW</strong> · an independent business analytics project<br>
        © 2026 Anjalo Theophine Wilson · All rights reserved
    </div>
    """
)
