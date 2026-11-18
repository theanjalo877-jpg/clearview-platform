from calendar import monthrange
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from io import BytesIO
from zipfile import ZipFile
import re

from openpyxl import load_workbook


@dataclass
class Entry:
    kind: str
    company: str
    currency: str
    source: str
    ref: str
    party: str
    day: date | None
    amount: Decimal
    paid: Decimal | None
    status: str
    category: str
    underbilled: Decimal = Decimal(0)


ALIASES = {
    'gross': ('total gross amount', 'amount payable gross', 'invoice amount', 'invoice total', 'gross amount'),
    'agreed': ('agreed gross amount', 'agreed invoice amount', 'contracted gross amount'),
    'paid': ('deposit paid', 'amount paid', 'amount received', 'payments received'),
    'balance': ('balance outstanding', 'balance due', 'outstanding balance', 'outstanding amount'),
    'amount': ('amount', 'total paid gross', 'transaction amount'),
    'flow': ('cash flow', 'direction', 'transaction type'),
    'signed': ('signed amount', 'signed planned', 'net cash movement'),
    'date': ('date', 'expected date', 'transaction date'),
    'due': ('due date', 'payment due date'),
    'party': ('client name', 'customer name', 'supplier name', 'customer / supplier', 'customer', 'supplier', 'vendor'),
    'ref': ('invoice number', 'invoice no', 'reference / invoice', 'reference', 'voucher no'),
    'company': ('brand select', 'company', 'business unit', 'brand'),
    'status': ('payment status', 'probability / status', 'status'),
    'category': ('category', 'description'),
    'debit': ('debit', 'money out', 'withdrawal'),
    'credit': ('credit', 'money in', 'deposit'),
}
LABELS = {label: key for key, labels in ALIASES.items() for label in labels}
CLOSED = {'paid', 'fully paid', 'completed', 'settled', 'received', 'cancelled', 'canceled', 'void', 'refunded'}
BRANDS = {'ANT': 'Antarria', 'TDH': 'The Design House'}


def clean(value):
    return str(value).strip() if value is not None else ''


def money(value):
    value = clean(value).replace(',', '')
    value = re.sub(r'^(AED|USD|EUR|GBP|[$€£])\s*', '', value, flags=re.I)
    if value.startswith('(') and value.endswith(')'):
        value = '-' + value[1:-1]
    try:
        result = Decimal(value)
        return result if result.is_finite() and abs(result) < 10**14 else None
    except InvalidOperation:
        return None


def day_of(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for pattern in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
        try:
            return datetime.strptime(clean(value), pattern).date()
        except ValueError:
            pass
    return None


def layout(title, row):
    text = title.lower()
    fields = {}
    currencies = set()
    for i, value in enumerate(row):
        label = clean(value).lower()
        unit = re.search(r'\((aed|usd|eur|gbp|sar|inr|cad|aud)\)\s*$', label)
        key = LABELS.get(re.sub(r'\s*\([^)]*\)\s*$', '', label))
        if key:
            fields.setdefault(key, i)
            if unit and key in {'gross', 'paid', 'balance', 'amount', 'signed', 'debit', 'credit', 'agreed'}:
                currencies.add(unit[1].upper())
    if ('balance' in fields or 'gross' in fields) and 'ref' in fields:
        party_key = clean(row[fields['party']]).lower() if 'party' in fields else ''
        if 'supplier' in text or 'payable' in text or party_key in {'supplier', 'supplier name', 'vendor'}:
            return 'payable', fields, currencies
        if 'customer' in text or 'receivable' in text or party_key in {'client name', 'customer', 'customer name'}:
            return 'invoice', fields, currencies
    if 'date' in fields and (
        ('amount' in fields and 'flow' in fields) or 'signed' in fields or {'debit', 'credit'} <= fields.keys()
    ):
        return ('forecast' if 'forecast' in text or 'expected date' in {clean(v).lower() for v in row} else 'actual'), fields, currencies
    if 'expense' in text and {'date', 'amount'} <= fields.keys():
        return 'expense', fields, currencies
    return None


def load(data, fallback_currency):
    with ZipFile(BytesIO(data)) as archive:
        if sum(item.file_size for item in archive.infolist()) > 60_000_000:
            raise ValueError('Workbook expands beyond 60 MB.')
    workbook = load_workbook(BytesIO(data), read_only=True, data_only=True, keep_links=False)
    entries, coverage, issues = [], [], []
    try:
        for sheet in workbook:
            if (sheet.max_row or 0) > 100_000 or (sheet.max_column or 0) > 256:
                raise ValueError('Use sheets with at most 100,000 rows and 256 columns.')
            found = None
            accepted = skipped = 0
            for index, row in enumerate(sheet.values, 1):
                if index > 100_000 or len(row) > 256:
                    raise ValueError("Use sheets with at most 100,000 rows and 256 columns.")
                if found is None:
                    if index > 35:
                        break
                    found = layout(sheet.title, row)
                    continue
                kind, fields, currencies = found

                def get(key):
                    return row[fields[key]] if key in fields and fields[key] < len(row) else None

                if not any(get(key) not in (None, '') for key in
                           ('date', 'due', 'ref', 'party', 'amount', 'gross', 'debit', 'credit', 'signed')):
                    continue
                source = f'{sheet.title}!{index}'
                status = clean(get('status')).lower()
                gross, paid, balance = (money(get(k)) for k in ('gross', 'paid', 'balance'))
                amount = money(get('amount'))
                when = day_of(get('due' if kind in {'invoice', 'payable'} else 'date'))
                if kind in {'invoice', 'payable'}:
                    if gross is not None and (paid is not None or (balance is None and 'paid' in fields and get('paid') is None)):
                        balance = gross - (paid or 0)
                    amount = balance if balance is not None else gross
                    if status in CLOSED and status not in {'cancelled', 'canceled', 'void', 'refunded'} and amount and amount > 0:
                        issues.append((source, 'Marked paid but the recorded amounts leave a balance. Check the bank receipt or payment record.'))
                    if status in CLOSED:
                        amount = Decimal(0)
                elif kind in {'actual', 'forecast'}:
                    flow = clean(get('flow')).lower()
                    signed = money(get('signed'))
                    debit, credit = money(get('debit')), money(get('credit'))
                    if amount is not None and flow in {'inflow', 'outflow', 'income', 'expense', 'receipt', 'payment'}:
                        amount = abs(amount) * (-1 if flow in {'outflow', 'expense', 'payment'} else 1)
                    elif signed is not None:
                        amount = signed
                    elif debit is not None or credit is not None:
                        amount = (credit or 0) - (debit or 0)
                    else:
                        amount = None
                    if kind == 'forecast' and status in CLOSED:
                        amount = Decimal(0)
                if amount is None or len(currencies) > 1 or (kind in {'actual', 'forecast'} and when is None):
                    skipped += 1
                    details = 'Check the amount, currency, date and cash direction in this row.'
                    issues.append((source, details))
                    continue
                prefix = sheet.title.upper().split()[0]
                company = clean(get('company')) or BRANDS.get(prefix, 'Unassigned')
                company = next((brand for code, brand in BRANDS.items() if company.casefold() in (code.casefold(), brand.casefold())), company)
                agreed = money(get('agreed'))
                underbilled = (
                    agreed - gross if kind == 'invoice' and status not in {'cancelled', 'canceled', 'void', 'refunded'}
                    and agreed is not None and gross is not None and 0 <= gross < agreed else Decimal(0)
                )
                entries.append(Entry(kind, company, next(iter(currencies), fallback_currency), source,
                                     clean(get('ref')), clean(get('party')), when, amount, paid, status,
                                     clean(get('category')), underbilled))
                accepted += 1
            coverage.append((sheet.title, found[0] if found else 'No supported table', accepted, skipped))
    finally:
        workbook.close()
    return entries, coverage, issues


def findings(entries, as_of):
    results = []
    invoices = [e for e in entries if e.kind == 'invoice' and e.amount > 0]
    keys = Counter((e.party.casefold(), e.ref.casefold()) for e in invoices if e.party and e.ref)
    repeated = {key for key, count in keys.items() if count > 1}
    overdue = [e for e in invoices if e.day and e.day < as_of and
               (e.party.casefold(), e.ref.casefold()) not in repeated]
    for e in sorted(entries, key=lambda row: row.underbilled, reverse=True):
        if e.underbilled:
            results.append(('Possible underbilling', e.underbilled, f'{e.party or "Customer"} • {e.ref or "invoice"}',
                            'Check the agreed amount for this invoice, including approved discounts and changes. If the difference is valid, issue the missing charge.', e.source))
    for e in sorted(overdue, key=lambda item: item.amount, reverse=True):
        results.append(('Overdue customer balance', e.amount,
                        f'{e.party or "Customer"} • {e.ref or "invoice"} • {(as_of-e.day).days} days overdue',
                        'Check the bank for a recent receipt, then send the client the invoice and ask for a payment date.', e.source))
    for e in invoices:
        if not e.day:
            results.append(('Missing invoice due date', e.amount, f'{e.party or "Customer"} • {e.ref or "invoice"}',
                            'Add the agreed due date so this balance can be aged and tracked.', e.source))
    groups = defaultdict(list)
    for e in entries:
        if e.party and e.ref and e.amount and e.status not in CLOSED and (e.kind in {'invoice', 'payable'} or e.kind == 'actual' and e.amount < 0):
            key = (e.kind, e.party.casefold(), e.ref.casefold())
            if e.kind == 'actual':
                key += (e.day, e.amount)
            groups[key].append(e)
    for group in groups.values():
        if len(group) > 1:
            e = group[0]
            results.append(('Possible duplicate ' + {'invoice': 'invoice', 'payable': 'supplier bill', 'actual': 'payment'}[e.kind],
                            sum(abs(row.amount) for row in group[1:]), f'{e.party} • {e.ref} • {len(group)} matching rows',
                            'Compare the original invoice and bank references before correcting entries or requesting a refund.',
                            ', '.join(row.source for row in group)))
    return overdue, results


def loan_schedule(principal, annual_pct, months, first_due):
    rate = annual_pct / 1200
    from math import expm1, log1p
    payment = principal / months if not rate else principal * rate / -expm1(-months * log1p(rate))
    balance = principal
    rows = []
    for offset in range(months):
        month_number = first_due.month - 1 + offset
        year, month = first_due.year + month_number // 12, month_number % 12 + 1
        due = date(year, month, min(first_due.day, monthrange(year, month)[1]))
        interest = round(balance * rate, 2)
        capital = balance if offset == months - 1 else min(balance, round(payment - interest, 2))
        balance = round(balance - capital, 2)
        rows.append((due, round(capital + interest, 2), capital, interest, balance))
    return rows


def projection(events, opening, start, days, delay, buffer):
    base_moves, delayed_moves = defaultdict(float), defaultdict(float)
    for due, amount in events:
        value = float(amount)
        base_moves[due] += value
        delayed_moves[due + timedelta(days=delay) if value > 0 else due] += value
    base = delayed = opening
    rows = []
    for offset in range(days + 1):
        today = start + timedelta(days=offset)
        base += base_moves[today]
        delayed += delayed_moves[today]
        rows.append((today, round(base, 2), round(delayed, 2), buffer))
    return rows
