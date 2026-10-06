from datetime import date
import io
import json
import openpyxl
from openpyxl.styles import Alignment
import pandas as pd

COLUMNS = [
    'תאריך',
    'מדד ידוע',
    'ימי ריבית',
    'קרן יתרת פתיחה',
    'צבירת קרן',
    'סילוק קרן',
    'קרן יתרת סגירה',
    'ריבית יתרת פתיחה',
    'צבירת ריבית',
    'סילוק ריבית',
    'ריבית יתרת סגירה',
    'הפרשי הצמדה קרן יתרת פתיחה',
    'צבירת הפרשי הצמדה קרן',
    'סילוק הפרשי הצמדה קרן',
    'הפרשי הצמדה קרן יתרת סגירה',
    'הפרשי הצמדה ריבית יתרת פתיחה',
    'צבירת הפרשי הצמדה ריבית',
    'סילוק הפרשי הצמדה ריבית',
    'הפרשי הצמדה ריבית יתרת סגירה',
]

CURRENT_DF_TEMP = None


def load_df_from_file(file_source):
  """טוענת קובץ מנתיב/בייטס ומנסה את כל קידודי העברית והמפרידים הנפוצים."""
  if isinstance(file_source, pd.DataFrame):
    return file_source

  if isinstance(file_source, str):
    with open(file_source, 'rb') as f:
      raw_bytes = f.read()
  elif isinstance(file_source, bytes):
    raw_bytes = file_source
  elif hasattr(file_source, 'to_py'):
    raw_bytes = bytes(file_source.to_py())
  elif isinstance(file_source, (memoryview, bytearray)):
    raw_bytes = bytes(file_source)
  else:
    raw_bytes = bytes(file_source)

  encodings = ['utf-8-sig', 'cp1255', 'iso-8859-8', 'utf-8', 'latin1']
  separators = [None, ',', '\t', ';']

  for enc in encodings:
    for sep in separators:
      try:
        df = pd.read_csv(
            io.BytesIO(raw_bytes), sep=sep, engine='python', encoding=enc
        )
        if not df.empty and len(df.columns) >= 1:
          cleaned_cols = []
          for col in df.columns:
            c_str = str(col).strip()
            if c_str.startswith('\ufeff'):
              c_str = c_str[1:]
            cleaned_cols.append(c_str)
          df.columns = cleaned_cols
          return df
      except Exception:
        continue

  return pd.read_csv(
      io.BytesIO(raw_bytes), sep=None, engine='python', encoding='iso-8859-8'
  )


def extract_file_metadata_from_file(file_path):
  """חילוץ כותרות עמודות לצורך הממשק הדינמי."""
  global CURRENT_DF_TEMP
  CURRENT_DF_TEMP = load_df_from_file(file_path)
  cols = [str(c).strip() for c in CURRENT_DF_TEMP.columns]
  return CURRENT_DF_TEMP, cols


def get_unique_type_values(col_name):
  """חילוץ ערכים ייחודיים מעמודת סוג התנועה."""
  global CURRENT_DF_TEMP
  if CURRENT_DF_TEMP is None or col_name not in CURRENT_DF_TEMP.columns:
    return []
  vals = (
      CURRENT_DF_TEMP[col_name]
      .dropna()
      .astype(str)
      .str.strip()
      .unique()
      .tolist()
  )
  return [v for v in vals if v]


def clean_input_data(df_tx, df_cpi, col_map, type_map):
  """מנרמלת את הנתונים לפי המיפוי שהתקבל מהממשק."""
  df_tx = df_tx.copy()
  df_cpi = df_cpi.copy()

  # ניקוי CPI
  df_cpi.columns = [str(col).strip() for col in df_cpi.columns]
  cpi_month_col = [
      col
      for col in df_cpi.columns
      if 'חודש' in str(col) or 'month' in str(col).lower()
  ]
  if cpi_month_col:
    df_cpi = df_cpi.rename(columns={cpi_month_col[0]: 'חודש'})

  cpi_val_col = [
      col
      for col in df_cpi.columns
      if 'מדד' in str(col) or 'cpi' in str(col).lower()
  ]
  if cpi_val_col:
    df_cpi = df_cpi.rename(columns={cpi_val_col[0]: 'מדד בגין'})

  # מיפוי עמודות תזרים
  df_tx = df_tx.rename(
      columns={
          col_map['date']: 'תאריך',
          col_map['amount']: 'סכום',
          col_map['type']: 'סוג תנועה',
      }
  )

  dep_val = str(type_map.get('deposit', '')).strip()
  wth_val = str(type_map.get('withdraw', '')).strip()

  def map_sug(val):
    s_val = str(val).strip()
    if s_val == dep_val:
      return 'הפקדה'
    elif s_val == wth_val:
      return 'משיכה'
    if 'הפקדה' in s_val or 'deposit' in s_val.lower():
      return 'הפקדה'
    if 'משיכה' in s_val or 'withdraw' in s_val.lower():
      return 'משיכה'
    return s_val

  df_tx['סוג תנועה'] = df_tx['סוג תנועה'].apply(map_sug)

  # המרת תאריכים
  df_tx['תאריך'] = pd.to_datetime(df_tx['תאריך'], dayfirst=True).dt.date
  df_cpi['חודש'] = pd.to_datetime(df_cpi['חודש'], dayfirst=True).dt.date

  # המרת סכומים ומדדים
  if df_tx['סכום'].dtype == 'object':
    df_tx['סכום'] = (
        df_tx['סכום'].astype(str).str.replace(',', '').astype(float)
    )
  else:
    df_tx['סכום'] = df_tx['סכום'].astype(float)

  if df_cpi['מדד בגין'].dtype == 'object':
    df_cpi['מדד בגין'] = (
        df_cpi['מדד בגין'].astype(str).str.replace(',', '').astype(float)
    )
  else:
    df_cpi['מדד בגין'] = df_cpi['מדד בגין'].astype(float)

  df_tx = df_tx.sort_values('תאריך').reset_index(drop=True)
  df_cpi = df_cpi.sort_values('חודש').reset_index(drop=True)

  return df_tx, df_cpi


def get_known_cpi(current_date, cpi_df):
  if isinstance(current_date, str):
    current_date = pd.to_datetime(current_date).date()
  elif isinstance(current_date, pd.Timestamp):
    current_date = current_date.date()

  if current_date.day >= 15:
    target_month_end = current_date.replace(day=1) - pd.Timedelta(days=1)
  else:
    first_of_this_month = current_date.replace(day=1)
    target_month_end = (first_of_this_month - pd.Timedelta(days=1)).replace(
        day=1
    ) - pd.Timedelta(days=1)

  target_year = target_month_end.year
  target_month = target_month_end.month

  cpi_months = pd.to_datetime(cpi_df['חודש'])
  matched = cpi_df[
      (cpi_months.dt.year == target_year)
      & (cpi_months.dt.month == target_month)
  ]

  if not matched.empty:
    return matched['מדד בגין'].iloc[0]
  else:
    return None


def build_timeline(min_date, max_date, tx_dates):
  min_dt = pd.to_datetime(min_date)
  max_dt = pd.to_datetime(max_date)

  first_of_months = pd.date_range(
      start=min_dt.replace(day=1), end=max_dt.replace(day=1), freq='MS'
  )

  month_ends = set(
      first_of_months.map(lambda d: d.replace(day=d.days_in_month)).date
  )

  all_dates = set(tx_dates).union(month_ends)
  timeline = sorted(
      [d for d in all_dates if d >= pd.to_datetime(min_date).date()]
  )

  return timeline


def get_consolidated_state(tables):
  rows = []
  for t_id, t in tables.items():
    tot = t['p'] + t['i'] + t['adj_p'] + t['adj_i']
    if abs(tot) > 0.0001:
      rows.append({
          'tranche_id': t_id,
          'p': t['p'],
          'i': t['i'],
          'adj_p': t['adj_p'],
          'adj_i': t['adj_i'],
          'tot': tot,
          'type': t.get('type', 'deposit'),
      })
  if not rows:
    return pd.DataFrame(
        columns=['tranche_id', 'p', 'i', 'adj_p', 'adj_i', 'tot', 'type']
    )
  return pd.DataFrame(rows)


def allocate_waterfall(df_summary, pmt_amount, method='FIFO'):
  rem = abs(pmt_amount)
  df_alloc = df_summary.copy()

  df_alloc['pay_i'] = 0.0
  df_alloc['pay_adj_i'] = 0.0
  df_alloc['pay_adj_p'] = 0.0
  df_alloc['pay_p'] = 0.0

  if df_alloc.empty or rem <= 0:
    return df_alloc, rem

  ascending = True if method == 'FIFO' else False
  df_alloc = df_alloc.sort_values('tranche_id', ascending=ascending)

  for idx in df_alloc.index:
    if rem <= 0:
      break
    p_i = min(rem, abs(df_alloc.loc[idx, 'i']))
    df_alloc.loc[idx, 'pay_i'] = p_i
    rem -= p_i

  for idx in df_alloc.index:
    if rem <= 0:
      break
    p_ai = min(rem, abs(df_alloc.loc[idx, 'adj_i']))
    df_alloc.loc[idx, 'pay_adj_i'] = p_ai
    rem -= p_ai

  for idx in df_alloc.index:
    if rem <= 0:
      break
    p_ap = min(rem, abs(df_alloc.loc[idx, 'adj_p']))
    df_alloc.loc[idx, 'pay_adj_p'] = p_ap
    rem -= p_ap

  for idx in df_alloc.index:
    if rem <= 0:
      break
    p_p = min(rem, abs(df_alloc.loc[idx, 'p']))
    df_alloc.loc[idx, 'pay_p'] = p_p
    rem -= p_p

  return df_alloc, rem


def run_full_calculation(
    tx_file_path,
    cpi_file_path,
    col_map,
    type_map,
    annual_rate=6.0,
    compounding_freq='חודשית',
    method='FIFO',
):
  r = annual_rate / 100.0
  if compounding_freq == 'יומית':
    r_daily = r / 365.0
  elif compounding_freq == 'חודשית':
    r_daily = ((1.0 + r / 12.0) ** (12.0 / 365.0)) - 1.0
  elif compounding_freq == 'רבעונית':
    r_daily = ((1.0 + r / 4.0) ** (4.0 / 365.0)) - 1.0
  elif compounding_freq == 'שנתית':
    r_daily = ((1.0 + r) ** (1.0 / 365.0)) - 1.0

  df_tx = load_df_from_file(tx_file_path)
  df_cpi = load_df_from_file(cpi_file_path)

  df_tx, df_cpi = clean_input_data(df_tx, df_cpi, col_map, type_map)

  tx_dates = set(pd.to_datetime(df_tx['תאריך']).dt.date)
  timeline = build_timeline(min(tx_dates), max(tx_dates), tx_dates)

  tables = {}
  tranche_counter = 0
  consolidated_history = []
  prev_timeline_dt = None

  for dt in timeline:
    idx = get_known_cpi(dt, df_cpi)
    day_txs = df_tx[df_tx['תאריך'] == dt]
    days_from_prev = (dt - prev_timeline_dt).days if prev_timeline_dt else 0

    for t_id, t in tables.items():
      if dt <= t['last_dt']:
        continue

      tot_balance = t['p'] + t['i'] + t['adj_p'] + t['adj_i']
      days = (dt - t['last_dt']).days

      if abs(tot_balance) <= 0.0001:
        t['pending_row'] = {
            'תאריך': dt,
            'מדד ידוע': idx,
            'ימי ריבית': days,
            'קרן יתרת פתיחה': 0.0,
            'צבירת קרן': 0.0,
            'סילוק קרן': 0.0,
            'קרן יתרת סגירה': 0.0,
            'ריבית יתרת פתיחה': 0.0,
            'צבירת ריבית': 0.0,
            'סילוק ריבית': 0.0,
            'ריבית יתרת סגירה': 0.0,
            'הפרשי הצמדה קרן יתרת פתיחה': 0.0,
            'צבירת הפרשי הצמדה קרן': 0.0,
            'סילוק הפרשי הצמדה קרן': 0.0,
            'הפרשי הצמדה קרן יתרת סגירה': 0.0,
            'הפרשי הצמדה ריבית יתרת פתיחה': 0.0,
            'צבירת הפרשי הצמדה ריבית': 0.0,
            'סילוק הפרשי הצמדה ריבית': 0.0,
            'הפרשי הצמדה ריבית יתרת סגירה': 0.0,
        }
        t['last_dt'] = dt
        continue

      new_i = t['p'] * (((1.0 + r_daily) ** days) - 1.0)
      idx_factor = (idx / t['last_idx']) - 1.0 if t['last_idx'] > 0 else 0.0

      new_adj_p = (t['p'] + t['adj_p']) * idx_factor
      new_adj_i = (t['i'] + new_i + t['adj_i']) * idx_factor

      t['pending_row'] = {
          'תאריך': dt,
          'מדד ידוע': idx,
          'ימי ריבית': days,
          'קרן יתרת פתיחה': t['p'],
          'צבירת קרן': 0.0,
          'סילוק קרן': 0.0,
          'קרן יתרת סגירה': t['p'],
          'ריבית יתרת פתיחה': t['i'],
          'צבירת ריבית': new_i,
          'סילוק ריבית': 0.0,
          'ריבית יתרת סגירה': t['i'] + new_i,
          'הפרשי הצמדה קרן יתרת פתיחה': t['adj_p'],
          'צבירת הפרשי הצמדה קרן': new_adj_p,
          'סילוק הפרשי הצמדה קרן': 0.0,
          'הפרשי הצמדה קרן יתרת סגירה': t['adj_p'] + new_adj_p,
          'הפרשי הצמדה ריבית יתרת פתיחה': t['adj_i'],
          'צבירת הפרשי הצמדה ריבית': new_adj_i,
          'סילוק הפרשי הצמדה ריבית': 0.0,
          'הפרשי הצמדה ריבית יתרת סגירה': t['adj_i'] + new_adj_i,
      }

      t['i'] += new_i
      t['adj_i'] += new_adj_i
      t['adj_p'] += new_adj_p
      t['last_dt'] = dt
      t['last_idx'] = idx

    for _, tx in day_txs.iterrows():
      tx_type = str(tx['סוג תנועה']).strip()
      amt = abs(float(tx['סכום']))
      df_summary = get_consolidated_state(tables)

      if 'משיכה' in tx_type or 'withdraw' in tx_type:
        if not df_summary.empty and df_summary['tot'].sum() > 0:
          df_alloc, rem_surplus = allocate_waterfall(
              df_summary, amt, method=method
          )
          for _, row in df_alloc.iterrows():
            t_id = int(row['tranche_id'])
            t = tables[t_id]
            r_dict = t.get('pending_row', t['rows'][-1])

            r_dict['סילוק ריבית'] -= row['pay_i']
            r_dict['סילוק הפרשי הצמדה ריבית'] -= row['pay_adj_i']
            r_dict['סילוק הפרשי הצמדה קרן'] -= row['pay_adj_p']
            r_dict['סילוק קרן'] -= row['pay_p']

            t['i'] -= row['pay_i']
            r_dict['ריבית יתרת סגירה'] = t['i']
            t['adj_i'] -= row['pay_adj_i']
            r_dict['הפרשי הצמדה ריבית יתרת סגירה'] = t['adj_i']
            t['adj_p'] -= row['pay_adj_p']
            r_dict['הפרשי הצמדה קרן יתרת סגירה'] = t['adj_p']
            t['p'] -= row['pay_p']
            r_dict['קרן יתרת סגירה'] = t['p']
        else:
          rem_surplus = amt

        if rem_surplus > 0.0001:
          tranche_counter += 1
          tables[tranche_counter] = {
              'type': 'overdraft',
              'p': -rem_surplus,
              'i': 0.0,
              'adj_p': 0.0,
              'adj_i': 0.0,
              'last_dt': dt,
              'last_idx': idx,
              'rows': [{
                  'תאריך': dt,
                  'מדד ידוע': idx,
                  'ימי ריבית': 0,
                  'קרן יתרת פתיחה': 0.0,
                  'צבירת קרן': -rem_surplus,
                  'סילוק קרן': 0.0,
                  'קרן יתרת סגירה': -rem_surplus,
                  'ריבית יתרת פתיחה': 0.0,
                  'צבירת ריבית': 0.0,
                  'סילוק ריבית': 0.0,
                  'ריבית יתרת סגירה': 0.0,
                  'הפרשי הצמדה קרן יתרת פתיחה': 0.0,
                  'צבירת הפרשי הצמדה קרן': 0.0,
                  'סילוק הפרשי הצמדה קרן': 0.0,
                  'הפרשי הצמדה קרן יתרת סגירה': 0.0,
                  'הפרשי הצמדה ריבית יתרת פתיחה': 0.0,
                  'צבירת הפרשי הצמדה ריבית': 0.0,
                  'סילוק הפרשי הצמדה ריבית': 0.0,
                  'הפרשי הצמדה ריבית יתרת סגירה': 0.0,
              }],
          }

      elif 'הפקדה' in tx_type or 'deposit' in tx_type:
        if not df_summary.empty and df_summary['tot'].sum() < 0:
          df_alloc, rem_surplus = allocate_waterfall(
              df_summary, amt, method=method
          )
          for _, row in df_alloc.iterrows():
            t_id = int(row['tranche_id'])
            t = tables[t_id]
            r_dict = t.get('pending_row', t['rows'][-1])

            r_dict['סילוק ריבית'] += row['pay_i']
            r_dict['סילוק הפרשי הצמדה ריבית'] += row['pay_adj_i']
            r_dict['סילוק הפרשי הצמדה קרן'] += row['pay_adj_p']
            r_dict['סילוק קרן'] += row['pay_p']

            t['i'] += row['pay_i']
            r_dict['ריבית יתרת סגירה'] = t['i']
            t['adj_i'] += row['pay_adj_i']
            r_dict['הפרשי הצמדה ריבית יתרת סגירה'] = t['adj_i']
            t['adj_p'] += row['pay_adj_p']
            r_dict['הפרשי הצמדה קרן יתרת סגירה'] = t['adj_p']
            t['p'] += row['pay_p']
            r_dict['קרן יתרת סגירה'] = t['p']
        else:
          rem_surplus = amt

        if rem_surplus > 0.0001:
          tranche_counter += 1
          tables[tranche_counter] = {
              'type': 'deposit',
              'p': rem_surplus,
              'i': 0.0,
              'adj_p': 0.0,
              'adj_i': 0.0,
              'last_dt': dt,
              'last_idx': idx,
              'rows': [{
                  'תאריך': dt,
                  'מדד ידוע': idx,
                  'ימי ריבית': 0,
                  'קרן יתרת פתיחה': 0.0,
                  'צבירת קרן': rem_surplus,
                  'סילוק קרן': 0.0,
                  'קרן יתרת סגירה': rem_surplus,
                  'ריבית יתרת פתיחה': 0.0,
                  'צבירת ריבית': 0.0,
                  'סילוק ריבית': 0.0,
                  'ריבית יתרת סגירה': 0.0,
                  'הפרשי הצמדה קרן יתרת פתיחה': 0.0,
                  'צבירת הפרשי הצמדה קרן': 0.0,
                  'סילוק הפרשי הצמדה קרן': 0.0,
                  'הפרשי הצמדה קרן יתרת סגירה': 0.0,
                  'הפרשי הצמדה ריבית יתרת פתיחה': 0.0,
                  'צבירת הפרשי הצמדה ריבית': 0.0,
                  'סילוק הפרשי הצמדה ריבית': 0.0,
                  'הפרשי הצמדה ריבית יתרת סגירה': 0.0,
              }],
          }

    for t in tables.values():
      if 'pending_row' in t:
        t['rows'].append(t['pending_row'])
        del t['pending_row']

    consolidated_history.append({
        'תאריך': dt,
        'מדד ידוע': idx,
        'ימי ריבית': days_from_prev,
        'קרן יתרת פתיחה': sum(
            t['rows'][-1]['קרן יתרת פתיחה'] for t in tables.values()
        ),
        'צבירת קרן': sum(t['rows'][-1]['צבירת קרן'] for t in tables.values()),
        'סילוק קרן': sum(t['rows'][-1]['סילוק קרן'] for t in tables.values()),
        'קרן יתרת סגירה': sum(
            t['rows'][-1]['קרן יתרת סגירה'] for t in tables.values()
        ),
        'ריבית יתרת פתיחה': sum(
            t['rows'][-1]['ריבית יתרת פתיחה'] for t in tables.values()
        ),
        'צבירת ריבית': sum(
            t['rows'][-1]['צבירת ריבית'] for t in tables.values()
        ),
        'סילוק ריבית': sum(
            t['rows'][-1]['סילוק ריבית'] for t in tables.values()
        ),
        'ריבית יתרת סגירה': sum(
            t['rows'][-1]['ריבית יתרת סגירה'] for t in tables.values()
        ),
        'הפרשי הצמדה קרן יתרת פתיחה': sum(
            t['rows'][-1]['הפרשי הצמדה קרן יתרת פתיחה'] for t in tables.values()
        ),
        'צבירת הפרשי הצמדה קרן': sum(
            t['rows'][-1]['צבירת הפרשי הצמדה קרן'] for t in tables.values()
        ),
        'סילוק הפרשי הצמדה קרן': sum(
            t['rows'][-1]['סילוק הפרשי הצמדה קרן'] for t in tables.values()
        ),
        'הפרשי הצמדה קרן יתרת סגירה': sum(
            t['rows'][-1]['הפרשי הצמדה קרן יתרת סגירה'] for t in tables.values()
        ),
        'הפרשי הצמדה ריבית יתרת פתיחה': sum(
            t['rows'][-1]['הפרשי הצמדה ריבית יתרת פתיחה']
            for t in tables.values()
        ),
        'צבירת הפרשי הצמדה ריבית': sum(
            t['rows'][-1]['צבירת הפרשי הצמדה ריבית'] for t in tables.values()
        ),
        'סילוק הפרשי הצמדה ריבית': sum(
            t['rows'][-1]['סילוק הפרשי הצמדה ריבית'] for t in tables.values()
        ),
        'הפרשי הצמדה ריבית יתרת סגירה': sum(
            t['rows'][-1]['הפרשי הצמדה ריבית יתרת סגירה']
            for t in tables.values()
        ),
    })

    prev_timeline_dt = dt

  return consolidated_history, tables
