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

SAVED_DF_TX = None


def read_csv_bulletproof(file_path):
  """קוראת קובץ CSV מהדיסק הווירטואלי ומנסה את כל קידודי העברית והמפרידים הנפוצים."""
  with open(file_path, 'rb') as f:
    raw = f.read()

  if raw.startswith(b'\xef\xbb\xbf'):
    raw = raw[3:]

  encodings = ['utf-8-sig', 'cp1255', 'iso-8859-8', 'utf-8', 'utf-16', 'latin1']
  separators = [None, ',', '\t', ';']

  for enc in encodings:
    for sep in separators:
      try:
        df = pd.read_csv(
            io.BytesIO(raw), sep=sep, engine='python', encoding=enc
        )
        if df is not None and not df.empty and len(df.columns) >= 1:
          cleaned_cols = []
          for c in df.columns:
            s = str(c).strip().replace('\ufeff', '')
            cleaned_cols.append(s)
          df.columns = cleaned_cols
          return df
      except Exception:
        continue

  return pd.read_csv(
      io.BytesIO(raw), sep=None, engine='python', encoding='iso-8859-8'
  )


def inspect_tx_file(tx_file_path):
  global SAVED_DF_TX
  SAVED_DF_TX = read_csv_bulletproof(tx_file_path)
  cols = [str(c) for c in SAVED_DF_TX.columns]
  return json.dumps(cols, ensure_ascii=False)


def inspect_type_values(type_col_name):
  global SAVED_DF_TX
  if SAVED_DF_TX is None or type_col_name not in SAVED_DF_TX.columns:
    return json.dumps([], ensure_ascii=False)
  vals = (
      SAVED_DF_TX[type_col_name]
      .dropna()
      .astype(str)
      .str.strip()
      .unique()
      .tolist()
  )
  return json.dumps([v for v in vals if v], ensure_ascii=False)


def clean_input_data(df_tx, df_cpi, col_map, type_map):
  df_tx = df_tx.copy()
  df_cpi = df_cpi.copy()

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

  date_col = col_map.get('date')
  amt_col = col_map.get('amount')
  type_col = col_map.get('type')

  if not date_col or date_col not in df_tx.columns:
    raise ValueError(f"עמודת התאריך '{date_col}' לא נמצאה בקובץ.")
  if not amt_col or amt_col not in df_tx.columns:
    raise ValueError(f"עמודת הסכום '{amt_col}' לא נמצאה בקובץ.")
  if not type_col or type_col not in df_tx.columns:
    raise ValueError(f"עמודת סוג התנועה '{type_col}' לא נמצאה בקובץ.")

  df_tx = df_tx.rename(
      columns={date_col: 'תאריך', amt_col: 'סכום', type_col: 'סוג תנועה'}
  )

  if df_tx['סכום'].dtype == 'object':
    df_tx['סכום'] = (
        df_tx['סכום']
        .astype(str)
        .str.replace('₪', '', regex=False)
        .str.replace(',', '', regex=False)
        .str.strip()
        .astype(float)
    )
  else:
    df_tx['סכום'] = df_tx['סכום'].astype(float)

  if df_cpi['מדד בגין'].dtype == 'object':
    df_cpi['מדד בגין'] = (
        df_cpi['מדד בגין']
        .astype(str)
        .str.replace(',', '', regex=False)
        .str.strip()
        .astype(float)
    )
  else:
    df_cpi['מדד בגין'] = df_cpi['מדד בגין'].astype(float)

  df_tx['תאריך'] = pd.to_datetime(
      df_tx['תאריך'], dayfirst=True, errors='coerce'
  ).dt.date
  df_tx = df_tx.dropna(subset=['תאריך'])
  if df_tx.empty:
    raise ValueError("לא נמצאו תאריכים תקינים בקובץ התזרים.")

  df_cpi['חודש'] = pd.to_datetime(
      df_cpi['חודש'], dayfirst=True, errors='coerce'
  ).dt.date
  df_cpi = df_cpi.dropna(subset=['חודש'])

  dep_val = str(type_map.get('deposit', '')).strip()
  wth_val = str(type_map.get('withdraw', '')).strip()

  def map_sug(val):
    s_val = str(val).strip()
    if s_val == dep_val:
      return 'הפקדה'
    elif s_val == wth_val:
      return 'משיכה'
    if any(k in s_val.lower() for k in ['הפקדה', 'deposit', 'זכות', 'cr']):
      return 'הפקדה'
    if any(
        k in s_val.lower() for k in ['משיכה', 'withdraw', 'חובה', 'dr', 'סילוק']
    ):
      return 'משיכה'
    return 'הפקדה'

  df_tx['סוג תנועה'] = df_tx['סוג תנועה'].apply(map_sug)

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


def run_calculation_json(params_json):
  params = json.loads(params_json)

  tx_file_path = params['tx_file_path']
  cpi_file_path = params['cpi_file_path']
  col_map = params['col_map']
  type_map = params['type_map']
  annual_rate = float(params['annual_rate'])
  compounding_freq = params['compounding_freq']
  method = params['method']

  r = annual_rate / 100.0
  if compounding_freq == 'יומית':
    r_daily = r / 365.0
  elif compounding_freq == 'חודשית':
    r_daily = ((1.0 + r / 12.0) ** (12.0 / 365.0)) - 1.0
  elif compounding_freq == 'רבעונית':
    r_daily = ((1.0 + r / 4.0) ** (4.0 / 365.0)) - 1.0
  elif compounding_freq == 'שנתית':
    r_daily = ((1.0 + r) ** (1.0 / 365.0)) - 1.0

  df_tx = read_csv_bulletproof(tx_file_path)
  df_cpi = read_csv_bulletproof(cpi_file_path)

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

    # 1. חישוב צבירות תקופתיות לכל השכבות הקימות (עד ליום החישוב הנוכחי)
    for t_id, t in tables.items():
      if dt <= t['last_dt']:
        continue

      days = (dt - t['last_dt']).days
      tot_balance = t['p'] + t['i'] + t['adj_p'] + t['adj_i']

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

      # צבירת ריבית לא-מוצמדת
      new_i = t['p'] * (((1.0 + r_daily) ** days) - 1.0)
      closing_i_unadjusted = t['i'] + new_i

      # הצמדה מצטברת מיום יצירת השכבה المקורי
      base_idx = t['base_idx']
      cum_idx_factor = (idx / base_idx) - 1.0 if base_idx > 0 else 0.0

      target_adj_p_closing = t['p'] * cum_idx_factor
      target_adj_i_closing = closing_i_unadjusted * cum_idx_factor

      new_adj_p = target_adj_p_closing - t['adj_p']
      new_adj_i = target_adj_i_closing - t['adj_i']

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
          'ריבית יתרת סגירה': closing_i_unadjusted,
          'הפרשי הצמדה קרן יתרת פתיחה': t['adj_p'],
          'צבירת הפרשי הצמדה קרן': new_adj_p,
          'סילוק הפרשי הצמדה קרן': 0.0,
          'הפרשי הצמדה קרן יתרת סגירה': target_adj_p_closing,
          'הפרשי הצמדה ריבית יתרת פתיחה': t['adj_i'],
          'צבירת הפרשי הצמדה ריבית': new_adj_i,
          'סילוק הפרשי הצמדה ריבית': 0.0,
          'הפרשי הצמדה ריבית יתרת סגירה': target_adj_i_closing,
      }

      t['i'] = closing_i_unadjusted
      t['adj_p'] = target_adj_p_closing
      t['adj_i'] = target_adj_i_closing
      t['last_dt'] = dt
      t['last_idx'] = idx

    # 2. הזרמת התזרימים של היום והפעלת ה-Waterfall בסוף התקופה (לאחר הצבירה)
    for _, tx in day_txs.iterrows():
      tx_type = str(tx['סוג תנועה']).strip()
      amt = abs(float(tx['סכום']))
      df_summary = get_consolidated_state(tables)

      if 'משיכה' in tx_type or 'withdraw' in tx_type:
        # אם יש חוב/יתרות זכות חיוביות, מפעילים Waterfall לסילוק
        if not df_summary.empty and df_summary['tot'].sum() > 0:
          df_alloc, rem_surplus = allocate_waterfall(
              df_summary, amt, method=method
          )
          for _, row in df_alloc.iterrows():
            t_id = int(row['tranche_id'])
            t = tables[t_id]
            r_dict = t.get('pending_row', t['rows'][-1])

            # סילוק בסוף התקופה על היתרות הצבורות
            r_dict['סילוק ריבית'] += row['pay_i']
            r_dict['סילוק הפרשי הצמדה ריבית'] += row['pay_adj_i']
            r_dict['סילוק הפרשי הצמדה קרן'] += row['pay_adj_p']
            r_dict['סילוק קרן'] += row['pay_p']

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

        # יתרת משיכה בלתי-מסולקת פותחת שכבת משיכת יתר (Overdraft)
        if rem_surplus > 0.0001:
          tranche_counter += 1
          tables[tranche_counter] = {
              'type': 'overdraft',
              'p': -rem_surplus,
              'i': 0.0,
              'adj_p': 0.0,
              'adj_i': 0.0,
              'base_idx': idx,
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
        # אם יש שכבות משיכת יתר שליליות, ההפקדה מסלקת אותן תחילה
        if not df_summary.empty and df_summary['tot'].sum() < 0:
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

        # יתרת הפקדה פותחת שכבת הפקדה רגילה (Deposit)
        if rem_surplus > 0.0001:
          tranche_counter += 1
          tables[tranche_counter] = {
              'type': 'deposit',
              'p': rem_surplus,
              'i': 0.0,
              'adj_p': 0.0,
              'adj_i': 0.0,
              'base_idx': idx,
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

    # שמירת השורות שהסתיימו בפרק הזמן
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

  output_path = 'output_results.xlsx'
  with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    pd.DataFrame(consolidated_history)[COLUMNS].to_excel(
        writer, sheet_name='טבלה מאוחדת', index=False
    )
    for t_id, t in tables.items():
      label = (
          f'הפקדה_{t_id}'
          if t.get('type') == 'deposit'
          else f'משיכת_יתר_{t_id}'
      )
      pd.DataFrame(t['rows'])[COLUMNS].to_excel(
          writer, sheet_name=label, index=False
      )

  wb = openpyxl.load_workbook(output_path)
  for ws in wb.worksheets:
    ws.sheet_view.rightToLeft = True
    ws.freeze_panes = 'B2'
    for cell in ws[1]:
      cell.alignment = Alignment(
          wrap_text=True, horizontal='center', vertical='center'
      )
    for row in ws.iter_rows(
        min_row=2, max_row=ws.max_row, max_col=ws.max_column
    ):
      for cell in row:
        if cell.column == 1:
          cell.number_format = 'YYYY-MM-DD'
        elif cell.column in [2, 3]:
          cell.number_format = '#,##0.00' if cell.column == 2 else '#,##0'
        else:
          cell.number_format = '#,##0.00'

  wb.save(output_path)

  last = consolidated_history[-1]
  as_of_date_str = pd.to_datetime(last['תאריך']).strftime('%d/%m/%Y')
  tot_debt = (
      last['קרן יתרת סגירה']
      + last['ריבית יתרת סגירה']
      + last['הפרשי הצמדה קרן יתרת סגירה']
      + last['הפרשי הצמדה ריבית יתרת סגירה']
  )

  summary = {
      'as_of_date': as_of_date_str,
      'p': float(last['קרן יתרת סגירה']),
      'i': float(last['ריבית יתרת סגירה']),
      'adj': float(
          last['הפרשי הצמדה קרן יתרת סגירה']
          + last['הפרשי הצמדה ריבית יתרת סגירה']
      ),
      'tot': float(tot_debt),
      'annual_rate': annual_rate,
      'compounding_freq': compounding_freq,
      'method': method,
  }

  return json.dumps(summary, ensure_ascii=False)
