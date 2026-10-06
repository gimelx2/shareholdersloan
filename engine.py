from datetime import date
import io
import json
import openpyxl
from openpyxl.styles import Font, PatternFill
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

  df_cpi.columns = [str(col).strip().replace('\ufeff', '') for col in df_cpi.columns]

  cpi_month_col = [
      col for col in df_cpi.columns
      if 'חודש' in str(col) or 'month' in str(col).lower() or 'date' in str(col).lower() or 'תאריך' in str(col)
  ]
  if cpi_month_col:
    df_cpi = df_cpi.rename(columns={cpi_month_col[0]: 'חודש'})
  else:
    df_cpi = df_cpi.rename(columns={df_cpi.columns[0]: 'חודש'})

  cpi_val_col = [
      col for col in df_cpi.columns
      if 'מדד' in str(col) or 'cpi' in str(col).lower() or 'index' in str(col).lower() or 'val' in str(col).lower()
  ]
  if cpi_val_col:
    df_cpi = df_cpi.rename(columns={cpi_val_col[0]: 'מדד בגין'})
  else:
    if len(df_cpi.columns) >= 2:
      df_cpi = df_cpi.rename(columns={df_cpi.columns[1]: 'מדד בגין'})
    else:
      raise ValueError("לא נמצאה עמודת מדד תקינה בקובץ המדדים cpi.csv")

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
    return float(cpi_df.iloc[-1]['מדד בגין'])


def build_timeline(min_date, target_as_of_date, tx_dates):
  """בונה ציר זמן יעיל הכולל אך ורק את ימי התזרים בפועל ואת תאריך הנכונות."""
  min_dt = pd.to_datetime(min_date).date()
  max_dt = pd.to_datetime(target_as_of_date).date()

  all_dates = set([d for d in tx_dates if min_dt <= d <= max_dt])
  all_dates.add(max_dt)

  timeline = sorted(list(all_dates))
  return timeline


def get_consolidated_state(tables, filter_type=None):
  rows = []
  for t_id, t in tables.items():
    if filter_type and t.get('type') != filter_type:
      continue
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
  original_filename = params.get('original_filename', 'קובץ תזרים')
  col_map = params['col_map']
  type_map = params['type_map']
  annual_rate = float(params['annual_rate'])
  compounding_freq = params['compounding_freq']
  method = params['method']
  user_as_of_date_str = params.get('as_of_date', '')

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

  if user_as_of_date_str:
    target_as_of_dt = pd.to_datetime(user_as_of_date_str, dayfirst=True).date()
  else:
    target_as_of_dt = date.today()

  df_tx = df_tx[df_tx['תאריך'] <= target_as_of_dt]

  tx_dates = set(pd.to_datetime(df_tx['תאריך']).dt.date)
  min_tx_date = min(tx_dates) if tx_dates else target_as_of_dt
  timeline = build_timeline(min_tx_date, target_as_of_dt, tx_dates)

  tables = {}
  tranche_counter = 0
  consolidated_history = []
  prev_timeline_dt = None

  for dt in timeline:
    idx = get_known_cpi(dt, df_cpi)
    day_txs = df_tx[df_tx['תאריך'] == dt]
    days_from_prev = (dt - prev_timeline_dt).days if prev_timeline_dt else 0

    # 1. צבירות לתקופה (לפני תזרימי היום)
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

      # 1.1 הצמדת הקרן מיום בסיס השכבה (base_idx)
      base_idx = t['base_idx']
      cum_idx_factor = (idx / base_idx) - 1.0 if base_idx > 0 else 0.0
      target_adj_p_closing = t['p'] * cum_idx_factor
      new_adj_p = target_adj_p_closing - t['adj_p']

      # 1.2 הצמדת הריבית (שרשור מדדים תקופתי + הצמדת הריבית החדשה):
      # א' - קידום יתרת הריבית הקיימת הצמודה (i + adj_i) בשינוי המדד של התקופה הנוכחית בלבד!
      period_idx_factor = (idx / t['last_idx']) - 1.0 if t['last_idx'] > 0 else 0.0
      adj_existing_i = (t['i'] + t['adj_i']) * period_idx_factor

      # ב' - הצמדת הריבית החדשה שנצברה בתקופה (new_i) מיום בסיס השכבה המקורי
      adj_new_i = new_i * cum_idx_factor

      new_adj_i = adj_existing_i + adj_new_i
      target_adj_i_closing = t['adj_i'] + new_adj_i

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

    # 2. הזרמת התזרימים היומיים (משיכות / הפקדות)
    for _, tx in day_txs.iterrows():
      tx_type = str(tx['סוג תנועה']).strip()
      amt = abs(float(tx['סכום']))

      if 'משיכה' in tx_type or 'withdraw' in tx_type:
        df_deposits = get_consolidated_state(tables, filter_type='deposit')
        if not df_deposits.empty and df_deposits['tot'].sum() > 0:
          df_alloc, rem_surplus = allocate_waterfall(
              df_deposits, amt, method=method
          )
          for _, row in df_alloc.iterrows():
            t_id = int(row['tranche_id'])
            t = tables[t_id]
            r_dict = t.get('pending_row', t['rows'][-1])

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

            if row['pay_p'] > 0:
              t['p'] -= row['pay_p']
              t['base_idx'] = idx

            r_dict['קרן יתרת סגירה'] = t['p']
            t['last_idx'] = idx
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
        df_overdrafts = get_consolidated_state(tables, filter_type='overdraft')
        if not df_overdrafts.empty and abs(df_overdrafts['tot'].sum()) > 0:
          df_alloc, rem_surplus = allocate_waterfall(
              df_overdrafts, amt, method=method
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

            if row['pay_p'] > 0:
              t['p'] += row['pay_p']
              t['base_idx'] = idx

            r_dict['קרן יתרת סגירה'] = t['p']
            t['last_idx'] = idx
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

  last = consolidated_history[-1]
  as_of_date_str = pd.to_datetime(last['תאריך']).strftime('%d/%m/%Y')
  last_cpi_val = float(last['מדד ידוע'])
  tot_debt = (
      last['קרן יתרת סגירה']
      + last['ריבית יתרת סגירה']
      + last['הפרשי הצמדה קרן יתרת סגירה']
      + last['הפרשי הצמדה ריבית יתרת סגירה']
  )

  # יצירת קובץ Excel מעוצב
  output_path = 'output_results.xlsx'
  wb = openpyxl.Workbook()
  wb.remove(wb.active)

  ws_info = wb.create_sheet(title='עקרונות והנחות חישוב')
  ws_info.sheet_view.rightToLeft = True

  header_fill = PatternFill(
      start_color='2F3A5B', end_color='2F3A5B', fill_type='solid'
  )
  header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
  title_font = Font(name='Calibri', size=16, bold=True, color='2F3A5B')
  bold_font = Font(name='Calibri', size=11, bold=True)

  ws_info['A1'] = 'דו"ח חישוב הפרשי הצמדה וריבית - Wise Consulting Group'
  ws_info['A1'].font = title_font

  info_rows = [
      ('תאריך הפקת הדוח:', date.today().strftime('%d/%m/%Y')),
      ('תאריך נכונות החישוב שנבחר:', as_of_date_str),
      ('מדד המחירים לצרכן הידוע לתאריך הנכונות:', f'{last_cpi_val:,.2f}'),
      ('שם קובץ הנתונים שהועלה:', original_filename),
      ('שיעור ריבית שנתית:', f'{annual_rate}%'),
      ('תדירות חישוב הריבית:', compounding_freq),
      ('שיטת סילוק משיכות:', f'{method} (Waterfall)'),
      ('', ''),
      ('סיכום יתרות לתאריך הנכונות:', ''),
      ('יתרת קרן לסגירה:', float(last['קרן יתרת סגירה'])),
      ('יתרת ריבית נצברת:', float(last['ריבית יתרת סגירה'])),
      ('יתרת הפרשי הצמדה (קרן + ריבית):', float(
          last['הפרשי הצמדה קרן יתרת סגירה']
          + last['הפרשי הצמדה ריבית יתרת סגירה']
      )),
      ('סה"כ חוב/יתרה מחושבת:', float(tot_debt)),
      ('', ''),
      ('מתודולוגיה ועקרונות החישוב במערכת:', ''),
      (
          '1. ניהול שכבות (Tranches):',
          'כל הפקדה או משיכה פותחת שכבת תזרים נפרדת הצוברת ריבית והצמדה באופן'
          ' עצמאי.',
      ),
      (
          '2. סדר קדימות הפירעון (Waterfall):',
          'סילוקים מנוכים לפי הסדר הבא: ריבית נצברת -> הפרשי הצמדה ריבית ->'
          ' הפרשי הצמדה קרן -> קרן.',
      ),
      (
          '3. מנוע ההצמדה למדד (CPI):',
          'ההצמדה מחושבת על בסיס מדד המחירים לצרכן הידוע. הצמדת הקרן מחושבת'
          ' מיום הבסיס/הסילוק האחרון, והצמדת הריבית מורכבת מקידום תקופתי של'
          ' היתרה הצמודה + הצמדת הריבית החדשה שנצברה.',
      ),
      (
          '4. משיכות יתר (Overdraft):',
          'תזרימי משיכה העולים על סך היתרות הקיימות פותחים שכבת חוב שלילית'
          ' הצוברת ריבית והצמדה לפי אותם הפרמטרים.',
      ),
  ]

  for r_idx, (k, v) in enumerate(info_rows, start=3):
    ws_info.cell(row=r_idx, column=1, value=k).font = bold_font
    cell_v = ws_info.cell(row=r_idx, column=2, value=v)
    if isinstance(v, float):
      cell_v.number_format = '#,##0.00'

  start_tx_row = len(info_rows) + 5
  ws_info.cell(row=start_tx_row, column=1, value='פירוט נתוני התזרים שהועלו ונבדקו:').font = title_font

  tx_headers = ['תאריך תזרים', 'סוג תנועה', 'סכום (₪)']
  for c_idx, h in enumerate(tx_headers, start=1):
    cell = ws_info.cell(row=start_tx_row + 1, column=c_idx, value=h)
    cell.fill = header_fill
    cell.font = header_font

  for r_offset, (_, row_tx) in enumerate(df_tx.iterrows(), start=start_tx_row + 2):
    ws_info.cell(row=r_offset, column=1, value=pd.to_datetime(row_tx['תאריך']).strftime('%Y-%m-%d'))
    ws_info.cell(row=r_offset, column=2, value=str(row_tx['סוג תנועה']))
    c3 = ws_info.cell(row=r_offset, column=3, value=float(row_tx['סכום']))
    c3.number_format = '#,##0.00'

  ws_main = wb.create_sheet(title='טבלה מאוחדת')
  ws_main.sheet_view.rightToLeft = True

  df_main = pd.DataFrame(consolidated_history)[COLUMNS]
  for c_idx, col_name in enumerate(COLUMNS, start=1):
    cell = ws_main.cell(row=1, column=c_idx, value=col_name)
    cell.fill = header_fill
    cell.font = header_font

  for r_idx, row_data in enumerate(df_main.values, start=2):
    for c_idx, val in enumerate(row_data, start=1):
      cell = ws_main.cell(row=r_idx, column=c_idx, value=val)
      if c_idx == 1:
        cell.number_format = 'YYYY-MM-DD'
      elif c_idx in [2, 3]:
        cell.number_format = '#,##0.00' if c_idx == 2 else '#,##0'
      else:
        cell.number_format = '#,##0.00'

  for t_id, t in tables.items():
    label = (
        f'הפקדה_{t_id}'
        if t.get('type') == 'deposit'
        else f'משיכת_יתר_{t_id}'
    )
    ws_t = wb.create_sheet(title=label)
    ws_t.sheet_view.rightToLeft = True
    df_t = pd.DataFrame(t['rows'])[COLUMNS]

    for c_idx, col_name in enumerate(COLUMNS, start=1):
      cell = ws_t.cell(row=1, column=c_idx, value=col_name)
      cell.fill = header_fill
      cell.font = header_font

    for r_idx, row_data in enumerate(df_t.values, start=2):
      for c_idx, val in enumerate(row_data, start=1):
        cell = ws_t.cell(row=r_idx, column=c_idx, value=val)
        if c_idx == 1:
          cell.number_format = 'YYYY-MM-DD'
        elif c_idx in [2, 3]:
          cell.number_format = '#,##0.00' if c_idx == 2 else '#,##0'
        else:
          cell.number_format = '#,##0.00'

  for ws in wb.worksheets:
    ws.freeze_panes = 'A2' if ws.title == 'עקרונות והנחות חישוב' else 'B2'
    for col in ws.columns:
      max_len = max(len(str(cell.value or '')) for cell in col)
      col_letter = openpyxl.utils.get_column_letter(col[0].column)
      ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

  wb.save(output_path)

  summary = {
      'as_of_date': as_of_date_str,
      'last_cpi_val': last_cpi_val,
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
      'tx_count': len(df_tx),
  }

  return json.dumps(summary, ensure_ascii=False)
