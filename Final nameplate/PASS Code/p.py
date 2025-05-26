import streamlit as st
import pandas as pd
import io
import re

# ------------------------- 1.  helper -------------------------
def is_positive_oltc(value: object) -> bool:
    """
    Return True only when content clearly indicates that an OLTC exists.
    Blank values and typical negative tokens return False.
    """
    if pd.isna(value):
        return False

    v = str(value).strip().lower()
    negative_tokens = {
        '', '-', '—', 'n/a', 'na', 'none', 'null', 'no', 'without oltc',
        'without', 'fixed', 'f', '0', 'off'
    }
    if v in negative_tokens:
        return False

    positive_kw = {'oltc', 'on‑load', 'on load', 'yes', 'y'}
    return any(kw in v for kw in positive_kw) or v not in negative_tokens


# ------------------------- 2.  POWTR‑CODE validation -------------------------
def validate_powtr_code(row: pd.Series) -> pd.Series:
    """Validate one row and return (is_correct, expected_code)."""
    current_class = str(row.get('Classification', '')).strip()

    # 1) Phase – assumed 3 unless you later supply it in your data.
    phase_char = '3'

    # 2) Voltage level – find the highest side voltage present
    voltage_char = 'M'          # default
    high_v = None
    for col in row.index:
        c = str(col).lower()
        if ('voltage' in c and ('high' in c or 'high side' in c)) \
           or col == 'Rated Voltage ( kV ).1':
            if pd.notna(row[col]) and row[col] not in {'High Side', 'Low Side'}:
                m = re.search(r'(\d+\.?\d*)', str(row[col]))
                if m:
                    high_v = float(m.group(1))
                    break

    # Fallback: take the maximum among any “voltage” columns
    if high_v is None:
        for col in row.index:
            if 'voltage' in str(col).lower():
                if pd.notna(row[col]) and row[col] not in {'High Side', 'Low Side', ''}:
                    m = re.search(r'(\d+\.?\d*)', str(row[col]))
                    if m:
                        v = float(m.group(1))
                        if high_v is None or v > high_v:
                            high_v = v

    # Decide the voltage character *unless* it exceeds 765 kV
    if high_v is not None:
        if high_v > 765:
            # New rule – over‑limit voltage: POWTR‑3‑OO (dash replaces voltage char)
            correct_code = 'POWTR-3-OO'
            return pd.Series([current_class == correct_code, correct_code])
        voltage_char = 'E' if high_v >= 345 else 'H' if high_v >= 100 else 'M' if high_v >= 1 else 'L'
    else:
        voltage_char = '-'   # unknown voltage

    # 3) Cooling type
    type_char = 'D'
    for col in row.index:
        if 'cooling' in str(col).lower() and pd.notna(row[col]):
            if 'O' in str(row[col]).upper():
                type_char = 'O'
            break

    # 4) Tap‑changer (OLTC)
    tap_char = 'F'
    has_oltc = False

    oltc_specific = [c for c in row.index if any(t in str(c).lower()
                      for t in ('oltc manuf', 'oltc type'))]
    for c in oltc_specific:
        if is_positive_oltc(row[c]):
            has_oltc = True
            break

    if not has_oltc and 'OLTC' in row.index and is_positive_oltc(row['OLTC']):
        has_oltc = True

    if not has_oltc:
        for c in row.index:
            if 'oltc' in str(c).lower() and c not in oltc_specific and c != 'OLTC':
                if is_positive_oltc(row[c]):
                    has_oltc = True
                    break

    if has_oltc:
        tap_char = 'O'

    correct_code = f'POWTR-{phase_char}{voltage_char}{type_char}{tap_char}'
    return pd.Series([current_class == correct_code, correct_code])


# ------------------------- 3.  Process whole DataFrame -------------------------
def process_excel(df: pd.DataFrame) -> pd.DataFrame:
    res = df.apply(validate_powtr_code, axis=1)
    df['Is_Correct'] = res[0]
    df['Correct_POWTR_CODE'] = res[1]

    # place the two new columns right after 'Classification'
    if 'Classification' in df.columns:
        cols = list(df.columns)
        cols.remove('Is_Correct')
        cols.remove('Correct_POWTR_CODE')
        idx = cols.index('Classification')
        cols[idx + 1:idx + 1] = ['Is_Correct', 'Correct_POWTR_CODE']
        df = df[cols]

    return df


# ------------------------- 4.  Streamlit UI -------------------------
st.title('POWTR‑CODE Validator')

st.write("""
ตรวจสอบรหัส **POWTR‑CODE** ตามเกณฑ์ (British spelling used):

1. **Phase** – first digit (normally 3)  
2. **Voltage level**  
   * E (345–765 kV) H (100–345 kV) M (1–100 kV) L (<1 kV)  
   * **If the high‑side voltage exceeds 765 kV → code becomes POWTR‑3‑OO**  
3. **Type** – O (oil‑immersed) D (dry‑type)  
4. **Tap‑changer** – O (with OLTC) F (without OLTC)
""")

uploaded = st.file_uploader('Upload an Excel file', ['xlsx', 'xls'])
if uploaded:
    try:
        df = pd.read_excel(uploaded)
        result = process_excel(df)

        st.subheader('Validation results')
        st.dataframe(result)

        # download button
        buff = io.BytesIO()
        with pd.ExcelWriter(buff, engine='openpyxl') as writer:
            result.to_excel(writer, index=False)
        buff.seek(0)
        st.download_button('Download validated file', buff,
                           'validated_powtr_codes.xlsx',
                           'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        # summary
        st.subheader('Summary')
        st.write(f"Total {len(result)} rows | Correct {result['Is_Correct'].sum()} | Incorrect {(~result['Is_Correct']).sum()}")

    except Exception as e:
        st.error(f'Error: {e}')