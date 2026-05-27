#!/usr/bin/env python3
"""
Extract Rawat Jalan and Rawat Gigi premium data from Excel files.
Uses zipfile + xml.etree.ElementTree (built-in) since openpyxl is not available.
The xlsx files use inlineStr format (NOT shared strings).
"""

import zipfile
import xml.etree.ElementTree as ET
import json
import re

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def parse_xlsx(filepath):
    """Parse an xlsx file and return list of rows (each row is a list of cell values)."""
    rows = []
    with zipfile.ZipFile(filepath, 'r') as z:
        sheet_xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
        root = ET.fromstring(sheet_xml)
        for row_el in root.findall(f'.//{{{NS}}}row'):
            cells = row_el.findall(f'{{{NS}}}c')
            row_data = []
            for cell in cells:
                value = ''
                # Check for inlineStr
                is_el = cell.find(f'{{{NS}}}is')
                if is_el is not None:
                    t_el = is_el.find(f'{{{NS}}}t')
                    value = t_el.text if t_el is not None and t_el.text else ''
                else:
                    # Check for regular value
                    v_el = cell.find(f'{{{NS}}}v')
                    if v_el is not None and v_el.text:
                        value = v_el.text
                row_data.append(value)
            rows.append(row_data)
    return rows


def parse_numeric(val):
    """Convert a string value to integer. Return 0 for empty/invalid."""
    if not val or val.strip() == '':
        return 0
    # Remove commas, spaces
    val = val.strip().replace(',', '').replace(' ', '')
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return 0


def extract_konv_rj(rows):
    """
    Extract Rawat Jalan data from Konvensional Excel.
    Columns: [Jenis Kelamin, Usia, Kategori Manfaat, Tipe Plan, Diamond, Ruby, Emerald, Topaz, Topaz ID, Jade, Jade ID, Sapphire]
    Structure: 8 Normal + 7 Smart (Sapphire Smart is always empty so 7 usable smart values)
    For ages 80-85, use age 79 values (corrupted data).
    """
    # Group by age: collect Normal and Smart rows for "Rawat Jalan"
    rj_data = {}  # age -> {'Normal': [8 values], 'Smart': [7 values]}

    for row in rows[1:]:  # Skip header
        if len(row) < 12:
            continue
        kategori = row[2].strip() if row[2] else ''
        if kategori != 'Rawat Jalan':
            continue

        age = row[1].strip() if row[1] else ''
        tipe = row[3].strip() if row[3] else ''

        # Plan values: columns 4-11 (Diamond, Ruby, Emerald, Topaz, Topaz ID, Jade, Jade ID, Sapphire)
        values = [parse_numeric(row[i]) for i in range(4, 12)]

        if age not in rj_data:
            rj_data[age] = {}
        rj_data[age][tipe] = values

    # Build the final structure: {age: [8 Normal + 7 Smart]}
    # Smart has 8 columns but Sapphire Smart (last) is always empty = 7 usable
    result = {}
    for age_str in sorted(rj_data.keys(), key=lambda x: int(x)):
        age_int = int(age_str)
        normal = rj_data[age_str].get('Normal', [0]*8)
        smart = rj_data[age_str].get('Smart', [0]*8)
        # Take first 7 of Smart (exclude Sapphire Smart which is always 0/empty)
        combined = normal + smart[:7]
        result[age_str] = combined

    # For ages 80-85, use age 79 values (corrupted in Excel)
    if '79' in result:
        for age in range(80, 86):
            result[str(age)] = result['79'][:]

    return result


def extract_syariah_rj(rows):
    """
    Extract Rawat Jalan data from Syariah Excel.
    Use "Wanita" data for ages 0-79. For ages 80-85, use age 79 values.
    Structure: 8 Normal + 5 Smart (no Topaz ID Smart, Jade ID Smart, Sapphire Smart)
    """
    # Group by gender+age: collect Normal and Smart rows for "Rawat Jalan"
    rj_data = {}  # (gender, age) -> {'Normal': [8 values], 'Smart': [8 values]}

    for row in rows[1:]:  # Skip header
        if len(row) < 12:
            continue
        kategori = row[2].strip() if row[2] else ''
        if kategori != 'Rawat Jalan':
            continue

        gender = row[0].strip() if row[0] else ''
        age = row[1].strip() if row[1] else ''
        tipe = row[3].strip() if row[3] else ''

        values = [parse_numeric(row[i]) for i in range(4, 12)]

        key = (gender, age)
        if key not in rj_data:
            rj_data[key] = {}
        rj_data[key][tipe] = values

    # Use Wanita data as the primary source (ages 0-79)
    # For Pria age 0, use actual Pria data if available
    result = {}

    # First, get Wanita data for all ages
    for (gender, age_str), data in rj_data.items():
        if gender == 'Wanita':
            normal = data.get('Normal', [0]*8)
            smart = data.get('Smart', [0]*8)
            # Syariah Smart: only 5 usable (Diamond, Ruby, Emerald, Topaz, Jade)
            # Indices 0-4 from smart, but Topaz ID Smart (idx 4), Jade ID Smart (idx 6), Sapphire Smart (idx 7) are empty
            # So usable smart values: Diamond(0), Ruby(1), Emerald(2), Topaz(3), Jade(5)
            smart_5 = [smart[0], smart[1], smart[2], smart[3], smart[5]]
            combined = normal + smart_5
            result[age_str] = combined

    # For ages 80-85, use age 79 values
    if '79' in result:
        for age in range(80, 86):
            result[str(age)] = result['79'][:]

    return result


def extract_syariah_gigi(rows):
    """
    Extract Rawat Gigi data from Syariah Excel.
    Same structure as Syariah RJ.
    """
    rj_data = {}

    for row in rows[1:]:
        if len(row) < 12:
            continue
        kategori = row[2].strip() if row[2] else ''
        if kategori != 'Rawat Gigi':
            continue

        gender = row[0].strip() if row[0] else ''
        age = row[1].strip() if row[1] else ''
        tipe = row[3].strip() if row[3] else ''

        values = [parse_numeric(row[i]) for i in range(4, 12)]

        key = (gender, age)
        if key not in rj_data:
            rj_data[key] = {}
        rj_data[key][tipe] = values

    # Use Wanita data as primary source
    result = {}
    for (gender, age_str), data in rj_data.items():
        if gender == 'Wanita':
            normal = data.get('Normal', [0]*8)
            smart = data.get('Smart', [0]*8)
            # Same smart structure as RJ: usable = Diamond(0), Ruby(1), Emerald(2), Topaz(3), Jade(5)
            smart_5 = [smart[0], smart[1], smart[2], smart[3], smart[5]]
            combined = normal + smart_5
            result[age_str] = combined

    # For ages 80-85, use age 79 values
    if '79' in result:
        for age in range(80, 86):
            result[str(age)] = result['79'][:]

    return result


def format_js_object(data, var_name):
    """Format data as a JavaScript object literal."""
    # data is {age_str: [values]}
    # Output: var_name = {PRIA: {age: [values]}, WANITA: {age: [values]}}
    # Use same data for both PRIA and WANITA
    lines = []
    lines.append(f'const {var_name} = {{"PRIA":{{')

    ages_sorted = sorted(data.keys(), key=lambda x: int(x))
    age_entries = []
    for age in ages_sorted:
        vals = ','.join(str(v) for v in data[age])
        age_entries.append(f'"{age}":[{vals}]')

    lines.append(','.join(age_entries))
    lines.append('},"WANITA":{')
    lines.append(','.join(age_entries))
    lines.append('}};')

    return ''.join(lines)


if __name__ == '__main__':
    # Parse Konvensional Excel
    konv_rows = parse_xlsx('Data_Premi_Lengkap.xlsx')
    print(f"Konvensional rows: {len(konv_rows)}", flush=True)

    # Parse Syariah Excel
    syariah_rows = parse_xlsx('Data_Premi_Syariah_Lengkap.xlsx')
    print(f"Syariah rows: {len(syariah_rows)}", flush=True)

    # Extract data
    konv_rj = extract_konv_rj(konv_rows)
    syariah_rj = extract_syariah_rj(syariah_rows)
    syariah_gigi = extract_syariah_gigi(syariah_rows)

    print(f"\nKonv RJ ages: {len(konv_rj)}", flush=True)
    print(f"Syariah RJ ages: {len(syariah_rj)}", flush=True)
    print(f"Syariah Gigi ages: {len(syariah_gigi)}", flush=True)

    # Sample verification
    if '45' in konv_rj:
        print(f"\nKonv RJ age 45: {konv_rj['45']}", flush=True)
    if '45' in syariah_rj:
        print(f"Syariah RJ age 45: {syariah_rj['45']}", flush=True)
    if '45' in syariah_gigi:
        print(f"Syariah Gigi age 45: {syariah_gigi['45']}", flush=True)

    # Output JavaScript
    print("\n\n// === DATA_KONV_RJ ===", flush=True)
    print(format_js_object(konv_rj, 'DATA_KONV_RJ'), flush=True)

    print("\n// === DATA_SYARIAH_RJ ===", flush=True)
    print(format_js_object(syariah_rj, 'DATA_SYARIAH_RJ'), flush=True)

    print("\n// === DATA_SYARIAH_GIGI ===", flush=True)
    print(format_js_object(syariah_gigi, 'DATA_SYARIAH_GIGI'), flush=True)
