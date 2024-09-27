import re
from openpyxl.utils import range_boundaries, get_column_letter, column_index_from_string

# CELLS
def extract_cells(formula):
    cell_pattern = r'(?<![:$\w])(\$?[A-Z]+\$?\d+)(?![:$\w])'
    cell_refs = [re.sub(r'\$', '', ref) for ref in re.findall(cell_pattern, formula)]
    return cell_refs    

def cell_to_context(cell, cell_lookup):
    if cell not in cell_lookup:
        return f"*{cell}* is not found in the cell lookup table"

    cell_dict = cell_lookup[cell]
    context = f"*{cell}* points to row: '{cell_dict['row_descrip']}' in col: '{cell_dict['col_descrip']}' in table: '{cell_dict['title']}'"
    return context

# RANGES
def extract_ranges(formula):
    range_pattern = r'(?<![:$\w])(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)(?![:$\w])'
    range_refs = [re.sub(r'\$', '', ref) for ref in re.findall(range_pattern, formula)]
    return range_refs

def range_to_context(range_str, cell_lookup):
    min_col, min_row, max_col, max_row = range_boundaries(range_str)
    errors = []

    row_descrips = {}
    for row in range(min_row, max_row + 1):
        key = f"row_{row}"
        reference_cell = f'{get_column_letter(min_col)}{row}'
        if reference_cell not in cell_lookup:
            errors.append(f"*{reference_cell}* is not found in the cell lookup table during row descrip construction")
            continue
        value = cell_lookup[reference_cell]['row_descrip']
        row_descrips[key] = value

    col_descrips = {}
    for col in range(min_col, max_col + 1):
        key = f"col_{get_column_letter(col)}"
        reference_cell = f'{get_column_letter(col)}{min_row}'
        if reference_cell not in cell_lookup:
            errors.append(f"*{reference_cell}* is not found in the cell lookup table during col descrip construction")
        value = cell_lookup[reference_cell]['col_descrip']
        col_descrips[key] = value

    first_cell = f"{get_column_letter(min_col)}{min_row}"
    range_title = cell_lookup[first_cell]['title']

    context = f"<{range_str}>*{range_str}* is defined by the following row and column descriptions: {str(row_descrips)} {str(col_descrips)} in the broader table '{range_title}'</{range_str}>"

    if errors:
        print(errors)

    return context

# PARSE FORMULA
def formula_context(formula, cell_lookup):
    range_explanations = ""
    ranges = (extract_ranges(formula))
    for r in ranges:
        range_context = range_to_context(r,cell_lookup)
        range_explanations += "\n" + range_context
    
    cell_explanations = ""
    cells =  extract_cells(formula)
    for c in cells:
        cell_context = cell_to_context(c,cell_lookup)
        cell_explanations += "\n" + cell_context

    full_context = range_explanations+"\n"+cell_explanations
    return full_context

# Test cases
cell_lookup = {
    "A1": {"col_descrip": "Region", "row_descrip": "Header", "title": "Sales Data"},
    "A2": {"col_descrip": "Region", "row_descrip": "North America", "title": "Sales Data"},
    "A3": {"col_descrip": "Region", "row_descrip": "Europe", "title": "Sales Data"},
    "A4": {"col_descrip": "Region", "row_descrip": "Asia", "title": "Sales Data"},
    "B1": {"col_descrip": "Q1 Sales", "row_descrip": "Header", "title": "Sales Data"},
    "B2": {"col_descrip": "Q1 Sales", "row_descrip": "North America", "title": "Sales Data"},
    "B3": {"col_descrip": "Q1 Sales", "row_descrip": "Europe", "title": "Sales Data"},
    "B4": {"col_descrip": "Q1 Sales", "row_descrip": "Asia", "title": "Sales Data"},
    "C1": {"col_descrip": "Q2 Sales", "row_descrip": "Header", "title": "Sales Data"},
    "C2": {"col_descrip": "Q2 Sales", "row_descrip": "North America", "title": "Sales Data"},
    "C3": {"col_descrip": "Q2 Sales", "row_descrip": "Europe", "title": "Sales Data"},
    "C4": {"col_descrip": "Q2 Sales", "row_descrip": "Asia", "title": "Sales Data"},
    "D1": {"col_descrip": "Q3 Sales", "row_descrip": "Header", "title": "Sales Data"},
    "D2": {"col_descrip": "Q3 Sales", "row_descrip": "North America", "title": "Sales Data"},
    "D3": {"col_descrip": "Q3 Sales", "row_descrip": "Europe", "title": "Sales Data"},
    "D4": {"col_descrip": "Q3 Sales", "row_descrip": "Asia", "title": "Sales Data"}
    }
formula = "A1 + B2 + $C3+$A1:$C2 + INDEXMATCH(D3:D4) + SUM($A1:A3) + D3 + A1"

print(formula_context(formula, cell_lookup))

#step 1: get parse formula into list
#step 2: deal with $ or $
