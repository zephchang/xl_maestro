from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill
import re
from openai import OpenAI
from dotenv import load_dotenv
import os
import json

# Load the workbook_map.json file
with open('workbook_map.json', 'r') as f:
    workbook_map = json.load(f)

load_dotenv('keys.env')
openai_api_key = os.getenv('OPENAI_API_KEY')
client = OpenAI(api_key=openai_api_key)


def parse_cell_reference(cell_ref):
    # Use regex to split the cell reference into letters and numbers
    match = re.match(r'([A-Za-z]+)(\d+)', cell_ref)
    if match:
        letters, numbers = match.groups()
        return letters, int(numbers)
    else:
        raise ValueError(f"Invalid cell reference: {cell_ref}")

def parse_cell_numbers(cell_ref):
    # Use regex to split the cell reference into letters and numbers
    match = re.match(r'([A-Za-z]+)(\d+)', cell_ref)
    if match:
        letters, numbers = match.groups()
        # Convert column letters to number
        column_number = sum((ord(char) - 64) * (26 ** i) for i, char in enumerate(reversed(letters.upper())))
        return column_number, int(numbers)
    else:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    
def parse_range_for_llm(range_start, range_end,formula_ws):
    start_col, start_row = parse_cell_reference(range_start)
    end_col, end_row = parse_cell_reference(range_end)

    for row in value_ws.iter_rows(min_row=start_row, max_row = end_row, min_col=start_col, max_col=end_col):
        for cell in row:
            print(f"Cell {cell.coordinate}:")
            print(f"   Value: {cell.value}")
            print(f"   Data type: {cell.data_type}")
            formula_cell = formula_ws[cell.coordinate] #testing committing from git directly more testing
            if formula_cell.data_type == 'f':
                print(f"    HAS FORMULA: Formula: {formula_cell.value}")
        print()

def l_to_n(column_letter):
    """
    Convert a column letter (e.g., 'A', 'B', 'AA') to its corresponding column number (e.g., 1, 2, 27).
    """
    column_number = 0
    for char in column_letter:
        column_number = column_number * 26 + (ord(char.upper()) - ord('A') + 1)
    return column_number

def n_to_l(column_number):
    """
    Convert a column number (e.g., 1, 2, 27) to its corresponding column letter (e.g., 'A', 'B', 'AA').
    """
    column_letter = ""
    while column_number > 0:
        column_number, remainder = divmod(column_number - 1, 26)
        column_letter = chr(65 + remainder) + column_letter
    return column_letter

def guess_cell_formula(cell, range_start, range_end, formula_ws):
    cell_col, cell_row = parse_cell_reference(cell)
    start_col_letter, start_row = parse_cell_reference(range_start)
    end_col_letter, end_row = parse_cell_reference(range_end)

    cell_col_number = l_to_n(cell_col)
    start_col_number = l_to_n(start_col_letter)
    end_col_number = l_to_n(end_col_letter)

    workbook_purpose = "This workbook calculates and analyzes the financial performance of a business in several scenarios, focusing on metrics like revenue, gross margin, variable profit, and fixed costs. It also computes pre-tax profit, net margins, and cash flows for three individuals, taking into account tax rates, debt payments, and profit-sharing percentages."

    col_headers = {} #for range end_col - start_col
    for col_number in range(start_col_number,end_col_number+1):
        cell_value = formula_ws.cell(row=start_row,column=col_number).value
        col_headers[f'col_{n_to_l(col_number)}'] = cell_value

    row_headers = {}
    for row_number in range(start_row, end_row + 1):
        cell_value = formula_ws.cell(row=row_number, column=start_col_number).value
        row_headers[f'row_{row_number}'] = cell_value

    cell_row_header = formula_ws.cell(row=cell_row,column=start_col_number).value
    cell_col_header = formula_ws.cell(row=start_row,column=cell_col_number).value

    prompt = f"""
    You are a keen-eyed excel expert who is skilled at catching subtle formula errors or typos.
    \n
    You are looking at a workbook which has the following purpose: {workbook_purpose}
    \n
    The workbook has the following rows and collumns:

    ROWS = {row_headers}

    COLUMNS = {col_headers}

    You are currently determining the correct formula for {cell}. {cell} is defined as {cell_row_header} (Row {cell_row}) in {cell_col_header} (Column {cell_col})

    What would you expect {cell}'s formula to be? Answer in 2 sentences.
    """
    
    messages = [{"role": "user", "content": prompt}]
    completion = client.chat.completions.create(
        model="gpt-4o",
        messages=messages
    )
    ai_guess = completion.choices[0].message.content
    messages.append({"role":"assistant", "content":ai_guess})
    messages.append({"role":"user", "content":f"The employee has attempted to write the formula, they wrote: {formula_ws.cell(row=cell_row, column=cell_col_number).value}\n\n\nDoes the employee's formula match with yours? If it doesn't match, why might that be? Does it need to be fixed? Respond YES it's fine to leave as is. or NO it needs to be fixed. Format you response as [Y] or [N]"})

    evaluation = client.chat.completions.create(
        model="gpt-4o",
        messages=messages
    )

    evaluation = evaluation.choices[0].message.content
    print(evaluation)

    match = re.search(r'\[(Y|N)\]', evaluation)
    if match:
        verdict = match.group(1)
    else:
        verdict = "Unknown"
    
    reasoning = f"AI GUESS: {ai_guess}\n\n EVALUATION: {evaluation}\n\nVERDICT: {verdict}\n\nPROMPT: {prompt}"

    return verdict, reasoning

def check_range(start_cell, end_cell, formula_ws):
    # Parse start and end cells
    start_col, start_row = parse_cell_reference(start_cell)
    end_col, end_row = parse_cell_reference(end_cell)

    # Convert column letters to numbers
    start_col_num = l_to_n(start_col)
    end_col_num = l_to_n(end_col)

    # Iterate over the range
    for row in range(start_row, end_row + 1):
        for col_num in range(start_col_num, end_col_num + 1):
            # Convert column number back to letter
            col_letter = n_to_l(col_num)
            cell = f"{col_letter}{row}"
            # Check if the cell contains a formula
            if formula_ws[cell].data_type == 'f':
                # Guess the formula for the cell
                verdict, reasoning = guess_cell_formula(cell, start_cell, end_cell, formula_ws)
                # Add the response as a comment
                cell_obj = formula_ws[cell]
                cell_obj.comment = Comment(text=reasoning, author="Maestro")
                if verdict == 'Y':
                    cell_obj.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")  # Green
                elif verdict == 'I':
                    cell_obj.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")  # Orange
                elif verdict == 'N':
                    cell_obj.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # Red
            else:
                pass

            print(f"{cell} processed")

        # Save the workbook after each row to balance between safety and performance
        formula_wb.save('output.xlsx')

    # Final save after processing all cells
    formula_wb.save('output.xlsx')
    print("Range check completed")


def semantic_map_table(config, modify_dict): #look I don't love that we're using mutation, but we need some way to deal with overwriting so I prefer this than some kind of dictionary combining utility function
    workbook = config['workbook']
    worksheet = config['worksheet']
    col_descriptors = config['col_descriptors']
    row_descriptors = config['row_descriptors']
    check_cell_range = config['check_cell_range']
    table_title = config['table_title']
    
    value_wb = load_workbook(workbook, data_only=True)
    ws = value_wb[worksheet]

    col_start_col, col_start_row, col_end_col, col_end_row = range_boundaries(col_descriptors)
    row_start_col, row_start_row, row_end_col, row_end_row = range_boundaries(row_descriptors)

    cells_start_col, cells_start_row, cells_end_col, cells_end_row = range_boundaries(check_cell_range)

    col_headers = {} #this should be {1: "decription", 2: "description"}
    row_headers = {} #this is also {1: "decription", 2: "description"}


    for col in range(col_start_col, col_end_col+1):
        col_headers[col] = ws.cell(row=col_start_row, column=col).value or "NULL"
    
    for row in range(row_start_row, row_end_row+1):
        row_headers[row] = ws.cell(row=row, column=row_start_col).value or "NULL"
    
    for row in range(cells_start_row,cells_end_row+1):
        for col in range(cells_start_col, cells_end_col+1):
            modify_dict[f"{get_column_letter(col)}{row}"] = {"col_descrip":col_headers[col], "row_descrip":row_headers[row],"title":table_title}
        #should throw an error if col/row already exists
    print("Table title:",table_title,"worksheet:",worksheet, "complete")

def semantic_map_workbook(workbook_map):
    workbook_tree = {}
    for worksheet in workbook_map["worksheets"]:
        worksheet_tree = {}
        for table in worksheet["tables"]:
            table_dic = {
                "workbook": workbook_map["wb_title"],
                "worksheet": worksheet["ws_title"],
                "table_title": table["title"],
                "col_descriptors": table["col_descriptors"],
                "row_descriptors": table["row_descriptors"],
                "check_cell_range": table["check_cell_range"]
            }
            semantic_map_table(table_dic, worksheet_tree)
        workbook_tree[worksheet["ws_title"]] = worksheet_tree
    return workbook_tree

        #ok so in theory, at this point we've gone through every worksheet and for each worksheet gone through every table and for each table constructed a dict and for each dict we added all of the cells to the semantic_map_table
semantic_map = semantic_map_workbook(workbook_map)

# Save the result to test.json
with open('test.json', 'w') as f:
    json.dump(semantic_map, f, indent=4)

print("Results saved to test.json")

print(semantic_map["S9-13, 29-36 | Ratio Summaries"]["G36"])