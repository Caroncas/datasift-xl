from openpyxl import load_workbook

excel_sheet = load_workbook(filename = "TestExcelSheet.xlsx")
SECTION_INDEX = 1
LESSON_INDEX = 2
SOURCE_INDEX = 3

same_source_lesson = []
modified_source_lesson = []

def main():
    excel_sheet.active = 0
    # sheets = excel_sheet.active
    # print(sheets.title)
    # to print a specific cell: 
    # print(sheet['A:C'], values_only=True)

    #USER INPUT GETTERS:
    # course_name = input("Which course are you working with? ")
    # lesson_name = input(f"What lesson are you looking for in {course_name}? ") 
    lesson_name = 'MyPlate'

    for sheet in excel_sheet.worksheets: #iterates through the sheets
        print(sheet.title) #prints the name of the title
        for row in sheet.iter_rows(min_col=LESSON_INDEX, max_col=LESSON_INDEX):
            cell = row[0]
            if cell.value != None:
                if cell.value.lower() == lesson_name.lower():
                    print(f"{cell.value} fount at: \nRow: {cell.row}")
                    print(Retrieving_Section(cell.row, sheet))

            
'''
In terms of functions:
    - retrieve section number
    - row iterator
    - check against word
    - if word == user input
    - check SOURCE index if it matches add to list
        - is modified
        - is original
'''
def Retrieve_User_Section():

    return "Source of the user's input lesson"

def Retrieving_Section(row_number, sheet):
    #Getting the source of derived course
    row = row_number
    column = SECTION_INDEX
    cell = sheet.cell(row=row, column=column)
    while cell.value == None:
        if cell.value == None:
            row = row - 1
        else:
            break   
        cell = sheet.cell(row=row, column=column)
    return cell.value

def Checking_Source(source_row, sheet, original_source):
    cell = sheet.cell(row = source_row, column = SOURCE_INDEX)
    cell_contents = cell.value
    if Is_Original(cell_contents):
        return False
    # elif Is_Modified(cell_contents):
    #     cell_pieces = cell_contents.split() 
    #     to_add_to_modified = f"{cell_pieces[0]} {cell_pieces[1]}"
    #     modified_source_lesson.append(to_add_to_modified)
    #     return False
    elif cell_contents == original_source:
        return True
    else:
        return False

def Is_Modified(cell_contents):
    cell = cell_contents.strip().lower()
    split_cell = cell.split()
    if "modified" in split_cell:
        return True
    else: 
        return False

def Is_Original(cell_contents):
    cell = cell_contents.strip().lower()
    split_cell = cell.split()
    if "original" in split_cell:
        return True
    else: 
        return False
    
if __name__ == "__main__":
    main()

# to access the title of the sheet you can set the original sheet = excel_sheet.active then iterate through sheets and make "sheet.title" the dictionaries being stored
   

# column_values_list = []
# for cell in cell.column[3]:
#     print(cell.value, end=" ")

# def iter_excel_openpyxl(file: IO[bytes]) -> Iterator[dict[str, object]]:
#     workbook = openpyxl.load_workbook(file, read_only=True)
#     rows = workbook.active.rows
#     headers = [str(cell.value) for cell in next(rows)]
#     for row in rows:
#         yield dict(zip(headers, (cell.value for cell in row)))

# with open(excel_sheet, 'rb') as f:
#     rows = iter_excel_openpyxl(f)
#     row = next(rows)
#     print(row)
    

# if __name__ == "__iter_excel_openpyxl__":
#     inter_excel_openpyxl()