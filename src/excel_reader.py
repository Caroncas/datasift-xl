from openpyxl import load_workbook

excel_sheet = load_workbook(filename = "TestExcelSheet.xlsx")

excel_sheet.active = 0
# sheets = excel_sheet.active
# print(sheets.title)
# to print a specific cell: 
# print(sheet['A:C'], values_only=True)
same_source_lesson = []
modified_source_lesson = []
lesson_name = 'MyPlate'
SECTION_INDEX = 1
LESSON_INDEX = 2
SOURCE_INDEX = 3

for sheet in excel_sheet.worksheets: #iterates through the sheets
    print(sheet.title) #prints the name of the title
    for row in sheet.iter_rows(min_col=LESSON_INDEX, max_col=LESSON_INDEX):
        cell = row[0]
        if cell.value == lesson_name:
            print(f"{cell.value} fount at: \nRow: {cell.row}\nColumn: {cell.column}")
        #     print(Retrieving_Section(cell.row, sheet))
        # elif Is_Modified():
            
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
    source = " "
    cell = sheet.cell(row=row, column=column)
    while row > 0:
        if cell.value == None:
            row -= 1
        else:
            source = cell.value
            return source

def Checking_Source(source_row):
    cell = sheet.cell(row = source_row, column = SOURCE_INDEX)
    cell_contents = cell.value
    if Is_Original(cell_contents):
        return
    if Is_Modified(cell_contents):
        cell_pieces = cell_contents.split() 
        to_add_to_modified = f"{cell_pieces[0]} {cell_pieces[1]}"
        modified_source_lesson.append(to_add_to_modified)

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

"""
Ultimate goal:
1. Iterate through every sheet
    a. Will need to skip Master Page
    b. Will need to skip PHF quizzes
2. Each Course (sheet) will be stored in a list as its own dictionary
    a. Maybe there's an easier way to just search through each course code since we're in read-only mode?
        I. no that wouldn't work because of the ones that say original, we need them to also show up with the correct code
    b. Do we want each name to be stored AS THE NAME (ex. Advanced PE) or AS THE CODE (ex. ADV PE)?
3. Sections will be stored as dictionary items
4. A "section code" (ex. FF1 3.1) will be given in each section alongside a list of individual lessons
5. Each lesson will be its own dictionary containing its "source" or where it comes from
    a. We will need to create a code that will adjust the code if it is labled as "original" or "modified from" if it is modified, it is an original lesson
       so if it contains 'modified' OR 'original' it will be given the 'section code' as a source

Once sheet is saved into dictionaries:
1. Get user will input which lesson they want to see all other course lessons that use it
    a. UserInput for Course Name
    b. UserInput for section
    c. UserInput for lesson name
    d. get source code for specific section
2. iterate through courses list
3. >iterate through each section
4. >>iterate through lessons list
    a. if 'lessons' key == userinput_name then 
        I. >>>view source, if source == userinput_source_code
            1. print(current source code)
    

When searching it would first look for course name > section > lesson (if lesson name = same then > source)
    It won't need to search through its own course though, that should be omitted, because they won't be calling it anywhere else in that course except for where the lesson is
Need to figure out the user side of it, maybe have it linked to an HTML, the JavaScript will just run the python file if x or y is done, ya know? 
Will have to create the try and fail sequences so the program doesn't crash.
"""

'''
Example
courses = [
    'Advanced PE Combo' : {
        'introduction' : {
            'section code' : 'AdvPE intro'
            'lessons' : [
                'Course Introduction' = {
                    'source' : 'original'
                }, 
                'Course Tasks' : {
                    'source' : 'original'
                }   
            ]
            ...
        }
        '1.2' : {
            'section code' : 'AdvPE 1.3
            'lessons' : [
                'Getting Started' : {'source' : 'AdvPE1 1.2'}
                'Exercise Programming : {'source' : 'AdvPE 1.2'}
                'Mode' : {'source' : 'ES 5.1'}
            ]
        }
    }

]
'''