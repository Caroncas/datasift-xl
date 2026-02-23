from excel_reader import Retrieving_Section, Checking_Source
import pytest
from openpyxl import load_workbook

excel_sheet = load_workbook(filename = "TestExcelSheet.xlsx")
# excel_sheet.active = 0
# sheet = excel_sheet.active

def test_retrieving_section():
    '''
    Parameters: row_number, sheet
    Return: lesson source
    '''
    excel_sheet.active = 0
    sheet = excel_sheet.active
    retrieve = Retrieving_Section(7, sheet)
    assert isinstance(retrieve, str)

    #Assertions:
    assert Retrieving_Section(7, sheet) == "6a"
    assert Retrieving_Section(5, sheet) == "3a"
    excel_sheet.active = 1
    sheet = excel_sheet.active
    assert Retrieving_Section(17, sheet) == "SECTION 2.1: TERMINOLOGY"
    assert Retrieving_Section(26, sheet) == "SECTION 2.2: SKELETAL & MUSCULAR SYSTEMS"

def test_checking_source():
    '''
    Parameters: source_row, sheet, original_source
    Return: string
    '''

pytest.main(["-v", "--tb=line", "-rN", __file__])