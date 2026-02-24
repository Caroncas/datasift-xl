from excel_reader import retrieving_section, checking_source, is_modified, is_original, split_section
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
    retrieve = retrieving_section(7, sheet)
    assert isinstance(retrieve, str)

    #Assertions:
    assert retrieving_section(7, sheet) == "6a"
    assert retrieving_section(5, sheet) == "3a"
    excel_sheet.active = 1
    sheet = excel_sheet.active
    assert retrieving_section(17, sheet) == "SECTION 2.1: TERMINOLOGY"
    assert retrieving_section(26, sheet) == "SECTION 2.2: SKELETAL & MUSCULAR SYSTEMS"

def test_checking_source():
    '''
    Parameters: source_row, sheet, original_source
    Return: string
    '''

def test_is_modified():
    '''
    Parameters: Cell_contents
    Return: bool
    '''
    mod = is_modified("modified")
    assert isinstance(mod, bool)

    #Assertions:
    assert is_modified("GSHS 1.2 modified") == True
    assert is_modified("GSHS 1.3 ModIfIed") == True
    assert is_modified("GSHS 1.3") == False
    assert is_modified("modified") == True
    assert is_modified("modified GSHS 1.3") == True
    assert is_modified("GSHS 1.2 mod") == False

def test_is_original():
    '''
    Parameters: Cell_contents
    Return: bool
    '''
    original = is_original("original")
    assert isinstance(original, bool)

    #Assertions:
    assert is_original("original") == True
    assert is_original("Original") == True
    assert is_original("ORIginal") == True
    assert is_original("GSHS 1.3") == False

def test_split_section():
    '''
    Parameters: section_line
    Return: section_number_string
    '''
    sect = split_section("Section 2.1: something something")
    assert isinstance(sect, str)

    #Assertions:
    assert split_section("Section 2.1: something something") == "2.1"
    assert split_section("a Section 2.1: something something") == "Section"
    assert split_section("Section 5.1: Another One Bites the Dust") == "5.1"

pytest.main(["-v", "--tb=line", "-rN", __file__])