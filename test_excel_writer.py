import os
import shutil
import pandas as pd
import pytest
from ExcelWriter import ExcelWriter

# Sample data for testing
sample_data = [
    {"ID": 1, "Name": "Alice", "Salary": 70000},
    {"ID": 2, "Name": "Bob", "Salary": 80000},
    {"ID": 3, "Name": "Charlie", "Salary": 90000},
]

df_sample = pd.DataFrame(sample_data)

@pytest.fixture
def test_file():
    filename = "test_excel_writer.xlsx"
    if os.path.exists(filename):
        os.remove(filename)
    yield filename
    if os.path.exists(filename):
        os.remove(filename)
    if os.path.exists("test_excel_writer_backup.xlsx"):
        os.remove("test_excel_writer_backup.xlsx")
    if os.path.exists("./test_exports/test_export.xlsx"):
        shutil.rmtree("./test_exports")

def test_create_and_save_workbook(test_file):
    writer = ExcelWriter(test_file)
    writer.create_new_workbook("Test Sheet")
    writer.save()
    assert os.path.exists(test_file)

def test_write_dict_and_edit(test_file):
    writer = ExcelWriter(test_file)
    writer.create_new_workbook("Dict Sheet")
    writer.write_dict(sample_data)
    writer.edit_cell("B2", "Updated Alice")
    writer.save()
    writer.load_workbook()
    assert writer.sheet["B2"].value == "Updated Alice"

def test_write_and_append_dataframe(test_file):
    writer = ExcelWriter(test_file)
    writer.create_new_workbook("DF Sheet")
    writer.write_dataframe(df_sample)
    writer.save()
    writer.load_workbook()
    writer.create_new_sheet("Appended")
    writer.add_dataframe(df_sample)
    writer.save()
    assert writer.sheet.max_row > len(df_sample)

def test_add_single_dict(test_file):
    writer = ExcelWriter(test_file)
    writer.create_new_workbook("Add Dict")
    writer.write_dict(sample_data)
    writer.add_dict({"ID": 4, "Name": "Dana", "Salary": 95000})
    writer.save()
    assert writer.sheet.cell(row=5, column=2).value == "Dana"

def test_save_to_location(test_file):
    writer = ExcelWriter(test_file)
    writer.create_new_workbook("Save Location")
    writer.write_dict(sample_data)
    export_path = "./test_exports"
    export_file = "test_export.xlsx"
    writer.save_to_location(export_path, export_file)
    full_path = os.path.join(export_path, export_file)
    assert os.path.exists(full_path)

def test_edit_invalid_cell(test_file):
    writer = ExcelWriter(test_file)
    writer.create_new_workbook("Invalid Cell")
    with pytest.raises(ValueError):
        writer.edit_cell("ZZZ9999", "Oops")


def test_write_empty_dict(test_file):
    writer = ExcelWriter(test_file)
    writer.create_new_workbook("Empty Dict")
    with pytest.raises(IndexError):
        writer.write_dict([])


def test_write_empty_dataframe(test_file):
    writer = ExcelWriter(test_file)
    writer.create_new_workbook("Empty DF")
    empty_df = pd.DataFrame()
    writer.write_dataframe(empty_df)
    writer.save()
    assert writer.sheet.cell(row=1, column=1).value is None



def test_add_dict_missing_keys(test_file):
    writer = ExcelWriter(test_file)
    writer.create_new_workbook("Partial Dict")
    writer.write_dict(sample_data)
    partial_data = {"ID": 4, "Salary": 50000}  # Missing 'Name'
    writer.add_dict(partial_data)
    writer.save()
    # Expecting 'None' where 'Name' should be
    assert writer.sheet.cell(row=5, column=2).value is None
