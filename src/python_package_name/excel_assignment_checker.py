import openpyxl
import os
import pandas as pd
from collections import defaultdict


class CheckExcelFiles:
    def __init__(self, folder_path, formula_answer_key, values_answer_key):
        self.values_answer_key = values_answer_key
        self.path = folder_path
        self.formula_answer_key = formula_answer_key


def get_all_excel_filenames(folder_path):
    files = os.listdir(folder_path)
    excel_files = []
    for f in files:
        if f.endswith('.xlsx'):
            excel_files.append(f)
    return excel_files


def evaluate_all_excel_files(folder_path, column_name, row_number, number_of_question, formula_answer_key=dict(),
                             values_answer_key=dict()):
    data = defaultdict(list)

    excel_filenames = get_all_excel_filenames(folder_path)

    for filename in excel_filenames:
        file_path = os.path.join(folder_path, filename)
        result_data = evaluate_excel_file(file_path, column_name, row_number, number_of_question, data,
                                          formula_answer_key, values_answer_key)
        df = pd.DataFrame(result_data)

    return df


def check_value(cell_address, worksheet):
    return worksheet[cell_address].value


def check_data_type(cell_address, worksheet):
    return worksheet[cell_address].data_type


def check_answers(column_name, row_number, number_of_question, formula_worksheet, value_worksheet,
                  formula_answer_key=dict(), values_answer_key=dict()):
    result = 0
    formula_sum = 0
    value_sum = 0
    error_sum = 0

    for i in range(number_of_question):
        cell_address = column_name + str(row_number + i)
        print(f"checking the {cell_address}")
        value = check_value(cell_address, value_worksheet)
        if value in (
                "#N/A", "#DIV/0", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#####",
                "Circular Reference"):
            print(
                f" Not checking even the formula because it shows {cell_address} value contains {check_value(cell_address, value_worksheet)}")
            error_sum += 1
            continue

        if cell_address in values_answer_key:
            print(f"This {cell_address} must be checked for values also")
            if value == values_answer_key[cell_address]:
                value_sum += 1
                print("value correct")
            else:
                continue

        print("checking the formula")

        if cell_address in formula_answer_key:
            print(f"This {cell_address} must be checked for formula also")
            if check_value(cell_address, formula_worksheet) is str and check_value(cell_address,
                                                                                   formula_worksheet).startswith(
                formula_answer_key[cell_address]):
                formula_sum += 1
                print("correct formula used")
            else:
                continue
        if not cell_address in values_answer_key and not cell_address in formula_answer_key:
            continue
        result += 1
    return result, error_sum, value_sum, formula_sum


def evaluate_excel_file(file_path, column_name, row_number, number_of_question, data=defaultdict(list),
                        formula_answer_key=dict(), values_answer_key=dict(), worksheet_name='Answers',
                        student_name_cell='B1', student_roll_no_cell='B2'):
    name, roll_no, result, error_sum, value_sum, formula_sum = None, None, None, None, None, None
    if file_path:
        print(f"working with {file_path}")
    else:
        print("No file_path")
    print("-" * 40)

    try:
        formula_wb = openpyxl.load_workbook(file_path)
    except:
        print(f"Could not find the {file_path}")
        return

    try:
        formula_ws = formula_wb[worksheet_name]
    except KeyError:
        print("sheet named 'Answers' must be there and contain answers.")
        return None

    try:
        value_wb = openpyxl.load_workbook(file_path, data_only=True)
    except:
        print(f"Could not find the {file_path}")

    try:
        value_ws = value_wb[worksheet_name]
    except KeyError:
        print("sheet named 'Answers' must be there and contain answers.")
        return None

    try:
        name = check_value(student_name_cell, formula_ws)
        roll_no = check_value(student_roll_no_cell, formula_ws)
    except:
        print(f'could not find  your official Name in B1 or  your roll number in B2.')
    print(name)
    print(roll_no)

    result, error_sum, value_sum, formula_sum = check_answers(column_name, row_number, number_of_question, formula_ws,
                                                              value_ws, formula_answer_key, values_answer_key)

    data['Name'].append(name)
    data['roll_no'].append(roll_no)
    data['result'].append(result)
    data['error_sum'].append(error_sum)
    data['value_sum'].append(value_sum)
    data['formula_sum'].append(formula_sum)

    print("*" * 40)

    return data
