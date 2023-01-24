import openpyxl
import os
import pandas as pd


def write_to_file(string_data, file_path="logfile.txt"):
    with open(file_path, 'a+') as file:
        file.writelines(string_data)


def get_all_excel_filenames(folder_path):
    files = os.listdir(folder_path)
    excel_files = []
    for f in files:
        if f.endswith('.xlsx'):
            excel_files.append(f)
    write_to_file(f"Found {len(excel_files)} excel files to evaluate. \n\n\n\n ")
    return excel_files


def check_value(cell_address, worksheet, log=True):
    if log:
        write_to_file(f" \n {worksheet[cell_address].value} \n ")
    return worksheet[cell_address].value


def check_data_type(cell_address, worksheet):
    write_to_file(f" \n {worksheet[cell_address].data_type} (__data_type) \n ")
    return worksheet[cell_address].data_type


def empty_data_structure():
    data_structure = dict()
    data_structure['Name'] = None
    data_structure['Roll Number'] = None
    data_structure['Result'] = 0
    data_structure['Error Sum'] = 0
    data_structure['Value Sum'] = 0
    data_structure['Formula Sum'] = 0

    return data_structure


def check_specific_answers(empty_data, column_name, row_number, number_of_question, formula_worksheet, value_worksheet,
                           values_answer_key, formula_answer_key):
    for i in range(number_of_question):
        write_to_file(f"{'-' * 4} \n \n Checking question number: {i + 1}")

        cell_address = column_name + str(row_number + i)
        write_to_file(f"        checking cell:     [{cell_address}]  \n ")

        write_to_file(f"Checking formula output for any errors  \n ")
        value = check_value(cell_address, value_worksheet, False)
        if value in (
                "#N/A", "#DIV/0", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#####",
                "Circular Reference"):
            write_to_file(
                f"Error: Not even checking the formula because it shows >>{check_value(cell_address, value_worksheet)}<<  \n \n\n")
            empty_data['Error Sum'] += 1
            write_to_file(f"Your ERROR count increased: {empty_data['Error Sum']}  \n ")
            continue

        if cell_address in values_answer_key:
            if value == values_answer_key[cell_address]:
                empty_data['Value Sum'] += 1
                write_to_file(f"Your VALUE count increased: {empty_data['Value Sum']}  \n ")
            else:
                write_to_file(f"   Warning: Output must be as specified in assignment. \n ")
                continue
        else:
            write_to_file(f"          No output compulsion for this cell. \n ")

        if cell_address in formula_answer_key:
            formula_string = check_value(cell_address, formula_worksheet, False)
            try:
                formula_string = str(formula_string)
            except:
                write_to_file("Could not convert formula to string.  \n ")
                continue

            expected_result = formula_answer_key[cell_address]

            if expected_result.upper() in formula_string.upper():
                empty_data['Formula Sum'] += 1
                write_to_file(f"Your FORMULA count increased: {empty_data['Formula Sum']}  \n ")
            else:
                write_to_file(
                    f"   Warning: You should have used ={expected_result}(...) function but you used {formula_string}.\n ")
                continue
        else:
            write_to_file(f"          No formula compulsion for this cell. \n ")

        if cell_address in values_answer_key or cell_address in formula_answer_key:
            empty_data['Result'] += 1
            write_to_file(f"#####Your RESULT count increased: {empty_data['Result']} ###### \n ")

    return empty_data


def evaluate_excel_file(file_path, column_name, row_number, number_of_question, values_answer_key, formula_answer_key,
                        worksheet_name='Answers', student_name_cell='B1', student_roll_no_cell='B2'):
    write_to_file(f"{'~' * 80} \n\n \n ")

    if file_path:
        file_name = file_path.split('/')[-1]
        write_to_file(f" \n \n {'+' * 25} \n Working with {file_name}   \n \n")
    else:
        write_to_file("No file_path  \n ")
        return

    try:
        formula_wb = openpyxl.load_workbook(file_path)
    except:
        write_to_file(f"Could not find the {file_path}  \n ")
        return

    try:
        formula_worksheet = formula_wb[worksheet_name]
    except KeyError:
        write_to_file("sheet named 'Answers' must be there and contain answers.  \n ")
        return

    try:
        value_wb = openpyxl.load_workbook(file_path, data_only=True)
    except:
        write_to_file(f"Could not find the {file_path}  \n ")
        return

    try:
        value_worksheet = value_wb[worksheet_name]
    except KeyError:
        write_to_file("sheet named 'Answers' must be there and contain answers.  \n ")
        return

    try:
        name = check_value(student_name_cell, formula_worksheet)
    except:
        write_to_file(f'Could not find your official Name in B1.  \n ')
        return

    try:
        roll_no = check_value(student_roll_no_cell, formula_worksheet)
    except:
        write_to_file(f'Could not find your roll number in B2.  \n ')
        return

    result_data = check_specific_answers(empty_data_structure(), column_name, row_number, number_of_question,
                                         formula_worksheet, value_worksheet, values_answer_key, formula_answer_key)
    result_data['Name'] = name
    result_data['Roll Number'] = roll_no

    write_to_file(str(result_data))
    write_to_file(f"{'*' * 80} \n\n \n ")

    return result_data


class CheckExcelFiles:
    def __init__(self, folder_path, formula_answer_key, values_answer_key):
        self.values_answer_key = values_answer_key
        self.path = folder_path
        self.formula_answer_key = formula_answer_key

    def evaluate_all_excel_files(self, column_name, row_number, number_of_question):
        result_data_list = []
        excel_filenames = get_all_excel_filenames(self.path)

        for filename in excel_filenames:
            file_path = os.path.join(self.path, filename)
            data = evaluate_excel_file(file_path, column_name, row_number, number_of_question, self.values_answer_key,
                                       self.formula_answer_key)
            if data:
                result_data_list.append(data)
        df = pd.DataFrame.from_records(result_data_list)

        return df
