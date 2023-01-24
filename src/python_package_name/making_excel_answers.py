import openpyxl

from python_package_name.excel_assignment_checker import check_value


def create_answer_key(excel_file_path, column_name, row_number, number_of_question, worksheet_name='Answers'):
    formula_wb = openpyxl.load_workbook(excel_file_path)
    formula_ws = formula_wb[worksheet_name]

    value_wb = openpyxl.load_workbook(excel_file_path, data_only=True)
    value_ws = value_wb[worksheet_name]

    formula_answer_key = create_formula_answers_loop(column_name, row_number, number_of_question, formula_ws)
    value_answer_key = create_answers_loop(column_name, row_number, number_of_question, value_ws)

    return formula_answer_key, value_answer_key


def create_answers_loop(column_name, row_number, number_of_question, worksheet):
    result = dict()
    for i in range(number_of_question):
        cell_address = column_name + str(row_number + i)
        result[cell_address] = check_value(cell_address, worksheet)

    return result


def create_formula_answers_loop(column_name, row_number, number_of_question, worksheet):
    result = dict()
    for i in range(number_of_question):
        cell_address = column_name + str(row_number + i)
        result[cell_address] = check_value(cell_address, worksheet).split("(")[0]

    return result
