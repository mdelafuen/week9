
import openpyxl
import openpyxl.utils


def get_data_rows(file_name):
    excel_file = openpyxl.load_workbook(file_name)
    first_sheet= excel_file.active
    all_data = first_sheet.rows
    return all_data

def main():
    employment_data = get_data_rows("MAEmplyomentData.xlsx")
    for current_row in employment_data:
        town_cell = current_row[0]
        town_name = town_cell.value
        town_name = town_name[1:]
        population_data = get_data_rows("massachusetts_population_1980-2010.xlsx")
        for population_row in population_data:

            pop_town_name = population_row[3].value
            if pop_town_name is None:
                continue
            if town_name.lower() == pop_town_name.lower():
                labor_force = current_row[1].value
                pop_column_num = openpyxl.utils.cell.column_index_from_string("I")-1
                pop_number = population_row[pop_column_num].value
                part_rate = labor_force/pop_number*100
                print(f"{town_name} has {part_rate:.2f} % labor participation")


main()