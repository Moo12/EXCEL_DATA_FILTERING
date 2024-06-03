import  openpyxl
from  openpyxl.utils import get_column_letter 
from pathlib import Path
import re
from ConfigManager.config_manager import ConfigManager

def set_column_names_to_index(sheet, column_names):

    column_index_map = dict.fromkeys(column_names, 0)

    flag_reached_title_row = False #flag to skip all blanket

    for row in sheet.iter_rows(min_row=1, min_col=0, max_row=sheet.max_row, max_col=sheet.max_column):
        if not flag_reached_title_row:
            for column in range(sheet.max_column):
                if row[column].value is not None:
                    cell_value = "".join(row[column].value.rstrip().lstrip())
                    if cell_value in column_index_map.keys():
                        flag_reached_title_row = True
                        column_index_map[cell_value] = column
                        print("column title: ",cell_value, "column index: ", column)
        else:
            break

    return column_index_map

def get_filtered_data(wb_obj, sheet_name, columns_names, ids, match_id_column_title):

    point_sheet = wb_obj[sheet_name]

    column_index_map = set_column_names_to_index(point_sheet, columns_names)

    id_column = column_index_map[match_id_column_title]

    relevant_excel_data = {}

    for row in point_sheet.iter_rows(min_row=2, min_col=0, max_row=point_sheet.max_row, max_col=point_sheet.max_column): 
        if row[id_column].value is not None and "".join(row[id_column].value.rstrip().lstrip()) in ids.keys():
            for column_name, column_index in column_index_map.items():
                if not column_name in relevant_excel_data.keys():
                    relevant_excel_data[column_name]= []
                relevant_excel_data[column_name].append(row[column_index].value)

    return relevant_excel_data

def get_data_by_columns_name(wb_obj, sheet_name, columns_title, regexPattern):
    sheet_obj = wb_obj[sheet_name]
    columns_index = set_column_names_to_index(sheet_obj, columns_title)

    data = {}

    for row in sheet_obj.iter_rows(min_row=2, min_col=0, max_row=sheet_obj.max_row, max_col=sheet_obj.max_column):
        for column_name, column_index in columns_index.items():
            if not column_name in data.keys():
                data[column_name]= []

            cell_value = row[column_index].value
            #print(f"colour: {row[column_index].fill.start_color.index}")
            if cell_value:
                pattern = re.compile(regexPattern)
                cell_value = "".join(row[column_index].value.rstrip().lstrip())
                
                if pattern.match(cell_value):
                    data[column_name].append(cell_value)

    return data

def load_workbook(file_path, toPrint=False, error_string=""):
    try:
        wb_obj = openpyxl.load_workbook(file_path)
        return wb_obj

    except Exception as e:
        print(e)
        if toPrint:
            print(f"-----------------------------------------------------------\n \
Error in {file_path}. {error_string} \n \
-----------------------------------------------------------")
            return None


def main():

    config_manager = ConfigManager("project_config.toml")

    if config_manager.error != 0:
        print("\n-----------------------------------------------------------/n \
              error in id config file. stoping process \
              ------------------------------------------------------------")
        exit()
    
    # config  general
    config_general = config_manager.general_config

    root_data = config_general.root_dir

    # config points id 

    ids_config = config_manager.point_id_config

    title = []
    title.append(ids_config.id_column_title)

    id_file_path = root_data / ids_config.id_path

    wb_obj = load_workbook(id_file_path, True, "Exit Process")
    #process points id sheet
    if not wb_obj:
        exit()

    point_ids = get_data_by_columns_name(wb_obj, ids_config.sheet_name, title, ids_config.id_pattern)
    
    wb_obj.close()

    ids = { item : None for item in point_ids[ids_config.id_column_title] }
    
    # config  source data

    source_data_config = config_manager.source_data_config

    source_data_path = root_data / source_data_config.data_path

    # config  target data

    target_data_config = config_manager.target_data_config
    
    target_folder = root_data / target_data_config.folder_path

    #process data sheet

    wb_obj = load_workbook(source_data_path, True, "Exit Process")

    if not wb_obj:
        exit()

    target_wb = openpyxl.Workbook()

    for sheet_name in source_data_config.sheets_names:

        columns_names = source_data_config.sheet_to_columns[sheet_name]
        columns_names_target = dict.fromkeys(columns_names, "")

        formula_excels_saved_words = { "IF": 0}

        if sheet_name in target_data_config.sheets.keys():
            for columnName, formula in target_data_config.sheets[sheet_name].items():
                if not columnName in columns_names_target.keys():
                    columns_names_target.setdefault(columnName)
                
                formula_columns = re.findall(r'[A-Za-z]+[-_0-9]?[a-zA-Z]', formula)
                for column in formula_columns:
                    if not column in formula_excels_saved_words and column not in columns_names_target:
                        columns_names_target.setdefault(column)
                        list(columns_names).append(column)
                        print (f"column from formula add to list {column}")
                    print (f"column from formula. already exists {column}")

        data = get_filtered_data(wb_obj, sheet_name, columns_names, ids, source_data_config.id_column_title)

        target_wb.create_sheet(sheet_name)
        sheet = target_wb[sheet_name]

        sheet_column_to_letters = {}
        for i, title_column in enumerate(data.keys()):
            sheet_column_to_letters[title_column] = get_column_letter(i + 1)

        num_of_rows = len(data[source_data_config.id_column_title])
        print(f"num of rows {num_of_rows}")

        # replace fromulas column names with excel column letter
        if sheet_name in target_data_config.sheets.keys():
            for column_title, raw_formula in target_data_config.sheets[sheet_name].items():
                for source_column, column_letter in sheet_column_to_letters.items():
                    raw_formula = re.sub(source_column, column_letter, raw_formula)
                    target_data_config.sheets[sheet_name][column_title] = raw_formula
                # add destination columns if not exists in data
                if column_title not in data.keys():
                    print(f"title {column_title} formula {raw_formula}")
                    data[column_title] = range(num_of_rows)
                else:
                    print(f"title { column_title} exists in data keys. formula: {raw_formula}")

        # write data
        for column, (title_column, values) in enumerate(data.items(), 1):
            sheet.cell(1 , column, title_column) # first row for titles
            for row, val in enumerate(values, 2):
                if sheet_name in target_data_config.sheets.keys() and title_column in target_data_config.sheets[sheet_name].keys():
                    formula_raw = target_data_config.sheets[sheet_name][title_column]
                    formula = re.sub(r'(?<=[^A-Z])[A-Z](?=[^A-Z])',  lambda g: g.group(0) + f"{row}", formula_raw)

                    formula = f"={formula}"
                    sheet.cell(row=row, column=column, value=formula)

                    print(f"formula with row: {formula}")
                else:
                   sheet.cell(row=row, column=column, value=val)
        

    wb_obj.close()

    target_wb.remove(target_wb.active) #remove auto created sheet

    target_folder.mkdir(parents=True, exist_ok=True)

    target_file = Path(target_data_config.file_prefix + "_" + ids_config.id_column_title +  source_data_path.suffix) if target_data_config.file_prefix else \
                    Path(ids_config.id_column_title +  source_data_path.suffix)
    
    target_full_path = target_folder / target_file
    try:
        target_wb.save(target_full_path)
    except Exception as e:
        print(e)
        print(  f"\n***********************************************************\n \
    Error Saving file {target_full_path}\n \
    ************************************************************")
        exit()

    # success
    print(  f"\n***********************************************************\n \
Processed Succesfully {config_general.success_strig}\n \
************************************************************")

main()