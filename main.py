import openpyxl 
from pathlib import Path
import re
from ConfigManager.config_manager import ConfigManager

def set_column_names_to_index(sheet, column_names):

    column_index_map = {}
    
    for item in column_names:
        column_index_map.setdefault(item)

    flag_reached_title_row = False

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

def get_data_by_columns_name(path, sheet_name, columns_title, regexPattern):
    wb_obj = openpyxl.load_workbook(path) 
    sheet_obj = wb_obj[sheet_name]
    columns_index = set_column_names_to_index(sheet_obj, columns_title)

    data = {}

    for row in sheet_obj.iter_rows(min_row=2, min_col=0, max_row=sheet_obj.max_row, max_col=sheet_obj.max_column):
        for column_name, column_index in columns_index.items():
            if not column_name in data.keys():
                data[column_name]= []

            cell_value = row[column_index].value
            if cell_value:
                pattern = re.compile(regexPattern)
                cell_value = "".join(row[column_index].value.rstrip().lstrip())
                
                if pattern.match(cell_value):
                    data[column_name].append(cell_value)

    wb_obj.close()

    return data

def main():

    config_manager = ConfigManager("project_config.toml")

    if config_manager.error != 0:
        print("***********************************************************/n \
              error in id config file. stoping process \
              ************************************************************")
        exit()
    
    # config  general
    config_general = config_manager.general_config

    root_data = config_general.root_dir

    # config points id 

    ids_config = config_manager.point_id_config

    title = []
    title.append(ids_config.id_column_title)

    id_file_path = root_data / ids_config.id_path

    #process points id sheet

    point_ids = get_data_by_columns_name(id_file_path , ids_config.sheet_name, title, ids_config.id_pattern)

    ids = { item : None for item in point_ids[ids_config.id_column_title] }
    
    # config  source data

    source_data_config = config_manager.source_data_config

    source_data_path = root_data / source_data_config.data_path

    # config  target data

    target_data_config = config_manager.target_data_config
    
    target_folder = root_data / target_data_config.folder_path

    #process data sheet

    wb_obj = openpyxl.load_workbook(source_data_path) 

    target_wb = openpyxl.Workbook()

    for sheet_name in source_data_config.sheets_names:

        columns_names = source_data_config.sheet_to_columns[sheet_name]
        data = get_filtered_data(wb_obj, sheet_name, columns_names, ids, source_data_config.id_column_title)

        target_wb.create_sheet(sheet_name)
        sheet = target_wb[sheet_name]
        column = 1

        for title_column, values in data.items():
            row = 1
            sheet.cell(row, column, title_column) # first row for titles
            row += 1
            for val in values: 
                sheet.cell(row=row, column=column, value=val)
                row +=1
            column +=1 

    wb_obj.close()

    target_wb.remove(target_wb.active) #remove auto created sheet

    target_folder.mkdir(parents=True, exist_ok=True)

    target_file = Path(target_data_config.file_prefix + "_" + ids_config.id_column_title +  source_data_path.suffix)
    target_wb.save(target_folder / target_file)

    print(  "***********************************************************\n \
Processed Succesfully Put your hands up in the air\n \
************************************************************")



main()