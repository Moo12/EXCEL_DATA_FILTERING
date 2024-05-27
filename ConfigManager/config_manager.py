import toml
from pathlib import Path

class ConfigManager(): 
    
    def __init__(self, file_name):
        self.error = 0 
        try:
            with open(file_name, 'r') as f:
                self.config_obj = toml.load(f)

                f.close()
        except:
            self.error = 1 

            print(f"error reading file {file_name}")


class ConfigPointsIdsManager(ConfigManager):
    ids_main_entry = 'IdsSheet'

    def __init__(self):
        super().__init__("project_config.toml")

        try:
                
            self.id_path = Path(self.config_obj[self.ids_main_entry]['path'])

            self.id_column_title = self.config_obj[self.ids_main_entry]['id_column_title']
            
            self.sheet_name = self.config_obj[self.ids_main_entry]['sheet_name']

            self.id_pattern = self.config_obj[self.ids_main_entry]['id_pattern']
        except:
            self.error = 1 

class ConfigDataSourceManager(ConfigManager):
    data_main_entry = 'DataSheets'

    def __init__(self):
        super().__init__("project_config.toml")

        try:
            self.path = Path(self.config_obj[self.data_main_entry]['source_data_path'])
            self.target_path = self.config_obj[self.data_main_entry]['destination_data_path_prefix'] + "_" + self.path.stem + self.path.suffix
            self.id_column_title = self.config_obj[self.data_main_entry]['id_column_title']
            self.sheets_names = self.config_obj[self.data_main_entry]["sheet_names"]

            self.sheet_to_columns = {}

            for sheet in self.sheets_names:
                self.sheet_to_columns[sheet] = self.config_obj[self.data_main_entry][sheet]

        except:
            self.error = 1 
