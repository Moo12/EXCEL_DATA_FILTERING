import toml
from pathlib import Path

class ConfigSegment:
    def __init__(self):
        self.error = 0

class ConfigGeneral(ConfigSegment):
    def __init__(self, raw_data):
        super().__init__()

        self.general_main_entry = "General"
        try:
                
            self.root_dir = Path(raw_data[self.general_main_entry]['root_dir'])
            self.success_strig = raw_data[self.general_main_entry]['success_strig']
        except Exception as e:
            print(e)
            self.error = 1

class ConfigPointsIdsManager(ConfigSegment):
    ids_main_entry = 'IdsSheet'

    def __init__(self, raw_data):
        super().__init__()

        try:
                
            self.id_path = Path(raw_data[self.ids_main_entry]['path'])

            self.id_column_title = raw_data[self.ids_main_entry]['id_column_title']
            
            self.sheet_name = raw_data[self.ids_main_entry]['sheet_name']

            self.id_pattern = raw_data[self.ids_main_entry]['id_pattern']
        
        except Exception as e:
            print(e)
            self.error = 1 

class ConfigDataSourceManager(ConfigSegment):
    data_main_entry = 'SourceDataSheets'

    def __init__(self, raw_data):
        super().__init__()

        try:
            self.data_path = Path(raw_data[self.data_main_entry]['data_path'])

            self.id_column_title = raw_data[self.data_main_entry]['id_column_title']
            self.sheets_names = raw_data[self.data_main_entry]["sheet_names"]

            self.sheet_to_columns = {}

            for sheet in self.sheets_names:
                self.sheet_to_columns[sheet] = raw_data[self.data_main_entry][sheet]

        except Exception as e:
            print("error ", e)
            self.error = 1

class ConfigDataDestinationManager(ConfigSegment):
    data_main_entry = 'DestinationDataSheets'

    def __init__(self, raw_data):
        super().__init__()
        try:
            self.folder_path = Path(raw_data[self.data_main_entry]['folder_path'])
            self.file_prefix = raw_data[self.data_main_entry]['file_prefix']
            self.sheets = raw_data[self.data_main_entry]['Sheets']

            if not isinstance(self.sheets, dict):
                self.sheets = {}
                print("destination sheets is empty")

        except Exception as e:
            print("error ", e.__str__())

            if  e.__str__() != "\'Sheets\'":
                self.error = 1
                print("error in Sheets key")
            else:  #bypass for default value
                print("destination sheets is empty")
                self.sheets = {}

class ConfigManager(): 
    
    def __init__(self, file_name):
        self.error = 0
        try:
            with open(file_name, 'r') as f:
                self.config_obj = toml.load(f)

                f.close()
        except Exception as e:
            self.error = 1 

            print(f"error: {e} {file_name}")
            return

        self.general_config = ConfigGeneral(self.config_obj)
        self.point_id_config = ConfigPointsIdsManager(self.config_obj)
        self.source_data_config = ConfigDataSourceManager(self.config_obj)
        self.target_data_config = ConfigDataDestinationManager(self.config_obj)

        if self.general_config.error != 0 or self.point_id_config.error or 0 | self.source_data_config.error or 0:
            print("error var")
            self.error = 1



