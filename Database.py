from os import replace
import pyodbc
import pandas as pd
from Config import *
import csv
control = 0

def DataCollection(DataName, SelectedSite, StartTime, EndTime):
    server = f"Driver=SQL Server;Server={SelectedSite}MESSQLODS;Database=ODS;Trusted_Connection=Yes"
    datasource = sql_sourcecode_list[DataName]\
                .replace("__SelectedSite__", SelectedSite)\
                .replace("__StartTime__", StartTime)\
                .replace("__EndTime__", EndTime)
    connection = pyodbc.connect(server)
    connection.cursor()
    DataCollected = pd.read_sql_query(datasource, connection, index_col=None)
    print(DataCollected)
    return DataCollected

def Manipulating(DataCollected, DataName, **kwargs):
    if len(DataCollected.index) == 0:
        pass
    else:
        # Check master folder is existed or not then create master folder
        if os.path.isdir(Master_Folder):
            print(f"{Master_Folder} Storage Has Been Existed : {Master_Folder}")
        else:
            print(f"{Master_Folder} Storage Has Not Existed => Create A New Folder To Save Data")
            os.mkdir(Master_Folder)
        # Check folder CSV is existed or not then create new folder
        CSVFileDataStorage = location_drive_dict[DataName + "_CSVFile"].replace("__Drive_Folder__", drive_path)
        if os.path.isdir(CSVFileDataStorage):
            print(f"{CSVFileDataStorage} Storage Has Been Existed : {CSVFileDataStorage}")
        else:
            print(f"{CSVFileDataStorage} Storage Has Not Existed => Create A New Folder To Save Data")
            os.mkdir(CSVFileDataStorage)
        # Check folder JMP is existed or not then create new folder
        JMPFileDataStorage = location_drive_dict[DataName + "_JMPFile"].replace("__Drive_Folder__", drive_path)
        if os.path.isdir(JMPFileDataStorage):
            print(f"{JMPFileDataStorage} Storage Has Been Existed : {JMPFileDataStorage}")
        else:
            print(f"{JMPFileDataStorage} Storage Has Not Existed => Create A New Folder To Save Data")
            os.mkdir(JMPFileDataStorage)
        # Check folder IMG is existed or not then create new folder
        IMGFileDataStorage = location_drive_dict[DataName + "_IMGFile"].replace("__Drive_Folder__", drive_path)
        if os.path.isdir(IMGFileDataStorage):
            print(f"{IMGFileDataStorage} Storage Has Been Existed : {IMGFileDataStorage}")
        else:
            print(f"{IMGFileDataStorage} Storage Has Not Existed => Create A New Folder To Save Data")
            os.mkdir(IMGFileDataStorage)
        # Prevent multi varible has name is same and impact to name of file saved
        datasaving = os.path.join(CSVFileDataStorage, DataName + f"{kwargs['ext_name'] if 'ext_name' in kwargs.keys() else ''}.csv") 
        root, *_ = os.walk(CSVFileDataStorage)
        img_list = [_ for _ in root[-1] if _.endswith(".png") or _.endswith(".jpg") or _.endswith(".jpeg")]
        for img in img_list:
            os.remove(os.path.join(CSVFileDataStorage, img))
        if DataName + f"{kwargs['ext_name'] if 'ext_name' in kwargs.keys() else '' }.csv" in root[-1]:
            os.remove(os.path.join(CSVFileDataStorage, DataName + f"{kwargs['ext_name'] if 'ext_name' in kwargs.keys() else ''}.csv"))
        img_list = []
        DataCollected.to_csv(datasaving, index=True)
    
    
    
