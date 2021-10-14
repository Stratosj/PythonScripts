import pandas as pd
import os

CHUNK_SIZE = 50000
FILE_LIST = os.listdir()
MERGED_FILE = "C://Users/cen68276/Documents/05 AdHoc/OrgChart/Source/result_merge/merged_file.csv"
OUTPUT_FILE = "C://Users/cen68276/Documents/05 AdHoc/OrgChart/Source/result_merge/output_file.csv"


def skip_headers():
    """if it is not the first csv file then skip the header row (row 0) of that file"""
    if not first_file: 
        return [0]
    else:
        return []


first_file = True
for csv_file in FILE_LIST:
    if os.path.isfile(csv_file) and ".py" not in csv_file and ".exe" not in csv_file: 
        chunk_container = pd.read_csv(csv_file, sep = ';', chunksize = CHUNK_SIZE, skiprows = skip_headers())
        for chunk in chunk_container:
            chunk.to_csv(MERGED_FILE, mode = 'a', sep = ';', index = False)
        first_file = False

SEMI_PRODUCT = pd.read_csv(MERGED_FILE, sep = ';')
SEMI_PRODUCT["Position Code_Employee"] = ""
SEMI_PRODUCT["Start Date_Position_Employee"] = ""
SEMI_PRODUCT.to_csv(OUTPUT_FILE, sep = ';', index = False)
