import os
import pandas as pd

FILE_LIST = os.listdir(".") # Lists current directory when scanning for .xlsx files
CURRENT_DIR = os.getcwd() # Fetches current working directory to make script working in any computer

######################################## CHANGE SETTINGS HERE ########################################
TARGET_FOLDER_NAME = "Source" # Folder inside current directory where generated CSVs will be stored
ENCODING = "utf 16" # Encoding for generated CSV
NA_REP = "NA" # Replacing null (not available) values with this string (Currently not used)


class ExcelFileConverter(): # Class containing a constructor used for .xlsx files in folder so that we can work with them as objects 
    # file = ExcelFileConverter(address = CURRENT_DIR, name=file): -> Constructs an object for each .xlsx file, with .address attribute CURRENT_DIR and .name attribute = file name.


    def __init__(self, address, name): # Constructor method that will launch whenever we write code like this: nazev_objektu = ExcelFileConverter(address = "....", name = "...")
        self.name = name # name of the file - TODO: Is it necessary? Can code be written more clearly without it?
        self.address = address # adress of the current working directory
        self.xls_data_frame = pd.ExcelFile(self.name) # Sets up a data frame for each file
        self.sheets = self.xls_data_frame.sheet_names # Looks for all sheets of every data frame by .sheets_name function pointing to line above
        self.encode_sheets() # Executes encode_sheets method for each constructed object (method defined below)


    def make_original_folder():
        """Checks for "Original XLSX" folder and creates it if necessary."""
        if not os.path.isdir(TARGET_FOLDER_NAME): # Checks if TARGET_FOLDER_NAME exists in current working directory
            os.mkdir(TARGET_FOLDER_NAME) # Makes folder if TARGET_FOLDER_NAME does not exist


    def encode_sheets(self):
        """Renders all sheets of Excel file using PD's .sheet_names for DF
        and converts them to .csv UTF_16, replacing null values with 'NA'"""

        for sheet in self.sheets: # Loops through all sheets in object.sheets (=self.xls_data_frame.sheet_names)
            sheet_data_frame = pd.read_excel(self.xls_data_frame, sheet_name = sheet) # loads individual sheet in object.sheets as a data frame
            sheet_data_frame = sheet_data_frame.replace('\s', ' ', regex=True) # removes white spaces (such as line breaks) #TODO: Confirm everything it does.
            sheet_data_frame.to_csv(f"{self.address}\\{TARGET_FOLDER_NAME}\\{self.name}_{sheet}.csv", encoding=ENCODING, index=False) # add "na_rep=NA_REP" after encoding to include value for N/A values.
            # transcribes data frame to csv. 
    

terminate_program = False
while not terminate_program:

    print(f"This script will re-write any files already existing in the {TARGET_FOLDER_NAME}.")
    user_ready = input(f"Enter 'R' if you want to export .xlsx files as .csv using {ENCODING} encoding:\n")
    if user_ready.upper() == "R":

        ExcelFileConverter.make_original_folder()

        file_count = 0
        for file in FILE_LIST:
            if file.endswith('.xlsx'): 
                file = ExcelFileConverter(address = CURRENT_DIR, name=file) # constructs ExcelFileConverter object
                file_count += 1 # Adds 1 to file_count for each 

        print(f"{file_count} files encoded to CSV, {ENCODING}.")
    
    user_quits = input("Do you want to quit? Y/N:\n")
    if user_quits.upper() == "Y":
        terminate_program = True

input("""
       ..                                   ,--------.
      / /                                 ,' /.|    /
    RED'                                ,'    ||   /
   DWARF                                \     ||  /
  / /                                    \_____---.  _---.
  `'                                .--~~~     ---:,'   / \.
                   ._--~~~--_.    ,'           ---:     | |
    _--~~~--_    ,'   |    :  `. /   __---~~~~~---'`.   \ /
  ,:_     ;  `._/     | .  :  ..\--~~         ;   |__~---`
 //  ~~-_. ;   \=   '---'..:     i STARBUG  1  ;    :  \.
i '~~~~~~' ;    i=               |.....        ;    :...i
|     ::   ;    |=........            :        ;    :   |
`.    ::   ;### !=       :       |    :        ;        |
 \)       ;    /=        :      !     :.. _---_         |  
  `. 0  _--_ ,'~\ :=====;:      /        /     \       .!
    ~--/    \    `.`.___|:    ,:        |       |    ,'/
       \_--_/      `~--___--~'  \        \_---_/    / /
        \,./                     `.    :::\   /    /,'
         | |                       `-_.   |   | ,_-'
         ||:                           ~~~(  |'~
         ||:                               | |
         ||'                               | |
       '"~~~"`                           '"~~~"`   FRM

    """)