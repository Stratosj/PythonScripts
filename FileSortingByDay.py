import os
import time
from pathlib import Path
import shutil
import fnmatch

# TODO: Check if same logic applies as in Excel date comparison
# TODO: Is there any other file date accessible (encoding date? How to access it?)

FILE_LIST = os.listdir()
CURRENT_DIR = os.getcwd()

class Picture:
  
  
    def __init__(self, address, name, time_created, time_modified): # object constructer
        self.name = name # name of the file
        self.address = address # adress of the current working directory
        self.time_created = time_created # file time created
        self.time_modified = time_modified # file time modified
        self.oldest_time = self.find_oldest_time() # compares time created and time modified and returns the older one
 
 
    def find_oldest_time(self):
        if self.time_modified > self.time_created: 
            oldest_time = time.strftime('%Y%m%d', time.localtime(self.time_created)) # must use .localtime (if .gmtime is used it returns shifted value for some days)
            return oldest_time
        elif self.time_modified <= self.time_created: 
            oldest_time = time.strftime('%Y%m%d', time.localtime(self.time_modified))
            return oldest_time

    
    def move_to_directory(self):
        CHECK_FOLDER = os.path.isdir(self.oldest_time)
        if not CHECK_FOLDER:
            os.mkdir(self.oldest_time)
        Path(f"{self.address}\\{self.name}").rename(f"{self.address}\\{self.oldest_time}\\{self.name}") # Changes file position to folder.
        # TODO: Check if file exists in folder?
        

# TODO: Doesn't override or return any error if file is already in folder.

print("Are you sure you want to run this script? All the files in the same folder as the .exe will be sorted into folders by date.\nThere is no (easy) way to reverse this process.")
user_ready = input("Enter 'R' if you want to run this script or anything else to quit.")

if user_ready.lower() == "r":
    file_count = 0
    for file in FILE_LIST:
        if os.path.isfile(file) and ".py" not in file and ".exe" not in file: # prevents .py or .exe or folders from being sorted into folders
            file = Picture(address = CURRENT_DIR, name=file, time_created = os.path.getctime(file), time_modified = os.path.getmtime(file))
            file.move_to_directory()
            file_count += 1

input(f"{file_count} files sorted based on the oldest date found.")
