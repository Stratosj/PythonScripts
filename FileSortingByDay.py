import os
import time
from pathlib import Path
import shutil
import fnmatch
from datetime import datetime
from PIL import Image
from PIL.ExifTags import TAGS


# TODO: Check if same logic applies as in Excel date comparison
# TODO: Adjust for video files (problematic - many formats, get sample files)

FILE_LIST = os.listdir()
CURRENT_DIR = os.getcwd()

class Picture:
  
  
    def __init__(self, address, name, time_created, time_modified): # object constructer
        self.name = name
        self.address = address
        self.time_created = time_created 
        self.time_modified = time_modified 
        self.fn = f"{self.address}\\{self.name}"
        self.time_taken = self.image_date(fn = self.fn)
        self.all_times = [self.time_created, self.time_modified, self.time_taken]
        self.oldest_time = self.find_oldest_time() # returns min date from all dates found
 

    def get_exif(self, fn):
        ret = {}
        i = Image.open(fn)
        info = i._getexif()
        for tag, value in info.items():
            decoded = TAGS.get(tag, tag)
            ret[decoded] = value
        return ret

    # TODO: Do this only for images, errors happen if presented with file like MP4
    # TODO: Test on other images that do not have date_taken
    # TODO: Find out if this is actualy local time or if it's possible to conver it (changing line 53 to .localtime does not work) 
    def image_date(self, fn): # -- gracefully stolen from: https://orthallelous.wordpress.com/2015/04/19/extracting-date-and-time-from-images-with-python/ 
        """Returns the date and time from image(if available)"""
        TTags = [('DateTimeOriginal', 'SubsecTimeOriginal'),  # when img taken
        ('DateTimeDigitized', 'SubsecTimeDigitized'),  # when img stored digitally
        ('DateTime', 'SubsecTime')]  # when img file was changed
        # for subsecond prec, see doi.org/10.3189/2013JoG12J126 , sect. 2.2, 2.3
        exif = self.get_exif(self.fn)
        for i in TTags:
            dat, sub = exif.get(i[0]), exif.get(i[1], 0)
            dat = dat[0] if type(dat) == tuple else dat  # PILLOW 3.0 returns tuples now
            sub = sub[0] if type(sub) == tuple else sub
            if dat != None: break  # got valid time
        if dat == None: return  # found no time tags
    
        # T = datetime.strptime('{}.{}'.format(dat, sub), '%Y:%m:%d %H:%M:%S.%f') # optional float
        T = time.mktime(time.strptime(dat, '%Y:%m:%d %H:%M:%S')) + float('0.%s' % sub)
        return T


    def find_oldest_time(self):
        oldest_time = time.strftime('%Y%m%d', time.localtime(min(self.all_times))) # must use .localtime (if .gmtime is used it returns shifted value for some days)
        return oldest_time


    def move_to_directory(self):
        CHECK_FOLDER = os.path.isdir(self.oldest_time)
        if not CHECK_FOLDER:
            os.mkdir(self.oldest_time)
        Path(f"{self.address}\\{self.name}").rename(f"{self.address}\\{self.oldest_time}\\{self.name}") # Changes file position to folder.
        # TODO: Check if file exists in folder?
        

# TODO: Doesn't override nor return any error if file is already in folder.

print("Are you sure you want to run this script? All the files in the same folder as the .exe will be sorted into folders by date.\nThere is no (easy) way to reverse this process.")
user_ready = input("Enter 'R' if you want to run this script or anything else to quit.")

if user_ready.lower() == "r":
    file_count = 0
    for file in FILE_LIST:
        if os.path.isfile(file) and ".py" not in file and ".exe" not in file: # prevents .py or .exe or folders from being sorted into folders
            file = Picture(address = CURRENT_DIR, name=file, time_created = os.path.getctime(file), time_modified = os.path.getmtime(file))
            file.move_to_directory()
            file_count += 1
            print(f"time_taken{file.time_taken}, time_created {file.time_created}, time_modified{file.time_modified}")

input(f"{file_count} files sorted based on the oldest date found in the file.")
