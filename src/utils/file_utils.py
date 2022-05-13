import os
import glob 
import sys
from log.app_logger import AppLogger

class FileUtils(object):
    
    log = AppLogger.get_logger()
    
    @classmethod
    def create_folder(cls, folder_name):
        
        if not os.path.exists(folder_name):
            cls.log.info("Creating directory: " + folder_name)
            os.makedirs(folder_name)
        else:
            cls.log.info(folder_name + " already exists")
    
    @classmethod
    def delete_files_and_directory(cls, directory_name: str):
        try:
            files = glob.glob(directory_name + "/**/*", recursive=True)
            for f in files:
                os.remove(f)
                
            os.rmdir(directory_name)
        except Exception as e:
            cls.log.error(str(e))
            sys.exit()
            