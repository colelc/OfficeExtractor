import os
import stat
from pathlib import Path
from src.log.app_logger import AppLogger

class FileUtils(object):
    
    
    @classmethod
    def create_folder(cls, folder_name):
        log = AppLogger.get_logger()
        
        if not os.path.exists(folder_name):
            log.info("Creating directory: " + folder_name)
            os.makedirs(folder_name)
        else:
            log.info(folder_name + " already exists")
    