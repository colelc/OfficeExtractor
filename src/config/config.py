import os 
import socket
import sys
from dotenv import dotenv_values
from log.app_logger import AppLogger
from pathlib import Path

class Config(object):
    _instance = None
    config = None
    HOME = None
   
    def __init__(self):
        logger = AppLogger.get_logger()
        
        if self.HOME is None:
            #logger.info("Calculating host name and IP Address")
            try:
                self.host_name = socket.gethostname()
                self.host_ip = socket.gethostbyname(self.host_name)
                #logger.info(self.host_ip + " " + self.host_name)
            except Exception as e:
                logger.error(str(e))
                sys.exit(99)
            
            #logger.info("Calculating the project HOME directory")
            path = Path(os.path.curdir.__str__())            
            src_path = Path(os.path.abspath(os.path.join(path, os.pardir)))
            self.HOME = os.path.abspath(os.path.join(src_path, os.pardir))
            
            self.config = dotenv_values(os.path.join(self.HOME, "resources", "OfficeExtractor.env"))
            
            logger.info("CONFIGURATION BEGIN : ******************************************************");
            logger.info("These are the OfficeExtractor configuration values")
            for key,value in self.config.items():
                logger.info("CONFIGURATION: " + key + " -> " + value)
            logger.info("CONFIGURATION END   : ******************************************************");
            logger.info("The machine host name and IP is: " + self.host_ip + " " + self.host_name)
            logger.info("The project HOME directory is " + self.HOME)

            self.supported_file_extensions = self.get_property("supported.file.extensions")
            self.extract_spec_headers = self.get_property("extract.spec.headers")
    
    @staticmethod
    def project_home(self):
        return self.HOME
    
    @staticmethod
    def get_config(self):
        return self.config
    
    def get_property(self, key:str):
        return self.config.get(key)
    
    # Singleton pattern- ensure HOME variable is assigned but once
    def __new__(cls, *args, **kwargs):
        if not isinstance(cls._instance, cls):
            cls._instance = object.__new__(cls)
        
        return cls._instance

Config()
Config()
#obj = Config()
#obj2 = Config()
#print(Config().project_home())