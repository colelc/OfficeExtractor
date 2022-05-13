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
            self.HOME = Path(os.path.abspath(os.path.join(path, os.pardir)))
             
            self.config = dotenv_values(os.path.join(self.HOME, "resources", "OfficeExtractor.env"))
            
            logger.info("CONFIGURATION BEGIN : ******************************************************");
            logger.info("These are the OfficeExtractor configuration values")
            for key,value in self.config.items():
                logger.info("CONFIGURATION: " + key + " -> " + value)
            logger.info("CONFIGURATION END   : ******************************************************");
            logger.info("ENVIRONMENT: The machine host name and IP is: " + self.host_ip + " " + self.host_name)

            self.supported_file_extensions = self.get_property("supported.file.extensions").replace(" ", "").split(",")#.split(",") if str else list()
            
            self.header_message = Config().get_property("header.message")
            self.header_host_ip = Config().get_property("header.host.ip")
            self.header_host_name = Config().get_property("header.host.name")
            self.header_in_filepath = Config().get_property("header.in.filepath")
            self.header_in_filename = Config().get_property("header.in.filename")
            self.header_in_file_ext = Config().get_property("header.in.file.ext")
            self.header_in_file_size = Config().get_property("header.in.file.size")
            self.header_out_filepath = Config().get_property("header.out.filepath")
            self.header_out_filename = Config().get_property("header.out.filename")
            self.header_out_file_ext = Config().get_property("header.out.file.ext")
            self.header_out_file_size = Config().get_property("header.out.file.size")
            self.input = "input"
            self.work_directory = Config().get_property("work.directory")
            
        # need a new copy of result for each extraction - this is an exception to the singleton pattern
        self.result = dict.fromkeys(self.get_property("extraction.headers").replace(" ", "").split(",")) #if str else list())
        
    
    def project_home(self):
        return self.HOME
    
    def get_config(self):
        return self.config
    
    def init_extraction_entry(self) -> dict:
        self.result[self.header_host_ip] = self.host_ip
        self.result[self.header_host_name] = self.host_name
        return self.result
    
    def extraction_entry_input(self, message: str, input_values: dict) -> dict:
        self.result = Config().init_extraction_entry()
        self.result[self.header_message] = message
        self.result[self.header_in_filepath] = input_values[self.header_in_filepath]
        self.result[self.header_in_filename] = input_values[self.header_in_filename]
        self.result[self.header_in_file_ext] = input_values[self.header_in_file_ext]
        self.result[self.header_in_file_size] = input_values[self.header_in_file_size]
        return dict(self.result)
    
    def extraction_entry_output(self, message: str, input_values: dict, output_values: dict) -> dict:
        self.result = Config().extraction_entry_input(message, input_values)
        self.result[self.header_out_filepath] = output_values[self.header_out_filepath]
        self.result[self.header_out_filename] = output_values[self.header_out_filename]
        self.result[self.header_out_file_ext] = output_values[self.header_out_file_ext]
        self.result[self.header_out_file_size] = output_values[self.header_out_file_size]
        return self.result
    
    def get_supported_file_extensions(self) -> list:
        return self.supported_file_extensions
    
    def get_property(self, key:str) -> str:
        return self.config.get(key)
    
    # Singleton pattern- ensure HOME variable is assigned but once
    def __new__(cls, *args, **kwargs):
        if not isinstance(cls._instance, cls):
            cls._instance = object.__new__(cls)
        
        return cls._instance

Config()
Config()
