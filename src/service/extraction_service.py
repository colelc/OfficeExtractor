import csv
import os
from datetime import datetime

from config.config import Config
from log.app_logger import AppLogger
from service.other_extraction_service import OtherExtractionService
from service.pdf_extraction_service import PdfExtractionService
from utils.file_utils import FileUtils
from utils.format_conversion_utils import FormatConversionUtils

class ExtractionService(object):
    
    def __init__(self):  
        self.config = Config()      
        self.log = AppLogger.get_logger()
        self.log.info("This is the ExtractionService")       
        
        self.supported_filetypes = self.config.get_supported_file_extensions()
        self.input_directory = self.config.get_property("input.directory")
        self.output_directory = self.config.get_property("output.directory")
        self.work_directory = self.config.work_directory
        
        FileUtils.create_folder(self.output_directory)
        FileUtils.create_folder(self.work_directory)
        
        # the log file where we enter 1 entry per image extraction 
        now = datetime.now()
        self.extractspecs = os.path.abspath(self.output_directory + "/extractspecs_" + now.strftime("%Y") + now.strftime("%m") + now.strftime("%d") + "_" + now.strftime("%H%M%S")  + ".csv")
        #self.log.info("The extractspecs.csv file: " + os.path.join(self.output_directory, self.extractspecs))

        self.fp = open(self.extractspecs, 'w', newline='', encoding='utf-8')
        self.writer = csv.DictWriter(self.fp, list(self.config.init_extraction_entry().keys())) 
        self.writer.writeheader()
       
    def extract(self):
        extraction_entry_list = list()
        
        # M A I N   P R O C E S S I N G   L O O P
        input_file_list = os.listdir(self.input_directory)
    
        for input_file_name in input_file_list:
            self.log.info("--> NEXT INPUT FILE: " + input_file_name)            
            file_extension = input_file_name.split(".")[-1]
            
            if not file_extension in self.supported_filetypes:
                self.log.warn(file_extension + " is not a supported file type: Skipping " + input_file_name)
                continue
        
            input_file_path = os.path.join(self.input_directory, input_file_name)
            image_directory = os.path.join(self.output_directory, "Extracted_"+ "".join(input_file_name.split(".")[:-1]))
             
            extraction_entry = Config().extraction_entry_input("", {
                self.config.header_in_filepath : os.path.dirname(input_file_path),
                self.config.header_in_filename : os.path.basename(os.path.splitext(input_file_path)[0]),
                self.config.header_in_file_ext : os.path.splitext(input_file_path)[1],
                self.config.header_in_file_size : os.stat(input_file_path).st_size
                })
    
            FileUtils.create_folder(image_directory)     
            
            try:           
                if file_extension == 'doc' or file_extension == 'dot':
                    pdf_work_file = FormatConversionUtils().doc2pdf(input_file_path, extraction_entry)
                    extraction_entry_list.extend(PdfExtractionService().extract(pdf_work_file, image_directory, extraction_entry))    
                elif file_extension == 'ppt' or file_extension == 'pot':
                    ppt_work_file = FormatConversionUtils().ppt2pdf(input_file_path, extraction_entry)
                    extraction_entry_list.extend(PdfExtractionService().extract(ppt_work_file, image_directory, extraction_entry))    
                elif file_extension == 'xls' or file_extension == 'xlt':
                    xls_work_file = FormatConversionUtils().xls2pdf(input_file_path, extraction_entry)
                    extraction_entry_list.extend(PdfExtractionService().extract(xls_work_file, image_directory, extraction_entry))    
                elif file_extension == 'pdf':
                    extraction_entry_list.extend(PdfExtractionService().extract(input_file_path, image_directory, extraction_entry))
                else:
                    extraction_entry_list.extend(OtherExtractionService().other_extraction(input_file_path, image_directory, extraction_entry))
            except Exception as e:
                self.log.error(str(e))
                    
        # write out the extraction results to CSV file
        try:
            self.writer.writerows(extraction_entry_list)
            self.fp.close()
        except Exception as e:
            self.log.error(str(type(e)))
            
        # delete the work directory
        FileUtils.delete_files_and_directory(self.work_directory)
        
        self.log.info("The results of this image extraction process are viewable at: ")
        self.log.info(os.path.join(self.output_directory, self.extractspecs))
              
