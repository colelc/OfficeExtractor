import csv
import json
import os
import sys
#import pdfplumber
from datetime import datetime
from pathlib import Path
#from pdf2image import convert_from_path
from config.config import Config
from src.log.app_logger import AppLogger
#from _csv import writer
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
        self.log.info("The extractspecs.csv file: " + os.path.join(self.output_directory, self.extractspecs))

        self.fp = open(self.extractspecs, 'w', newline='', encoding='utf-8')
        #self.writer = csv.writer(fp)       
        #self.writer.writerow(list(self.config.init_extraction_entry().keys()))
        self.writer = csv.DictWriter(self.fp, list(self.config.init_extraction_entry().keys())) 
        self.writer.writeheader()
       
    def extract(self):
        extraction_entry_list = list()
        
        #Extract audio, images, and video from common office files. 
        input_file_list = os.listdir(self.input_directory)
    
        for input_file_name in input_file_list:
                        
            file_extension = input_file_name.split(".")[-1]
            
            if not file_extension in self.supported_filetypes:
                self.log.warn(file_extension + " is not a supported file type: Skipping " + input_file_name)
                continue
        
            input_file_path = os.path.join(self.input_directory, input_file_name)
            image_directory = os.path.join(self.output_directory, "Extracted_"+ "".join(input_file_name.split(".")[:-1]))
             
            extraction_entry = Config().extraction_entry_input({
                self.config.header_in_filepath : os.path.dirname(input_file_path),
                self.config.header_in_filename : os.path.basename(os.path.splitext(input_file_path)[0]),
                self.config.header_in_file_ext : os.path.splitext(input_file_path)[1],
                self.config.header_in_file_size : os.stat(input_file_path).st_size
                })
            

            nimags = 0
   
            FileUtils.create_folder(image_directory)     
            
            try:           
                if file_extension == 'doc' or file_extension == 'dot':
                    pdf_work_file = FormatConversionUtils().doc2pdf(input_file_path, extraction_entry)
                    #extraction_entry_list.extend(PdfExtractionService().extract(input_file_path, image_directory, extraction_entry))
                    extraction_entry_list.extend(PdfExtractionService().extract(pdf_work_file, image_directory, extraction_entry))
                    #file_name_path = self.doc2pdf(os.path.abspath(file_name_path))
                    #nimags = self.pdf_extract(file_name_path, output_image_dir)
                    # os.remove(file_name_path)
                    #filesize = os.stat(file_name_path).st_size
                    #print(f'\\\\n{nimags} media files extracted from {os.path.basename(file_name_path)}:')
                    #print(f'- {nimags} png')
                    #self.writer.writerow([os.path.splitext(os.path.basename(file_name_path))[0], os.path.splitext(file_name_path)[1], filesize, nimags])
    
                elif file_extension == 'ppt' or file_extension == 'pot':
                    try:
                        file_name_path = self.ppt2pdf(os.path.abspath(file_name_path))
                        nimags = self.pdf_extract(file_name_path, output_image_dir)
                        # os.remove(file_name_path)
                        filesize = os.stat(file_name_path).st_size
                        print(f'\\\\n{nimags} media files extracted from {os.path.basename(file_name_path)}:')
                        print(f'- {nimags} png')
                        self.writer.writerow([os.path.splitext(os.path.basename(file_name_path))[0], os.path.splitext(file_name_path)[1], filesize, nimags])
                    except:
                        print(f'Failed to extract media from {os.path.basename(file_name_path)} after conversion to PDF')
    
                elif file_extension == 'xls' or file_extension == 'xlt':
                    try:
                        file_name_path = self.xls2pdf(os.path.abspath(file_name_path))
                        nimags = self.pdf_extract(file_name_path, output_image_dir)
                        # os.remove(file_name_path)
                        filesize = os.stat(file_name_path).st_size
                        print(f'\\\\n{nimags} media files extracted from {os.path.basename(file_name_path)}:')
                        print(f'- {nimags} png')
                        self.writer.writerow([os.path.splitext(os.path.basename(file_name_path))[0], os.path.splitext(file_name_path)[1], filesize, nimags])
                    except:
                        print(f'Failed to extract media from {os.path.basename(file_name_path)} after conversion to PDF')
    
                elif file_extension == 'pdf':
                    extraction_entry_list.extend(PdfExtractionService().extract(input_file_path, image_directory, extraction_entry))
                else:
                    extraction_entry_list.extend(OtherExtractionService().other_extraction(input_file_path, image_directory, extraction_entry))
            except Exception as e:
                self.log.error(str(e))
                sys.exit()
                    
        # write out the extraction results to CSV file
        try:
            self.writer.writerows(extraction_entry_list)
            self.fp.close()
        except Exception as e:
            self.log.error(str(type(e)))
            
        # delete the work directory
        FileUtils.delete_files_and_directory(self.work_directory)
              
