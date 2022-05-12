import csv
import os
#import pdfplumber
from pathlib import Path
#from pdf2image import convert_from_path
from config.config import Config
from src.log.app_logger import AppLogger
#from _csv import writer
from service.other_extraction_service import OtherExtractionService
from service.pdf_extraction_service import PdfExtractionService

from utils.file_utils import FileUtils

class ExtractionService(object):
    
    def __init__(self, inputDirectory: str, outputDirectory: str):        
        self.log = AppLogger.get_logger()
        self.log.info("This is the ExtractionService")
        
        """
        self.supported_filetypes = [
        'docx', 'docm', 'dotx', 'dotm', 'xlsx', 'xlsm', 'xltx', 'xltm',
        'pptx', 'pptm', 'potm', 'potx', 'ppsx', 'ppsm', 'odt',  'ott',
        'ods',  'ots',  'odp',  'otp',  'odg',  'doc',  'dot',  'ppt',
        'pot',  'xls',  'xlt',  'pdf'
        ]
        """
        self.log.info("Input directory: " + inputDirectory)
        if not os.path.exists(outputDirectory):
            self.log.info("Creating an output directory: "  + str(outputDirectory))
            os.makedirs(outputDirectory)
        else:
            self.log.info("Output directory verfied: " + str(outputDirectory))

        #self.log.info("The supported file types are: ")
        #self.log.info(str(self.supported_filetypes))
        
        self.inputDirectory = inputDirectory
        self.outputDirectory = outputDirectory
        
        extractspecs = os.path.abspath(outputDirectory + "/extractspecs.csv")
        self.log.info("The extractspecs.csv file has been created in " + str(outputDirectory))

        fp = open(extractspecs, 'w', newline='', encoding='utf-8')
        self.writer = csv.writer(fp)
        headers = ['input_filepath', 'input_filename','input_ext','input_filesize', 
                   'output_filepath', 'output_filename','output_ext','output_filesize'];
        
        self.writer.writerow(headers)
        self.log.info("Headers written to " + str(extractspecs) + " : " + str(headers))
        
        results = dict.fromkeys(headers)
        self.log.info(str(results))

    
    #def extract(self, inputDirectory:str, outputDirectory:str, writer):
    def extract(self):
        #Extract audio, images, and video from common office files.
          
        input_file_list = os.listdir(self.inputDirectory)
    
        for input_file_name in input_file_list:
                        
            file_extension = input_file_name.split(".")[-1]
            
            if not file_extension in self.supported_filetypes:
                self.log.warn(file_extension + " is not a supported file type")
                continue
        
            input_file_path = os.path.join(self.inputDirectory, input_file_name)
            image_directory = os.path.join(self.outputDirectory, "Extracted_"+ "".join(input_file_name.split(".")[:-1]))
            
            writer_input_file_path = os.path.dirname(input_file_path)
            writer_input_file_name = os.path.basename(os.path.splitext(input_file_path)[0])
            writer_input_file_extension = os.path.splitext(input_file_path)[1]
            writer_input_file_size = os.stat(input_file_path).st_size

            writer_output_list = list()
            nimags = 0
   
            FileUtils.create_folder(image_directory)     
            
            if file_extension == 'doc' or file_extension == 'dot':
                try:
                    file_name_path = self.doc2pdf(os.path.abspath(file_name_path))
                    nimags = self.pdf_extract(file_name_path, output_image_dir)
                    # os.remove(file_name_path)
                    filesize = os.stat(file_name_path).st_size
                    print(f'\\\\n{nimags} media files extracted from {os.path.basename(file_name_path)}:')
                    print(f'- {nimags} png')
                    self.writer.writerow([os.path.splitext(os.path.basename(file_name_path))[0], os.path.splitext(file_name_path)[1], filesize, nimags])
                except:
                    print(f'Failed to extract media from {os.path.basename(file_name_path)} after conversion to PDF')

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
                writer_output_list = PdfExtractionService.extract(input_file_path, image_directory)
            else:
                try:
                    #self.log.info("Skipping this file for now")
                    # skip for right now
                    #self.office_extract(os.path.abspath(file_name_path), output_folder)
                    writer_output_list = OtherExtractionService.other_extraction(input_file_path, image_directory)
                    #print("SKIP otherextractionservice for now...")
                except:
                    print(f'Failed to extract media from {os.path.basename(file_name_path)}')
                  
            self.log.info(str(nimags) + " media files extracted from " + input_file_path)   
            
            
            for dct in writer_output_list:                                   
                self.writer.writerow([writer_input_file_path, writer_input_file_name, writer_input_file_extension, writer_input_file_size, 
                                      str(dct["writer_output_file_path"]),
                                      dct["writer_output_file_name"],
                                      dct["writer_output_file_ext"],
                                      dct["writer_output_file_size"]
                                      ])
