import os
import sys
from config.config import Config
import comtypes.client
from log.app_logger import AppLogger

class FormatConversionUtils(object):
    
    def __init__(self):
        self.log = AppLogger().get_logger()
        self.config = Config()
        self.wdFormatPDF = 17
        
    def doc2pdf(self, input_file_path: str, extraction_entry: dict) -> str:
        output_file_path = os.path.join(self.config.work_directory, extraction_entry[self.config.header_in_filename] + ".pdf")
        
        #Convert Word document files (1997-2003 *.xls) to Portable Document Format (*.pdf)
        try:
            word = comtypes.client.CreateObject("Word.Application")
            doc = word.Documents.Open(input_file_path)
            doc.SaveAs(output_file_path, FileFormat=self.wdFormatPDF)
        except Exception as e:
            self.log.error(str(e))
            sys.exit()
            
        self.log.info(output_file_path + " converted to PDF file " + output_file_path)
        return output_file_path
      
