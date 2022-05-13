import os
from config.config import Config
import comtypes.client
from log.app_logger import AppLogger

class FormatConversionUtils(object):
    
    def __init__(self):
        self.log = AppLogger().get_logger()
        self.config = Config()
        self.docFormatPDF = 17
        self.pptFormatPDF = 32
        self.xlsFormatPDF = 57
        self.config = Config()
        
    def doc2pdf(self, input_file_path: str, extraction_entry: dict) -> str:
        output_file_path = os.path.join(self.config.work_directory, extraction_entry[self.config.header_in_filename] + ".pdf")
        
        try:
            word = comtypes.client.CreateObject("Word.Application")
            word.DisplayAlerts = 0
            doc = word.Documents.Open(input_file_path)#, WithWindow=False)
            doc.SaveAs(output_file_path, FileFormat=self.docFormatPDF)
            doc.Close()
        except Exception as e:
            self.log.error(str(e))
            return None
            #sys.exit()
            
        self.log.info(output_file_path + " converted to PDF file " + output_file_path)
        return output_file_path
      
    def ppt2pdf(self, input_file_path: str, extraction_entry: dict) -> str:
        output_file_path = os.path.join(self.config.work_directory, extraction_entry[self.config.header_in_filename] + ".pdf")
        
        try:
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            #powerpoint.Visible = 1
            powerpoint.DisplayAlerts = 0
            ppt = powerpoint.Presentations.open(input_file_path, WithWindow=False)
            ppt.SaveAs(output_file_path, FileFormat=self.pptFormatPDF)
            ppt.Close()
        except Exception as e:
            self.log.error(str(e))
            return None
            #sys.exit()
            
        self.log.info(output_file_path + " converted to PDF file " + output_file_path)
        return output_file_path
    
    def xls2pdf(self, input_file_path: str, extraction_entry: dict) -> str:
        output_file_path = os.path.join(self.config.work_directory, extraction_entry[self.config.header_in_filename] + ".pdf")
        
        try:
            # to solve: this will only grab the active sheet within a workbook
            excel = comtypes.client.CreateObject("Excel.Application")
            #excel.Visible = True
            excel.DisplayAlerts = 0
            wb = excel.Workbooks.Open(input_file_path)#, WithWindow=False)
            wb.SaveAs(output_file_path, FileFormat=self.xlsFormatPDF)
            wb.Close()
        except Exception as e:
            self.log.error(str(e))
            return None
            #sys.exit()
            
        self.log.info(output_file_path + " converted to PDF file " + output_file_path)
        return output_file_path
    
    #def doc2docx(self, path):
        """
        Convert Word document files (1997-2003 *.doc) to the new Word document
        format (*.docx). This modifies the older proprietary document format
        to utilize the XML format used by Microsoft.

        Parameters
        ----------
        path : string
            Full path and name to file to be converted.

        Returns
        -------
        new_file_abs : string
            Return path and name to converted file.

        """

        #word = win32.gencache.EnsureDispatch('Word.Application')
        #word.DisplayAlerts = 0
        #doc = word.Documents.Open(path)
        #doc.Activate ()
    
        # Rename path with .docx
        #new_file_abs = os.path.abspath(path)
        #new_file_abs = re.sub(r'\\\\.\\\\w+$', '.docx', new_file_abs)
        
        # Save and Close
        #word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
        #doc.Close(False)
        #word.Quit()
        
        #return new_file_abs
        

      
