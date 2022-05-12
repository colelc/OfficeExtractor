import fitz
import os
import pdfplumber
from config.config import Config
from pdf2image import convert_from_path as PDF2Image_convert_from_path
from log.app_logger import AppLogger


class PdfExtractionService(object):
    
    def __init__(self):
        self.log = AppLogger.get_logger()
        self.config = Config()

    #extraction_entry_list = PdfExtractionService.extract(input_file_path, image_directory, extraction_entry)
    def extract(self, input_file_path: str, image_directory: str, extraction_entry: dict) -> list:
        
        extraction_list = list()
  
        try:
            with pdfplumber.open(os.path.abspath(input_file_path)) as pdf:
                page = pdf.pages[0]
                text = page.extract_text()
                
                if text is not None: # searchable
                    #image_extractions = cls.pdf_extract(input_file_path, image_directory, extraction_entry)
                    extraction_list.extend(self.pdf_extract(input_file_path, image_directory, extraction_entry))
                else: # scanned
                    #  should be returning a list of images ?
                    num_images = PDF2Image_convert_from_path(os.path.abspath(input_file_path), output_folder=image_directory, fmt="jpg")
        except Exception as e:
            self.log.error(str(type(e)))
            self.log.error(str(e))
            
        return extraction_list

    def pdf_extract(self, input_file_path: str, image_directory: str, extraction_entry: dict) -> list:
        # open pdf file
        pdf_file = fitz.open(input_file_path)
                
        # finding number of pages in the pdf
        number_of_pages = len(pdf_file)

        # image counter
        #nimags = 0
        
        # list of result dictionaries
        return_list = list()

        #iterating through each page in the pdf
        for current_page_index in range(number_of_pages):
            #iterating through each image in every page of PDF
            for img_index, img in enumerate(pdf_file.getPageImageList(current_page_index)):
                xref = img[0]
                image = fitz.Pixmap(pdf_file, xref)
    
                image_file = os.path.join(image_directory, "image"+str(current_page_index) + "-" + str(img_index)+".png")
                #if it is a is GRAY or RGB image
                if image.n < 5:        
                    #image.writePNG("{}/image{}-{}.png".format(image_directory,current_page_index, img_index)) 
                    image.writePNG(image_file)  
                #if it is CMYK: convert to RGB first
                else:                
                    new_image = fitz.Pixmap(fitz.csRGB, image)
                    #new_image.writePNG("{}/image{}-{}.png".format(image_directory,current_page_index, img_index))
                    new_image.writePNG(image_file)
                    
                #self.log.info(str(image_directory) + " /image" + str(current_page_index) + "-" + str(img_index) + ".png")                      
                #nimags = nimags + 1
                
                extraction_entry = self.config.extraction_entry_output(extraction_entry, {
                    self.config.header_out_filepath : image_directory,
                    self.config.header_out_filename : os.path.basename(os.path.splitext(image_file)[0]),
                    self.config.header_out_file_ext : ".png",
                    self.config.header_out_file_size : str(os.stat(image_file).st_size)
                })
                
                self.log.info(str(extraction_entry))
                return_list.append(extraction_entry)
                
        return return_list

