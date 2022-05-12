import fitz
import os
import pdfplumber
from pdf2image import convert_from_path as PDF2Image_convert_from_path
from log.app_logger import AppLogger


class PdfExtractionService(object):
    log = AppLogger.get_logger()

    @classmethod
    def extract(cls, input_file_path: str, image_directory: str):
        
        num_images = 0
  
        try:
            with pdfplumber.open(os.path.abspath(input_file_path)) as pdf:
                page = pdf.pages[0]
                text = page.extract_text()
                
                if text is not None: # searchable
                    writer_image_dict = cls.pdf_extract(input_file_path, image_directory)
                else: # scanned
                    num_images = PDF2Image_convert_from_path(os.path.abspath(input_file_path), output_folder=image_directory, fmt="jpg")
        except Exception as e:
            cls.log.error(f'Failed to extract media from {input_file_path}')
            cls.log.error(str(type(e)))
            cls.log.error(str(e))
            num_images = "ERROR"
            
        return writer_image_dict


    @classmethod
    def pdf_extract(cls, input_file_path: str, image_directory: str):
         
        # open pdf file
        pdf_file = fitz.open(input_file_path)
                
        # finding number of pages in the pdf
        number_of_pages = len(pdf_file)

        # image counter
        nimags = 0
        
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
                    
                cls.log.info(str(image_directory) + " /image" + str(current_page_index) + "-" + str(img_index) + ".png")                      
                nimags = nimags + 1
                
                entry = {"writer_output_file_path": image_directory, 
                        "writer_output_file_name": os.path.basename(os.path.splitext(image_file)[0]), 
                        "writer_output_file_ext": ".png", 
                        "writer_output_file_size": str(os.stat(image_file).st_size)
                        }                
                return_list.append(entry)
                
        return return_list

