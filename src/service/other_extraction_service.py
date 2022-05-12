
import os
import sys
from zipfile import ZipFile
from config.config import Config
from log.app_logger import AppLogger


class OtherExtractionService(object):
    
    def __init__(self):
        self.config = Config()
        self.log = AppLogger.get_logger()
    
    def other_extraction(self, input_file_path: str, image_directory: str,  extraction_entry: dict) -> list:     
        # Do the actual extraction
        return_list = list()
        try:
            with ZipFile(input_file_path, "r") as zip_file:
                criterion_1 = lambda media: "." in media and ("media" in media or "Pictures" in media) 
                criterion_2 = lambda media: not media.split(".")[-1] in "emf"
                candidates = list(filter(criterion_1, zip_file.namelist()))
                media_list = list(filter(criterion_2, candidates))

                return_list.extend(self.extract_media(media_list, zip_file, image_directory, extraction_entry))                               
        except Exception as e:
            self.log.error(str(type(e)))
            self.log.error(str(e))
            sys.exit()
            
        return return_list
       
    def extract_media(self, media_list: list, zip_file, image_directory: str, extraction_entry: dict) -> list:

        #file_type_count: dict = {}  # Dict for the found media files
        return_list = list()

        if not media_list:
            return list()
        
        for media in media_list:
            media_file = media.split("/")[-1]  # Get file name
            #file_type = media_file.split(".")[-1]  # Get file type
            
            output_file_path = os.path.join(image_directory, media_file)

            # Add file_type to dict, if it doesn't exist yet and increase count by 1
            #file_type_count[file_type] = file_type_count.get(file_type, 0) + 1

            # Read the data from the media file
            media_data = zip_file.read(name=media)

            # Write media data back to a file in the image_directory
            try:
                with open(os.path.join(image_directory, media_file), "wb") as media_fh:
                    media_fh.write(media_data)
                    #self.log.info(str(image_directory) + "/" + media_file)   
            except Exception as e:
                self.log.error(str(e))
                sys.exit()
                
            extraction_entry = self.config.extraction_entry_output(extraction_entry, {
                self.config.header_out_filepath : image_directory,
                self.config.header_out_filename : os.path.basename(os.path.splitext(output_file_path)[0]),
                self.config.header_out_file_ext : ".png",
                self.config.header_out_file_size : os.stat(output_file_path).st_size
            })
        
            self.log.info(str(extraction_entry))
            return_list.append(extraction_entry)
                
        return return_list

    
        # Return the file_type_count as list, sorted first by values (frequency) and then by
        # key (file extension)
        #return sorted(file_type_count.items(), key=lambda i: (-i[1], i[0]))
 
