
import os
from zipfile import ZipFile
from log.app_logger import AppLogger
from utils.file_utils import FileUtils


class OtherExtractionService(object):
    log = AppLogger.get_logger()
    #OtherExtractionService.other_extraction(os.path.join(inputDirectory, input_file_name), image_directory, writer)

    @classmethod
    #def other_extraction(file, location):
    def other_extraction(cls, input_file_path: str, image_directory: str):
         
        # Do the actual extraction
        try:
            with ZipFile(input_file_path, "r") as zip_file:
                criterion_1 = lambda media: "." in media and ("media" in media or "Pictures" in media) 
                criterion_2 = lambda media: not media.split(".")[-1] in "emf"
                candidates = list(filter(criterion_1, zip_file.namelist()))
                media_list = list(filter(criterion_2, candidates))

                file_type_count = cls.extract_media(
                    media_list=media_list, zip_file=zip_file, output_folder=image_directory
                )
                
                cls.log.info(str(file_type_count))
            
            amount_files = sum(i[1] for i in file_type_count)
            
            cls.log.info("Number of media files detected: " + str(amount_files))
                        
            # Print amount of file types
            for i in file_type_count:
                cls.log.info(f"- {i[1]} {i[0]}")
        
        except Exception as e:
            cls.log.error(f'Failed to extract media from {input_file_path}')
            cls.log.error(str(type(e)))
            cls.log.error(str(e))
      
    @classmethod    
    def extract_media(cls, media_list, zip_file, output_folder):

        file_type_count: dict = {}  # Dict for the found media files

        if media_list:
            #FileUtils.create_folder(output_folder)  # Create subfolder
            for media in media_list:
                media_file = media.split("/")[-1]  # Get file name
                file_type = media_file.split(".")[-1]  # Get file type
    
                # Add file_type to dict, if it doesn't exist yet and increase count by 1
                file_type_count[file_type] = file_type_count.get(file_type, 0) + 1
    
                # Read the data from the media file
                media_data = zip_file.read(name=media)
    
                # Write media data back to a file in the output_folder
                try:
                    with open(os.path.join(output_folder, media_file), "wb") as media_fh:
                        media_fh.write(media_data)
                        cls.log.info(str(output_folder) + "/" + media_file)   
                except Exception as e:
                    cls.log.error(str(type(e)))
                    cls.log.error(str(e))
    
        # Return the file_type_count as list, sorted first by values (frequency) and then by
        # key (file extension)
        return sorted(file_type_count.items(), key=lambda i: (-i[1], i[0]))
 
