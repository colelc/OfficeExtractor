# -*- coding: utf-8 -*-
"""
Created on Thu May 14 13:51:15 2021

@author: shipleyd

conda install -c conda-forge poppler
pip install PyMuPDF pdfplumber pdf2image
"""
import os
import re
import argparse
#import win32com.client as win32
import csv
#from win32com.client import constants
#from pathlib import Path
from PIL import Image
from log.app_logger import AppLogger
from service.extraction_service import ExtractionService

class OfficeExtractor(object):
    
    log = AppLogger.get_logger()
    
    @classmethod
    def go(cls):
        cls.log.info("This is the OfficeExtractor process")
        #oe.log().info(f'oe.extract src={filelist}, dest={outputDirectory}')
        #oe.extract(src=filelist, dest=outputDirectory)
        ExtractionService().extract()
        #oe.image_summary(imgdir=outputDirectory)
        #oe.__del__()

          
    def __del__(self):
        """
        Deconstructor
        """
        self.fp.close()  

    #def __create_folder(self, folder_name):
    #    if not Path.exists(folder_name):
    #        Path.mkdir(folder_name, parents=True)
    """  
    def __get_media_list(self, zip_file):
        media_list = []

        for media in zip_file.namelist():
            if "." in media:
                file_type = media.split(".")[-1]
    
                # Microsoft Office stores media files in a folder called "media",
                # LibreOffice in a folder called "Pictures"; .emf files are added to
                # Microsoft Office documents, if media objects are embedded - these files
                # are skipped.
                if ("media" in media or "Pictures" in media) and file_type not in "emf":
                    media_list.append(media)
    
        return media_list   
    """ 
    """
    def __extract_media(self, media_list, zip_file, output_folder):
        file_type_count: dict = {}  # Dict for the found media files

        if media_list:
            self.__create_folder(output_folder)  # Create subfolder
    
            for media in media_list:
                media_file = media.split("/")[-1]  # Get file name
                file_type = media_file.split(".")[-1]  # Get file type
    
                # Add file_type to dict, if it doesn't exist yet and increase count by 1
                file_type_count[file_type] = file_type_count.get(file_type, 0) + 1
    
                # Read the data from the media file
                media_data = zip_file.read(name=media)
    
                # Write media data back to a file in the output_folder
                with open(Path.joinpath(output_folder, media_file), "wb") as media_fh:
                    media_fh.write(media_data)
    
        # Return the file_type_count as list, sorted first by values (frequency) and then by
        # key (file extension)
        return sorted(file_type_count.items(), key=lambda i: (-i[1], i[0]))
    """
    def ppt2pdf(self, path):
        """
        Convert PowerPoint Presentation files (1997-2003 *.ppt) to Portable
        Document Format (*.pdf)

        Parameters
        ----------
        path : string
            Full path and name to file to be converted.

        Returns
        -------
        new_file_abs : string
            Return path and name to converted file.

        """

        powerpoint = win32.Dispatch("Powerpoint.Application")
        # powerpoint.Visible = 1
        powerpoint.DisplayAlerts = 0

        new_file_abs = os.path.abspath(path)
        new_file_abs = re.sub(r'\\\\.\\\\w+$', '.pdf', new_file_abs)    

        # WithWindow needs to be set to false to run as headless. This is not
        # needed for Word or Excel.
        deck = powerpoint.Presentations.Open(path, WithWindow=False)
        formatType = 32 # ppt to pdf
        deck.SaveAs(new_file_abs, formatType) # formatType = 32 for ppt to pdf
        deck.Close()
        powerpoint.Quit()
        
        return new_file_abs

    def xls2pdf(self, path):
        """
        Convert Excel spreadsheet files (1997-2003 *.xls) to Portable
        Document Format (*.pdf)

        Parameters
        ----------
        path : string
            Full path and name to file to be converted.

        Returns
        -------
        new_file_abs : string
            Return path and name to converted file.

        """

        excel = win32.Dispatch("Excel.Application")
        # excel.Visible = 1
        excel.DisplayAlerts = 0

        new_file_abs = os.path.abspath(path)
        new_file_abs = re.sub(r'\\\\.\\\\w+$', '.pdf', new_file_abs)    

        wb = excel.Workbooks.Open(path)
        formatType = 57 # excel to pdf
        wb.SaveAs(new_file_abs, formatType) # formatType = 32 for ppt to pdf
        wb.Close()
        excel.Quit()
        
        return new_file_abs
    
    def doc2pdf(self, path):
        """
        Convert Word document files (1997-2003 *.xls) to Portable
        Document Format (*.pdf)

        Parameters
        ----------
        path : string
            Full path and name to file to be converted.

        Returns
        -------
        new_file_abs : string
            Return path and name to converted file.

        """

        word = win32.Dispatch("Word.Application")
        # word.Visible = 1
        word.DisplayAlerts = 0

        new_file_abs = os.path.abspath(path)
        new_file_abs = re.sub(r'\\\\.\\\\w+$', '.pdf', new_file_abs)    

        doc = word.Documents.Open(path)
        formatType = 17 # word to pdf
        doc.SaveAs(new_file_abs, formatType) # formatType = 32 for ppt to pdf
        doc.Close()
        word.Quit()
        
        return new_file_abs
    
    def doc2docx(self, path):
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

        word = win32.gencache.EnsureDispatch('Word.Application')
        word.DisplayAlerts = 0
        doc = word.Documents.Open(path)
        doc.Activate ()
    
        # Rename path with .docx
        new_file_abs = os.path.abspath(path)
        new_file_abs = re.sub(r'\\\\.\\\\w+$', '.docx', new_file_abs)
        
        # Save and Close
        word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
        doc.Close(False)
        word.Quit()
        
        return new_file_abs
        
    """
    def office_extract(self, file, location):
        
        filesize = os.stat(file).st_size

        # Do the actual extraction
        with ZipFile(file, "r") as zip_file:
            media_list = self.__get_media_list(zip_file=zip_file)
            file_type_count = self.__extract_media(
                media_list=media_list, zip_file=zip_file, output_folder=location
            )

        amount_files = sum(i[1] for i in file_type_count)

        if amount_files == 0:
            print(f"\\\\nNo media files found in {os.path.basename(file)}.")
        elif amount_files == 1:
            print(f"\\\\n1 media file extracted from {os.path.basename(file)}:")
            self.writer.writerow([os.path.splitext(os.path.basename(file))[0], os.path.splitext(file)[1], filesize, amount_files])
        else:
            print(f"\\\\n{amount_files} media files extracted from {os.path.basename(file)}")
            self.writer.writerow([os.path.splitext(os.path.basename(file))[0], os.path.splitext(file)[1], filesize, amount_files])

        # Print amount of file types
        for i in file_type_count:
            print(f"- {i[1]} {i[0]}")
    """

    def __get_file_list(self, dirName):
        """
        For the given path, get the List of all files in the directory tree 
        """
        # create a list of file and sub directories names in the given directory 
        listOfFile = os.listdir(dirName)
        allFiles = list()

        # Iterate over all the entries
        for entry in listOfFile:

            # Create full path
            fullPath = os.path.join(dirName, entry)

            # If entry is a directory then get the list of files in this directory 
            if os.path.isdir(fullPath):
                allFiles = allFiles + self.__get_file_list(fullPath)
            else:
                allFiles.append(fullPath)
                    
        return allFiles

    # def extract_embedded_files(self, file_path, save_path, sub_dir='xl'):
    #     """
    #     Extracts embedded files from Excel documents, it takes advantage of
    #     excel being a zipped collection of files. It creates a temporary folder,
    #     extracts all the contents of the excel folder there and then moves the
    #     embedded files to the requested save_path.
    
    #     Parameters:
    #     ----------
    #     file_path : str, 
    #         The path to the excel file to extract embedded files from.
        
    #     save_path : str,
    #         Path to save the extracted files to.
    
    #     sub_dir : str,
    #         one of 'xl' (for excel), 'word' , or 'ppt'. 
    #     """
        
    #     # make a temporary directory 
    #     temp_dir = tempfile.mkdtemp()
    
    #     # extract contents excel file to temporary dir
    #     zip_file = zipfile.ZipFile(file_path)
    #     zip_file.extractall(temp_dir)
    #     zip_file.close()
    
    #     # find all embedded files and copy to save_path
    #     embeddings_dir = f'{temp_dir}/{sub_dir}/embeddings/'
    #     embedded_files = list(glob.glob(embeddings_dir+'*'))
    
    #     # use mime type to identify binary data
    #     mime = magic.Magic(mime=True)

    #     self.__create_folder(save_path)
    
    #     for result_file in embedded_files:
    
    #         mime_type = mime.from_file(result_file)
        
    #         if mime_type == 'image/jpeg':
    #             os.rename(result_file, result_file + '.jpg')
    #         elif mime_type == 'image/png':
    #             os.rename(result_file, result_file + '.png')
    #         elif mime_type == 'image/gif':
    #             os.rename(result_file, result_file + '.gif')
    #         elif mime_type == 'image/bmp':
    #             os.rename(result_file, result_file + '.bmp')
    #         elif mime_type == 'image/tiff':
    #             os.rename(result_file, result_file + '.tiff')
    #         else:
    #             print('Not an image? %s' % mime_type)
    #             shutil.copy(result_file, save_path)

    def image_summary(self, imgdir, specfile=None):
        """
        Generates a log file of all images extracted from media.

        Parameters
        ----------
        imgdir : str
            Parent directory containing images and/or extracted image subfolders
            containing images.
        specfile : str, optional
            Log file name for image specs. The default is None.

        Returns
        -------
        None.

        """
        
        if specfile is None:
            imagespecs = os.path.abspath(imgdir + "/imagespecs.csv")
        else:
            imagespecs = specfile
    
        with open(imagespecs, 'w', newline='', encoding='utf-8') as fp:
            writer = csv.writer(fp)
            writer.writerow(['filename','ext','filesize','rows', 'cols', 'channels', 'ppi', 'format'])
            
            # Get the list of all files in directory tree at given path
            listOfFiles = self.__get_file_list(imgdir)
            
            # Get the list of all files in directory tree at given path
            listOfFiles = list()
        
            for (dirpath, dirnames, filenames) in os.walk(imgdir):
                listOfFiles += [os.path.join(dirpath, file) for file in filenames]
                        
            # Print the files
            for elem in listOfFiles:
                img_entry = []
                try:
                    img_entry.append(elem)
                    img_entry.append(os.path.splitext(os.path.dirname(elem))[-1])
                    # img_entry.append(elem.split('.')[-1])
                    img_entry.append(os.stat(elem).st_size)
        
                    with Image.open(elem) as im:
                        img_entry.append(im.size[1])
                        img_entry.append(im.size[0])
                        img_entry.append(len(im.getbands()))
                        img_entry.append(im.info['dpi'])
                        img_entry.append(im.mode)

                        writer.writerow(img_entry)
                except:
                    continue

def check_args(args=None):
    """
    Command line parsing.

    Parameters
    ----------
    args : TYPE, optional
        Command line arguments to run script. The default is None.

    Returns
    -------
    args.input_files (default: None)
        Input file(s) to process.
    args.input_directory (default: .\\\\data)
        Input directory containing files to be processed.
    args.output_directory (default: .\\\\results)
        Output directory of extracted media.
    args.logfile (default: .\\\\results\\\\extractor.log)
        Log file name.

    """
    
    parser = argparse.ArgumentParser(description='Extract images, audio, and video from media files.')
    parser.add_argument('-i', '--input_files',
                        help='Specify input files for media extraction',
                        nargs='+',
                        required=False,
                        default=None)
    parser.add_argument('-d', '--input_directory',
                        help='Specify the directory containing media files to be processed',
                        nargs='+',
                        required=False)
    parser.add_argument('-o', '--output_directory',
                        help='Specify the directory extracted media is to be stored',
                        required=False)
    
    args = parser.parse_args()

    return (args.input_files,
            args.input_directory,
            args.output_directory)
    
#inputDirectory = "C:\\Users\\oggie\\eclipse-workspace-python\\envs\\officeextractor\\input"
#outputDirectory = "C:\\Users\\oggie\\eclipse-workspace-python\\envs\\officeextractor\\output"

OfficeExtractor.go()

"""
if __name__ == '__main__':
    
    oe = OfficeExtractor(outputDirectory)
    oe.log().info("test")
    
    #logging.info(f'oe.extract src={filelist}, dest={outputDirectory}')
    #oe.log().info(f'oe.extract src={filelist}, dest={outputDirectory}')
    oe.extract(src=filelist, dest=outputDirectory)
    #oe.image_summary(imgdir=outputDirectory)
    #oe.__del__()
"""