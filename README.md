# Office Extractor #

This project extracts media images from Microsoft Office files.  This includes Word, Excel, Powerpoint, and PDF.  
*In the case of Excel, only the active worksheet is processed.  This is a limitation (bug) that will need resolution.*

This project currently does not have a GUI component but one will be added.

Required Packages:  python-dotenv, pymupdf, pdfplumber, pdf2image, zipfile, comtypes.client

Python Version: 3.9 *(but currently unknown why - should also be tried with 3.7)*

The driver program is `OfficeExtractor.py`

To execute: `python path-to-project/OfficeExtractor.py`


# Configuration #

There is a configuration file, `OfficeExtractor.env` under `PROJECT_HOME/resources`.

Edit `OfficeExtractor.env` to specify:

 	-Input directory where Microsoft files and PDF files are stored
 
 	-Work directory location which is a work area for the extractor program - it will be deleted at the end of program execution
 	
 	-Output directory where the extracted images will be stored
 	
 	-The PROJECT_HOME directory
 	
# Logging #

Two logging files are produced.

The first is a CSV file that details the SUCCESS/ERROR for every attempted image extraction.

This file has the name `extractspec_YYYYMMDD_HHMMS.csv` and is placed in the output directory.

The second is a program execution log, `MediaExtractor.log` and is placed in the PROJECT_HOME directory.

# TODOs #

Provide a summary of the total MB of image files extracted

Provide the extracted image metadata

Figure out how to extract images from all sheets in an Excel workbook, not just the active sheet.

Experiment with additional input file types

Implement a GUI


 


