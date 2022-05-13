
from log.app_logger import AppLogger
from service.extraction_service import ExtractionService

class OfficeExtractor(object):
    
    log = AppLogger.get_logger()
    
    @classmethod
    def go(cls):
        cls.log.info("This is the OfficeExtractor process")
        ExtractionService().extract()
        cls.log.info("DONE")


OfficeExtractor.go()
