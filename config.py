import os


class Config:
    # paths
    DOWNLOADS = r'C:\Users\ychaube\Downloads'
    DESTINATION = r'C:\Users\ychaube\Downloads\PerformanceTesting'
    # create the destination folder if it doesn't exist
    if not os.path.exists(DESTINATION):
        os.makedirs(DESTINATION)

    # file names
    # FILE_NAME = r'DIndia - Standard Inwards Template'
    # FILE_NAME = r'Inwards-Deloitte India'
    # FILE_NAME = r'GSTR-Dmart'
    OUTPUT_FILE_NAME = r'PerformanceTest'

    # sheet names
    SHEET_NAME = 'Sheet1'

    # script configurations
    ITERATIONS = 2 # number of files to create
    RECORD_COUNT = 30000 # list of number of records to create in each iteration
    REFERENCE_COLUMN = 'DocumentNo'
    INCREMENT_VALUE = 1