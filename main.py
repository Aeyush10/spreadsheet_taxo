import os
from step2 import create_keywords
INPUT_FOLDER = "orkney_spreadsheets_data"
OUTPUT_FOLDER = "orkney_spreadsheets_analysis"


for folder in os.listdir(INPUT_FOLDER):
    #get folder name
    #find sheetjson file in folder
    create_keywords(folder, os.path.join(INPUT_FOLDER, folder), OUTPUT_FOLDER)
    