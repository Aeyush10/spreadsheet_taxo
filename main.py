import os
import time
import logging
from llm_caller import create_keywords, create_codes, create_themes, create_concepts, create_conceptual_model
INPUT_FOLDER = "orkney_spreadsheets_data"
OUTPUT_FOLDER = "orkney_spreadsheets_analysis"
RUN_STEP2 = False  #generate keywords
RUN_STEP3 = True  #generate codes
RUN_STEP4 = True  #generate themes
RUN_STEP5 = True  #generate concepts
RUN_STEP6 = True  #generate conceptual model

for logger_name in ['llm_api_client', 'llm_api', 'llm_caller']:
    logging.getLogger(logger_name).setLevel(logging.WARNING)


data_sample = ""
DATA_SAMPLE_SIZE = 0
# i = 0
# for folder in os.listdir(INPUT_FOLDER):
#         i += 1
#         data_sample = data_sample + f"Content of spreadsheet {folder}:\n"
#         print(folder)
#         with open(os.path.join(INPUT_FOLDER,folder,'sheetjson.json'),"r", encoding="utf-8") as data_file:
#             spreadsheet_data = data_file.read()
#             #read json file into string

#             # spreadsheet_data = " "
#             data_sample = data_sample + spreadsheet_data + "\n"
#         if (i >= DATA_SAMPLE_SIZE):
#             break


# run step 2
if RUN_STEP2:
    for folder in os.listdir(INPUT_FOLDER):
        print(f"Creating keywords for {folder}")
        create_keywords(folder, os.path.join(INPUT_FOLDER, folder), OUTPUT_FOLDER)
        #wait 30 seconds
        time.sleep(30)

with open(os.path.join(OUTPUT_FOLDER, "keywords.txt"), "r", encoding="utf-8") as f:
    keywords = f.read().splitlines()

keywords_sample = ""
KEYWORD_CHUNK_SIZE = 100


if len(keywords) > KEYWORD_CHUNK_SIZE:
    keywords_sample = "\n".join(keywords[:KEYWORD_CHUNK_SIZE])
else:
    keywords_sample = "\n".join(keywords)

#run step 3
if RUN_STEP3:
    for i in range(0, len(keywords), KEYWORD_CHUNK_SIZE):
        if i + KEYWORD_CHUNK_SIZE > len(keywords):
            chunk = keywords[i:]
        else:
            chunk = keywords[i:i + KEYWORD_CHUNK_SIZE]
        print(f"Creating codes for keywords chunk {i//KEYWORD_CHUNK_SIZE + 1}")
        create_codes(chunk, data_sample, OUTPUT_FOLDER)
        #wait 30 seconds
        time.sleep(30)

#read codes from file
with open(os.path.join(OUTPUT_FOLDER, "codes.txt"), "r", encoding="utf-8") as f:
    codes = f.read().splitlines()

if RUN_STEP4:
    print("Creating themes from codes and keywords sample")
    create_themes(codes, keywords_sample, OUTPUT_FOLDER)

#read themes from file
with open(os.path.join(OUTPUT_FOLDER, "themes.txt"), "r", encoding="utf-8") as f:
    themes = f.read().splitlines()

if RUN_STEP5:
    print("Creating concepts from themes, codes and keywords sample")
    create_concepts(themes, codes, keywords_sample, OUTPUT_FOLDER)

if RUN_STEP6:
    print("Creating conceptual model from themes, codes and keywords sample")
    create_conceptual_model(themes, codes, keywords_sample, OUTPUT_FOLDER)