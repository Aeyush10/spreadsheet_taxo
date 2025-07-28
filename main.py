import os
import time
from step2 import create_keywords
from step3 import create_codes
INPUT_FOLDER = "orkney_spreadsheets_data"
OUTPUT_FOLDER = "orkney_spreadsheets_analysis"


data_sample = ""
for folder in os.listdir(INPUT_FOLDER):
        i += 1
        data_sample = data_sample + f"Content of spreadsheet {folder}:\n"
        print(folder)
        with open(os.path.join(INPUT_FOLDER,folder,'sheetjson.json'),"r", encoding="utf-8") as data_file:
            spreadsheet_data = data_file.read()
            #read json file into string

            # spreadsheet_data = " "
            data_sample = data_sample + spreadsheet_data + "\n"
        if (i > 3):
            break


#run step 2
# for folder in os.listdir(INPUT_FOLDER):
#     create_keywords(folder, os.path.join(INPUT_FOLDER, folder), OUTPUT_FOLDER)
#     #wait 30 seconds
#     time.sleep(30)

keywords_sample = ""
if len(keywords) > 100:
    keywords_sample = "\n".join(keywords[:100])
else:
    keywords_sample = "\n".join(keywords)

#run step 3
#read from keywords.txt
with open(os.path.join(OUTPUT_FOLDER, "keywords.txt"), "r", encoding="utf-8") as f:
    keywords = f.read().splitlines()

keywords_sample = ""
if len(keywords) > 100:
    keywords_sample = "\n".join(keywords[:100])
else:
    keywords_sample = "\n".join(keywords)


# #give 100 lines at a time
# for i in range(0, len(keywords), 100):
#     if i + 100 < len(keywords):
#         chunk = keywords[i:i+100]
#     else:
#         chunk = keywords[i:]
#     create_codes(chunk,INPUT_FOLDER, OUTPUT_FOLDER)
#     #wait 30 seconds
#     time.sleep(30)

#run step 4

    