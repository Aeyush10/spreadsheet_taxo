## Spreadsheet Taxonomy
Creating a taxonomy for spreadsheets from the ground up using LLMs.

## Usage

### Workbook Extraction
Use ```workbook_extraction.py```to extract details of the workbooks for the LLM to go through.   
Set the ```INPUT_FOLDER``` and ```OUTPUT_FOLDER``` variables and run the file.

### Running qualitative analysis
Use the ```main.py``` file to run qualitative analysis. This will run all steps of the analysis.
Again, set the ```INPUT_FOLDER``` and ```OUTPUT_FOLDER``` variables before running the file.

### Link to spreadsheets
https://dev.azure.com/msrcambridge/CalcIntel/_git/orkney?path=%2Ftests%2Fworkbooks


## TODOs
- Add support for the LLMs to use the chart/embedded image details by using a VLM
- Add support to extract comments from workbooks
- Add step 3-6 for qualitative analysis (based on step 2)