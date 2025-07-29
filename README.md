## Spreadsheet Taxonomy
Creating a taxonomy for spreadsheets from the ground up using LLMs.

## Usage

### Prerequisites
You must have the [substrate LLM API client](https://eng.ms/docs/experiences-devices/m365-core-msai/platform/substrate-intelligence/llm-api/llm-api-partner-docs/onboarding) installed   
You must have the [sheetjson package](https://msdata.visualstudio.com/PROSE/_artifacts/feed/PROSE/connect) installed

### Workbook Extraction
Use ```workbook_extraction.py```to extract details of the workbooks for the LLM to go through.   
Set the ```INPUT_FOLDER``` and ```OUTPUT_FOLDER``` variables and run the file.

### Running qualitative analysis
Use the ```main.py``` file to run qualitative analysis. This will run all steps of the analysis.
Again, set the ```INPUT_FOLDER``` and ```OUTPUT_FOLDER``` variables before running the file.  
Further, based on which steps of the analysis you wish to run, you can set the ```RUN_STEP``` variables to True or False.  
  
To change details of the LLM call (model, chat/completions, etc.) you can modify the ```run``` and ```run_variant``` functions in the ```llm_caller.py``` file.

### Link to spreadsheets
https://dev.azure.com/msrcambridge/CalcIntel/_git/orkney?path=%2Ftests%2Fworkbooks


## TODOs
- Add support for the LLMs to use the chart/embedded image details by using a VLM
- Add support to extract comments from workbooks

## Open issues
Too many keywords generated  
Number of codes is too small  
Should a keyword map to one or multiple codes? Similarly upstream  
Conceptual model is confusing  