## Spreadsheet Taxonomy
Creating a taxonomy for spreadsheets from the ground up using LLMs.

## Setup

### Prerequisites
- You must have the [substrate LLM API client](https://eng.ms/docs/experiences-devices/m365-core-msai/platform/substrate-intelligence/llm-api/llm-api-partner-docs/onboarding) installed   
- You must have the [sheetjson package](https://msdata.visualstudio.com/PROSE/_artifacts/feed/PROSE/connect) installed

### Installation Steps

1. **Create and activate a Python virtual environment:**
   ```bash
   python -m venv .venv
   .venv\Scripts\activate  # Windows
   # or
   source .venv/bin/activate  # Linux/Mac
   ```

2. **Install authentication tools for private packages:**
   ```bash
   pip install keyring artifacts-keyring
   ```

3. **Install public dependencies:**
   ```bash
   pip install openpyxl Pillow PyYAML tqdm xlwings rich webcolors
   ```

4. **Install private packages:**
   ```bash
   # Install llm-api-client from O365 feed
   pip install llm-api-client --index-url https://o365exchange.pkgs.visualstudio.com/_packaging/O365PythonPackagesV2/pypi/simple/
   
   # Install sheetjson from PROSE feed (specific version required)
   pip install sheetjson==10.13.17100 --index-url https://msdata.pkgs.visualstudio.com/_packaging/PROSE/pypi/simple/
   ```

   **Note:** You may need to authenticate with your Microsoft credentials when accessing private feeds.

## Usage

### Workbook Extraction
Use ```workbook_extractor.py``` (not `workbook_extraction.py`) to extract details of the workbooks for the LLM to go through.   
Set the ```INPUT_FOLDER``` and ```OUTPUT_FOLDER``` variables and run the file.

**Example:**
```python
INPUT_FOLDER = r"C:\path\to\your\excel\files"
OUTPUT_FOLDER = "extracted_data"
```

### Running qualitative analysis
Use the ```main.py``` file to run qualitative analysis. This will run all steps of the analysis.
Again, set the ```INPUT_FOLDER``` and ```OUTPUT_FOLDER``` variables before running the file.  
Further, based on which steps of the analysis you wish to run, you can set the ```RUN_STEP``` variables to True or False.  

**Important:** Set `RUN_STEP2 = True` if running for the first time, as it generates the keywords file needed for subsequent steps.

To change details of the LLM call (model, chat/completions, etc.) you can modify the ```run``` and ```run_variant``` functions in the ```llm_caller.py``` file.

### Example Workflow
```bash
# 1. First extract workbook data
python workbook_extractor.py

# 2. Then run qualitative analysis
python main.py
```

### Troubleshooting

#### Package Installation Issues
- **sheetjson not found:** This is a private Microsoft package. Ensure you have access to the PROSE feed and are authenticated. Use the specific version: `sheetjson==10.13.17100`
- **llm-api-client not found:** This requires access to the O365 Python packages feed.
- **Authentication errors:** You may need to set up Azure Artifacts Credential Provider or use a Personal Access Token.
- **Rich/webcolors dependencies:** Install these public packages first before installing sheetjson.

#### Runtime Issues
- **"No such file or directory: 'keywords.txt'":** Enable `RUN_STEP2 = True` in `main.py` to generate keywords first.
- **Empty output folders:** Check that your `INPUT_FOLDER` path is correct and contains .xlsx/.xls files.
- **xlwings errors:** Ensure Microsoft Excel is installed (required for chart extraction).
- **Path issues:** Use raw strings (r"path") or forward slashes for file paths.

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
