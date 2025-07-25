# Excel Data Extractor and Analyzer

A comprehensive Python tool for extracting ALL possible information from Excel spreadsheets (.xlsx and .xls files), including data, formulas, charts, images, macros, styling, and more.

## Features

### Data Extraction
- **Multi-sheet support**: Extract data from all sheets in a workbook
- **Multiple formats**: Save data as CSV, JSON, and Excel files
- **Data type analysis**: Identify and categorize different data types
- **Empty cell detection**: Track data density and empty cells

### Formula Analysis
- **Formula extraction**: Extract all formulas with their calculated values
- **Dependency analysis**: Map formula dependencies and references
- **Complexity scoring**: Identify complex formulas
- **Function usage**: Track which Excel functions are used
- **External references**: Detect references to other workbooks

### Visual Elements
- **Chart extraction**: Extract chart information and metadata
- **Image extraction**: Save embedded images as separate files
- **Conditional formatting**: Extract conditional formatting rules
- **Cell styling**: Extract font, fill, border, and alignment information

### Advanced Features
- **VBA/Macro extraction**: Extract VBA code and macros
- **Data validation**: Extract data validation rules
- **Pivot tables**: Analyze pivot table structures
- **Named ranges**: Extract and analyze named ranges
- **Protection analysis**: Detect worksheet and workbook protection
- **Raw XML access**: Extract underlying XML structure

## Installation

1. **Install required packages**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Required packages**:
   - pandas >= 1.3.0
   - openpyxl >= 3.0.0
   - xlrd >= 2.0.0
   - xlsxwriter >= 3.0.0
   - Pillow >= 8.0.0
   - lxml >= 4.6.0

## Usage

### Quick Start

1. **Place Excel files** in the `raw_spreadsheets/` folder
2. **Run the extractor**:
   ```bash
   python run_extractor.py
   ```
3. **Follow the interactive menu** to check dependencies and run extraction

### Manual Usage

#### Basic Extraction
```python
from excel_extractor import ExcelExtractor

extractor = ExcelExtractor("raw_spreadsheets", "spreadsheet_data")
extractor.extract_all_files()
```

#### Advanced Analysis
```python
from excel_analyzer import analyze_excel_file
from pathlib import Path

excel_file = Path("raw_spreadsheets/example.xlsx")
output_dir = Path("spreadsheet_data/example/analysis")
report = analyze_excel_file(excel_file, output_dir)
```

#### Batch Processing
```python
from batch_processor import BatchProcessor

processor = BatchProcessor()
processor.process_all_files()
```

## Output Structure

For each Excel file `filename.xlsx`, the following structure is created:

```
spreadsheet_data/
├── filename/
│   ├── data/                          # Sheet data in multiple formats
│   │   ├── Sheet1.csv
│   │   ├── Sheet1.json
│   │   ├── Sheet1.xlsx
│   │   └── sheet_info.json
│   ├── formulas/                      # Formula information
│   │   └── formulas.json
│   ├── styles/                        # Cell styling information
│   │   └── styles.json
│   ├── images/                        # Extracted images
│   │   ├── Sheet1_image_1.png
│   │   └── images_info.json
│   ├── charts/                        # Chart information
│   │   └── charts_info.json
│   ├── macros/                        # VBA/Macro information
│   │   ├── vba_archive.bin
│   │   └── macros_info.json
│   ├── metadata/                      # Workbook metadata
│   │   └── metadata.json
│   ├── raw/                          # Raw XML files
│   │   ├── xl_workbook.xml
│   │   ├── xl_styles.xml
│   │   └── file_structure.txt
│   ├── analysis/                      # Advanced analysis
│   │   └── comprehensive_analysis.json
│   ├── extraction_summary.json        # Extraction summary
│   └── file_summary.json             # Combined summary
└── batch_processing_summary.json      # Overall processing summary
```

## File Descriptions

### Core Files
- `excel_extractor.py`: Main extraction engine
- `excel_analyzer.py`: Advanced analysis utilities
- `batch_processor.py`: Batch processing with logging
- `run_extractor.py`: Interactive runner with menu system

### Output Files

#### Data Files
- `*.csv`: Sheet data in CSV format
- `*.json`: Sheet data in JSON format
- `*.xlsx`: Individual sheet as Excel file
- `sheet_info.json`: Metadata about each sheet

#### Analysis Files
- `formulas.json`: All formulas with complexity analysis
- `styles.json`: Cell formatting and styling information
- `images_info.json`: Information about embedded images
- `charts_info.json`: Chart metadata and configuration
- `macros_info.json`: VBA/Macro information
- `metadata.json`: Workbook properties and metadata

#### Advanced Analysis
- `comprehensive_analysis.json`: Complete analysis report including:
  - Data patterns and statistics
  - Formula dependencies
  - Data validation rules
  - Conditional formatting
  - Pivot table analysis
  - Named ranges
  - Protection settings

## Examples

### Extract Specific Information

```python
import json
from pathlib import Path

# Load analysis results
with open('spreadsheet_data/SS1/analysis/comprehensive_analysis.json', 'r') as f:
    analysis = json.load(f)

# Get formula complexity
formulas = analysis['formula_dependencies']
for sheet, data in formulas.items():
    complex_formulas = data['complex_formulas']
    print(f"Sheet {sheet} has {len(complex_formulas)} complex formulas")

# Get data patterns
patterns = analysis['data_patterns']
for sheet, data in patterns.items():
    density = data['data_density']
    print(f"Sheet {sheet} has {density:.1%} data density")
```

### Access Raw Data

```python
import pandas as pd

# Load sheet data
df = pd.read_csv('spreadsheet_data/SS1/data/Sheet1.csv')
print(f"Shape: {df.shape}")
print(f"Columns: {df.columns.tolist()}")

# Load formulas
with open('spreadsheet_data/SS1/formulas/formulas.json', 'r') as f:
    formulas = json.load(f)
    
print(f"Total formulas: {sum(len(sheet) for sheet in formulas.values())}")
```

## Error Handling

The system includes comprehensive error handling:

- **Individual file failures** don't stop batch processing
- **Detailed logging** to track processing status
- **Graceful degradation** when certain features aren't available
- **Validation** of input files and dependencies

## Limitations

- **Password-protected files**: Cannot extract from password-protected Excel files
- **Corrupted files**: May fail on severely corrupted Excel files
- **Memory usage**: Large files with many images/charts may require significant memory
- **VBA extraction**: VBA code is saved as binary archive (requires separate VBA tools for full analysis)

## Troubleshooting

### Common Issues

1. **Missing dependencies**: Run `pip install -r requirements.txt`
2. **No Excel files found**: Check that files are in `raw_spreadsheets/` folder
3. **Permission errors**: Ensure Excel files are not open in Excel
4. **Memory errors**: Process files individually for very large spreadsheets

### Getting Help

- Check the processing log files in `spreadsheet_data/`
- Review the `batch_processing_summary.json` for overall status
- Use the interactive menu in `run_extractor.py` for guided troubleshooting

## Advanced Usage

### Custom Processing

```python
from excel_extractor import ExcelExtractor
from excel_analyzer import ExcelAnalyzer

# Custom extraction with specific output folder
extractor = ExcelExtractor("my_input_folder", "my_output_folder")
extractor.extract_all_files()

# Custom analysis
analyzer = ExcelAnalyzer(Path("my_file.xlsx"))
patterns = analyzer.analyze_data_patterns()
dependencies = analyzer.analyze_formula_dependencies()
analyzer.close()
```

### Integration with Other Tools

The extracted data can be easily integrated with other tools:

- **Data analysis**: Use pandas with the extracted CSV/JSON files
- **Visualization**: Import data into plotting libraries
- **Database import**: Load structured data into databases
- **Migration**: Convert to other formats or systems

## Contributing

This tool is designed to be extensible. You can:

1. Add new extraction methods to `ExcelExtractor`
2. Add new analysis methods to `ExcelAnalyzer`
3. Modify output formats in the processing pipeline
4. Add new visualization or reporting features

## License

This project is provided as-is for educational and research purposes.
