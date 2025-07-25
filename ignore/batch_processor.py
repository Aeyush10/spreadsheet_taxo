"""
Batch processor for comprehensive Excel file analysis and extraction.
This script combines the extractor and analyzer to process all Excel files.
"""

import os
import sys
from pathlib import Path
import json
from datetime import datetime
import traceback

# Import our custom modules
from excel_extractor import ExcelExtractor
from excel_analyzer import analyze_excel_file


class BatchProcessor:
    """Batch processor for Excel files."""
    
    def __init__(self, input_folder: str = "raw_spreadsheets", output_folder: str = "spreadsheet_data"):
        """Initialize the batch processor."""
        self.input_folder = Path(input_folder)
        self.output_folder = Path(output_folder)
        self.output_folder.mkdir(exist_ok=True)
        
        # Create log file
        self.log_file = self.output_folder / f"processing_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        
    def log(self, message: str):
        """Log a message to both console and log file."""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_message = f"[{timestamp}] {message}"
        
        print(log_message)
        
        with open(self.log_file, 'a', encoding='utf-8') as f:
            f.write(log_message + '\n')
    
    def process_all_files(self):
        """Process all Excel files in the input folder."""
        self.log("Starting batch processing of Excel files")
        
        # Find all Excel files
        excel_files = list(self.input_folder.glob("*.xlsx")) + list(self.input_folder.glob("*.xls"))
        
        if not excel_files:
            self.log("No Excel files found in the input folder")
            return
        
        self.log(f"Found {len(excel_files)} Excel files to process")
        
        # Process each file
        processed_files = []
        failed_files = []
        
        for excel_file in excel_files:
            try:
                self.log(f"Processing: {excel_file.name}")
                success = self.process_single_file(excel_file)
                
                if success:
                    processed_files.append(excel_file.name)
                    self.log(f"Successfully processed: {excel_file.name}")
                else:
                    failed_files.append(excel_file.name)
                    self.log(f"Failed to process: {excel_file.name}")
                    
            except Exception as e:
                failed_files.append(excel_file.name)
                self.log(f"Error processing {excel_file.name}: {str(e)}")
                self.log(f"Traceback: {traceback.format_exc()}")
        
        # Generate summary
        self.generate_processing_summary(processed_files, failed_files)
    
    def process_single_file(self, excel_file: Path) -> bool:
        """Process a single Excel file."""
        try:
            file_stem = excel_file.stem
            output_dir = self.output_folder / file_stem
            output_dir.mkdir(exist_ok=True)
            
            # Step 1: Extract basic data using ExcelExtractor
            self.log(f"  Extracting data from {excel_file.name}")
            extractor = ExcelExtractor(str(self.input_folder), str(self.output_folder))
            extractor.extract_file(excel_file)
            
            # Step 2: Perform advanced analysis
            self.log(f"  Analyzing {excel_file.name}")
            analysis_dir = output_dir / "analysis"
            analysis_dir.mkdir(exist_ok=True)
            
            analysis_report = analyze_excel_file(excel_file, analysis_dir)
            
            # Step 3: Generate combined summary
            self.generate_file_summary(excel_file, output_dir, analysis_report)
            
            return True
            
        except Exception as e:
            self.log(f"  Error in process_single_file: {str(e)}")
            return False
    
    def generate_file_summary(self, excel_file: Path, output_dir: Path, analysis_report: dict):
        """Generate a comprehensive summary for a single file."""
        
        # Read extraction summary if it exists
        extraction_summary_path = output_dir / "extraction_summary.json"
        extraction_summary = {}
        
        if extraction_summary_path.exists():
            with open(extraction_summary_path, 'r', encoding='utf-8') as f:
                extraction_summary = json.load(f)
        
        # Create combined summary
        combined_summary = {
            'file_info': {
                'filename': excel_file.name,
                'file_size': excel_file.stat().st_size,
                'file_path': str(excel_file),
                'processing_timestamp': datetime.now().isoformat()
            },
            'extraction_summary': extraction_summary,
            'analysis_summary': {
                'data_patterns': analysis_report.get('data_patterns', {}),
                'formula_count': sum(len(sheet.get('formulas', {})) for sheet in analysis_report.get('formula_dependencies', {}).values()),
                'complex_formulas': sum(len(sheet.get('complex_formulas', [])) for sheet in analysis_report.get('formula_dependencies', {}).values()),
                'data_validation_rules': sum(len(sheet) for sheet in analysis_report.get('data_validation', {}).values()),
                'conditional_formatting_rules': sum(len(sheet) for sheet in analysis_report.get('conditional_formatting', {}).values()),
                'pivot_tables': sum(len(sheet) for sheet in analysis_report.get('pivot_tables', {}).values()),
                'named_ranges': len(analysis_report.get('named_ranges', {})),
                'protection_enabled': bool(analysis_report.get('protection', {}))
            },
            'output_structure': self.get_output_structure(output_dir)
        }
        
        # Save combined summary
        summary_path = output_dir / "file_summary.json"
        with open(summary_path, 'w', encoding='utf-8') as f:
            json.dump(combined_summary, f, indent=2, ensure_ascii=False)
        
        self.log(f"  Generated summary for {excel_file.name}")
    
    def get_output_structure(self, output_dir: Path) -> dict:
        """Get the structure of output directory."""
        structure = {}
        
        for item in output_dir.rglob('*'):
            if item.is_file():
                rel_path = item.relative_to(output_dir)
                structure[str(rel_path)] = {
                    'size': item.stat().st_size,
                    'modified': datetime.fromtimestamp(item.stat().st_mtime).isoformat()
                }
        
        return structure
    
    def generate_processing_summary(self, processed_files: list, failed_files: list):
        """Generate overall processing summary."""
        summary = {
            'processing_timestamp': datetime.now().isoformat(),
            'total_files': len(processed_files) + len(failed_files),
            'processed_successfully': len(processed_files),
            'failed_processing': len(failed_files),
            'success_rate': len(processed_files) / (len(processed_files) + len(failed_files)) if (processed_files or failed_files) else 0,
            'processed_files': processed_files,
            'failed_files': failed_files,
            'output_folder': str(self.output_folder),
            'log_file': str(self.log_file)
        }
        
        summary_path = self.output_folder / "batch_processing_summary.json"
        with open(summary_path, 'w', encoding='utf-8') as f:
            json.dump(summary, f, indent=2, ensure_ascii=False)
        
        self.log(f"Processing complete: {len(processed_files)} successful, {len(failed_files)} failed")
        self.log(f"Summary saved to: {summary_path}")


def main():
    """Main function to run batch processing."""
    
    # Check if input folder exists
    input_folder = Path("raw_spreadsheets")
    if not input_folder.exists():
        print(f"Input folder '{input_folder}' does not exist. Please create it and add Excel files.")
        return
    
    # Check if required packages are installed
    try:
        import pandas
        import openpyxl
    except ImportError as e:
        print(f"Required package not installed: {e}")
        print("Please install required packages using:")
        print("pip install -r requirements.txt")
        return
    
    # Create and run batch processor
    processor = BatchProcessor()
    processor.process_all_files()


if __name__ == "__main__":
    main()
