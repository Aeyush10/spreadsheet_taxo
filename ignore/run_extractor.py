"""
Simple runner script for Excel extraction and analysis.
This script provides an easy way to extract data from Excel files.
"""

import os
import sys
from pathlib import Path


def check_dependencies():
    """Check if required packages are installed."""
    required_packages = ['pandas', 'openpyxl', 'xlrd']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print("❌ Missing required packages:")
        for package in missing_packages:
            print(f"  - {package}")
        print("\n📦 Please install missing packages using:")
        print("pip install -r requirements.txt")
        return False
    
    print("✅ All required packages are installed")
    return True


def check_folder_structure():
    """Check and create necessary folder structure."""
    input_folder = Path("raw_spreadsheets")
    output_folder = Path("spreadsheet_data")
    
    # Create input folder if it doesn't exist
    if not input_folder.exists():
        input_folder.mkdir()
        print(f"📁 Created input folder: {input_folder}")
    
    # Create output folder if it doesn't exist
    if not output_folder.exists():
        output_folder.mkdir()
        print(f"📁 Created output folder: {output_folder}")
    
    # Check for Excel files
    excel_files = list(input_folder.glob("*.xlsx")) + list(input_folder.glob("*.xls"))
    
    if not excel_files:
        print(f"⚠️  No Excel files found in {input_folder}")
        print("Please add Excel files (.xlsx or .xls) to the raw_spreadsheets folder")
        return False
    
    print(f"📊 Found {len(excel_files)} Excel files:")
    for file in excel_files:
        print(f"  - {file.name}")
    
    return True


def run_extraction():
    """Run the Excel extraction process."""
    print("\n🚀 Starting Excel extraction and analysis...")
    
    try:
        from batch_processor import BatchProcessor
        
        processor = BatchProcessor()
        processor.process_all_files()
        
        print("\n✅ Processing complete!")
        print(f"📁 Results saved in: spreadsheet_data/")
        print(f"📄 Check the batch_processing_summary.json for details")
        
    except Exception as e:
        print(f"❌ Error during processing: {str(e)}")
        return False
    
    return True


def show_results():
    """Show summary of extracted results."""
    output_folder = Path("spreadsheet_data")
    
    if not output_folder.exists():
        print("No results found. Please run extraction first.")
        return
    
    print("\n📊 Extraction Results:")
    print("=" * 50)
    
    for item in output_folder.iterdir():
        if item.is_dir() and not item.name.startswith('.'):
            print(f"\n📁 {item.name}/")
            
            # Count files in each subdirectory
            for subdir in item.iterdir():
                if subdir.is_dir():
                    file_count = len(list(subdir.glob("*")))
                    print(f"  📂 {subdir.name}/ ({file_count} files)")
    
    # Show summary file if it exists
    summary_files = list(output_folder.glob("batch_processing_summary.json"))
    if summary_files:
        import json
        with open(summary_files[0], 'r', encoding='utf-8') as f:
            summary = json.load(f)
        
        print(f"\n📈 Processing Summary:")
        print(f"  Total files processed: {summary['total_files']}")
        print(f"  Successful: {summary['processed_successfully']}")
        print(f"  Failed: {summary['failed_processing']}")
        print(f"  Success rate: {summary['success_rate']:.1%}")


def main():
    """Main function with interactive menu."""
    print("=" * 60)
    print("🔍 Excel Data Extractor and Analyzer")
    print("=" * 60)
    
    while True:
        print("\nOptions:")
        print("1. Check dependencies and setup")
        print("2. Run extraction and analysis")
        print("3. Show results summary")
        print("4. Exit")
        
        choice = input("\nEnter your choice (1-4): ").strip()
        
        if choice == '1':
            print("\n🔍 Checking dependencies...")
            deps_ok = check_dependencies()
            
            print("\n🔍 Checking folder structure...")
            folders_ok = check_folder_structure()
            
            if deps_ok and folders_ok:
                print("\n✅ All checks passed! You can now run the extraction.")
            else:
                print("\n❌ Please fix the issues above before proceeding.")
        
        elif choice == '2':
            print("\n🔍 Pre-flight checks...")
            if not check_dependencies():
                continue
            
            if not check_folder_structure():
                continue
            
            run_extraction()
        
        elif choice == '3':
            show_results()
        
        elif choice == '4':
            print("\n👋 Goodbye!")
            break
        
        else:
            print("❌ Invalid choice. Please enter 1, 2, 3, or 4.")


if __name__ == "__main__":
    main()
