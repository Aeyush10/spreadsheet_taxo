import pandas as pd
import json
import os
from pathlib import Path
from typing import List

# Get the directory where this script is located
SCRIPT_DIR = Path(__file__).parent.absolute()

# Configuration - all paths relative to script directory
MOONSHOT_BASIC_FILE = SCRIPT_DIR / "MoonshotBasic.xlsx"
CONTEXT_DATA_FILE = SCRIPT_DIR / "Context_Data.xlsm"
WORKBOOK_DIR = SCRIPT_DIR  # Same as script directory
OUTPUT_DIR = (
    SCRIPT_DIR.parent.parent / "test_cases" / "moonshotBasic"
)  # Go up to tests/test_cases/moonshotBasic
OUTPUT_FILE = "MoonshotBasic.json"

# Excel reading settings
MOONSHOT_SHEET_NAME = "Examples"
MOONSHOT_HEADER_ROW = 4  # Row 5 is index 4 (0-based)


def load_excel_file_info(
    file_path, sheet_name=None, header_row=None, file_description=None
):
    """
    Load Excel file information including sheet names and optionally data from a specific sheet.

    Args:
        file_path: Path to the Excel file
        sheet_name: Specific sheet to load data from (optional)
        header_row: Header row for data loading (optional, only used if sheet_name provided)
        file_description: Description for error messages (optional)

    Returns:
        tuple: (data, sheet_names) where data is None if sheet_name not provided
    """
    if file_description is None:
        file_description = str(file_path)

    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names

        data = None
        if sheet_name:
            # Prevent pandas from converting "n/a" to NaN by specifying keep_default_na=False
            # and na_values as an empty list
            data = pd.read_excel(
                excel_file,
                sheet_name=sheet_name,
                header=header_row,
                keep_default_na=False,  # Don't use default NA values
                na_values=[""],  # Only treat empty strings as NaN
            )

        excel_file.close()
        return data, sheet_names

    except Exception as e:
        print(f"Warning: Could not read from {file_description}: {e}")
        return None, []


def load_moonshot_basic_data_and_sheets():
    """
    Load MoonshotBasic Excel file data for cleared entries AND get sheet names.

    Returns:
        tuple: (filtered_data, sheet_names)
    """
    # Use the generic function to load both data and sheet names
    data, sheet_names = load_excel_file_info(
        MOONSHOT_BASIC_FILE,
        MOONSHOT_SHEET_NAME,
        MOONSHOT_HEADER_ROW,
        "MoonshotBasic.xlsx",
    )

    if data is None:
        raise ValueError("Failed to load MoonshotBasic data")

    # Use the exact column name
    cleared_col = "Cleared for Eval?"

    if cleared_col not in data.columns:
        raise ValueError(f"Column '{cleared_col}' not found in the data")

    cleared_values = ["YES", "YES - New", "n/a"]
    filtered_data = data[data[cleared_col].isin(cleared_values)].copy()

    print(f"Found {len(filtered_data)} rows with cleared status: {cleared_values}")
    print(f"Total rows in file: {len(data)}")

    # Show a breakdown of all values
    value_counts = data[cleared_col].value_counts(dropna=False)
    print(f"\nBreakdown of all 'Cleared for Eval?' values:")
    for value, count in value_counts.items():
        included = "✅" if value in cleared_values else "❌"
        print(f"  {included} '{value}': {count} rows")

    return filtered_data, sheet_names


def determine_tags(row):
    """
    Determine appropriate tags for a test case based on original capability columns.

    Args:
        row: Pandas Series with row data (contains original capabilities)

    Returns:
        list: List of tags for this test case
    """
    tags = ["MoonshotBasic"]

    # Use original capability columns from the input file
    original_capability_mapping = {
        "Analysis": "analysis",
        "Advanced Analysis": "python_in_excel",
        "Text Analysis": "text_analysis",
        "Editing": "editing",
        "Calculation": "calculation",
        "Clean Data": "data_cleaning",
        "Search": "search",
        "Create": "creation",
        "M365": "m365_integration",
    }

    # Add tags based on original capabilities that are True
    for capability_col, tag_name in original_capability_mapping.items():
        if capability_col in row:
            capability_value = row[capability_col]
            # Check if the capability is marked as True (handle various True representations)
            if pd.notna(capability_value) and (
                capability_value is True
                or (
                    isinstance(capability_value, str)
                    and capability_value.lower() in ["true", "yes", "1"]
                )
                or capability_value == 1
            ):
                tags.append(tag_name)

    return tags


def determine_workbook_and_sheets(
    data_name, sheet_context, moonshot_sheets=None, context_data_sheets=None
):
    """
    Determine both workbook and sheets to import based on Data and Sheet/Context ID.

    Returns:
        tuple: (workbook_name, sheets_to_import)
    """
    sheets_to_import = []

    # Default workbook
    workbook = None

    # STEP 1: Handle Data column (workbook or sheet-in-workbook references)
    if not pd.isna(data_name) and data_name != "":
        # Special case: "Blank sheet" means no data to import
        if data_name.lower().strip() in {"blank sheet", "blank", "n/a"}:
            workbook = ""
            return workbook, sheets_to_import

        sheet_references = [sheet.strip() for sheet in str(data_name).split(",")]

        moonshot_matches = []
        context_data_matches = []

        # Check where each sheet exists and collect them
        for sheet_ref in sheet_references:
            if moonshot_sheets and sheet_ref in moonshot_sheets:
                moonshot_matches.append(sheet_ref)
                sheets_to_import.append(sheet_ref)
            elif context_data_sheets and sheet_ref in context_data_sheets:
                context_data_matches.append(sheet_ref)
                sheets_to_import.append(sheet_ref)

        # Determine workbook based on where sheets were found
        if context_data_matches and not moonshot_matches:
            workbook = CONTEXT_DATA_FILE.name
        elif moonshot_matches and not context_data_matches:
            workbook = MOONSHOT_BASIC_FILE.name
        elif context_data_matches and moonshot_matches:
            # Mixed - prefer Context Data if more matches
            workbook = (
                CONTEXT_DATA_FILE.name
                if len(context_data_matches) >= len(moonshot_matches)
                else MOONSHOT_BASIC_FILE.name
            )
        elif len(sheet_references) == 1:
            # Single reference, no matches - could be custom workbook
            original_ref = sheet_references[0]

            # Check if it looks like a filename
            if any(ext in original_ref.lower() for ext in [".xlsx", ".xls", ".xlsm"]):
                # This is definitely a filename
                normalized_filename = original_ref.replace(" ", "_")
            else:
                # Not a filename - treat as workbook reference (not sheet reference)
                normalized_filename = f"{original_ref.replace(' ', '_')}.xlsx"

            # Don't add original_ref to sheets_to_import - let sheet context handle all sheets
            potential_path = WORKBOOK_DIR / normalized_filename
            if potential_path.exists():
                workbook = normalized_filename

        else:
            # Multiple references but no matches - treat as custom sheets
            if sheet_references:
                first_ref = sheet_references[0]
                workbook = f"{first_ref.replace(' ', '_')}.xlsx"
                sheets_to_import = sheet_references  # Only for multiple references

    # STEP 2: Handle Sheet/Context ID (additional selection sheets)
    if not pd.isna(sheet_context) and sheet_context != "":
        context_sheets = parse_sheet_context(
            sheet_context, workbook, moonshot_sheets, context_data_sheets
        )

        # Add context sheets that exist in known workbooks and aren't already imported
        for sheet in context_sheets:
            if sheet and sheet not in sheets_to_import:
                if moonshot_sheets and sheet in moonshot_sheets:
                    sheets_to_import.append(sheet)
                elif context_data_sheets and sheet in context_data_sheets:
                    sheets_to_import.append(sheet)
                elif workbook not in [MOONSHOT_BASIC_FILE.name, CONTEXT_DATA_FILE.name]:
                    # For custom workbooks, ADD the sheet even if we can't verify it exists
                    # (we'll validate existence later in the validation step)
                    sheets_to_import.append(sheet)

    # STEP 3: Handle missing case - custom workbook with no sheet context
    elif workbook and workbook not in [
        MOONSHOT_BASIC_FILE.name,
        CONTEXT_DATA_FILE.name,
    ]:
        # We have a custom workbook but no sheet context specified
        # Load all visible sheets from the custom workbook
        if (
            not sheets_to_import
        ):  # Only if we haven't already added sheets from Data column
            visible_sheets = load_visible_sheets_from_workbook(workbook)
            sheets_to_import = visible_sheets[:3]  # Limit to first 3 visible sheets

    return workbook, sheets_to_import


def parse_sheet_context(
    sheet_context, workbook_name=None, moonshot_sheets=None, context_data_sheets=None
):
    """Extract sheet names from Sheet/Context ID column."""
    if "both sheets" in sheet_context.lower():
        # "Both sheets" only applies to custom workbooks
        if workbook_name and workbook_name not in [
            MOONSHOT_BASIC_FILE.name,
            CONTEXT_DATA_FILE.name,
        ]:
            # For custom workbooks, load the visible sheets and return first 2
            visible_sheets = load_visible_sheets_from_workbook(workbook_name)
            return visible_sheets[:2] if len(visible_sheets) >= 2 else visible_sheets
        else:
            # For main workbooks, "both sheets" doesn't make sense - return empty
            # Users should specify explicit sheet names for MoonshotBasic.xlsx and Context_Data.xlsm
            return []
    elif "&" in sheet_context:
        return [sheet.strip() for sheet in sheet_context.split("&")]
    elif " and " in sheet_context.lower():
        sheets = [sheet.strip() for sheet in sheet_context.split(" and ")]
        return [sheet for sheet in sheets]
    elif "," in sheet_context:
        return [sheet.strip() for sheet in sheet_context.split(",")]
    else:
        return [sheet_context.strip()]


def load_visible_sheets_from_workbook(workbook_name):
    """Load only visible sheets from a custom workbook, excluding hidden and very hidden sheets."""
    try:
        import openpyxl

        workbook_path = WORKBOOK_DIR / workbook_name
        if workbook_path.exists():
            wb = openpyxl.load_workbook(workbook_path, read_only=True)
            visible_sheets = []
            hidden_sheets = []

            for sheet in wb.worksheets:
                if sheet.sheet_state == "visible" and not sheet.title.startswith("_"):
                    visible_sheets.append(sheet.title)
                else:
                    hidden_sheets.append(sheet.title)

            wb.close()

            return visible_sheets
    except Exception as e:
        print(f"   ⚠️  Error loading sheets from {workbook_name}: {e}")
    return []


def create_test_case(query_index, item, workbook, sheets_to_import, tags, supplementary:List[str]=None, selection_sheet=None):
    """
    Create a test case entry in the format similar to CurlingTournament.
    """

    test_case = {
        "query_index": query_index,
        "tags": tags,
        "query": item,
        "sheetsToImport": sheets_to_import,  # Use the passed sheets
    }
    if workbook != "":
        test_case["workbook"] = workbook
    if selection_sheet:
        test_case["selection"] = {
            "sheet": selection_sheet,
            "range": "A1",
        }
    if supplementary:
        test_case["supplementary"] = {
            supplementary[0]: supplementary[1]
        }

    return test_case


def print_missing_workbooks_summary(skipped_cases, validation_results=None):
    """
    Print a detailed summary of all missing workbooks and their associated IDs.

    Args:
        skipped_cases: List of skipped test case dictionaries
        validation_results: Results from workbook validation (optional)
    """
    print(f"\n" + "=" * 80)
    print("MISSING/PROBLEMATIC WORKBOOKS DETAILED SUMMARY")
    print("=" * 80)

    # Collect all problematic workbooks
    all_problematic_workbooks = {}

    # Add skipped cases (workbooks that don't exist)
    for case in skipped_cases:
        workbook = case["workbook"]
        if workbook not in all_problematic_workbooks:
            all_problematic_workbooks[workbook] = {
                "cases": [],
                "issue_type": "File not found",
                "error": "Workbook file does not exist",
            }
        all_problematic_workbooks[workbook]["cases"].append(case)

    # Add workbooks with validation errors (exist but can't be read)
    if validation_results:
        for workbook, info in validation_results.items():
            if info.get("error") and workbook not in all_problematic_workbooks:
                # This workbook exists but has reading errors
                # We need to find which test cases reference it
                all_problematic_workbooks[workbook] = {
                    "cases": [],
                    "issue_type": "Reading error",
                    "error": info["error"],
                }

    if not all_problematic_workbooks:
        print(
            "✅ No missing or problematic workbooks - all test cases were successfully generated!"
        )
        return

    print(f"Total problematic workbooks: {len(all_problematic_workbooks)}")
    total_affected_cases = sum(
        len(wb_info["cases"]) for wb_info in all_problematic_workbooks.values()
    )
    print(f"Total affected test cases: {total_affected_cases}")

    # Sort workbooks by number of affected test cases (most impactful first)
    sorted_workbooks = sorted(
        all_problematic_workbooks.items(),
        key=lambda x: len(x[1]["cases"]),
        reverse=True,
    )

    for workbook, workbook_info in sorted_workbooks:
        cases = workbook_info["cases"]
        issue_type = workbook_info["issue_type"]
        error = workbook_info["error"]

        print(f"\n📋 PROBLEMATIC WORKBOOK: {workbook}")
        print(f"   Issue type: {issue_type}")
        print(f"   Error: {error}")

        if cases:  # Only show case details if we have them
            print(f"   Affected test cases: {len(cases)}")

            # Sort cases by ID for easier reference
            sorted_cases = sorted(cases, key=lambda x: x["id"])

            # Show all IDs associated with this workbook
            ids = [str(case["id"]) for case in sorted_cases]
            print(f"   Associated IDs: {', '.join(ids)}")

            # Show data names that led to this workbook
            data_names = list(
                set(case["data_name"] for case in sorted_cases if case["data_name"])
            )
            if data_names:
                print(f"   Data references: {', '.join(sorted(data_names))}")

            # Show a few example queries
            print(f"   Example queries:")
            for case in sorted_cases[:3]:  # Show first 3 examples
                print(f"     ID {case['id']}: {case['query']}")

            if len(sorted_cases) > 3:
                print(f"     ... and {len(sorted_cases) - 3} more cases")
        else:
            print(f"   Affected test cases: Unable to determine (validation error)")

    # Create a comprehensive quick reference list
    print(f"\n" + "=" * 50)
    print("QUICK REFERENCE - ALL PROBLEMATIC WORKBOOKS")
    print("=" * 50)

    for workbook, workbook_info in sorted_workbooks:
        cases = workbook_info["cases"]
        issue_type = workbook_info["issue_type"]

        if cases:
            # Ensure all IDs are integers before sorting
            ids = sorted([int(case["id"]) for case in cases])
            id_ranges = []

            # Group consecutive IDs into ranges for more compact display
            if ids:
                start = ids[0]
                end = ids[0]

                for i in range(1, len(ids)):
                    if ids[i] == end + 1:
                        end = ids[i]
                    else:
                        if start == end:
                            id_ranges.append(str(start))
                        else:
                            id_ranges.append(f"{start}-{end}")
                        start = end = ids[i]

                # Add the last range
                if start == end:
                    id_ranges.append(str(start))
                else:
                    id_ranges.append(f"{start}-{end}")

            print(
                f"{workbook}: IDs {', '.join(id_ranges)} ({len(cases)} cases) - {issue_type}"
            )
        else:
            print(f"{workbook}: (validation error) - {issue_type}")


def collect_test_cases_for_workbook(test_cases, workbook):
    """
    Helper function to collect test case info for a specific workbook.

    Args:
        test_cases: List of all generated test cases
        workbook: Workbook name to search for

    Returns:
        List of case info dictionaries
    """
    matching_cases = []
    for case in test_cases:
        if case["workbook"] == workbook:
            matching_cases.append(
                {
                    "id": case["query_index"],
                    "data_name": "",  # We don't have this in the test case, would need original data
                    "workbook": workbook,
                    "query": case["query"][:50] + "...",
                }
            )
    return matching_cases


def generate_moonshot_test_cases():
    """
    Generate test cases from MoonshotBasic data and save to JSON.
    Only create test cases where the workbook exists.
    """
    try:
        # Load the data and sheet names
        print("Loading MoonshotBasic data and sheet names...")
        filtered_data, moonshot_sheets = load_moonshot_basic_data_and_sheets()

        if filtered_data.empty:
            print("No cleared entries found!")
            return

        print("Loading Context Data sheet names...")
        _, context_data_sheets = load_excel_file_info(
            CONTEXT_DATA_FILE, file_description="Context_Data.xlsm"
        )

        # Initialize known workbooks
        existing_workbooks = {MOONSHOT_BASIC_FILE.name, CONTEXT_DATA_FILE.name}

        print(f"Generating test cases for {len(filtered_data)} cleared entries...")
        valid_test_cases = []
        skipped_cases = []

        for idx, (_, row) in enumerate(filtered_data.iterrows()):
            moonshot_id = int(row.get("ID", 0))
            item = row.get("Item", "").strip()
            data_name = str(row.get("Data", "")).strip()
            sheet_context = row.get("Sheet/ Context ID", "")
            tags = determine_tags(row)

            # Skip if no query/item
            if not item or item.strip() == "":
                skipped_cases.append(
                    {
                        "id": moonshot_id,
                        "data_name": data_name,
                        "workbook": "No query provided",
                        "query": "Empty query",
                    }
                )
                print(f"❌ Skipped test case ID {moonshot_id}: No query provided")
                continue

            # STEP 1: Determine workbook and sheets
            workbook, sheets_to_import = determine_workbook_and_sheets(
                data_name, sheet_context, moonshot_sheets, context_data_sheets
            )

            # STEP 1.2: Handle special case - no workbook determined but we have a query
            if workbook == "" and not sheets_to_import:
                test_case = create_test_case(
                    moonshot_id, item, workbook, sheets_to_import, tags
                )
                valid_test_cases.append(test_case)

            # STEP 2: Check if workbook was determined and exists
            workbook_exists = False
            if workbook is not None and workbook != "":
                if workbook in existing_workbooks:
                    workbook_exists = True
                else:
                    workbook_path = WORKBOOK_DIR / workbook
                    if workbook_path.exists():
                        workbook_exists = True
                        existing_workbooks.add(workbook)

            # STEP 3: Create test case with pre-determined values
            if workbook_exists:
                test_case = create_test_case(
                    moonshot_id, item, workbook, sheets_to_import, tags
                )

                # Prefix workbook with "moonshotBasic/" for the JSON output
                test_case["workbook"] = f"moonshotBasic/{workbook}"

                valid_test_cases.append(test_case)

                sheets_info = (
                    f" -> Sheets: {sheets_to_import}"
                    if sheets_to_import
                    else " (no sheet imports)"
                )

                print(
                    f"✅ Valid test case {len(valid_test_cases)}: ID {moonshot_id} - Data: '{data_name}' -> Workbook: '{workbook}'{sheets_info}"
                )
            else:
                # Skip this case - workbook couldn't be determined or doesn't exist
                workbook_display = (
                    workbook if workbook else "Unable to determine workbook"
                )
                skipped_cases.append(
                    {
                        "id": moonshot_id,
                        "data_name": data_name,
                        "workbook": workbook_display,
                        "query": row.get("Item", "")[:50] + "...",
                    }
                )

                print(
                    f"❌ Skipped test case ID {moonshot_id}: Data: '{data_name}' -> {workbook_display}"
                )

        # Sort valid test cases by query_index (ID) to maintain order
        valid_test_cases.sort(key=lambda x: x["query_index"])

        # Create output directory if it doesn't exist
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

        # Save only valid test cases to JSON file
        output_path = OUTPUT_DIR / OUTPUT_FILE
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(valid_test_cases, f, indent=2, ensure_ascii=False)

        # Print comprehensive summary
        print(f"\n" + "=" * 60)
        print("GENERATION SUMMARY")
        print("=" * 60)
        print(f"Total entries processed: {len(filtered_data)}")
        print(f"Valid test cases (workbook exists): {len(valid_test_cases)}")
        print(f"Skipped test cases (missing workbook): {len(skipped_cases)}")

        if len(filtered_data) > 0:
            print(f"Success rate: {len(valid_test_cases)/len(filtered_data)*100:.1f}%")

        print(f"Saved to: {output_path}")

        # At the end, return all needed data
        return valid_test_cases, skipped_cases, filtered_data

    except Exception as e:
        print(f"Error generating test cases: {e}")
        raise


def diagnose_context_data_file():
    """
    Diagnose the Context Data file issue by checking what files actually exist.
    """
    import os
    import glob

    workbook_dir = WORKBOOK_DIR

    print("=" * 60)
    print("CONTEXT DATA FILE DIAGNOSIS")
    print("=" * 60)
    print(f"Checking directory: {workbook_dir}")
    print()

    # Check if directory exists
    if not os.path.exists(workbook_dir):
        print(f"❌ Directory does not exist: {workbook_dir}")
        return

    # List all files in the directory
    print("All files in directory:")
    try:
        all_files = os.listdir(workbook_dir)
        for file in sorted(all_files):
            file_path = os.path.join(workbook_dir, file)
            if os.path.isfile(file_path):
                print(f"  📄 {file}")
            else:
                print(f"  📁 {file}/")
    except Exception as e:
        print(f"❌ Error listing files: {e}")
        return

    print()

    # Look for files containing "context" (case insensitive)
    print("Files containing 'context' (case insensitive):")
    context_files = [f for f in all_files if "context" in f.lower()]
    for file in context_files:
        print(f"  🎯 {file}")

    print()

    # Look for .xlsm files
    print("All .xlsm files:")
    xlsm_files = [f for f in all_files if f.lower().endswith(".xlsm")]
    for file in xlsm_files:
        print(f"  📊 {file}")

    print()

    # Test specific variations
    test_names = [
        "Context Data.xlsm",
        "Context_Data.xlsm",
        "context_data.xlsm",
        "context data.xlsm",
        "ContextData.xlsm",
        "Context-Data.xlsm",
    ]

    print("Testing specific filename variations:")
    for name in test_names:
        file_path = os.path.join(workbook_dir, name)
        exists = os.path.exists(file_path)
        print(f"  {'✅' if exists else '❌'} {name}")
        if exists:
            try:
                # Try to read it
                import pandas as pd

                excel_file = pd.ExcelFile(file_path)
                sheets = excel_file.sheet_names
                print(
                    f"      🔍 Found {len(sheets)} sheets: {sheets[:5]}{'...' if len(sheets) > 5 else ''}"
                )
            except Exception as e:
                print(f"      ⚠️  File exists but error reading: {e}")

    print()

    # Search using glob patterns
    print("Using glob patterns:")
    patterns = ["*context*.xlsm", "*Context*.xlsm", "*CONTEXT*.xlsm", "*.xlsm"]

    for pattern in patterns:
        full_pattern = os.path.join(workbook_dir, pattern)
        matches = glob.glob(full_pattern)
        print(f"  Pattern '{pattern}': {len(matches)} matches")
        for match in matches:
            filename = os.path.basename(match)
            print(f"    📄 {filename}")


def print_original_capability_tags_summary(test_cases, filtered_data):
    """
    Print summary of original capability-based tags.

    Args:
        test_cases: List of test case objects
        filtered_data: Original DataFrame with capability columns
    """
    print("\n" + "=" * 60)
    print("ORIGINAL CAPABILITY TAGS SUMMARY")
    print("=" * 60)

    capability_columns = [
        "Analysis",
        "Advanced Analysis",
        "Text Analysis",
        "Editing",
        "Calculation",
        "Clean Data",
        "Search",
        "Create",
        "M365",
    ]

    available_capabilities = [
        col for col in capability_columns if col in filtered_data.columns
    ]
    print(f"Using original capability columns: {available_capabilities}")

    # Count capability-based tags
    capability_tag_counts = {}
    cases_with_capability_tags = 0

    for case in test_cases:
        tags = case["tags"]

        # Count original capability tags
        capability_tags = [
            tag
            for tag in tags
            if tag
            in [
                "analysis",
                "python_in_excel",
                "text_analysis",
                "editing",
                "calculation",
                "data_cleaning",
                "search",
                "creation",
                "m365_integration",
            ]
        ]

        if capability_tags:
            cases_with_capability_tags += 1
            for tag in capability_tags:
                capability_tag_counts[tag] = capability_tag_counts.get(tag, 0) + 1

    print(
        f"Test cases with capability tags: {cases_with_capability_tags}/{len(test_cases)} ({cases_with_capability_tags/len(test_cases)*100:.1f}%)"
    )
    print(
        f"Test cases without capability tags: {len(test_cases) - cases_with_capability_tags}"
    )

    if capability_tag_counts:
        print("\nOriginal capability tag distribution:")
        for tag, count in sorted(
            capability_tag_counts.items(), key=lambda x: x[1], reverse=True
        ):
            percentage = (count / len(test_cases)) * 100
            print(f"  {tag}: {count} ({percentage:.1f}%)")

    # Show some examples
    print("\nExample test cases with capability tags:")
    examples_shown = 0
    for case in test_cases:
        capability_tags = [
            tag
            for tag in case["tags"]
            if tag != "MoonshotBasic"
            and tag
            not in [
                "fully_specified",
                "partially_specified",
                "no_reference",
                "content",
                "semantics",
                "location",
            ]
        ]
        if capability_tags and examples_shown < 3:
            print(f"\nID {case['query_index']}:")
            print(f"  Query: {case['query'][:60]}...")
            print(f"  Capability tags: {capability_tags}")
            print(f"  Full tags: {case['tags']}")

            examples_shown += 1

    # Show examples without capability tags
    print("\nExample test cases without capability tags:")
    examples_shown = 0
    for case in test_cases:
        capability_tags = [
            tag
            for tag in case["tags"]
            if tag
            in [
                "analysis",
                "python_in_excel",
                "text_analysis",
                "editing",
                "calculation",
                "data_cleaning",
                "search",
                "creation",
                "m365_integration",
            ]
        ]

        if not capability_tags and examples_shown < 2:
            print(f"\nID {case['query_index']}:")
            print(f"  Query: {case['query'][:60]}...")
            print(f"  Tags: {case['tags']}")

            examples_shown += 1


def print_summary_statistics(test_cases):
    """
    Print summary statistics about the generated test cases.

    Args:
        test_cases: List of test case objects
    """
    print("\n" + "=" * 60)
    print("TEST CASE SUMMARY STATISTICS")
    print("=" * 60)

    # Count by tags
    tag_counts = {}
    workbook_counts = {}

    for case in test_cases:
        # Count tags
        for tag in case["tags"]:
            tag_counts[tag] = tag_counts.get(tag, 0) + 1

        # Count workbooks
        workbook = case["workbook"]
        workbook_counts[workbook] = workbook_counts.get(workbook, 0) + 1

    print(f"Total test cases: {len(test_cases)}")
    print(
        f"Average tags per case: {sum(len(case['tags']) for case in test_cases) / len(test_cases):.1f}"
    )

    # Show ID range
    ids = [case["query_index"] for case in test_cases]
    print(f"ID range: {min(ids)} to {max(ids)}")

    print("\nTag distribution:")
    for tag, count in sorted(tag_counts.items(), key=lambda x: x[1], reverse=True):
        percentage = (count / len(test_cases)) * 100
        print(f"  {tag}: {count} ({percentage:.1f}%)")

    print("\nWorkbook distribution:")
    for workbook, count in sorted(
        workbook_counts.items(), key=lambda x: x[1], reverse=True
    ):
        percentage = (count / len(test_cases)) * 100
        print(f"  {workbook}: {count} ({percentage:.1f}%)")

    # Show some example test cases
    print("\nExample test cases:")
    for i in range(min(3, len(test_cases))):
        case = test_cases[i]
        print(f"\nTest case {i}:")
        print(f"  Query Index (Original ID): {case['query_index']}")
        print(f"  Query: {case['query'][:80]}...")
        print(f"  Tags: {case['tags']}")
        print(f"  Workbook: {case['workbook']}")


def print_workbook_mapping_summary(filtered_data, moonshot_sheets, context_data_sheets):
    """
    Print summary of how data names were mapped to workbooks.

    Args:
        filtered_data: Original filtered DataFrame
        moonshot_sheets: List of MoonshotBasic sheet names
        context_data_sheets: List of Context Data sheet names
    """
    print("\n" + "=" * 60)
    print("WORKBOOK MAPPING SUMMARY")
    print("=" * 60)

    mapping_counts = {
        "MoonshotBasic.xlsx": 0,
        "Context Data.xlsm": 0,
        "Other/Unknown": 0,
    }

    data_mappings = {}

    for _, row in filtered_data.iterrows():
        data_name = row.get("Data", "")

        if pd.isna(data_name) or data_name == "":
            workbook = "MoonshotBasic.xlsx"
            mapping_type = "Default (no data specified)"
        elif data_name in moonshot_sheets:
            workbook = "MoonshotBasic.xlsx"
            mapping_type = "Found in MoonshotBasic"
        elif data_name in context_data_sheets:
            workbook = "Context Data.xlsm"
            mapping_type = "Found in Context Data"
        else:
            workbook = f"{data_name}.xlsx"
            mapping_type = "Unknown - created new workbook name"
            mapping_counts["Other/Unknown"] += 1
            continue

        if workbook in mapping_counts:
            mapping_counts[workbook] += 1

        if data_name not in data_mappings:
            data_mappings[data_name] = mapping_type

    print(f"MoonshotBasic sheets found: {moonshot_sheets}")
    print(f"Context Data sheets found: {context_data_sheets}")
    print()

    for workbook, count in mapping_counts.items():
        print(f"{workbook}: {count} test cases")

    print("\nData name mappings:")
    for data_name, mapping_type in sorted(data_mappings.items()):
        if data_name:  # Skip empty data names
            print(f"  '{data_name}' -> {mapping_type}")


def print_sheets_to_import_summary(test_cases):
    """
    Print summary of sheetsToImport usage.

    Args:
        test_cases: List of test case objects
    """
    print("\n" + "=" * 60)
    print("SHEETS TO IMPORT SUMMARY")
    print("=" * 60)

    import_counts = {}
    cases_with_imports = 0

    for case in test_cases:
        sheets = case.get("sheetsToImport", [])
        if sheets:
            cases_with_imports += 1
            for sheet in sheets:
                import_counts[sheet] = import_counts.get(sheet, 0) + 1

    print(
        f"Test cases with sheet imports: {cases_with_imports}/{len(test_cases)} ({cases_with_imports/len(test_cases)*100:.1f}%)"
    )
    print(f"Test cases without sheet imports: {len(test_cases) - cases_with_imports}")

    if import_counts:
        print("\nMost referenced sheets:")
        for sheet, count in sorted(
            import_counts.items(), key=lambda x: x[1], reverse=True
        )[:10]:
            print(f"  {sheet}: {count} test cases")

    # Show some examples
    print("\nExample test cases with sheet imports:")
    examples_shown = 0
    for case in test_cases:
        if case.get("sheetsToImport") and examples_shown < 3:
            print(f"\nID {case['query_index']}:")
            print(f"  Query: {case['query'][:60]}...")
            print(f"  Workbook: {case['workbook']}")
            print(f"  Sheets to import: {case['sheetsToImport']}")
            print(f"  Selection sheet: {case['selection']['sheet']}")
            examples_shown += 1


def print_workbook_mapping_summary_for_valid_cases(test_cases):
    """
    Print summary of workbook mappings for the valid test cases that were actually written.

    Args:
        test_cases: List of valid test case objects (only those with existing workbooks)
    """
    print("\n" + "=" * 60)
    print("VALID WORKBOOK MAPPING SUMMARY")
    print("=" * 60)

    workbook_counts = {}

    for case in test_cases:
        workbook = case["workbook"]
        workbook_counts[workbook] = workbook_counts.get(workbook, 0) + 1

    print("Workbook distribution (valid cases only):")
    for workbook, count in sorted(
        workbook_counts.items(), key=lambda x: x[1], reverse=True
    ):
        percentage = (count / len(test_cases)) * 100
        print(f"  {workbook}: {count} cases ({percentage:.1f}%)")

    return workbook_counts


def print_successful_workbooks_summary(test_cases, validation_results):
    """
    Print a summary of all unique workbooks that were successfully found and used.

    Args:
        test_cases: List of valid test case objects (only those with existing workbooks)
        validation_results: Results from workbook validation
    """
    print(f"\n" + "=" * 80)
    print("SUCCESSFUL WORKBOOKS SUMMARY")
    print("=" * 80)

    # Get unique workbooks from test cases (these are guaranteed to exist)
    workbooks_in_json = set(case["workbook"] for case in test_cases)

    print(f"Total unique workbooks in output JSON: {len(workbooks_in_json)}")
    print(f"These workbooks contain {len(test_cases)} total test cases")

    # Categorize workbooks by their status
    fully_successful = []
    readable_with_sheet_issues = []

    for workbook in sorted(workbooks_in_json):
        workbook_info = validation_results.get(workbook, {})

        # Count test cases for this workbook
        test_case_count = sum(1 for case in test_cases if case["workbook"] == workbook)

        # Count test cases with sheet imports for this workbook
        cases_with_imports = [
            case
            for case in test_cases
            if case["workbook"] == workbook and case.get("sheetsToImport")
        ]
        import_count = len(cases_with_imports)

        # Check if there were any sheet validation issues for this workbook
        has_sheet_issues = any(
            case.get("sheetsToImport")
            and any(
                sheet not in workbook_info.get("sheets", [])
                for sheet in case["sheetsToImport"]
            )
            for case in cases_with_imports
        )

        workbook_status = {
            "name": workbook,
            "test_cases": test_case_count,
            "cases_with_imports": import_count,
            "available_sheets": workbook_info.get("sheets", []),
            "sheet_count": len(workbook_info.get("sheets", [])),
            "has_sheet_issues": has_sheet_issues,
            "error": workbook_info.get("error"),
        }

        if workbook_info.get("error"):
            # Workbook exists but has reading errors - shouldn't happen since we filtered these out
            pass
        elif has_sheet_issues:
            readable_with_sheet_issues.append(workbook_status)
        else:
            fully_successful.append(workbook_status)

    # Print fully successful workbooks
    if fully_successful:
        print(f"\n✅ FULLY SUCCESSFUL WORKBOOKS ({len(fully_successful)}):")
        print("   (Workbook exists, readable, all referenced sheets found)")
        for wb in fully_successful:
            print(f"\n   📁 {wb['name']}")
            print(f"      Test cases: {wb['test_cases']}")
            print(f"      Cases with sheet imports: {wb['cases_with_imports']}")
            print(
                f"      Available sheets ({wb['sheet_count']}): {wb['available_sheets']}"
            )

    # Print workbooks with sheet issues
    if readable_with_sheet_issues:
        print(f"\n⚠️  WORKBOOKS WITH SHEET ISSUES ({len(readable_with_sheet_issues)}):")
        print("   (Workbook exists and readable, but some referenced sheets missing)")
        for wb in readable_with_sheet_issues:
            print(f"\n   📁 {wb['name']}")
            print(f"      Test cases: {wb['test_cases']}")
            print(f"      Cases with sheet imports: {wb['cases_with_imports']}")
            print(
                f"      Available sheets ({wb['sheet_count']}): {wb['available_sheets']}"
            )

            # Show which sheets are missing
            missing_sheets = set()
            for case in test_cases:
                if case["workbook"] == wb["name"] and case.get("sheetsToImport"):
                    for sheet in case["sheetsToImport"]:
                        if sheet not in wb["available_sheets"]:
                            missing_sheets.add(sheet)

            if missing_sheets:
                print(f"      Missing sheets: {sorted(missing_sheets)}")

    # Quick reference list
    print(f"\n" + "=" * 50)
    print("QUICK REFERENCE - ALL SUCCESSFUL WORKBOOKS")
    print("=" * 50)

    all_successful = fully_successful + readable_with_sheet_issues
    all_successful.sort(key=lambda x: x["test_cases"], reverse=True)

    for wb in all_successful:
        status_indicator = "✅" if wb not in readable_with_sheet_issues else "⚠️"
        sheet_info = f"{wb['sheet_count']} sheets"
        if wb["cases_with_imports"] > 0:
            sheet_info += f", {wb['cases_with_imports']} cases import sheets"

        print(
            f"{status_indicator} {wb['name']}: {wb['test_cases']} test cases ({sheet_info})"
        )

    return {
        "fully_successful": fully_successful,
        "readable_with_sheet_issues": readable_with_sheet_issues,
        "total_workbooks": len(workbooks_in_json),
        "total_test_cases": len(test_cases),
    }


if __name__ == "__main__":
    try:
        # Generate the test cases (now filtered for existing workbooks)
        test_cases, skipped_cases, filtered_data = generate_moonshot_test_cases()

        if test_cases:
            # Print summary statistics for valid cases only
            print_summary_statistics(test_cases)
            print_workbook_mapping_summary_for_valid_cases(test_cases)
            # print_sheets_to_import_summary(test_cases)
            # print_original_capability_tags_summary(test_cases, filtered_data)

            # Add the successful workbooks summary
            # print_successful_workbooks_summary(test_cases, validation_results)

            # Add workbooks with validation errors to the summary
            # print_missing_workbooks_summary(skipped_cases, validation_results)

            print(f"\n✅ Successfully generated MoonshotBasic test cases!")
            print(f"📁 Output location: {os.path.join(OUTPUT_DIR, OUTPUT_FILE)}")

    except Exception as e:
        print(f"❌ Error: {e}")
