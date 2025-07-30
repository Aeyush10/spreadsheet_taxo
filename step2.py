"""
Module for Step 2 of the spreadsheet taxonomy process - keywords generation.
"""

from llm_base import run as base_run
from utils import get_prompt

# Custom scenario GUID for step2
SCENARIO_GUID = "4d89af25-54b8-414a-807a-0c9186ff7539"

def run(
    data_file: str,
    output_file: str,
    model: str = "gpt-41-longco-2025-04-14",
    scenario_guid: str = SCENARIO_GUID,
) -> None:
    """Run the step 2 process - generate keywords from data file"""
    prompt = get_prompt("step2").replace("[data]", data_file)
    base_run(
        prompt=prompt,
        output_file=output_file,
        model=model,
        scenario_guid=scenario_guid
    )
    
def create_keywords(spreadsheet_name, spreadsheet_dir, output_folder):
    print(f"Creating keywords for {spreadsheet_name}")
    run(
        data_file=f"{spreadsheet_dir}/sheetjson.json",
        output_file=f"{output_folder}/keywords.txt",
    )


# def main() -> None:
#     # set to scenario id onboarded to the async API
#     parser = argparse.ArgumentParser(description="Basic Usage Example")
#     parser.add_argument("--model", help="Model to query", required=False)
#     parser.add_argument("--scenario_guid", help="Scenario GUID", required=False)
#     args = vars(parser.parse_args())
#     run(**{k: v for k, v in args.items() if v is not None})


# if __name__ == "__main__":
#     main()