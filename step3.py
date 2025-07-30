"""
Module for Step 3 of the spreadsheet taxonomy process - codes generation.
"""

from llm_base import run as base_run
from utils import get_prompt

# Custom scenario GUID for step3
SCENARIO_GUID = "4d89af25-54b8-414a-807a-0c9186ff7539"

def run(
    keywords,
    data,
    output_file: str,
    model: str = "gpt-41-longco-2025-04-14",
    scenario_guid: str = SCENARIO_GUID,
) -> None:
    """Run the step 3 process - generate codes from keywords and data"""
    prompt = get_prompt("step3").replace("[keywords]", keywords).replace("[data]", data)
    base_run(
        prompt=prompt,
        output_file=output_file,
        model=model,
        scenario_guid=scenario_guid
    )
    
def create_codes(keywords, data_sample, output_folder):
    # Convert keywords list to a single string
    keywords_str = "\n".join(keywords)
    run(
        keywords=keywords_str,
        data=data_sample,
        output_file=f"{output_folder}/codes.txt",
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