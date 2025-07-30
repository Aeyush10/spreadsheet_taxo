"""
Module for calling LLM API for various tasks in the spreadsheet taxonomy project.
"""

import os
from llm_base import run

# Use custom scenario GUID for llm_caller
SCENARIO_GUID = "fd004048-ba97-46c8-9b09-6f566bdcd2d7"

# Import utility for getting prompts
from utils import get_prompt

def create_keywords(spreadsheet_name, spreadsheet_dir, output_folder):
    data_file=f"{spreadsheet_dir}/sheetjson.json"
    prompt = get_prompt("step2").replace("[data]", data_file)
    run(
        prompt=prompt,
        output_file=f"{output_folder}/keywords.txt",
        scenario_guid=SCENARIO_GUID
    )
    
def create_codes(keywords, data_sample, output_folder):
    # Convert keywords and data to properly formatted strings
    keywords_str = "\n".join(keywords) + "\n"
    data_sample_str = data_sample + "\n"
    prompt = get_prompt("step3").replace("[keywords]", keywords_str).replace("[data]", data_sample_str)
    run(
        prompt=prompt,
        output_file=f"{output_folder}/codes.txt",
        scenario_guid=SCENARIO_GUID
    )

def create_themes(codes, keywords_sample, output_folder):
    # Convert inputs to properly formatted strings
    keywords_str = "\n".join(keywords_sample) + "\n"
    codes_str = "\n".join(codes) + "\n"
    prompt = get_prompt("step4").replace("[codes]", codes_str).replace("[keywords]", keywords_str)
    run(
        prompt=prompt,
        output_file=f"{output_folder}/themes.txt",
        scenario_guid=SCENARIO_GUID
    )

def create_concepts(themes, codes, keywords_sample, output_folder):
    # Convert inputs to properly formatted strings
    themes_str = "\n".join(themes) + "\n"
    keywords_str = "\n".join(keywords_sample) + "\n"
    codes_str = "\n".join(codes) + "\n"
    prompt = get_prompt("step5").replace("[codes]", codes_str).replace("[keywords]", keywords_str).replace("[themes]", themes_str)
    run(
        prompt=prompt,
        output_file=f"{output_folder}/concepts.txt",
        scenario_guid=SCENARIO_GUID
    )

def create_conceptual_model(themes, codes, keywords_sample, output_folder):
    # Convert inputs to properly formatted strings
    themes_str = "\n".join(themes) + "\n"
    keywords_str = "\n".join(keywords_sample) + "\n"
    codes_str = "\n".join(codes) + "\n"
    prompt = get_prompt("step6").replace("[codes]", codes_str).replace("[keywords]", keywords_str).replace("[themes]", themes_str)
    run(
        prompt=prompt,
        output_file=f"{output_folder}/conceptual_model.txt",
        scenario_guid=SCENARIO_GUID
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