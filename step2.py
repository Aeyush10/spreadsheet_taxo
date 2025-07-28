"""
Basic sample which shows how to call the LLM API using the llm-api-client library.
"""

import argparse
import time
import os

from llm_api_client.llm_call import llm_call
from llm_api_client.structured_processing.post_process import (
    PassthroughResponseProcessorFactory,
)
from llm_api_client.structured_processing.prompt_data import PromptData, PromptSpec
from utils import get_prompt


def run_variant(
    model: str, 
    scenario_guid: str, 
    enable_async: bool, 
    output_file,
    data_file: str = None
) -> None:
    """
    Runs some prompts against the LLM API to demonstrate basic usage.

    Arguments:
        model: Which model to query.
        scenario_guid: The scenario GUID to use when querying the model.
        enable_async: Whether to use asynchronous mode.
        use_chat: If `True`, chat mode is used. Otherwise, completions mode is used.
        output_file: File handle to write outputs to.
    """

    # Read questions from questions.txt
    # questions = read_questions_from_file("questions.txt")
    
    # if not questions:
    #     print("No questions found in questions.txt")
    #     return

    # print(f"\n{'='*60}")
    # print(f"Testing: Model={model}, Async={enable_async}")
    # print(f"{'='*60}")
    
    # Write header to output file
    # output_file.write(f"\n{'='*60}\n")
    # output_file.write(f"Testing: Model={model}, Async={enable_async}, Chat={use_chat}\n")
    # output_file.write(f"{'='*60}\n")

    # For chat payloads, the primary part of the payload is "messages". This is a list of dictionaries. Each
    # dictionary contains a "role" and "content".
    payloads = [
        {
            "messages": [
                {
                    "role": "system",
                    "content": get_prompt("system"),
                },
                {"role": "user", "content": get_prompt("step2").replace("[data]", data_file)},
            ],
            "temperature": 0,
            "top_p": 1,
            "max_tokens": 500,
            "presence_penalty": 0,
        }
    ]

    # Wrap each AOAI payload in a `PromptData`, which allows you to attach arbitrary metadata. Metadata is never sent
    # to the server. In this case we set the city as the metadata. If you don't need metadata, you can set this to
    # `None`. Each `PrompData` object is then wrapped in a `PromptSpec`. The `PromptSpec` allows you to specify, for
    # each prompt, whether the cache should be disabled and how many times the prompt should be retried if
    # post-processing fails. In this example, we don't use these options.
    prompts = [
        PromptSpec(PromptData(prompt=payload, metadata=None))
        for payload in payloads
    ]

    # send the prompts through the LLM API
    for result in llm_call(
        model=model,
        model_path="/chat/completions",
        prompts=prompts,
        response_processor_factory=PassthroughResponseProcessorFactory(),
        scenario_guid=scenario_guid,
        cache_path="",
        disable_cache=True,
        enable_async=enable_async,
        sync_if_fewer_minutes_than=0,  # turn off dynamic use of sync mode to force async usage for the demo
    ):
        # print(f"\n--- Response for {result.metadata} ---")
        print(f"Keywords: {result.response['choices'][0]['message']['content']}")
        # print(f"Response Type: {type(result.response)}")
        
        # Write to output file
        # output_file.write(f"\n--- Response for {result.metadata} ---\n")
        output_file.write(f"{result.response['choices'][0]['message']['content']}\n")
        # output_file.write(f"Response Type: {type(result.response)}\n")
        # output_file.write("-" * 40 + "\n")


def run(
    data_file: str,
    output_file: str,
    model: str = "gpt-4o-2024-05-13",
    # model: str = "gpt-45-preview",
    scenario_guid: str = "4d89af25-54b8-414a-807a-0c9186ff7539",
) -> None:
    # Open output file for writing
    # with open("outputs.txt", "w", encoding="utf-8") as output_file:
    #create output file if it doesn't exist

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    with open(output_file, "a", encoding="utf-8") as output_file:
        #open in append mode
        # output_file.write("LLM API Testing Results\n")
        # output_file.write("=" * 50 + "\n\n")
        
        # for enable_async in (False, True):
        #     for use_chat in (False, True):
        enable_async = False
        run_variant(
            model=f"dev-{model}",
            scenario_guid=scenario_guid,
            enable_async=enable_async,
            output_file=output_file,
            data_file=data_file
        )
    
def create_keywords(spreadsheet_name, spreadsheet_dir, output_folder):
    print(f"Creating keywords for {spreadsheet_name}")
    run(
        data_file=f"{spreadsheet_dir}/sheetjson.json",
        output_file=f"{output_folder}/keywords.txt",
    )
    # print("\nAll outputs have been written to 'outputs.txt'")


# def main() -> None:
#     # set to scenario id onboarded to the async API
#     parser = argparse.ArgumentParser(description="Basic Usage Example")
#     parser.add_argument("--model", help="Model to query", required=False)
#     parser.add_argument("--scenario_guid", help="Scenario GUID", required=False)
#     args = vars(parser.parse_args())
#     run(**{k: v for k, v in args.items() if v is not None})


# if __name__ == "__main__":
#     main()