"""
Base module for LLM API calls - contains shared functionality for all LLM calling modules.
"""

import os
from typing import Optional, Any, Dict, List

from llm_api_client.llm_call import llm_call
from llm_api_client.structured_processing.post_process import (
    PassthroughResponseProcessorFactory,
)
from llm_api_client.structured_processing.prompt_data import PromptData, PromptSpec
from utils import get_prompt

# Default model and scenario GUID that can be overridden by specific implementations
DEFAULT_MODEL = "gpt-41-longco-2025-04-14"
DEFAULT_SCENARIO_GUID = "fd004048-ba97-46c8-9b09-6f566bdcd2d7"

def run_variant(
    model: str, 
    scenario_guid: str, 
    enable_async: bool, 
    output_file: Any,
    prompt: str,
) -> None:
    """
    Runs a prompt against the LLM API and writes the result to the output file.

    Arguments:
        model: Which model to query.
        scenario_guid: The scenario GUID to use when querying the model.
        enable_async: Whether to use asynchronous mode.
        output_file: File handle to write outputs to.
        prompt: The prompt content to send to the LLM.
    """
    payloads = [
        {
            "messages": [
                {
                    "role": "system",
                    "content": get_prompt("system"),
                },
                {"role": "user", "content": prompt},
            ],
            "temperature": 0,
            "top_p": 1,
            "presence_penalty": 0,
        }
    ]

    # Wrap each payload in a PromptData and PromptSpec
    prompts = [
        PromptSpec(PromptData(prompt=payload, metadata=None))
        for payload in payloads
    ]

    # Send the prompts through the LLM API
    for result in llm_call(
        model=model,
        model_path="/chat/completions",
        prompts=prompts,
        response_processor_factory=PassthroughResponseProcessorFactory(),
        scenario_guid=scenario_guid,
        cache_path="",
        disable_cache=True,
        enable_async=enable_async,
        sync_if_fewer_minutes_than=0,  # turn off dynamic use of sync mode to force async usage
    ):
        print(f"Response: {result.response['choices'][0]['message']['content']}")
        output_file.write(f"{result.response['choices'][0]['message']['content']}\n")

def run(
    prompt: str,
    output_file: str,
    model: str = DEFAULT_MODEL,
    scenario_guid: str = DEFAULT_SCENARIO_GUID,
) -> None:
    """
    Main function to run an LLM prompt and save the output to a file.
    
    Arguments:
        prompt: The prompt content to send to the LLM.
        output_file: Path to the output file where results will be saved.
        model: Which model to query (defaults to DEFAULT_MODEL).
        scenario_guid: The scenario GUID to use (defaults to DEFAULT_SCENARIO_GUID).
    """
    # Create output directory if it doesn't exist
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    # Open in append mode
    with open(output_file, "a", encoding="utf-8") as f:
        enable_async = False
        run_variant(
            model=f"dev-{model}",
            scenario_guid=scenario_guid,
            enable_async=enable_async,
            output_file=f,
            prompt=prompt
        )
