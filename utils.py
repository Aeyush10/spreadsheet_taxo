import yaml
import re

def get_max_tokens_for_step(step):
    """
    Returns appropriate max_tokens value based on the analysis step.
    
    Args:
        step (str): The step identifier (e.g., 'step2', 'step3', etc.)
    
    Returns:
        int: The maximum number of tokens for the given step
    """
    step_token_limits = {
        'step2': 1000,   # Keywords: comma-separated list, should be concise
        'step3': 2000,   # Codes: multiple codes with related keywords
        'step4': 2000,   # Themes: themes with related codes
        'step5': 3000,   # Concepts: concept definitions, more detailed
        'step6': 4000,   # Conceptual model: comprehensive model description
        'system': 500,   # System prompts should be short
    }
    
    # Default to 2000 if step not found
    return step_token_limits.get(step, 2000)

def get_prompt(step):
    """
    Retrieves the prompt for a given step from the prompts.yaml file and fills in 
    details from prompt_details.yaml when [..] placeholders are found.
    
    Args:
        step (str): The step number or identifier to retrieve the prompt for.
    
    Returns:
        str: The prompt for the specified step with placeholders filled in.
    """
    # Load the main prompts
    with open("prompts.yaml", "r") as file:
        prompts = yaml.safe_load(file)
        prompt = prompts.get(step, "Prompt not found.")
    
    # Load the prompt details
    with open("prompt_details.yaml", "r") as file:
        details = yaml.safe_load(file)
    
    # Find all placeholders in the format [key] and replace them
    def replace_placeholder(match):
        key = match.group(1)  # Extract the key from [key]
        return details.get(key, f"[{key}]")  # Replace with value or keep original if not found
    
    # Use regex to find and replace all [key] patterns
    filled_prompt = re.sub(r'\[([^\]]+)\]', replace_placeholder, prompt)
    # filled_prompt - add_data_to_prompt(filled_prompt)
    return filled_prompt

