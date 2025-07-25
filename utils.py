import yaml
import re

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

