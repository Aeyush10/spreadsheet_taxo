"""
Main entry point for running LLM API tests.
"""

from step2 import run


def main():
    """Run the LLM API testing."""
    print("Starting LLM API Testing...")
    
    print("Running full test suite...")
    run()
    
    # Or run a specific variant
    # print("Running specific test variant...")
    # run_variant(
    #     model=f"dev-{model}",
    #     scenario_guid=scenario_guid,
    #     enable_async=True,
    #     use_chat=True,
    #     output_file=open("outputs.txt", "w", encoding="utf-8")
    # )


if __name__ == "__main__":
    main()
