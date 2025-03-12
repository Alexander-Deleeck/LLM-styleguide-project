import os
import logging
from typing import List, Dict, Any
import random
from datetime import datetime
from utils_dataset import (
    load_style_rules,
    load_publication_text,
    get_random_text_sample,
    generate_example,
    get_azure_openai_client,
    save_test_dataset
)


def setup_logging(log_dir: str) -> None:
    """
    Set up logging configuration with both file and console handlers.
    
    Args:
        log_dir: Directory where log files will be stored
    """
    # Create logs directory if it doesn't exist
    os.makedirs(log_dir, exist_ok=True)
    
    # Create a timestamp for the log file
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(log_dir, f'dataset_creation_{timestamp}.log')
    
    # Define formatters for different levels of detail
    detailed_formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(funcName)s\n'
        'Message: %(message)s\n'
    )
    console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    
    # Set up file handler with detailed formatting
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(detailed_formatter)
    
    # Set up console handler with simpler formatting
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(console_formatter)
    
    # Configure root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)
    
    logging.info(f"Logging setup complete. Log file: {log_file}")


def create_test_dataset(
    rules_file: str,
    publication_file: str,
    output_file: str,
    examples_per_rule: int = 2,
    limit: int | None = None
) -> None:
    """
    Create a test dataset by generating examples for each rule.
    
    Args:
        rules_file: Path to the XML file containing style guide rules
        publication_file: Path to the publication text file
        output_file: Path where to save the test dataset
        examples_per_rule: Number of examples to generate per rule (default: 2)
        limit: Maximum total number of examples to generate (default: None, meaning no limit)
    """
    try:
        # Load rules and publication text
        logging.info(f"Loading rules from {rules_file}")
        rules = load_style_rules(rules_file)
        logging.info(f"Loaded {len(rules)} rules")
        
        logging.info(f"Loading publication text from {publication_file}")
        publication_text = load_publication_text(publication_file)
        logging.info(f"Loaded publication text ({len(publication_text)} characters)")
        
        # Initialize Azure OpenAI client
        logging.info("Initializing Azure OpenAI client")
        client = get_azure_openai_client()
        
        # List to store all examples
        all_examples = []
        
        # Generate examples for each rule
        for rule_id, rule_info in rules.items():
            # Check if we've hit the limit
            if limit and len(all_examples) >= limit:
                logging.info(f"Reached example limit of {limit}. Stopping generation.")
                break
                
            logging.info(f"Processing rule {rule_id}: {rule_info['title']}")
            logging.debug(f"Rule description: {rule_info['description']}")
            
            # Combine rule description with cases for a complete rule description
            full_description = f"{rule_info['description']}\n\nSpecific cases:\n"
            for case in rule_info['cases']:
                full_description += f"- {case['title']}: {case['description']}\n"
            
            # Calculate how many examples we can still generate within the limit
            remaining_examples = min(
                examples_per_rule,
                limit - len(all_examples) if limit else examples_per_rule
            )
            
            # Generate examples for this rule
            for i in range(remaining_examples):
                # Randomly decide if this example should adhere to the rule
                should_adhere = random.choice([True, False])
                adherence_str = "adherent" if should_adhere else "non-adherent"
                logging.info(f"Generating {adherence_str} example {i+1}/{remaining_examples} for rule {rule_id}")
                
                # Get a random sample of text from the publication
                sample_text = get_random_text_sample(publication_text)
                logging.debug(f"Using text sample of {len(sample_text)} characters")
                
                # Generate the example
                example_text, explanation = generate_example(
                    client,
                    full_description,
                    sample_text,
                    should_adhere
                )
                
                if example_text and explanation:
                    example = {
                        "rule_id": rule_id,
                        "rule_description": full_description,
                        "example_text": example_text,
                        "is_correct": should_adhere,
                        "explanation": explanation
                    }
                    all_examples.append(example)
                    logging.info(f"Successfully generated example {i+1} for rule {rule_id}")
                    logging.debug(f"Example text: {example_text[:100]}...")
                else:
                    logging.error(f"Failed to generate example {i+1} for rule {rule_id}")
                
                # Check if we've hit the limit after adding an example
                if limit and len(all_examples) >= limit:
                    logging.info(f"Reached example limit of {limit}. Stopping generation.")
                    break
        
        # Save the dataset
        logging.info(f"Saving dataset to {output_file}")
        save_test_dataset(all_examples, output_file)
        logging.info(f"Successfully created test dataset with {len(all_examples)} examples")
        
    except Exception as e:
        logging.error(f"Error creating test dataset: {str(e)}", exc_info=True)
        raise


if __name__ == "__main__":
    # Define paths relative to the script location
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Set up logging
    logs_dir = os.path.join(script_dir, "logs")
    setup_logging(logs_dir)
    
    rules_file = os.path.join(base_dir, "style-guide-simple-rules.xml")
    publication_file = os.path.join(
        base_dir, 
        "sample-publications", 
        "text-extractions",
        "sample_sustainability_reporting_standards.txt"
    )
    
    # Create test-datasets directory if it doesn't exist
    datasets_dir = os.path.join(script_dir, "test-datasets")
    os.makedirs(datasets_dir, exist_ok=True)
    
    # Find the next available dataset ID
    existing_datasets = [f for f in os.listdir(datasets_dir) if f.startswith("test_dataset_") and f.endswith(".json")]
    next_id = 1
    if existing_datasets:
        max_id = max(int(f.split("_")[-1].split(".")[0]) for f in existing_datasets)
        next_id = max_id + 1
        
    output_file = os.path.join(datasets_dir, f"test_dataset_{next_id}.json")
    
    # Create the test dataset with a limit of 10 examples
    create_test_dataset(
        rules_file=rules_file,
        publication_file=publication_file,
        output_file=output_file,
        examples_per_rule=2,  # Generate 2 examples per rule (1 adherent, 1 non-adherent)
        limit=10  # Set total example limit to 10
    )
