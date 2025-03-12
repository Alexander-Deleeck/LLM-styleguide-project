import xml.etree.ElementTree as ET
import os
from typing import List, Dict, Any, Tuple
import json
import random
from openai import AzureOpenAI
import logging
from dotenv import load_dotenv


def load_style_rules(xml_path: str) -> Dict[str, Any]:
    """Load and parse style guide rules from XML file."""
    tree = ET.parse(xml_path)
    root = tree.getroot()
    
    rules_dict = {}
    for rule in root.findall('.//rule'):
        section_id = rule.find('section_id').text
        title = rule.find('title').text
        description = rule.find('description').text if rule.find('description') is not None else ""
        
        cases = []
        for case in rule.findall('.//case'):
            case_dict = {
                'title': case.find('title').text,
                'description': case.find('description').text if case.find('description') is not None else ""
            }
            cases.append(case_dict)
        
        rules_dict[section_id] = {
            'title': title,
            'description': description,
            'cases': cases
        }
    
    return rules_dict


def load_publication_text(file_path: str) -> str:
    """Load publication text from file."""
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()


def get_random_text_sample(text: str, num_lines: int = 40) -> str:
    """Get a random sample of lines from the text."""
    lines = text.split('\n')
    if len(lines) <= num_lines:
        return text
    
    start_idx = random.randint(0, len(lines) - num_lines)
    sample_lines = lines[start_idx:start_idx + num_lines]
    return '\n'.join(sample_lines)


def create_example_prompt(rule_description: str, sample_text: str, should_adhere: bool) -> str:
    """Create a prompt for the LLM to generate an example."""
    adherence = "adheres to" if should_adhere else "violates"
    return f"""Task: Create a synthetic example text based on the provided sample that {adherence} a specific style guide rule.

Style Guide Rule:
{rule_description}

Sample Text (for context and style):
{sample_text}

Please create a new example text that {adherence} the style guide rule above. The example should be similar in style to the sample text but modified to specifically {adherence} the rule.

Return the response in the following JSON format:
{{
    "example_text": "your synthetic example here",
    "explanation": "detailed explanation of why this example {adherence} the rule"
}}"""


def get_azure_openai_client() -> AzureOpenAI:
    """Initialize and return Azure OpenAI client."""
    load_dotenv()
    
    client = AzureOpenAI(
        api_key=os.getenv("AZURE_OPENAI_API_KEY"),
        api_version="2024-02-15-preview",
        azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT")
    )
    return client


def generate_example(client: AzureOpenAI, 
                    rule_description: str, 
                    sample_text: str, 
                    should_adhere: bool) -> Tuple[str, str]:
    """Generate a synthetic example using Azure OpenAI."""
    prompt = create_example_prompt(rule_description, sample_text, should_adhere)
    
    try:
        response = client.chat.completions.create(
            model=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
            messages=[
                {"role": "system", "content": "You are a helpful assistant that creates examples for style guide rule checking."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=500,
            response_format={"type": "json_object"}
        )
        
        try:
            result = json.loads(response.choices[0].message.content)
            if not isinstance(result, dict) or 'example_text' not in result or 'explanation' not in result:
                logging.error(f"Invalid response format. Expected dict with 'example_text' and 'explanation'. Got: {result}")
                return "", ""
            return result["example_text"], result["explanation"]
            
        except json.JSONDecodeError as e:
            logging.error(f"Failed to parse JSON response: {str(e)}")
            logging.error(f"Raw response content: {response.choices[0].message.content}")
            return "", ""
            
    except Exception as e:
        logging.error(f"API call failed: {str(e)}")
        logging.error(f"Prompt used:\n{prompt}")
        if hasattr(e, 'response'):
            logging.error(f"API Response: {e.response}")
        return "", ""


def save_test_dataset(examples: List[Dict[str, Any]], output_file: str):
    """Save the test dataset to a JSON file."""
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump({"examples": examples}, f, indent=2, ensure_ascii=False)