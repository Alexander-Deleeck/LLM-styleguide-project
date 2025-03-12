from hyphenation_utils import load_markdown_file, extract_text_by_page, find_hyphenated_terms, add_context_to_matches, unique_hyphenated_terms, check_adherence_to_guidelines, get_azure_client, append_adherence_to_cases
import os
from dotenv import load_dotenv
import datetime
import json
import time

load_dotenv()

TEST_FILEPATH = os.path.join(os.path.dirname(__file__), 'test-docs/base-doc-hyphenation.docx')
RULE_FILEPATH = os.path.join(os.path.dirname(__file__), 'hyphenation-rule-new.md')
print(f"TEST_FILEPATH: {TEST_FILEPATH}")
print(f"RULE_FILEPATH: {RULE_FILEPATH}")
BATCH_SIZE = 4

def run_hyphenation_pipeline():
    pages = extract_text_by_page(TEST_FILEPATH)

    matches = find_hyphenated_terms(pages[0])

    context_matches = add_context_to_matches(pages[0], matches)
    unique_matches = [value for key, value in unique_hyphenated_terms(context_matches).items()]

    azure_client = get_azure_client()
    lexical_guideline = load_markdown_file(RULE_FILEPATH)

    merged_cases = []

    for i in range(0, len(unique_matches), BATCH_SIZE):
        try:
            adherence_cases = check_adherence_to_guidelines(lexical_guideline, unique_matches[i:i+BATCH_SIZE], azure_client)
            print(f"Adherence cases: {adherence_cases}\n\n")
            try:
                merged_cases = append_adherence_to_cases(unique_matches[i:i+BATCH_SIZE], adherence_cases, final_cases=merged_cases)
            except Exception as e:
                print(f"Error with merging {i}: {e}")
                time.sleep(30)
                continue
        except Exception as e:
            print(f"Error in batch {i}: {e}")
            time.sleep(30)
            continue
        

    output_filename = f"hyphenation-results-{datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.json"
    with open(output_filename, 'w') as f:
        json.dump(merged_cases, f, indent=4)
        
    return "Hyphenation pipeline completed successfully"
        

if __name__ == "__main__":
    run_hyphenation_pipeline()

