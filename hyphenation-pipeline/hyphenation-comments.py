from hyphenation_utils_spire import add_comments_to_docx
import os


TEST_FILEPATH = os.path.join(os.path.dirname(__file__), 'test-docs/base-doc-hyphenation.docx')
RULE_FILEPATH = os.path.join(os.path.dirname(__file__), 'hyphenation-rule-new.md')
RESULTS_JSON_FILEPATH = os.path.join(os.path.dirname(__file__), 'hyphenation-results-2025-03-11-19-33-33.json')

if __name__ == "__main__":
    docx_filepath = TEST_FILEPATH
    json_filepath = RESULTS_JSON_FILEPATH

    add_comments_to_docx(docx_filepath, json_filepath)