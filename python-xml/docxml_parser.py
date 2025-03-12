import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re

current_dir = os.path.dirname(os.path.abspath(__file__))
project_dir = os.path.dirname(current_dir)
proof_dir = os.path.join(project_dir, "dataset-proofreading")
results_dir = os.path.join(current_dir, "results")

def extract_text_from_runs(runs):
    """ Extracts and joins text from <w:t> elements inside <w:r> runs. """
    return "".join(run.text for run in runs if run.text)

def get_surrounding_text(text, target, window=5):
    """ Extracts a few words before and after the target text for context. """
    if not target.strip():  # Check for empty or whitespace-only target
        return text  # If no valid target, return full text
    
    words = text.split()
    target_words = target.split()
    
    if not target_words:  # Edge case where splitting fails
        return text
    
    try:
        index = words.index(target_words[0])  # Use first word of target for indexing
        start = max(0, index - window)
        end = min(len(words), index + window + len(target_words))
        return "".join(words[start:end])
    except ValueError:
        return text  # Return full text if not found


def is_partial_word_modification(original, corrected):
    """ Detects if the correction is a partial-word change (e.g., color → colour). """
    if len(original) != len(corrected) and original.lower().startswith(corrected.lower()):
        return True
    return False

def merge_deletions(deletions):
    """ Merges logically connected deletions into a single entry. """
    if not deletions:
        return []

    merged = []
    current_del = deletions[0]

    for next_del in deletions[1:]:
        # Merge if they are in the same sentence and consecutive in text flow
        if current_del[0] == next_del[0]:
            current_del = (current_del[0], current_del[1] + " " + next_del[1])
        else:
            merged.append(current_del)
            current_del = next_del

    merged.append(current_del)
    return merged




def extract_tracked_changes(docx_path):
    """ Extracts tracked changes from a .docx file and returns a structured dataset. """

    # Unzip the .docx file to extract the document XML
    with zipfile.ZipFile(docx_path, 'r') as docx:
        xml_content = docx.read('word/document.xml', 'utf-8')

    # Parse the XML
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    root = ET.fromstring(xml_content)

    data = []
    unique_id = 1

    # Iterate through paragraphs to find changes
    for paragraph in root.findall('.//w:p', ns):
        full_text = extract_text_from_runs(paragraph.findall('.//w:t', ns))
        deletions = []
        insertions = []

        # Find deletions
        for del_element in paragraph.findall('.//w:del', ns):
            del_id = del_element.attrib.get('w:id', str(unique_id))
            del_text = extract_text_from_runs(del_element.findall('.//w:delText', ns))
            if del_text:
                deletions.append((del_id, del_text))

        # Merge related deletions
        deletions = merge_deletions(deletions)

        # Find insertions
        for ins_element in paragraph.findall('.//w:ins', ns):
            ins_id = ins_element.attrib.get('w:id', str(unique_id))
            ins_text = extract_text_from_runs(ins_element.findall('.//w:t', ns))
            if ins_text:
                insertions.append((ins_id, ins_text))

        # Process changes
        for del_id, del_text in deletions:
            matching_insert = next((ins_text for ins_id, ins_text in insertions if ins_id == del_id), None)

            if matching_insert:
                # Replacement = (deletion + insertion)
                change_type = "both"  
                correction_single = matching_insert
            else:
                change_type = "deletion"
                correction_single = None

            if del_text.strip():
                original_partial = get_surrounding_text(full_text, del_text)
                corrected_sentence = full_text.replace(del_text, correction_single if correction_single else "")

            # Detect partial-word modifications
            if correction_single and is_partial_word_modification(del_text, correction_single):
                change_type = "partial-word modification"

            data.append({
                "ID": unique_id,
                "original_single": del_text,
                "original_partial": original_partial,
                "original_sentence": full_text,
                "correction_single": correction_single,
                "correction_partial": corrected_sentence, 
                "corrected_sentence": corrected_sentence,
                "change_type": change_type
            })

            unique_id += 1  

        for ins_id, ins_text in insertions:
            if not any(del_id == ins_id for del_id, _ in deletions):  
                change_type = "insertion"
                original_partial = get_surrounding_text(full_text, ins_text)
                corrected_sentence = full_text.replace(ins_text, ins_text)  

                data.append({
                    "ID": unique_id,
                    "original_single": None,
                    "original_partial": original_partial,
                    "original_sentence": full_text,
                    "correction_single": ins_text,
                    "correction_partial": original_partial.replace(ins_text, ins_text),
                    "corrected_sentence": corrected_sentence,
                    "change_type": change_type
                })

                unique_id += 1  

    # Convert to DataFrame
    df = pd.DataFrame(data)
    return df

if __name__ == "__main__":
    # Example usage
    #test_foldername = "2023.00274"
    max_files = 10
    files_processed = 0
    
    for publication_folder in os.listdir(proof_dir):
        for file in os.listdir(os.path.join(proof_dir, publication_folder)):
            if file.endswith(".docx") and file.startswith("CORR_"):
                try:
                    test_filepath = os.path.join(proof_dir, publication_folder, file)
                    df_changes = extract_tracked_changes(test_filepath)

                    # Save to CSV for further analysis
                    df_changes.to_excel(os.path.join(results_dir, f"{publication_folder}-dataset-xml-corrections.xlsx"), index=False)

                    files_processed += 1
                    break
                except Exception as e:
                    print(f"Error processing {file}: {e}")
                    continue
            
        if files_processed >= max_files:
            break
