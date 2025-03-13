import json
import os
import uuid
from spire.doc import Document, Comment, CommentMark, CommentMarkType
from spire.doc import FileFormat

def add_comments_to_docx(docx_path, json_path, output_path=None):
    """
    Adds comments to a .docx file for terms marked as 'adheres_to_rule': False in the results.json file.

    Parameters:
        docx_path (str): Path to the .docx file.
        json_path (str): Path to the JSON results file.

    Output:
        Saves a new `.docx` file with comments added for incorrect hyphenations.
    """
    # Load the Word document using Spire.Doc
    doc = Document()
    doc.LoadFromFile(docx_path)

    # Load results JSON
    with open(json_path, 'r', encoding='utf-8') as f:
        results = json.load(f)

    # Filter out cases that do not adhere to the rule
    incorrect_cases = [case for case in results if not case["adheres_to_rule"]]

    # Process paragraphs to find exact matches
    for case in incorrect_cases:
        match_text = case["match_term"]

        # Find all instances of the match_text in the document
        text_selections = doc.FindAllString(match_text, True, True)

        for text_selection in text_selections:
            # Create a comment and set the content and author
            comment = Comment(doc)
            comment.Body.AddParagraph().Text = f"Incorrect hyphenation: '{match_text}' should be revised."
            comment.Format.Author = "Proofreader"

            # Get the text range and the paragraph it belongs to
            text_range = text_selection.GetAsOneRange()
            paragraph = text_range.OwnerParagraph

            # Insert the comment after the found text
            paragraph.ChildObjects.Insert(paragraph.ChildObjects.IndexOf(text_range) + 1, comment)

            # Create comment start and end marks and set their IDs
            comment_start = CommentMark(doc, CommentMarkType.CommentStart)
            comment_end = CommentMark(doc, CommentMarkType.CommentEnd)
            comment_start.CommentId = comment.Format.CommentId
            comment_end.CommentId = comment.Format.CommentId

            # Insert the comment start and end marks before and after the found text
            paragraph.ChildObjects.Insert(paragraph.ChildObjects.IndexOf(text_range), comment_start)
            paragraph.ChildObjects.Insert(paragraph.ChildObjects.IndexOf(text_range) + 1, comment_end)

    # Generate output filename if not provided
    if output_path is None:
        output_path = generate_output_filename(docx_path)

    # Save the modified document
    doc.SaveToFile(output_path, FileFormat.Docx)
    print(f"✅ Comments added and document saved to {output_path}")
    return

def generate_output_filename(docx_path):
    """
    Generates a new filename with `_hyphenation_comments` before the `.docx` extension.

    Parameters:
        docx_path (str): The original file path.

    Returns:
        str: The modified output file path.
    """
    uuid_str = str(uuid.uuid1()).split('-')[0]
    base, ext = os.path.splitext(docx_path)
    if 'hyphenation' in base.split('/')[-1]:
        return f"{base}_comments_{uuid_str}{ext}"
    else:
        return f"{base}_hyphenation_comments_{uuid_str}{ext}"
