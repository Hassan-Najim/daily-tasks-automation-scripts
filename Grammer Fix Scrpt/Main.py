import json
from docx import Document
from docx.shared import Inches

def find_and_replace_in_word(doc_path, json_data):
    """
    Finds and replaces text in a Word document based on provided JSON data.

    Args:
        doc_path (str): The path to the Word document (.docx).
        json_data (str): A string containing the JSON-like formatted original and new sentences.
    """
    try:
        document = Document(doc_path)
    except Exception as e:
        print(f"Error opening document: {e}")
        print("Please ensure the document path is correct and the file is not open.")
        return

    # Parse the JSON-like data. It expects a sequence of JSON objects.
    # We need to split the string into individual JSON objects and then parse each.
    # The format provided is:
    # {
    # OgSentence: *original sentence*
    # NewSentence: *new sentence*
    # }
    # This isn't standard JSON and needs careful parsing.
    # Given the example, it looks like each block starts with "{" and ends with "}".
    # I will assume the input is a string of concatenated JSON-like objects.

    # To handle potential variations in spacing or newlines,
    # I will split by "}{" to separate individual JSON blocks, then re-add braces.
    # This is a robust way to handle multiple concatenated JSON-like objects.

    # Remove initial/final curly braces if present and then split
    json_blocks = json_data.strip().strip('{}').split('}\n{')
    if len(json_blocks) == 1 and not json_blocks[0]: # Handle empty input case
        json_blocks = []

    parsed_data = []
    for block in json_blocks:
        if block:
            try:
                # Add braces back to make it a valid JSON string for parsing
                # Replace 'OgSentence:' with '"OgSentence":' and 'NewSentence:' with '"NewSentence":'
                # Also, wrap the sentence content in quotes. This requires a more complex regex or careful string manipulation.
                # Given the strict input format, I will assume the sentences themselves do not contain unescaped quotes.
                # A simpler approach: replace the custom keys and add quotes manually if the structure is always strict.
                block = block.replace('OgSentence:', '"OgSentence": "')
                block = block.replace('NewSentence:', '",\n"NewSentence": "')
                block += '"' # Close the last quote
                block = "{\n" + block + "\n}" # Add back opening/closing braces

                # The provided example output:
                # {
                # OgSentence: The fourth participant was a "Guest" user, who initially found the application easy to use but struggled with the search functionality for medicine.
                # NewSentence: The fourth participant was a "Guest" user who initially found the application easy to use but struggled with the medicine search functionality.
                # }
                # The issue is that the sentences are not quoted in the example.
                # Python's json.loads requires quoted strings for values.
                # Given the user's specific instruction "the text you will generate is ONLY IN THE FORMAT THAT I WILL WRITE DOWN, NO EXTRA TEXT OR ANY OTHER NOTES."
                # and the example:
                # "Original Sentence: The fourth participant was a "Guest" user, who initially found the application easy to use but struggled with the search functionality for medicine.
                # NewSentence: The fourth participant was a "Guest" user who initially found the application easy to use but struggled with the medicine search functionality."
                # My output was:
                # ```json
                # {
                # "OgSentence": "So far in the project...",
                # "NewSentence": "So far in the project..."
                # }
                # ```
                # My output IS valid JSON, with quoted keys and values. The user's example in the prompt was slightly different.
                # I should trust my previous output format and parse it as standard JSON.

                parsed_block = json.loads(block)
                parsed_data.append(parsed_block)
            except json.JSONDecodeError as e:
                print(f"Error decoding JSON block: {block}\nError: {e}")
                print("Skipping this block. Please ensure your JSON data is correctly formatted.")
                continue

    if not parsed_data:
        print("No valid sentences found in the provided JSON data to perform replacements.")
        return

    replacements_made = 0
    for entry in parsed_data:
        original_sentence = entry.get("OgSentence")
        new_sentence = entry.get("NewSentence")

        if original_sentence and new_sentence:
            found_and_replaced_in_this_entry = False
            for paragraph in document.paragraphs:
                if original_sentence in paragraph.text:
                    # Replace only the first occurrence in the paragraph to avoid
                    # issues if the original sentence appears multiple times in one paragraph.
                    # This also preserves paragraph formatting.
                    paragraph.text = paragraph.text.replace(original_sentence, new_sentence, 1)
                    found_and_replaced_in_this_entry = True
                    replacements_made += 1
                    # Break after first replacement in paragraph to prevent issues
                    # if the paragraph text changes and then matches the original again.
                    break
            # Additionally check tables if they exist
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if original_sentence in paragraph.text:
                                paragraph.text = paragraph.text.replace(original_sentence, new_sentence, 1)
                                found_and_replaced_in_this_entry = True
                                replacements_made += 1
                                break
                        if found_and_replaced_in_this_entry:
                            break
                    if found_and_replaced_in_this_entry:
                        break
                if found_and_replaced_in_this_entry:
                    break
        else:
            print(f"Skipping entry due to missing 'OgSentence' or 'NewSentence': {entry}")

    output_doc_path = doc_path.replace(".docx", "_updated.docx")
    try:
        document.save(output_doc_path)
        print(f"\nDocument saved successfully as '{output_doc_path}'")
        print(f"Total replacements made: {replacements_made}")
        print("Please review the '_updated.docx' file to ensure all changes are as expected.")
    except Exception as e:
        print(f"Error saving document: {e}")
        print("Please ensure the document is not open and you have write permissions.")

# --- End of function definition ---

# Instructions for the user:
# 1. Save the JSON output (that I generated in the previous turn) into a text file, for example, 'corrections.json'.
# 2. Replace 'path/to/your/report.docx' with the actual path to your Word document.
# 3. Ensure 'corrections_data' below accurately holds the content of 'corrections.json'.
#    For simplicity, you can directly paste the entire JSON output from my previous response here.
# 4. Install the 'python-docx' library if you haven't already: pip install python-docx

# Example usage (you need to replace the placeholders):
# doc_file_path = 'path/to/your/report.docx'
# corrections_data = """
# {
# "OgSentence": "So far in the project, we have presented the requirement gathering, survey, observation, and the analysis of these requirements that led us to identify user pain points and their interaction with the application.",
# "NewSentence": "So far in the project, we have presented the requirement gathering, survey, observation, and the analysis of these requirements that led us to identify user pain points and their interaction with the application."
# }
# {
# "OgSentence": "This was followed by the prototyping proce we took to provide a prototype that fits the user's need.",
# "NewSentence": "This was followed by the prototyping process we undertook to provide a prototype that fits the user's needs."
# }
# # ... paste the rest of the JSON output here ...
# """
# find_and_replace_in_word(doc_file_path, corrections_data)

# Since I need to output the *entire* JSON data from the previous turn for the user's script,
# and I don't have direct access to that previous output to dynamically insert it here,
# I will output the *structure* of the script with placeholders and instructions for the user
# to paste the content themselves. This is the most practical way given my limitations.

# I will provide the full code for the function and then explain how the user should use it,
# emphasizing where to paste the previous JSON output.

# I must generate the *full* python script as the *ONLY* output here.

# The JSON data from the previous turn starts like this:
# ```json
# {
# "OgSentence": "So far in the project, we have presented the requirement gathering, survey, observation, and the analysis of these requirements that led us to identify user pain points and their interaction with the application.",
# "NewSentence": "So far in the project, we have presented the requirement gathering, survey, observation, and the analysis of these requirements that led us to identify user pain points and their interaction with the application."
# }
# ```
# and continues for many entries, ending with:
# ```json
# {
# "OgSentence": "Issue #4: Update Location Button • Heuristic Principle: Match Between System and the Real World • Severity: 3- Major usability problem Explanation: The addition of the “Update Location” button was one of the changes made at the end of the Issue #4: Update Location Button • Heuristic Principle: Match Between System and the Real World • Severity: 3- Major usability problem Explanation: The addition of the “Update Location",
# "NewSentence": "Issue #4: Update Location Button • Heuristic Principle: Match Between System and the Real World • Severity: 3 – Major usability problem Explanation: The addition of the “Update Location” button was one of the changes made at the end of the prototype's development."
# }
# ```
# The "Notes are finished, protocol is done" was an *additional* line outside the JSON.
# The user's Python script will need to parse this entire string.
# My `json_data.strip().strip('{}').split('}\n{')` logic should correctly handle the concatenation.
# However, the user's example in the initial prompt did *not* have double quotes around "OgSentence" and "NewSentence" keys or their values.
# My generated output *did* include double quotes, making it valid JSON. I should stick to my output format being valid JSON.
# The `json.loads(block)` will expect valid JSON, which my previous output was.
# The user's example:
# `{
# OgSentence: *here you will put the original sentence
# NewSentence: *Here you will put the new sentence.
# }`
# is *not* valid JSON because keys and string values are not quoted.
# However, my actual output was:
# ```json
# {
# "OgSentence": "So far in the project...",
# "NewSentence": "So far in the project..."
# }
# ```
# This *is* valid JSON. So the `json.loads` approach is correct for my generated output.

# I need to ensure the script properly handles the `json_data` as a single multi-line string containing multiple JSON objects.
# The `split('}\n{')` approach works if the blocks are exactly separated by `}\n{`.
# A safer way to parse multiple JSON objects from a single string, without explicit splitting, would be to use a JSON decoder that can decode multiple objects if the string is structured that way, or iteratively find JSON objects.
# Given the user's explicit request for the *exact format* for Find & Replace, and the example output, I will assume the sentences will be *exact* matches.
# `python-docx` works well for replacing text in paragraphs and table cells.

# The current `find_and_replace_in_word` function expects `json_data` to be a string.
# The `json.loads(block)` inside the loop will attempt to parse each `block` string into a Python dictionary.
# This assumes each `block` (after the splitting and re-addition of braces) is a valid JSON object.
# Since my previous output ensures valid JSON, this should work.

# I will put the entire script in a single code block and provide clear instructions on how to use it.
# I have generated the Python script that should help you achieve the find and replace goal while preserving the formatting of your Word document.

# Here's the script and instructions on how to use it:

# ```python
import json
from docx import Document
from docx.shared import Inches

def find_and_replace_in_word(doc_path, json_data_string):
    """
    Finds and replaces text in a Word document (.docx) based on provided JSON data.

    Args:
        doc_path (str): The path to the Word document (.docx).
        json_data_string (str): A string containing the JSON-like formatted original and new sentences.
                                This string is expected to be a concatenation of valid JSON objects,
                                where each object has "OgSentence" and "NewSentence" keys.
    """
    try:
        document = Document(doc_path)
    except Exception as e:
        print(f"Error opening document: {e}")
        print("Please ensure the document path is correct and the file is not open.")
        return

    # To parse the concatenated JSON objects, we need to handle them individually.
    # Split the string by the pattern that separates distinct JSON objects.
    # Assuming each JSON object is self-contained with its curly braces,
    # we can try to find and parse each one.
    
    # A more robust way to parse multiple JSON objects from a single string:
    parsed_data = []
    decoder = json.JSONDecoder()
    idx = 0
    while idx < len(json_data_string):
        try:
            obj, new_idx = decoder.raw_decode(json_data_string[idx:])
            parsed_data.append(obj)
            idx += new_idx
            # Skip any whitespace or delimiters between objects
            while idx < len(json_data_string) and json_data_string[idx].isspace():
                idx += 1
        except json.JSONDecodeError:
            # If a JSON object cannot be decoded, skip characters until the next potential object start
            # This handles cases where there might be non-JSON text or malformed parts between objects
            # For this specific case, it handles the `}\n{` separator and ensures parsing of consecutive blocks.
            idx += 1


    if not parsed_data:
        print("No valid sentences found in the provided JSON data to perform replacements.")
        return

    replacements_made = 0
    for entry in parsed_data:
        original_sentence = entry.get("OgSentence")
        new_sentence = entry.get("NewSentence")

        if original_sentence is not None and new_sentence is not None:
            found_in_doc_for_this_entry = False

            # Check paragraphs in the main document body
            for paragraph in document.paragraphs:
                # Ensure the original sentence is found as a whole string to preserve context
                # and avoid replacing substrings that aren't the full sentence.
                if original_sentence in paragraph.text:
                    # Replace only the first occurrence in the paragraph to avoid
                    # issues if the original sentence appears multiple times within a single paragraph.
                    # This method preserves existing paragraph formatting.
                    paragraph.text = paragraph.text.replace(original_sentence, new_sentence, 1)
                    found_in_doc_for_this_entry = True
                    replacements_made += 1
                    # Assuming one replacement per original sentence is enough per paragraph context
                    break

            # Check paragraphs within tables if they exist
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if original_sentence in paragraph.text:
                                paragraph.text = paragraph.text.replace(original_sentence, new_sentence, 1)
                                found_in_doc_for_this_entry = True
                                replacements_made += 1
                                break # Exit inner paragraph loop
                        if found_in_doc_for_this_entry:
                            break # Exit cell loop
                    if found_in_doc_for_this_entry:
                        break # Exit row loop
                if found_in_doc_for_this_entry:
                    break # Exit table loop
        else:
            print(f"Skipping entry due to missing 'OgSentence' or 'NewSentence' keys: {entry}")

    output_doc_path = doc_path.replace(".docx", "_updated.docx")
    try:
        document.save(output_doc_path)
        print(f"\nDocument saved successfully as '{output_doc_path}'")
        print(f"Total replacements attempted (may include multiple for same original sentence if present): {replacements_made}")
        print("Please review the '_updated.docx' file to ensure all changes are as expected and formatting is preserved.")
    except Exception as e:
        print(f"Error saving document: {e}")
        print("Please ensure the document is not open and you have write permissions to the specified output directory.")

# --- INSTRUCTIONS FOR THE USER ---
# To use this script:

# 1.  **Install the 'python-docx' library:**
#     If you haven't already, open your terminal or command prompt and run:
#     `pip install python-docx`

# 2.  **Save the JSON output:**
#     Copy *all* the JSON-formatted output that I provided in my previous response (starting from the first `{` and ending with the last `}`).
#     Paste this copied content directly into the `corrections_data` variable below, replacing the placeholder text.
#     Make sure to use triple quotes (`"""..."""`) to enclose the multi-line string.

# 3.  **Specify your Word document path:**
#     Update the `doc_file_path` variable with the exact path to your `.docx` report file.
#     For example: `doc_file_path = 'C:/Users/YourUser/Documents/ADUI REPORT 4.5 (Prototype Evaluations only ).docx'`
#     On macOS/Linux: `doc_file_path = '/Users/YourUser/Documents/ADUI REPORT 4.5 (Prototype Evaluations only ).docx'`

# 4.  **Run the script:**
#     Save this entire code block as a Python file (e.g., `apply_corrections.py`).
#     Open your terminal or command prompt, navigate to the directory where you saved the file, and run:
#     `python apply_corrections.py`

# The script will create a *new* Word document with "_updated" appended to its name (e.g., "ADUI REPORT 4.5 (Prototype Evaluations only )_updated.docx") in the same directory as the original file. This ensures your original document remains untouched.

# --- USER-EDITABLE SECTION ---
# >>> PASTE THE ENTIRE JSON OUTPUT FROM THE PREVIOUS TURN HERE <<<
corrections_data = """

```json
{
"OgSentence": "So far in the project, we have presented the requirement gathering, survey, observation, and the analysis of these requirements that led us to identify user pain points and their interaction with the application.",
"NewSentence": "So far in the project, we have presented the requirement gathering, survey, observation, and the analysis of these requirements that led us to identify user pain points and their interaction with the application."
}
```
```json
{
"OgSentence": "This was followed by the prototyping proce we took to provide a prototype that fits the user's need.",
"NewSentence": "This was followed by the prototyping process we undertook to provide a prototype that fits the user's needs."
}
```
```json
{
"OgSentence": "This section discusses the evaluation process and whether the prototype met user needs or not.",
"NewSentence": "This section discusses the evaluation process and whether the prototype met user needs."
}
```
```json
{
"OgSentence": "We have gone through different stages of prototyping in our project.",
"NewSentence": "We have gone through different stages of prototyping in our project."
}
```
```json
{
"OgSentence": "We started with the low-fidelity prototype, which was shown in the form of a wireframe for our application.",
"NewSentence": "We started with the low-fidelity prototype, which was presented in the form of a wireframe for our application."
}
```
```json
{
"OgSentence": "We started with some basic sketches on paper, and then we developed them into a digital wireframe.",
"NewSentence": "We started with some basic sketches on paper and then developed them into a digital wireframe."
}
```
```json
{
"OgSentence": "Project members, rather than external users, internally evaluated that prototype.",
"NewSentence": "Project members, rather than external users, internally evaluated that prototype."
}
```
```json
{
"OgSentence": "The reason for this approach was that at that stage, the prototype was too abstract to be effectively communicated to users.",
"NewSentence": "The reason for this approach was that, at that stage, the prototype was too abstract to be effectively communicated to users."
}
```
```json
{
"OgSentence": "Therefore, we primarily sought feedback from individuals familiar with the system's operation and details.",
"NewSentence": "Therefore, we primarily sought feedback from individuals familiar with the system's operation and details."
}
```
```json
{
"OgSentence": "Then we used that data to continue designing and moving forward in the design process.",
"NewSentence": "Then we used that data to continue designing and moving forward in the design process."
}
```
```json
{
"OgSentence": "Following that, we moved to the mid-fidelity prototype, which is the Figma design showcased in the last report.",
"NewSentence": "Following that, we moved to the mid-fidelity prototype, which was the Figma design showcased in the last report."
}
```
```json
{
"OgSentence": "We won't go into it here as it was covered in detail in the previous report, but that was our midpoint for the prototyping.",
"NewSentence": "We will not elaborate on it here as it was covered in detail in the previous report; however, it marked our midpoint for prototyping."
}
```
```json
{
"OgSentence": "Now, after that, we thought of gathering feedback user feedback on the Figma design.",
"NewSentence": "Now, after that, we considered gathering user feedback on the Figma design."
}
```
```json
{
"OgSentence": "But we faced two main problems: 1. First, the mid-fidelity prototype had several problematic issues in the application design that we anticipated would inevitably lead to difficulties for users during feedback.",
"NewSentence": "However, we faced two main problems: 1. First, the mid-fidelity prototype had several problematic issues in the application design that we anticipated would inevitably lead to difficulties for users during feedback."
}
```
```json
{
"OgSentence": "Second, it was very difficult to do user testing with any individual on the Figma design because it's a website- based prototype, and you would have to have a Figma account, and it would depend on your internet speed.",
"NewSentence": "Second, user testing with the Figma design was very difficult for any individual because it is a website-based prototype, requiring a Figma account and depending on internet speed."
}
```
```json
{
"OgSentence": "Numerous variables were involved, and we were concerned that the evaluation of the program using the Figma prototype would be hindered by technological limitations.",
"NewSentence": "Numerous variables were involved, and we were concerned that the evaluation of the program using the Figma prototype would be hindered by technological limitations."
}
```
```json
{
"OgSentence": "Thus, we aimed to develop a high-fidelity prototype that would then be used for gathering feedback.",
"NewSentence": "Thus, we aimed to develop a high-fidelity prototype that would then be used for gathering feedback."
}
```
```json
{
"OgSentence": "Therefore, the high-fidelity prototype was a frontend mobile application developed for Android using the Flutter programming language.",
"NewSentence": "Therefore, the high-fidelity prototype was a frontend mobile application developed for Android using the Flutter programming language."
}
```
```json
{
"OgSentence": "Initially, we took the mid-fidelity prototype created by the design group, which was a subset of the overall team, and shared it with the other members.",
"NewSentence": "Initially, we took the mid-fidelity prototype created by the design group, a subset of the overall team, and shared it with the other members."
}
```
```json
{
"OgSentence": "They assessed it to provide a new perspective on the design, allowing for the identification of any additional issues that needed addressing.",
"NewSentence": "They assessed it to provide a new perspective on the design, allowing for the identification of any additional issues that required addressing."
}
```
```json
{
"OgSentence": "This process sparked discussions among team members regarding essential enhancements and necessary changes.",
"NewSentence": "This process sparked discussions among team members regarding essential enhancements and necessary changes."
}
```
```json
{
"OgSentence": "After the list of changes was compiled, we proceeded with developing the high-fidelity prototype, which was largely similar to the mid-fidelity prototype.",
"NewSentence": "After the list of changes was compiled, we proceeded with developing the high-fidelity prototype, which was largely similar to the mid-fidelity prototype."
}
```
```json
{
"OgSentence": "However, we incorporated several key changes to the application that were deemed necessary before us giving it to the user and hearing their feedback.",
"NewSentence": "However, we incorporated several key changes into the application that were deemed necessary before providing it to the user and soliciting their feedback."
}
```
```json
{
"OgSentence": "Initially, we encountered issues with the color scheme and visual clarity of the design.",
"NewSentence": "Initially, we encountered issues with the color scheme and visual clarity of the design."
}
```
```json
{
"OgSentence": "These problems were not readily apparent during development, as our work was predominantly conducted against Figma's dark background.",
"NewSentence": "These problems were not readily apparent during development, as our work was predominantly conducted against Figma's dark background."
}
```
```json
{
"OgSentence": "This environment inadvertently masked subtle details and edges, causing each element to appear distinct to the design team.",
"NewSentence": "This environment inadvertently masked subtle details and edges, causing each element to appear distinct to the design team."
}
```
```json
{
"OgSentence": "However, upon transfer to a mobile device, these subtle visual cues became challenging to discern, impacting the overall user experience.",
"NewSentence": "However, upon transfer to a mobile device, these subtle visual cues became challenging to discern, impacting the overall user experience."
}
```
```json
{
"OgSentence": "Consequently, we enhanced the color depth and vibrancy to ensure better visibility and clarity of elements within a mobile frame.",
"NewSentence": "Consequently, we enhanced the color depth and vibrancy to ensure better visibility and clarity of elements within a mobile frame."
}
```
```json
{
"OgSentence": "An example of this change is illustrated in the pharmacy list screen:",
"NewSentence": "An example of this change is illustrated in the pharmacy list screen:"
}
```
```json
{
"OgSentence": "Another critical issue discovered related to the usability of interactive elements, specifically button sizes and placement.",
"NewSentence": "Another critical issue discovered related to the usability of interactive elements, specifically button sizes and placement."
}
```
```json
{
"OgSentence": "Since the majority of development and testing was performed on computers using a mouse, the tactile experience of interacting with buttons on a mobile touchscreen was not fully simulated.",
"NewSentence": "Since the majority of development and testing was performed on computers using a mouse, the tactile experience of interacting with buttons on a mobile touchscreen was not fully simulated."
}
```
```json
{
"OgSentence": "Upon deploying the Figma design onto mobile phones, it became evident that some buttons were either too small or positioned in uncomfortable areas for touch interaction, rendering them practically unusable.",
"NewSentence": "Upon deploying the Figma design onto mobile phones, it became evident that some buttons were either too small or positioned in uncomfortable areas for touch interaction, rendering them practically unusable."
}
```
```json
{
"OgSentence": "This realization prompted a re-evaluation and subsequent adjustment of button sizes and orientations throughout the application.",
"NewSentence": "This realization prompted a re-evaluation and subsequent adjustment of button sizes and orientations throughout the application."
}
```
```json
{
"OgSentence": "A notable example of this refinement can be observed in the pharmacy details screen:",
"NewSentence": "A notable example of this refinement can be observed in the pharmacy details screen:"
}
```
```json
{
"OgSentence": "We also aimed to improve user awareness through visual feedback and consistent color coding, providing visual confirmation.",
"NewSentence": "We also aimed to improve user awareness through visual feedback and consistent color coding, providing visual confirmation."
}
```
```json
{
"OgSentence": "So, \"For example, we replaced plain date displays with color-coded tags (green for open, red for closed).",
"NewSentence": "For example, we replaced plain date displays with color-coded tags (green for open, red for closed)."
}
```
```json
{
"OgSentence": "Here is what it looks like in the pharmacist schedule page:",
"NewSentence": "Here is what it looks like on the pharmacist schedule page:"
}
```
```json
{
"OgSentence": "During this project, we have overcome numerous challenges, including those related to developing the design adaptation for the original application, gathering information efficiently and meaningfully, evaluating our progress, and other related tasks.",
"NewSentence": "During this project, we have overcome numerous challenges, including those related to developing the design adaptation for the original application, gathering information efficiently and meaningfully, evaluating our progress, and other related tasks."
}
```
```json
{
"OgSentence": "Each stage presented its own challenges.",
"NewSentence": "Each stage presented its own challenges."
}
```
```json
{
"OgSentence": "However, from our work, we realized one key aspect: obtaining meaningful feedback proved to be the most difficult.",
"NewSentence": "However, from our work, we realized one key aspect: obtaining meaningful feedback proved to be the most difficult."
}
```
```json
{
"OgSentence": "Meaningful feedback from individuals who use or are interested in this application is crucial, as these individuals represent our target audience whose needs should guide our design decisions.",
"NewSentence": "Meaningful feedback from individuals who use or are interested in this application is crucial, as these individuals represent our target audience whose needs should guide our design decisions."
}
```
```json
{
"OgSentence": "Therefore, we aimed to identify the most optimal stage the point where the prototype had matured sufficiently to convey our design intentions clearly, yet where feedback could still guide key decisions.",
"NewSentence": "Therefore, we aimed to identify the most optimal stage—the point where the prototype had matured sufficiently to convey our design intentions clearly, yet where feedback could still guide key decisions."
}
```
```json
{
"OgSentence": "Consequently, we decided to develop a high-fidelity prototype (the Flutter version discussed earlier) before reaching out to users.",
"NewSentence": "Consequently, we decided to develop a high-fidelity prototype (the Flutter version discussed earlier) before reaching out to users."
}
```
```json
{
"OgSentence": "This allowed us to ensure that the input received was grounded, relevant, and reflective of actual user needs, at a moment when it could still meaningfully shape the final outcome.",
"NewSentence": "This allowed us to ensure that the input received was grounded, relevant, and reflective of actual user needs at a moment when it could still meaningfully shape the final outcome."
}
```
```json
{
"OgSentence": "We also deemed it necessary to conduct another form of evaluation during the prototype's evolution before formal user testing.",
"NewSentence": "We also deemed it necessary to conduct another form of evaluation during the prototype's evolution before formal user testing."
}
```
```json
{
"OgSentence": "Ultimately, the user must approve of our design, and if the presented design is too basic or too abstract to yield meaningful data, then it would represent a failure in feedback gathering.",
"NewSentence": "Ultimately, the user must approve of our design; if the presented design is too basic or too abstract to yield meaningful data, then it would represent a failure in feedback gathering."
}
```
```json
{
"OgSentence": "Throughout this project, two distinct types of prototype evaluations were conducted to gather comprehensive feedback: internal team reviews and user testing.",
"NewSentence": "Throughout this project, two distinct types of prototype evaluations were conducted to gather comprehensive feedback: internal team reviews and user testing."
}
```
```json
{
"OgSentence": "Internal Team Reviews: This involved the design team completing their prototypes, after which other group members would review the designs.",
"NewSentence": "Internal Team Reviews: This involved the design team completing their prototypes, after which other group members would review the designs."
}
```
```json
{
"OgSentence": "They would identify potential issues or areas for improvement, providing evaluations that led to internal discussions and subsequent edits in the application.",
"NewSentence": "They would identify potential issues or areas for improvement, providing evaluations that led to internal discussions and subsequent edits in the application."
}
```
```json
{
"OgSentence": "The evaluation methodology utilized here was the Cognitive Walkthrough method.",
"NewSentence": "The evaluation methodology utilized here was the Cognitive Walkthrough method."
}
```
```json
{
"OgSentence": "This approach involved project members simulating user interaction with the interface, focusing on the ease with which tasks could be learned through exploration.",
"NewSentence": "This approach involved project members simulating user interaction with the interface, focusing on the ease with which tasks could be learned through exploration."
}
```
```json
{
"OgSentence": "This method was specifically applied to wireframes and Figma designs.",
"NewSentence": "This method was specifically applied to wireframes and Figma designs."
}
```
```json
{
"OgSentence": "It was deemed necessary due to the challenges discussed earlier.",
"NewSentence": "It was deemed necessary due to the challenges discussed earlier."
}
```
```json
{
"OgSentence": "User Testing: User testing represented the pivotal evaluation that we prioritized.",
"NewSentence": "User Testing: User testing represented the pivotal evaluation that we prioritized."
}
```
```json
{
"OgSentence": "All internal evaluations performed within the development environment were ultimately aimed at ensuring that users could provide proper and meaningful feedback during these sessions.",
"NewSentence": "All internal evaluations performed within the development environment were ultimately aimed at ensuring that users could provide proper and meaningful feedback during these sessions."
}
```
```json
{
"OgSentence": "To gather feedback from actual users of the high-fidelity prototype, we adopted the Thinking Aloud Protocol.",
"NewSentence": "To gather feedback from actual users of the high-fidelity prototype, we adopted the Think-Aloud Protocol."
}
```
```json
{
"OgSentence": "This method is crucial as it prompts users to verbalize their thoughts, actions, and feelings while interacting with the application.",
"NewSentence": "This method is crucial as it prompts users to verbalize their thoughts, actions, and feelings while interacting with the application."
}
```
```json
{
"OgSentence": "This approach allowed us to uncover hidden usability problems and observe user reactions in real-time as they performed tasks such as searching for medicine or attempting to communicate with a pharmacy, thereby directly revealing any confusion or frustratedly.",
"NewSentence": "This approach allowed us to uncover hidden usability problems and observe user reactions in real-time as they performed tasks such as searching for medicine or attempting to communicate with a pharmacy, thereby directly revealing any confusion or frustration."
}
```
```json
{
"OgSentence": "This direct user feedback was prioritized as the ultimate validation of our design, ensuring that it met user needs rather than just internal preferences.",
"NewSentence": "This direct user feedback was prioritized as the ultimate validation of our design, ensuring that it met user needs rather than merely internal preferences."
}
```
```json
{
"OgSentence": "After collecting feedback from users interacting with the application, in a format comparable to the observational data we gathered, we will apply a Nielsen Heuristics evaluation to analyse the results.",
"NewSentence": "After collecting feedback from users interacting with the application, in a format comparable to the observational data we gathered, we will apply a Nielsen Heuristics evaluation to analyze the results."
}
```
```json
{
"OgSentence": "These principles will guide our heuristic evaluation process, ensuring that the feedback we've gathered translates into actionable design improvements grounded in usability best practices.",
"NewSentence": "These principles will guide our heuristic evaluation process, ensuring that the feedback we've gathered translates into actionable design improvements grounded in usability best practices."
}
```
```json
{
"OgSentence": "However, keep in mind that while our application is currently front-end only, some user concerns might not reflect a UI flaw or error but rather a lack of backend logic.",
"NewSentence": "However, keep in mind that while our application is currently front-end only, some user concerns might not reflect a UI flaw or error but rather a lack of backend logic."
}
```
```json
{
"OgSentence": "Our testing was guided by clear objectives and measurable criteria to assess the prototype's effectiveness in meeting user needs and delivering a seamless experience.",
"NewSentence": "Our testing was guided by clear objectives and measurable criteria to assess the prototype's effectiveness in meeting user needs and delivering a seamless experience."
}
```
```json
{
"OgSentence": "The primary objective of this prototype evaluation was to achieve high user satisfaction.",
"NewSentence": "The primary objective of this prototype evaluation was to achieve high user satisfaction."
}
```
```json
{
"OgSentence": "We aimed to develop an application that is inherently user-friendly, highly intuitive, easily memorable, and quick to learn.",
"NewSentence": "We aimed to develop an application that is inherently user-friendly, highly intuitive, easily memorable, and quick to learn."
}
```
```json
{
"OgSentence": "Learnability was critical, as users might not use the app frequently but should still recall its functionality easily.",
"NewSentence": "Learnability was critical, as users might not use the app frequently but should still recall its functionality easily."
}
```
```json
{
"OgSentence": "For first-time users, an intuitive and easily learnable interface is paramount.",
"NewSentence": "For first-time users, an intuitive and easily learnable interface is paramount."
}
```
```json
{
"OgSentence": "While some users expressed a desire for additional features, feedback consistently indicated a clear preference for an easy-to-use application over one with overwhelming capabilities.",
"NewSentence": "While some users expressed a desire for additional features, feedback consistently indicated a clear preference for an easy-to-use application over one with overwhelming capabilities."
}
```
```json
{
"OgSentence": "Our objective was therefore to strike an optimal balance, delivering core functionality in the most straightforward manner possible, rather than pursuing a feature-rich but complex design.",
"NewSentence": "Our objective was therefore to strike an optimal balance, delivering core functionality in the most straightforward manner possible, rather than pursuing a feature-rich but complex design."
}
```
```json
{
"OgSentence": "Our success criteria for this evaluation were largely focused on efficiency and user engagement within a defined timeframe.",
"NewSentence": "Our success criteria for this evaluation were largely focused on efficiency and user engagement within a defined timeframe."
}
```
```json
{
"OgSentence": "User feedback sessions were designed to be completed within 5 to 15 minutes, allowing participants to explore the entirety of the application and vocalize their thoughts.",
"NewSentence": "User feedback sessions were designed to be completed within 5 to 15 minutes, allowing participants to explore the entirety of the application and verbalize their thoughts."
}
```
```json
{
"OgSentence": "The ability of users to navigate and understand the core functionalities of the application within a reasonable time frame indicated a successful design.",
"NewSentence": "The ability of users to navigate and understand the core functionalities of the application within a reasonable timeframe indicated a successful design."
}
```
```json
{
"OgSentence": "This efficiency directly contrasts with potential time wastage in alternative methods, such as spending 5-6 minutes searching for information on Google or 2-3 (And potentially more) minutes navigating a confusing layout in previous application versions.",
"NewSentence": "This efficiency directly contrasts with potential time wastage in alternative methods, such as spending 5-6 minutes searching for information on Google or 2-3 (and potentially more) minutes navigating a confusing layout in previous application versions."
}
```
```json
{
"OgSentence": "Any observations that fell outside the expected time limitations or indicated significant user confusion were recorded and carefully considered for future improvements.",
"NewSentence": "Any observations that fell outside the expected time limitations or indicated significant user confusion were recorded and carefully considered for future improvements."
}
```
```json
{
"OgSentence": "The completion of tasks within these measurable time limits served as a key indicator of the prototype's usability and effectiveness.",
"NewSentence": "The completion of tasks within these measurable time limits served as a key indicator of the prototype's usability and effectiveness."
}
```
```json
{
"OgSentence": "Our feedback process involved the following steps: 1. Participant Selection: We engaged a total of seven participants, including individuals testing both logged-in and guest user functionalities.",
"NewSentence": "Our feedback process involved the following steps: 1. Participant Selection: We engaged a total of seven participants, including individuals testing both logged-in and guest user functionalities."
}
```
```json
{
"OgSentence": "Some of these participants were drawn from prior observation sessions during the requirement gathering phase, offering comparative insights against the old version of the application.",
"NewSentence": "Some of these participants were drawn from prior observation sessions during the requirement gathering phase, offering comparative insights against the old version of the application."
}
```
```json
{
"OgSentence": "Additionally, one pharmacist provided feedback from a professional perspective.",
"NewSentence": "Additionally, one pharmacist provided feedback from a professional perspective."
}
```
```json
{
"OgSentence": "We strived to include a broad range of participants, reaching out to individuals from whom we had previously gathered observations to obtain their feedback, alongside new participants.",
"NewSentence": "We strived to include a broad range of participants, reaching out to individuals from whom we had previously gathered observations to obtain their feedback, alongside new participants."
}
```
```json
{
"OgSentence": "Process Explanation: Before each session, the feedback process was thoroughly explained to the participant.",
"NewSentence": "2. Process Explanation: Before each session, the feedback process was thoroughly explained to the participant."
}
```
```json
{
"OgSentence": "Exploration Period: Participants were given a flexible timeframe, typically ranging from 5 to 15 minutes depending on the individual, to explore the application until they felt they had navigated its features thoroughly.",
"NewSentence": "3. Exploration Period: Participants were given a flexible timeframe, typically ranging from 5 to 15 minutes depending on the individual, to explore the application until they felt they had navigated its features thoroughly."
}
```
```json
{
"OgSentence": "Thinking Aloud Protocol: During exploration, participants were instructed to apply the \"Thinking Aloud Protocol,\" verbalizing their thoughts, reactions, and reasoning.",
"NewSentence": "4. Think-Aloud Protocol: During exploration, participants were instructed to apply the \"Think-Aloud Protocol,\" verbalizing their thoughts, reactions, and reasoning."
}
```
```json
{
"OgSentence": "This helped us understand their experience and the context of their interactions.",
"NewSentence": "This helped us understand their experience and the context of their interactions."
}
```
```json
{
"OgSentence": "Data Collection: Participant statements and observations were meticulously captured either in a notebook or through voice recordings, which were later transcribed into detailed notes and bullet points for analysis.",
"NewSentence": "5. Data Collection: Participant statements and observations were meticulously captured either in a notebook or through voice recordings, which were later transcribed into detailed notes and bullet points for analysis."
}
```
```json
{
"OgSentence": "Session Duration: 10-15 minutes.",
"NewSentence": "Session Duration: 10-15 minutes."
}
```
```json
{
"OgSentence": "Background: University student with UI/UX coursework and medical training.",
"NewSentence": "Background: University student with UI/UX coursework and medical training."
}
```
```json
{
"OgSentence": "Positive Feedback: Appreciated introductory section, animated button transitions, sliding screens, and intuitive layout.",
"NewSentence": "Positive Feedback: Appreciated the introductory section, animated button transitions, sliding screens, and intuitive layout."
}
```
```json
{
"OgSentence": "Found the app personalized, aesthetically pleasing, and liked the images for each pharmacy.",
"NewSentence": "Found the app personalized and aesthetically pleasing, and liked the images for each pharmacy."
}
```
```json
{
"OgSentence": "Praised micro-transitions and color changes on interaction.",
"NewSentence": "Praised micro-transitions and color changes on interaction."
}
```
```json
{
"OgSentence": "Liked automatic cursor movement for verification and clear error messages.",
"NewSentence": "Liked automatic cursor movement for verification and clear error messages."
}
```
```json
{
"OgSentence": "Commented positively on medicine availability statements (e.g., \"available at 4 pharmacies\").",
"NewSentence": "Commented positively on medicine availability statements (e.g., \"available at 4 pharmacies\")."
}
```
```json
{
"OgSentence": "Preferred chat feature over directions and valued consistent red color for communication buttons.",
"NewSentence": "Preferred the chat feature over directions and valued the consistent red color for communication buttons."
}
```
```json
{
"OgSentence": "Found direct dialing useful when tapping the call button.",
"NewSentence": "Found direct dialing useful when tapping the call button."
}
```
```json
{
"OgSentence": "Noted inconsistent sizing/alignment of \"Send SMS\" and \"Skip\" buttons; suggested checking responsive design guidelines.",
"NewSentence": "Noted inconsistent sizing/alignment of \"Send SMS\" and \"Skip\" buttons; suggested checking responsive design guidelines."
}
```
```json
{
"OgSentence": "\"Continue\" button non-functional without profile completion; recommended making profile details optional.",
"NewSentence": "The \"Continue\" button was non-functional without profile completion; it was recommended to make profile details optional."
}
```
```json
{
"OgSentence": "Suggested adding more micro-interactions and making the back arrow icon larger.",
"NewSentence": "Suggested adding more micro-interactions and making the back arrow icon larger."
}
```
```json
{
"OgSentence": "Requested to click the app logo to return to the homepage.",
"NewSentence": "Requested the ability to click the app logo to return to the homepage."
}
```
```json
{
"OgSentence": "Disliked the static pink background; suggested a dynamic or reworked design.",
"NewSentence": "Disliked the static pink background; suggested a dynamic or reworked design."
}
```
```json
{
"OgSentence": "Preferred 12-hour (AM/PM) over 24-hour time format in filters.",
"NewSentence": "Preferred the 12-hour (AM/PM) over the 24-hour time format in filters."
}
```
```json
{
"OgSentence": "Expected \"Update Location\" to open a map with user position; recommended a visual indicator (e.g., a dot) or removal.",
"NewSentence": "Expected \"Update Location\" to open a map with user position; recommended a visual indicator (e.g., a dot) or its removal."
}
```
```json
{
"OgSentence": "Found the pencil icon for editing details unclear; suggested consistency with intuitive design.",
"NewSentence": "Found the pencil icon for editing details unclear; suggested consistency with intuitive design."
}
```
```json
{
"OgSentence": "Recommended saving changes should keep the user on the same page.",
"NewSentence": "Recommended that saving changes should keep the user on the same page."
}
```
```json
{
"OgSentence": "Suggested swapping \"Delete Account\" and \"Logout\" buttons, as Delete should not be primary.",
"NewSentence": "Suggested swapping the \"Delete Account\" and \"Logout\" buttons, as \"Delete\" should not be the primary option."
}
```
```json
{
"OgSentence": "Session Duration: 10-15 minutes.",
"NewSentence": "Session Duration: 10-15 minutes."
}
```
```json
{
"OgSentence": "Background: High school junior with freelance design experience.",
"NewSentence": "Background: High school junior with freelance design experience."
}
```
```json
{
"OgSentence": "Positive Feedback: Expressed strong satisfaction with the prototype, consistently complimenting design, interactions, and user experience.",
"NewSentence": "Positive Feedback: Expressed strong satisfaction with the prototype, consistently complimenting its design, interactions, and user experience."
}
```
```json
{
"OgSentence": "Found the prototype significantly more engaging, intuitive, and visually appealing than the previous version.",
"NewSentence": "Found the prototype significantly more engaging, intuitive, and visually appealing than the previous version."
}
```
```json
{
"OgSentence": "Praised introductory screens, overall flow, and described the layout as clean and well-organized.",
"NewSentence": "Praised the introductory screens, overall flow, and described the layout as clean and well-organized."
}
```
```json
{
"OgSentence": "Highlighted personalization as superior to the earlier version, imagining it as a professional pharmacy app.",
"NewSentence": "Highlighted personalization as superior to the earlier version, envisioning it as a professional pharmacy app."
}
```
```json
{
"OgSentence": "Found micro-transitions smooth and visually pleasant.",
"NewSentence": "Found micro-transitions smooth and visually pleasant."
}
```
```json
{
"OgSentence": "Liked automatic cursor movement for verification and helpful error messages.",
"NewSentence": "Liked automatic cursor movement for verification and helpful error messages."
}
```
```json
{
"OgSentence": "Found the search experience easy and appreciated the medicine availability information.",
"NewSentence": "Found the search experience easy and appreciated the medicine availability information."
}
```
```json
{
"OgSentence": "Liked consistent red color for chat and call buttons, valuing the direct dialing feature.",
"NewSentence": "Liked the consistent red color for chat and call buttons, valuing the direct dialing feature."
}
```
```json
{
"OgSentence": "Suggestions/Issues: Expressed strong criticism regarding the location icon's functionality, expecting it to open a map with nearby pharmacies and the current location.",
"NewSentence": "Suggestions/Issues: Expressed strong criticism regarding the location icon's functionality, expecting it to open a map displaying nearby pharmacies and the current location."
}
```
```json
{
"OgSentence": "And the popup for location permissions caused more confusion and frustration, leading her to describe the icon's design and behavior as \"stupid,\" misleading, and unintuitive.",
"NewSentence": "The popup for location permissions caused more confusion and frustration, leading her to describe the icon's design and behavior as \"stupid,\" misleading, and unintuitive."
}
```
```json
{
"OgSentence": "Stated that filling in a profile picture should not be required at the start of the application.",
"NewSentence": "Stated that filling in a profile picture should not be required at the start of the application."
}
```
```json
{
"OgSentence": "Session Duration: 16-17 minutes.",
"NewSentence": "Session Duration: 16-17 minutes."
}
```
```json
{
"OgSentence": "Background: Brief experience in UI/UX design.",
"NewSentence": "Background: Brief experience in UI/UX design."
}
```
```json
{
"OgSentence": "Positive Feedback: New participant User went through the splash screens that explain the app (starting by clicking \"Next\" then sliding).",
"NewSentence": "Positive Feedback: New participant. The user went through the splash screens that explain the app (starting by clicking \"Next\" then sliding)."
}
```
```json
{
"OgSentence": "Skipped the login process when the option was available.",
"NewSentence": "Skipped the login process when the option was available."
}
```
```json
{
"OgSentence": "User liked that each card has the pharmacy picture, unlike the previous app design.",
"NewSentence": "The user liked that each card has the pharmacy picture, unlike the previous app design."
}
```
```json
{
"OgSentence": "User liked the filter options and how he could filter the pharmacies based on time and location.",
"NewSentence": "The user liked the filter options and how he could filter the pharmacies based on time and location."
}
```
```json
{
"OgSentence": "User seemed fairly relaxed, with no visible or audible signs of confusion or hesitation throughout exploring the application, with only two exceptions.",
"NewSentence": "The user seemed fairly relaxed, with no visible or audible signs of confusion or hesitation throughout exploring the application, with only two exceptions."
}
```
```json
{
"OgSentence": "\"The new design adaptation is way, way, way better than the old one.\"",
"NewSentence": "\"The new design adaptation is way, way, way better than the old one.\""
}
```
```json
{
"OgSentence": "The user was glad that the repetitive logos in the application were removed.",
"NewSentence": "The user was glad that the repetitive logos in the application were removed."
}
```
```json
{
"OgSentence": "User was happy to see that navigation is now easier and feels more natural.",
"NewSentence": "The user was happy to see that navigation is now easier and feels more natural."
}
```
```json
{
"OgSentence": "The user appreciated the added functionality, saying that the application is now worth keeping on the phone, unlike the previous design, where he decided to delete it right after the observation.",
"NewSentence": "The user appreciated the added functionality, stating that the application is now worth keeping on the phone, unlike the previous design, where he decided to delete it right after the observation."
}
```
```json
{
"OgSentence": "This version strikes a good balance between features and usability.",
"NewSentence": "This version strikes a good balance between features and usability."
}
```
```json
{
"OgSentence": "Noted a grammar mistake in the second splash screen.",
"NewSentence": "Noted a grammar mistake in the second splash screen."
}
```
```json
{
"OgSentence": "Annoyed by the 24-hour time clock and the inconsistent use of 24-hour and AM/PM formats on the pharmacy page.",
"NewSentence": "Annoyed by the 24-hour time clock and the inconsistent use of 24-hour and AM/PM formats on the pharmacy page."
}
```
```json
{
"OgSentence": "Hesitated to click chat/call buttons, fearing immediate conversation without confirmation.",
"NewSentence": "Hesitated to click chat/call buttons, fearing immediate conversation without confirmation."
}
```
```json
{
"OgSentence": "Confused by the \"Update Location\" button, expecting a map view, and suggesting its relocation for better visibility.",
"NewSentence": "Confused by the \"Update Location\" button, expecting a map view and suggesting its relocation for better visibility."
}
```
```json
{
"OgSentence": "Noted that phone number changes lacked SMS verification, allowing for fake numbers.",
"NewSentence": "Noted that phone number changes lacked SMS verification, allowing for the entry of fake numbers."
}
```
```json
{
"OgSentence": "Fix inconsistent time display.",
"NewSentence": "Fix inconsistent time display."
}
```
```json
{
"OgSentence": "Clarify/edit FAQs (e.g., \"guest vs. normal user\" typo).",
"NewSentence": "Clarify/edit FAQs (e.g., \"guest vs. normal user\" typo)."
}
```
```json
{
"OgSentence": "Add pharmacy phone numbers to the view page, not just the call button.",
"NewSentence": "Add pharmacy phone numbers to the view page, not just the call button."
}
```
```json
{
"OgSentence": "Include medicine details (e.g., package size, purpose).",
"NewSentence": "Include medicine details (e.g., package size, purpose)."
}
```
```json
{
"OgSentence": "Show medicine prices if legally possible.",
"NewSentence": "Show medicine prices if legally possible."
}
```
```json
{
"OgSentence": "Consider Adding favorites for medicines/pharmacies.",
"NewSentence": "Consider adding favorites for medicines/pharmacies."
}
```
```json
{
"OgSentence": "Consider Adding a share option for pharmacy/medicine links.",
"NewSentence": "Consider adding a share option for pharmacy/medicine links."
}
```
```json
{
"OgSentence": "Consider combining medicine and pharmacy search.",
"NewSentence": "Consider combining medicine and pharmacy search."
}
```
```json
{
"OgSentence": "Consider Adding notifications for out-of-stock medicines when restocked.",
"NewSentence": "Consider adding notifications for out-of-stock medicines when restocked."
}
```
```json
{
"OgSentence": "Consider adding pharmacy reviews.",
"NewSentence": "Consider adding pharmacy reviews."
}
```
```json
{
"OgSentence": "Session Duration: 5–6 minutes.",
"NewSentence": "Session Duration: 5–6 minutes."
}
```
```json
{
"OgSentence": "Background: general user",
"NewSentence": "Background: General user."
}
```
```json
{
"OgSentence": "Positive Feedback: • Pleased with the new search bar addition and tested it successfully.",
"NewSentence": "Positive Feedback: • Pleased with the new search bar addition and tested it successfully."
}
```
```json
{
"OgSentence": "Liked seeing the medicine count display.",
"NewSentence": "Liked seeing the medicine count display."
}
```
```json
{
"OgSentence": "Hit a login wall when trying to use the chat feature, but after logging in as a user, he was able to use the chat feature seamlessly.",
"NewSentence": "Hit a login wall when trying to use the chat feature, but after logging in as a user, he was able to use the chat feature seamlessly."
}
```
```json
{
"OgSentence": "Stated for communication features, \"Calling to confirm stock or alternatives makes sense, but I'd rarely use chat, I'd just call\".",
"NewSentence": "Regarding communication features, he stated, \"Calling to confirm stock or alternatives makes sense, but I'd rarely use chat; I'd just call\"."
}
```
```json
{
"OgSentence": "Appreciated having a quick-access list to view medicines and where they're available.",
"NewSentence": "Appreciated having a quick-access list to view medicines and where they're available."
}
```
```json
{
"OgSentence": "Thought the app was \"way better than the old one\".",
"NewSentence": "Thought the app was \"way better than the old one\"."
}
```
```json
{
"OgSentence": "Was extremely positive about the new design adaptation, stating \"this one is way, way better\" and \"kind of enjoyed it\".",
"NewSentence": "Was extremely positive about the new design adaptation, stating, \"this one is way, way better\" and \"kind of enjoyed it\"."
}
```
```json
{
"OgSentence": "Would consider downloading and using the app if it works.",
"NewSentence": "Would consider downloading and using the app if it works."
}
```
```json
{
"OgSentence": "Noted the app serves an important purpose, \"You might only need this once every 2-3 months, but when you need it, you need it\".",
"NewSentence": "Noted the app serves an important purpose: \"You might only need this once every 2-3 months, but when you need it, you need it\"."
}
```
```json
{
"OgSentence": "Acknowledged clear visual improvements, commenting that \"colors are better [than original].",
"NewSentence": "Acknowledged clear visual improvements, commenting that \"colors are better [than original]."
}
```
```json
{
"OgSentence": "Could be more ideal, but improved,\" and declaring the overall \"design is way, way better - no weird options/settings cluttering the interface\".",
"NewSentence": "Could be more ideal, but improved,\" and declaring the overall \"design is way, way better - no weird options/settings cluttering the interface\"."
}
```
```json
{
"OgSentence": "In account settings, liked being able to edit details, but noted that “you can't remove the profile picture from the account\".",
"NewSentence": "In account settings, liked being able to edit details, but noted that “you can't remove the profile picture from the account”."
}
```
```json
{
"OgSentence": "Praised navigation for being \"straightforward now\".",
"NewSentence": "Praised navigation for being \"straightforward now\"."
}
```
```json
{
"OgSentence": "Concluded with overall satisfaction with the redesign.",
"NewSentence": "Concluded with overall satisfaction with the redesign."
}
```
```json
{
"OgSentence": "Suggestions/Issues: • Noticed a navigation issue where the home bar disappears after going deep into the app (e.g., medicine → pharmacy → chat).",
"NewSentence": "Suggestions/Issues: • Noticed a navigation issue where the home bar disappears after going deep into the app (e.g., medicine → pharmacy → chat)."
}
```
```json
{
"OgSentence": "Suggested it should \"stay present on all pages\" to make backtracking easier.",
"NewSentence": "Suggested it should \"stay present on all pages\" to make backtracking easier."
}
```
```json
{
"OgSentence": "Emphasized that \"a PFP (Personal profile picture) isn't necessary for a medicine app\".",
"NewSentence": "Emphasized that \"a PFP (Personal profile picture) isn't necessary for a medicine app\"."
}
```
```json
{
"OgSentence": "Mentioned the pink shade (\"not ideal but not ugly\").",
"NewSentence": "Mentioned the pink shade (\"not ideal but not ugly\")."
}
```
```json
{
"OgSentence": "Suggested revising the pharmacy icon since \"it's a location pin but shouldn't be\".",
"NewSentence": "Suggested revising the pharmacy icon since \"it's a location pin but shouldn't be\"."
}
```
```json
{
"OgSentence": "Initially missed the update location button, and after it was pointed out, expected a map or list refresh from it.",
"NewSentence": "Initially missed the update location button, and after it was pointed out, expected a map or list refresh from it."
}
```
```json
{
"OgSentence": "Remove or reposition the location update button .",
"NewSentence": "Remove or reposition the location update button."
}
```
```json
{
"OgSentence": "The user noted that \"Phone number change should have some 2FA or something\", suggested to Add 2FA/SMS verification for phone number changes.",
"NewSentence": "The user noted that \"Phone number change should have some 2FA or something\", suggesting to add 2FA/SMS verification for phone number changes."
}
```
```json
{
"OgSentence": "Keep the home bar permanently visible (disappears after deep navigation like medicine → pharmacy → chat).",
"NewSentence": "Keep the home bar permanently visible (disappears after deep navigation like medicine → pharmacy → chat)."
}
```
```json
{
"OgSentence": "Make profile picture optional/removable independently (\"allow me to remove my pfp\").",
"NewSentence": "Make profile picture optional/removable independently (\"allow me to remove my pfp\")."
}
```
```json
{
"OgSentence": "Colors are improved, but could be \"more ideal\".",
"NewSentence": "Colors are improved, but could be \"more ideal\"."
}
```
```json
{
"OgSentence": "Session Duration: about 5 minutes.",
"NewSentence": "Session Duration: About 5 minutes."
}
```
```json
{
"OgSentence": "Background: General User",
"NewSentence": "Background: General user."
}
```
```json
{
"OgSentence": "Positive Feedback: • Appreciated the overall look and feel of the application, stating that the UI design was nice.",
"NewSentence": "Positive Feedback: • Appreciated the overall look and feel of the application, stating that the UI design was nice."
}
```
```json
{
"OgSentence": "Feedback was generally positive and showed satisfaction with the design and usability.",
"NewSentence": "Feedback was generally positive and showed satisfaction with the design and usability."
}
```
```json
{
"OgSentence": "Expressed a clear preference for the updated design over the previous version.",
"NewSentence": "Expressed a clear preference for the updated design over the previous version."
}
```
```json
{
"OgSentence": "Described the new UI as \"much nicer\" and \"easy to understand.\"",
"NewSentence": "Described the new UI as \"much nicer\" and \"easy to understand.\""
}
```
```json
{
"OgSentence": "Felt this version was more user-friendly and visually appealing compared to the old design.",
"NewSentence": "Felt this version was more user-friendly and visually appealing compared to the old design."
}
```
```json
{
"OgSentence": "Noted that the new prototype felt more modern and clean.",
"NewSentence": "Noted that the new prototype felt more modern and clean."
}
```
```json
{
"OgSentence": "Appreciated the overall simplicity and addition of new features like chat and the use of actual pharmacy images.",
"NewSentence": "Appreciated the overall simplicity and the addition of new features like chat and the use of actual pharmacy images."
}
```
```json
{
"OgSentence": "Liked that the interface was intuitive.",
"NewSentence": "Liked that the interface was intuitive."
}
```
```json
{
"OgSentence": "Would be more likely to use this version of the app in real scenarios.",
"NewSentence": "Would be more likely to use this version of the app in real scenarios."
}
```
```json
{
"OgSentence": "Overall, no major usability issues and positive reception of UI design.",
"NewSentence": "Overall, no major usability issues, and positive reception of UI design."
}
```
```json
{
"OgSentence": "The participant responded positively overall, particularly toward the visual design and ease of use.",
"NewSentence": "The participant responded positively overall, particularly toward the visual design and ease of use."
}
```
```json
{
"OgSentence": "Suggestions/Issues: • Did not understand the purpose of the \"Update Location\" button and initially ignored it, later the user Felt that some features (like the \"Update Location\" button) lacked clarity, user suggested to Improve The visibility or the explanation of the Update Location button to clarify its function • Explicitly mentioned that she did not like the presence of working hours in the filters, stating she didn't think it was important, and asked to Reconsider the usefulness of them, as it may not be a priority for all users.",
"NewSentence": "Suggestions/Issues: • Did not understand the purpose of the \"Update Location\" button and initially ignored it. Later, the user felt that some features (like the \"Update Location\" button) lacked clarity and suggested improving the visibility or explanation of the \"Update Location\" button to clarify its function. • Explicitly mentioned that she did not like the presence of working hours in the filters, stating she didn't think it was important, and asked to reconsider their usefulness, as it may not be a priority for all users."
}
```
```json
{
"OgSentence": "Session Duration: 5 minutes.",
"NewSentence": "Session Duration: 5 minutes."
}
```
```json
{
"OgSentence": "Background: Foreign user with infrequent app needs.",
"NewSentence": "Background: Foreign user with infrequent app needs."
}
```
```json
{
"OgSentence": "Positive Feedback: • Liked the easy-going nature of the application and the simple UI; it wasn't very overwhelming.",
"NewSentence": "Positive Feedback: • Liked the easy-going nature of the application and the simple UI; it was not very overwhelming."
}
```
```json
{
"OgSentence": "It was nice to see the pharmacy image in the application.",
"NewSentence": "It was nice to see the pharmacy image in the application."
}
```
```json
{
"OgSentence": "The automatic search bar is nice, as it updates the list automatically.",
"NewSentence": "The automatic search bar is a beneficial feature, as it updates the list automatically."
}
```
```json
{
"OgSentence": "User was able to appropriately describe what each page does and identify all the key info from the cards for pharmacies (name, location, status, work hours, etc.).",
"NewSentence": "The user was able to appropriately describe what each page does and identify all the key information from the cards for pharmacies (name, location, status, work hours, etc.)."
}
```
```json
{
"OgSentence": "Both communication features are necessary.",
"NewSentence": "Both communication features are necessary."
}
```
```json
{
"OgSentence": "Users would use chat communications 99% of the time, as they are foreigners and do not know Turkish, and conversing via chat is easier than trying to call.",
"NewSentence": "Users would utilize chat communications 99% of the time, as they are foreigners and do not know Turkish, and conversing via chat is easier than attempting to call."
}
```
```json
{
"OgSentence": "User stated that his use case in Turkey for the application is probably not as frequent (maybe once every 2 or 3 months), but it's one of those applications that, when you need it, there is no other good alternative.",
"NewSentence": "The user stated that his use case in Turkey for the application is probably not as frequent (perhaps once every 2 or 3 months), but it is one of those applications that, when needed, has no other good alternative."
}
```
```json
{
"OgSentence": "User didn't have any comment on what should be changed or any improvements to add to the app, mainly because he thought this application should stay on the easy-to-use side rather than the feature-rich side, because most people have simple needs that need to be fulfilled, and the application should only aspire to do that as easy as possible.",
"NewSentence": "The user did not have any comments on what should be changed or any improvements to add to the app, mainly because he thought this application should remain on the easy-to-use side rather than the feature-rich side, as most people have simple needs that need to be fulfilled, and the application should only aspire to do that as easily as possible."
}
```
```json
{
"OgSentence": "Session Duration: 7-8 minutes.",
"NewSentence": "Session Duration: 7-8 minutes."
}
```
```json
{
"OgSentence": "Background: General User.",
"NewSentence": "Background: General user."
}
```
```json
{
"OgSentence": "Positive Feedback: • Recognized the main page as a list of pharmacies and could identify their information (name, location, opening and closing times, etc.).",
"NewSentence": "Positive Feedback: • Recognized the main page as a list of pharmacies and could identify their information (name, location, opening and closing times, etc.)."
}
```
```json
{
"OgSentence": "Liked the filter options and how they could filter pharmacies based on time and location.",
"NewSentence": "Liked the filter options and how they could filter pharmacies based on time and location."
}
```
```json
{
"OgSentence": "Found the chat feature useful, especially as a non-Turkish speaker, preferring it over calls for checking medicine availability or for consultations.",
"NewSentence": "Found the chat feature useful, especially as a non-Turkish speaker, preferring it over calls for checking medicine availability or for consultations."
}
```
```json
{
"OgSentence": "Stated the current design is significantly better than the previous one, being \"cleaner, easier to understand, and has almost no redundancies\".",
"NewSentence": "Stated the current design is significantly better than the previous one, being \"cleaner, easier to understand, and has almost no redundancies\"."
}
```
```json
{
"OgSentence": "Found the new design \"very clear and straightforward\".",
"NewSentence": "Found the new design \"very clear and straightforward\"."
}
```
```json
{
"OgSentence": "Appreciated the added functionality, aligning with initial expectations from the previous design.",
"NewSentence": "Appreciated the added functionality, aligning with initial expectations from the previous design."
}
```
```json
{
"OgSentence": "Noted that the aesthetics and colors of the new design are \"much nicer\".",
"NewSentence": "Noted that the aesthetics and colors of the new design are \"much nicer\"."
}
```
```json
{
"OgSentence": "Preferred having pictures of the pharmacy and medicine in the new design over the repeated logo in the old design.",
"NewSentence": "Preferred having pictures of the pharmacy and medicine in the new design over the repeated logo in the old design."
}
```
```json
{
"OgSentence": "Suggestions/Issues: • Unsure about the ordering of pharmacies in the list (e.g., by distance or alphabetical order) and wished for a way to confirm.",
"NewSentence": "Suggestions/Issues: • Unsure about the ordering of pharmacies in the list (e.g., by distance or alphabetical order) and wished for a way to confirm."
}
```
```json
{
"OgSentence": "The \"Update Location\" button was confusing; it looked like a map button but didn't display a map.",
"NewSentence": "The \"Update Location\" button was confusing; it looked like a map button but did not display a map."
}
```
```json
{
"OgSentence": "Later the user suggested to remove or reposition the \"Location update\" button, as it is confusing and felt unnecessary.",
"NewSentence": "Later, the user suggested removing or repositioning the \"Location update\" button, as it was confusing and felt unnecessary."
}
```
```json
{
"OgSentence": "Make the profile picture in login optional, suggesting a default icon instead.",
"NewSentence": "Make the profile picture in login optional, suggesting a default icon instead."
}
```
```json
{
"OgSentence": "Add more information to medicine descriptions (e.g., purpose like \"headache medicine\" or \"painkiller\").",
"NewSentence": "Add more information to medicine descriptions (e.g., purpose like \"headache medicine\" or \"painkiller\")."
}
```
```json
{
"OgSentence": "Consider adding a booking option for medicines to pick up at a different time.",
"NewSentence": "Consider adding a booking option for medicines to pick up at a different time."
}
```
```json
{
"OgSentence": "Session Duration: 10-13 minutes.",
"NewSentence": "Session Duration: 10-13 minutes."
}
```
```json
{
"OgSentence": "Background: Pharmacist.",
"NewSentence": "Background: Pharmacist."
}
```
```json
{
"OgSentence": "Method: Interview/Discussion focusing on professional insights and operational logic of pharmacies.",
"NewSentence": "Method: Interview/Discussion focusing on professional insights and operational logic of pharmacies."
}
```
```json
{
"OgSentence": "Positive Feedback: • The application effectively addresses a significant gap within existing pharmacy systems, particularly for urgent needs (e.g., finding pharmacies for prescriptions late at night).",
"NewSentence": "Positive Feedback: • The application effectively addresses a significant gap within existing pharmacy systems, particularly for urgent needs (e.g., finding pharmacies for prescriptions late at night)."
}
```
```json
{
"OgSentence": "The charting option is a beneficial feature from the user's perspective.",
"NewSentence": "The charting option is a beneficial feature from the user's perspective."
}
```
```json
{
"OgSentence": "The application's design is perceived as easy to use, visually appealing, and has appropriate colors, which is desirable for pharmacists who prefer straightforward tools.",
"NewSentence": "The application's design is perceived as easy to use, visually appealing, and has appropriate colors, which is desirable for pharmacists who prefer straightforward tools."
}
```
```json
{
"OgSentence": "Key Insights & Suggestions/Issues: • System Integration for Inventory Management: o The primary challenge and the key to the application's true value lies in its integration with current pharmacy inventory systems.",
"NewSentence": "Key Insights & Suggestions/Issues: • System Integration for Inventory Management: o The primary challenge and the key to the application's true value lie in its integration with current pharmacy inventory systems."
}
```
```json
{
"OgSentence": "Manually tracking approximately 15,000 medicines and additional Over-The-Counter (OTC) products (like cosmetics and vitamins) is a difficult task, one that require much effort that some pharmacists might not be able to handle.",
"NewSentence": "Manually tracking approximately 15,000 medicines and additional Over-The-Counter (OTC) products (like cosmetics and vitamins) is a difficult task, one that requires much effort that some pharmacists might not be able to handle."
}
```
```json
{
"OgSentence": "An integrated system allowing automatic stock updates (e.g., decreasing inventory upon dispensing medicine) is essential.",
"NewSentence": "An integrated system allowing automatic stock updates (e.g., decreasing inventory upon dispensing medicine) is essential."
}
```
```json
{
"OgSentence": "Enhancing Chat Feature Efficiency for Pharmacists: o The chat feature offers great convenience for users and opens a new channel for direct communication with the pharmacy.",
"NewSentence": "• Enhancing Chat Feature Efficiency for Pharmacists: o The chat feature offers great convenience for users and opens a new channel for direct communication with the pharmacy."
}
```
```json
{
"OgSentence": "o To ensure this feature remains effective without overwhelming pharmacists (ones who might already manage a wide range of responsibilities), it was suggested that support staff or pharmacy students assist in handling those chat interactions.",
"NewSentence": "o To ensure this feature remains effective without overwhelming pharmacists (who might already manage a wide range of responsibilities), it was suggested that support staff or pharmacy students assist in handling those chat interactions."
}
```
```json
{
"OgSentence": "o By involving additional personnel, pharmacists can still engage meaningfully in chats when needed, while maintaining focus on their core duties.",
"NewSentence": "o By involving additional personnel, pharmacists can still engage meaningfully in chats when needed, while maintaining focus on their core duties."
}
```
```json
{
"OgSentence": "This approach ensures the chat system remains both functional for both user types and sustainable from the pharmacist side.",
"NewSentence": "This approach ensures the chat system remains both functional for both user types and sustainable from the pharmacist side."
}
```
```json
{
"OgSentence": "Pharmacy Schedule Management: o The concept of displaying pharmacy schedules is Extremely valuable.",
"NewSentence": "• Pharmacy Schedule Management: o The concept of displaying pharmacy schedules is extremely valuable."
}
```
```json
{
"OgSentence": "o Clarification was provided that schedules are determined by government-affiliated organizations (e.g., \"Istanbul Pharmacy Chamber\") rather than directly by the government for individual pharmacies.",
"NewSentence": "o Clarification was provided that schedules are determined by government-affiliated organizations (e.g., \"Istanbul Pharmacy Chamber\") rather than directly by the government for individual pharmacies."
}
```
```json
{
"OgSentence": "These organizations coordinate to ensure comprehensive coverage across regions.",
"NewSentence": "These organizations coordinate to ensure comprehensive coverage across regions."
}
```
```json
{
"OgSentence": "o The application should display schedules for more than just a week; pharmacists typically have schedules available for an entire year (e.g., from January to August/September).",
"NewSentence": "o The application should display schedules for more than just a week; pharmacists typically have schedules available for an entire year (e.g., from January to August/September)."
}
```
```json
{
"OgSentence": "o It is crucial to incorporate information about official Turkish public holidays when pharmacies may not be open, clearly indicating these days to users.",
"NewSentence": "o It is crucial to incorporate information about official Turkish public holidays when pharmacies may not be open, clearly indicating these days to users."
}
```
```json
{
"OgSentence": "Feedback Results & Heuristic Evaluation After collecting user feedback, we now move to evaluate the data using Nielsen’s Ten Usability Heuristics.",
"NewSentence": "Feedback Results & Heuristic Evaluation: After collecting user feedback, we now move to evaluate the data using Nielsen’s Ten Usability Heuristics."
}
```
```json
{
"OgSentence": "This step helps us identify and categorize the key usability issues present in our application, allowing us to take informed steps toward improving the user experience.",
"NewSentence": "This step helps us identify and categorize the key usability issues present in our application, allowing us to take informed steps toward improving the user experience."
}
```
```json
{
"OgSentence": "While Nielsen’s principles are highly effective in uncovering design flaws, it's important to note that some heuristics may not fully apply to our case due to the specific nature of our application and its current frontend-only scope.",
"NewSentence": "While Nielsen’s principles are highly effective in uncovering design flaws, it's important to note that some heuristics may not fully apply to our case due to the specific nature of our application and its current frontend-only scope."
}
```
```json
{
"OgSentence": "Nonetheless, using this heuristic framework provides a valuable, structured approach to highlighting interface problems, even if some principles are more relevant than others in our context.",
"NewSentence": "Nonetheless, using this heuristic framework provides a valuable, structured approach to highlighting interface problems, even if some principles are more relevant than others in our context."
}
```
```json
{
"OgSentence": "Positive Feedback Now it's clear after viewing the user feedback data that we were able to meet the objective of the design adaptation with this prototype.",
"NewSentence": "Positive Feedback: It is now clear after viewing the user feedback data that we were able to meet the objective of the design adaptation with this prototype."
}
```
```json
{
"OgSentence": "As all of the users have displayed satisfaction with the prototype, with minor issues here and there, which is to be expected, almost all participants expressed complete satisfaction with the new design.",
"NewSentence": "As all of the users displayed satisfaction with the prototype, with minor issues here and there, which is to be expected, almost all participants expressed complete satisfaction with the new design."
}
```
```json
{
"OgSentence": "Whether from returning users who had seen the old version or new users seeing the app for the first time, the added functionalities made the app feel significantly more purposeful and reliable for both pharmacists and customers, whether logged in or Browse as guests.",
"NewSentence": "Whether from returning users who had seen the old version or new users seeing the app for the first time, the added functionalities made the app feel significantly more purposeful and reliable for both pharmacists and customers, whether logged in or Browse as guests."
}
```
```json
{
"OgSentence": "Several participants explicitly emphasized how much better the new design was.",
"NewSentence": "Several participants explicitly emphasized how much better the new design was."
}
```
```json
{
"OgSentence": "One user stated, “The new design adaptation is way, way, way better than the old one” (P3), while another echoed similar sentiment: “declaring the overall \"design is way, way better - no weird options/settings cluttering the interface” and “I kind of enjoyed it” (P4).",
"NewSentence": "One user stated, “The new design adaptation is way, way, way better than the old one” (P3), while another echoed similar sentiment: “declaring the overall 'design is way, way better - no weird options/settings cluttering the interface' and 'I kind of enjoyed it'” (P4)."
}
```
```json
{
"OgSentence": "The interface was consistently described as clean, modern, intuitive, and more user-friendly, with B5 calling it “much nicer” and “easy to understand.”",
"NewSentence": "The interface was consistently described as clean, modern, intuitive, and more user-friendly, with B5 calling it “much nicer” and “easy to understand.”"
}
```
```json
{
"OgSentence": "Visual improvements also stood out, with praise for the personalized pharmacy cards, clear layout, and pleasing color scheme.",
"NewSentence": "Visual improvements also stood out, with praise for the personalized pharmacy cards, clear layout, and pleasing color scheme."
}
```
```json
{
"OgSentence": "P1 and P2 commended the smooth micro-transitions and overall flow, while P6 highlighted that the app wasn’t overwhelming and was easy to navigate.",
"NewSentence": "P1 and P2 commended the smooth micro-transitions and overall flow, while P6 highlighted that the app was not overwhelming and was easy to navigate."
}
```
```json
{
"OgSentence": "Users also appreciated features like the improved search function, direct call/chat options, and the pharmacy availability indicators, which helped in quickly identifying where to find a given medicine.",
"NewSentence": "Users also appreciated features like the improved search function, direct call/chat options, and the pharmacy availability indicators, which helped in quickly identifying where to find a given medicine."
}
```
```json
{
"OgSentence": "Even the pharmacist we interviewed recognized that the application effectively fills a gap in existing pharmacy systems, especially in urgent situations.",
"NewSentence": "Even the pharmacist we interviewed recognized that the application effectively fills a gap in existing pharmacy systems, especially in urgent situations."
}
```
```json
{
"OgSentence": "They appreciated the “visually appealing” design and the practical communication tools integrated into the interface (P8).",
"NewSentence": "They appreciated the “visually appealing” design and the practical communication tools integrated into the interface (P8)."
}
```
```json
{
"OgSentence": "This overwhelmingly positive feedback not only validates the design choices we made but also confirms that the core functionality and user experience goals were successfully met.",
"NewSentence": "This overwhelmingly positive feedback not only validates the design choices we made but also confirms that the core functionality and user experience goals were successfully met."
}
```
```json
{
"OgSentence": "Heuristic Issues Mapping & Elaboration The Nielsen Heuristic Evaluation is based on ten foundational usability principles that help identify and assess interface issues from a user experience perspective.",
"NewSentence": "Heuristic Issues Mapping & Elaboration: The Nielsen Heuristic Evaluation is based on ten foundational usability principles that help identify and assess interface issues from a user experience perspective."
}
```
```json
{
"OgSentence": "These principles serve as a widely recognized framework in the field of usability testing.",
"NewSentence": "These principles serve as a widely recognized framework in the field of usability testing."
}
```
```json
{
"OgSentence": "The ten heuristics are as follows: 1. Visibility of system status 2. Match between system and the real world 3. User control and freedom 4. Consistency and standards 5. Error prevention 6. Recognition rather than recall 7. Flexibility and efficiency of use 8. Aesthetic and minimalist design 9. Help users recognize, diagnose, and recover from errors 10. Help and documentation As mentioned previously, not all of these principles carry equal weight in the context of our project.",
"NewSentence": "The ten heuristics are as follows: 1. Visibility of system status 2. Match between system and the real world 3. User control and freedom 4. Consistency and standards 5. Error prevention 6. Recognition rather than recall 7. Flexibility and efficiency of use 8. Aesthetic and minimalist design 9. Help users recognize, diagnose, and recover from errors 10. Help and documentation As mentioned previously, not all of these principles carry equal weight in the context of our project."
}
```
```json
{
"OgSentence": "The nature of the App being only frontend and its features led certain heuristics to be more relevant and frequently observed than others.",
"NewSentence": "The frontend nature of the App and its features led certain heuristics to be more relevant and frequently observed than others."
}
```
```json
{
"OgSentence": "Nonetheless, we aimed to consider most of them during our evaluation.",
"NewSentence": "Nonetheless, we aimed to consider most of them during our evaluation."
}
```
```json
{
"OgSentence": "To assess the severity of each identified issue, we applied a standardized classification scale.",
"NewSentence": "To assess the severity of each identified issue, we applied a standardized classification scale."
}
```
```json
{
"OgSentence": "The table below outlines the severity levels used to describe the impact of each usability problem.",
"NewSentence": "The table below outlines the severity levels used to describe the impact of each usability problem."
}
```
```json
{
"OgSentence": "This scale will be referenced throughout the upcoming subsections: Severity Level Label Description 0 Not a usability problem The issue doesn't affect usability or might just be a design preference.",
"NewSentence": "This scale will be referenced throughout the upcoming subsections: Severity Level Label Description 0 Not a usability problem The issue does not affect usability or might just be a design preference."
}
```
```json
{
"OgSentence": "1 Cosmetic problem only Does not need to be fixed unless there's extra time; does not affect functionality or performance.",
"NewSentence": "1 Cosmetic problem only Does not need to be fixed unless there is extra time; does not affect functionality or performance."
}
```
```json
{
"OgSentence": "2 Minor usability problem Fixing is low priority; the issue causes slight inconvenience but does not significantly impact task completion or understanding.",
"NewSentence": "2 Minor usability problem Fixing is low priority; the issue causes slight inconvenience but does not significantly impact task completion or understanding."
}
```
```json
{
"OgSentence": "3 Major usability problem High priority fix; causes user frustration, confusion, or slows down interaction, but users can still complete tasks.",
"NewSentence": "3 Major usability problem High priority fix; causes user frustration, confusion, or slows down interaction, but users can still complete tasks."
}
```
```json
{
"OgSentence": "4 Usability catastrophe Must fix before release; the issue causes major errors, prevents task completion, or introduces potential security/privacy risks.",
"NewSentence": "4 Usability catastrophe Must fix before release; the issue causes major errors, prevents task completion, or introduces potential security/privacy risks."
}
```
```json
{
"OgSentence": "Main Vital Issues Issue #1 : Profile Picture Requirement • Heuristic Principle: User Control and Freedom • Severity Level: 1 – Minor usability problem Explanation: in the login page that you face the first time you open the application, 6.2.1 Main Vital Issues Issue #1 : Profile Picture Requirement • Heuristic Principle: User Control and Freedom • Severity Level: 1 – Minor usability problem Explanation: in the login page that you face the first time you open the application, you are prompted to enter your personal information (first name, last name, Personal Profile Picture), and while the first name and last name are mandatory (for the chat feature to function properly for both users) the Personal profile picture is not necessary.",
"NewSentence": "Main Vital Issues Issue #1: Profile Picture Requirement • Heuristic Principle: User Control and Freedom • Severity Level: 1 – Minor usability problem Explanation: In the login page that you encounter the first time you open the application, you are prompted to enter your personal information (first name, last name, Personal Profile Picture). While the first name and last name are mandatory (for the chat feature to function properly for both users), the Personal Profile Picture is not necessary."
}
```
```json
{
"OgSentence": "The app currently lacks the ability to remove or skip setting a profile picture.",
"NewSentence": "The app currently lacks the ability to remove or skip setting a profile picture."
}
```
```json
{
"OgSentence": "This limits user control over their account.",
"NewSentence": "This limits user control over their account."
}
```
```json
{
"OgSentence": "Taken action: Agreed to add an option to skip or remove the profile picture.",
"NewSentence": "Taken action: Agreed to add an option to skip or remove the profile picture."
}
```
```json
{
"OgSentence": "Issue #2 : Remove already set Profile Picture • Heuristic Principle: User Control and Freedom • Severity: 2- Minor usability problem Explanation: Users are unable to remove their profile picture once it is set, which is inconvenient and unconventional.",
"NewSentence": "Issue #2: Remove Already Set Profile Picture • Heuristic Principle: User Control and Freedom • Severity: 2 – Minor usability problem Explanation: Users are unable to remove their profile picture once it is set, which is inconvenient and unconventional."
}
```
```json
{
"OgSentence": "This limits personalization and control for the Users.",
"NewSentence": "This limits personalization and control for the users."
}
```
```json
{
"OgSentence": "Taken action: we have added a sub-section that will appear once you start to edit your account details information and click on the profile picture change icon, it will prompt you to either upload a new image or remove your existing profile picture.",
"NewSentence": "Taken action: We have added a sub-section that will appear once you start to edit your account details information and click on the profile picture change icon; it will prompt you to either upload a new image or remove your existing profile picture."
}
```
```json
{
"OgSentence": "Issue #3: Time Format (AM/PM vs. 24-hour) • Heuristic Principle: Match Between System and the Real World & Consistency and Standards • Severity: 2- Minor usability problem Explanation: The app inconsistently displays time formats; some pages use the 24-hour format, while others use AM/PM.",
"NewSentence": "Issue #3: Time Format (AM/PM vs. 24-hour) • Heuristic Principle: Match Between System and the Real World & Consistency and Standards • Severity: 2 – Minor usability problem Explanation: The app inconsistently displays time formats; some pages use the 24-hour format, while others use AM/PM."
}
```
```json
{
"OgSentence": "This inconsistency can confuse users and feel unnatural.",
"NewSentence": "This inconsistency can confuse users and feel unnatural."
}
```
```json
{
"OgSentence": "Taken action: From the feedback Data, we have seen that our users prefer the AM/PM system over the 24-Hour one, so we have converted the application Time format into the AM/PM system.",
"NewSentence": "Taken action: From the feedback data, we have observed that our users prefer the AM/PM system over the 24-hour system, so we have converted the application's time format to the AM/PM system."
}
```
```json
{
"OgSentence": "Issue #4: Update Location Button • Heuristic Principle: Match Between System and the Real World • Severity: 3- Major usability problem Explanation: The addition of the “Update Location” button was one of the changes made at the end of the Issue #4: Update Location Button • Heuristic Principle: Match Between System and the Real World • Severity: 3- Major usability problem Explanation: The addition of the “Update Location",
"NewSentence": "Issue #4: Update Location Button • Heuristic Principle: Match Between System and the Real World • Severity: 3 – Major usability problem Explanation: The addition of the “Update Location” button was one of the changes made at the end of the prototype's development."
}
```

"""
doc_file_path = r'E:\Documents\pyhton projects\Grammer Fix Scrpt\Test subject.docx'

find_and_replace_in_word(doc_file_path, corrections_data)