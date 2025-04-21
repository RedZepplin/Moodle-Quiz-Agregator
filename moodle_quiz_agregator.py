import os
import email
from email import policy
from bs4 import BeautifulSoup
import base64
import pdfkit
import subprocess
import time
import re # <-- Import the re module
import argparse # <-- Import argparse
import sys

def sanitize_filename(name):
    """Removes or replaces characters invalid for filenames."""
    if not name:
        return "Untitled_Assessment"
    # Remove characters invalid in Windows/Linux/MacOS filenames
    # Including characters like '.', ':', etc., often found in Moodle titles
    name = re.sub(r'[<>:"/\\|?*.,;!]', '', name)
    # Replace consecutive whitespace with a single underscore
    name = re.sub(r'\s+', '_', name).strip('_')
    # Limit length (optional)
    max_len = 100
    if len(name) > max_len:
        # Try to truncate nicely at an underscore
        truncated_name = name[:max_len]
        if '_' in truncated_name:
            name = truncated_name.rsplit('_', 1)[0]
        else:
            name = truncated_name
    # Ensure filename is not empty after sanitization
    return name if name else "Untitled_Assessment"

# --- Helper function to determine state from grade ---
def determine_correctness_from_grade(grade_text):
    """
    Determines the correctness state ('Correct', 'Partially Correct', 'Incorrect')
    based on the grade text (e.g., "Mark 1.00 out of 1.00").
    """
    if not grade_text:
        return "Incorrect" # Default if grade is missing

    # Regex to extract mark and total points
    match = re.search(r"Mark\s*([\d\.]+)\s*out\s*of\s*([\d\.]+)", grade_text, re.IGNORECASE)

    if match:
        try:
            mark = float(match.group(1))
            total = float(match.group(2))

            if total <= 0: # Avoid division by zero or weird cases
                 return "Incorrect"

            if mark == total:
                return "Correct"
            elif mark > 0 and mark < total:
                return "Partially Correct"
            elif mark == 0:
                return "Incorrect"
            else: # Handle cases like mark > total? Treat as Correct for now, or adjust as needed.
                 # Or maybe treat as an error/unknown state? Let's stick to Correct if mark >= total.
                 if mark >= total:
                     return "Correct"
                 else: # Should not happen based on previous checks
                     return "Incorrect"

        except ValueError:
            # Handle cases where conversion to float fails
            print(f"Warning: Could not parse grade numbers in: '{grade_text}'")
            return "Incorrect" # Treat parsing errors as Incorrect
    else:
        # Handle cases where the regex doesn't match the expected format
        print(f"Warning: Could not parse grade format: '{grade_text}'")
        return "Incorrect" # Treat format errors as Incorrect

# --- Existing functions (extract_html_from_mhtml, extract_divs_from_html) remain the same ---
# ... (keep extract_html_from_mhtml and extract_divs_from_html as they are) ...
def extract_html_from_mhtml(mhtml_file):
    """
    Extracts HTML body content string, images, header content string, and document title
    from an MHTML file. Does NOT destructively modify the body content.
    """
    images = {}
    header_content_str = ""
    body_content_str = ""
    document_title = None # Initialize title

    try:
        with open(mhtml_file, 'rb') as f:
            msg = email.message_from_binary_file(f, policy=policy.default)
    except FileNotFoundError:
        print(f"Error: MHTML file not found: {mhtml_file}")
        return None, {}, "", None # Return None for body, empty dict, empty str, None title
    except Exception as e:
        print(f"Error reading MHTML file {mhtml_file}: {e}")
        return None, {}, "", None

    for part in msg.iter_parts():
        content_type = part.get_content_type()
        if content_type == 'text/html':
            try:
                # Decode the primary HTML content
                current_html_content = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                soup = BeautifulSoup(current_html_content, 'html.parser') # Parse once
            except Exception as e:
                print(f"Error parsing HTML from {mhtml_file}: {e}")
                continue # Skip this part if parsing fails

            # --- Find the body content for question extraction ---
            # Find body tag - use a more general approach first
            body = soup.find('body')
            # If specific ID is needed and reliable:
            # body = soup.find('body', id='page-mod-quiz-review')

            if body:
                # --- Get the body content string ---
                # IMPORTANT: Get string representation of the body as found, BEFORE any potential modifications
                body_content_str = str(body)

                # --- Extract header separately WITHOUT modifying the body object ---
                # Find header in the original soup (adjust selector if needed)
                header_div = soup.find('header', id='page-header')
                if header_div:
                    header_content_str = str(header_div) # Store header HTML string

                    # --- Extract Title from Header ---
                    # Look for h1, then h2, then h3 within the header_div
                    heading_tag = header_div.find(['h1', 'h2', 'h3','h4'])
                    if heading_tag:
                        # Try to get text more cleanly, removing potential inner tags like spans if needed
                        title_text = heading_tag.get_text(separator=' ', strip=True)
                        if title_text:
                            document_title = title_text
                            # print(f"  Found title: '{document_title}'") # Optional debug print
                    # else:
                        # print(f"  Warning: No h1, h2, or h3 found within header#page-header in {mhtml_file}")

            else:
                # Fallback if no body tag is found
                print(f"Warning: No <body> tag found in {mhtml_file}. Using full HTML for body content.")
                body_content_str = current_html_content
                # Still try to find header and title in the full soup
                header_div = soup.find('header', id='page-header')
                if header_div:
                    header_content_str = str(header_div)
                    heading_tag = header_div.find(['h1', 'h2', 'h3'])
                    if heading_tag:
                        title_text = heading_tag.get_text(separator=' ', strip=True)
                        if title_text:
                            document_title = title_text
                            # print(f"  Found title (no body tag): '{document_title}'") # Optional debug print

            # Only process the first HTML part found
            break # Assume only one main HTML part per MHTML

        elif content_type.startswith('image'):
            # Image handling
            try:
                image_data = part.get_payload(decode=True)
                content_disposition = part.get('Content-Disposition', '')
                # Get the filename of the image (usually used as the reference)
                filename = content_disposition.split('filename=')[-1].strip('"') if 'filename=' in content_disposition else 'image'
                # Content-Location will be the URL we need to map from the HTML <img src=...>
                content_location = part.get('Content-Location', '')
                if content_location: # Ensure content_location is not empty
                    image_base64 = base64.b64encode(image_data).decode('utf-8')
                    images[content_location] = (image_base64, content_type.split('/')[1])  # Store both base64 and MIME type
                # else:
                    # print(f"Warning: Image found without Content-Location in {mhtml_file}. Skipping.") # Optional debug
            except Exception as e:
                print(f"Error processing image in {mhtml_file}: {e}")

    # Return the original body content string, images, header string, and extracted title
    return body_content_str, images, header_content_str, document_title

def extract_divs_from_html(html_content):
    """Extract divs with class 'que' from the HTML content."""
    soup = BeautifulSoup(html_content, 'html.parser')
    return soup.find_all('div', class_='que')


# --- Modified deduplicate function ---
def deduplicate_and_replace_with_correct(questions):
    """
    Refine deduplication by using the 'grade' div to determine correctness
    and replacing less correct answers with more correct ones.
    Prioritization: Correct > Partially Correct > Incorrect.
    """
    question_map = {}  # Map to store unique questions with the best answer found so far

    # Loop through all the questions extracted from all files
    for question in questions:
        question_text_div = question.find('div', class_='qtext')
        if not question_text_div:
            print("Warning: Found a 'que' div without 'qtext'. Skipping.")
            continue # Skip if question text is missing

        question_text = question_text_div.get_text(strip=True)

        # Find the grade div and extract its text
        grade_div = question.find('div', class_='grade')
        grade_text = grade_div.get_text(strip=True) if grade_div else ""

        # Determine the state based on the grade text
        calculated_state = determine_correctness_from_grade(grade_text)

        # --- Deduplication Logic based on calculated_state ---
        if question_text in question_map:
            # A version of this question already exists in our map
            existing_entry = question_map[question_text]
            existing_state = existing_entry['state']

            # Define state priorities (higher number is better)
            state_priority = {"Correct": 3, "Partially Correct": 2, "Incorrect": 1}

            current_priority = state_priority.get(calculated_state, 0) # Default to 0 if state is unknown
            existing_priority = state_priority.get(existing_state, 0)

            # Replace if the current question's state is better than the existing one
            if current_priority > existing_priority:
                question_map[question_text] = {'question': question, 'state': calculated_state}
                # Optional: Log the replacement
                # print(f"Replacing '{existing_state}' with '{calculated_state}' for question: {question_text[:50]}...")

        else:
            # If the question is not in the map, add it
            question_map[question_text] = {'question': question, 'state': calculated_state}

    # Return the list of question divs corresponding to the best versions found
    return [value['question'] for value in question_map.values()]

# --- consolidate_mhtml_files function ---
# Needs a small adjustment to handle the re import if not already done globally
# Also, ensure the logic using seen_divs still makes sense after deduplication
def consolidate_mhtml_files(mhtml_files, output_html_file, first_file_header_str=""):
    """
    Consolidates divs with class 'que' from multiple MHTML files into one HTML document.
    Uses the provided header string from the first file.
    """
    # Use the title from the output filename for the HTML title tag
    html_title = os.path.splitext(os.path.basename(output_html_file))[0].replace('_', ' ')
    consolidated_html = f'<html><head><meta charset="UTF-8"><title>{html_title}</title>' # Use dynamic title

    # Extract CSS from the first MHTML file (still needed here)
    if mhtml_files:
        first_mhtml_file = mhtml_files[0]
        css_content = extract_css_from_mhtml(first_mhtml_file) # Assumes extract_css_from_mhtml exists and is correct
        if css_content:
            consolidated_html += f'<style>{css_content}</style>'

    # Add the header content string passed from the main block
    if first_file_header_str:
        consolidated_html += f'{first_file_header_str}'

    consolidated_html += '</head><body><section>' # Start the main content section

    question_number = 1
    all_images = {}
    questions_to_process = []

    # Process all files to gather questions and images
    for mhtml_file in mhtml_files:
        print(f'Processing {mhtml_file}...')
        # Get body content and images for this file
        # We ignore header_content and title returned here as they are handled elsewhere/already extracted
        body_content, images, _, _ = extract_html_from_mhtml(mhtml_file)

        if body_content is None: # Check if extraction failed
            print(f"Skipping file due to extraction error: {mhtml_file}")
            continue

        # Merge images, giving priority to existing ones
        for loc, data in images.items():
            if loc not in all_images:
                all_images[loc] = data

        # Parse the body content string to find questions
        try:
            soup = BeautifulSoup(body_content, 'html.parser')
            found_questions = soup.find_all('div', class_='que')
            if not found_questions:
                 print(f"Warning: No '<div class=\"que\">' elements found in the body of {mhtml_file}")
            for div in found_questions:
                questions_to_process.append(div)
        except Exception as e:
            print(f"Error parsing body content or finding questions in {mhtml_file}: {e}")
            continue

    print(f"Found {len(questions_to_process)} question divs in total.")

    # Deduplicate and prioritize
    final_questions = deduplicate_and_replace_with_correct(questions_to_process) # Assumes this function exists and is correct
    print(f"Processing {len(final_questions)} unique/best questions for output.")

    # Process the final list of questions
    for question in final_questions:
        # Renumber question
        qno_span = question.find('span', class_=re.compile(r'qno'))
        if qno_span:
            num_element = qno_span.find(string=re.compile(r'\d+'))
            if num_element:
                 num_element.replace_with(str(question_number))
            else:
                 # Fallback if number is directly in the span
                 qno_span.string = str(question_number)
            question_number += 1
        else:
             print(f"Warning: Question number span ('qno') not found in a question div.")

        # Embed images
        for img in question.find_all('img'):
            img_src = img.get('src', '')
            if img_src in all_images:
                image_base64, mime_type = all_images[img_src]
                img['src'] = f"data:image/{mime_type};base64,{image_base64}"

        consolidated_html += str(question)

    consolidated_html += '</section></body></html>'

    # Save the consolidated HTML
    try:
        with open(output_html_file, 'w', encoding='utf-8') as f:
            f.write(consolidated_html)
        print(f'Consolidated document saved as {output_html_file}')
    except Exception as e:
        print(f"Error writing consolidated HTML file {output_html_file}: {e}")


# --- Other functions (extract_css_from_mhtml, convert_html_to_pdf, find_wkhtmltopdf) remain the same ---
# ... (keep extract_css_from_mhtml, convert_html_to_pdf, find_wkhtmltopdf as they are) ...
def extract_css_from_mhtml(mhtml_file):
    """Extract CSS content (both internal and external) from the first MHTML file."""
    with open(mhtml_file, 'rb') as f:
        msg = email.message_from_binary_file(f, policy=policy.default)

    css_content = ""

    # Look for the CSS part in the MHTML file (internal CSS)
    for part in msg.iter_parts():
        if part.get_content_type() == 'text/css':
            css_content += part.get_payload(decode=True).decode('utf-8', errors='ignore')
        elif part.get_content_type() == 'text/html':
            html_content = part.get_payload(decode=True).decode('utf-8', errors='ignore')
            soup = BeautifulSoup(html_content, 'html.parser')
            for style_tag in soup.find_all('style'):
                css_content += style_tag.get_text()

    # Alternatively, if there are external stylesheets (like <link rel="stylesheet">), we would need to handle those
    return css_content

def convert_html_to_pdf(html_file, output_pdf):
    """Convert the consolidated HTML file to PDF with each question on a separate page."""
    # Check if wkhtmltopdf is installed and get its path
    wkhtmltopdf_path = find_wkhtmltopdf()
    if not wkhtmltopdf_path:
        raise OSError("wkhtmltopdf not found. Please install it and ensure it's in your system's PATH.")

    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

    options = {
        'no-outline': None,  # Disable outlines in the PDF
        'margin-top': '20mm',  # Set top margin
        'margin-right': '20mm',  # Set right margin
        'margin-bottom': '20mm',  # Set bottom margin
        'margin-left': '20mm',  # Set left margin
        'page-size': 'A4',  # Set page size to A4
        'footer-right': '[page]',  # Add page numbers to the footer
        'footer-font-size': '10',  # Set footer font size
        'disable-smart-shrinking': None,  # Avoid shrinking content to fit on the page
        'enable-local-file-access': None, # Enable local file access
    }

    # Make sure the questions are in individual divs, so each div becomes a page.
    # Ensure your HTML content has separate divs for each question
    with open(html_file, 'r', encoding='utf-8') as file:
        html_content = file.read()

    # Add a page break *before* each question div (except the first one)
    # We modify the BeautifulSoup object *before* writing to HTML for cleaner page breaks
    # This modification should ideally happen within consolidate_mhtml_files before writing the file,
    # or we re-parse the HTML here. Re-parsing is simpler for this example.

    soup = BeautifulSoup(html_content, 'html.parser')
    first_que = True
    for que_div in soup.find_all('div', class_='que'):
        if not first_que:
            # Add page break style
            if 'style' in que_div.attrs:
                que_div['style'] += '; page-break-before: always;'
            else:
                que_div['style'] = 'page-break-before: always;'
        first_que = False

    modified_html_content = str(soup)


    # Save the modified HTML as a temporary file and convert to PDF
    try:
        pdfkit.from_string(modified_html_content, output_pdf, options=options, configuration=config)
        print(f'PDF saved as {output_pdf}')
    except OSError as e:
        # Check if it's the specific "exit status 1" error often related to rendering issues
        if 'exit status 1' in str(e):
             print(f"Warning: wkhtmltopdf exited with status 1. This might indicate rendering issues (e.g., missing assets, complex JS/CSS). The PDF might be incomplete or malformed.")
             print(f"PDF saved as {output_pdf} (potentially with issues)") # Still save the potentially partial PDF
        else:
            print(f"Error during PDF conversion: {e}")
            # Optional: Add retry logic if needed, but often wkhtmltopdf errors are persistent
            # print("Retrying in 5 seconds...")
            # time.sleep(5)
            # try:
            #     pdfkit.from_string(modified_html_content, output_pdf, options=options, configuration=config)
            #     print(f'PDF saved as {output_pdf} after retry')
            # except OSError as e_retry:
            #     print(f"Error during PDF conversion after retry: {e_retry}")
            #     print("Please check wkhtmltopdf installation and HTML content.")


def find_wkhtmltopdf():
    """Finds the wkhtmltopdf executable in common locations or PATH."""
    # Common locations for wkhtmltopdf on Windows
    possible_paths = [
        r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe",
        r"C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe",
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return path

    # Check if wkhtmltopdf is in the system's PATH
    try:
        # Use --version which is a standard command and less likely to fail than converting empty input
        result = subprocess.run(["wkhtmltopdf", "--version"], capture_output=True, text=True, check=True, shell=True) # Added shell=True for Windows PATH resolution sometimes
        if "wkhtmltopdf" in result.stdout:
             return "wkhtmltopdf"  # If it's in PATH, just return the command name
    except (FileNotFoundError, subprocess.CalledProcessError) as e:
         print(f"wkhtmltopdf not found via subprocess check: {e}")
         pass # Continue searching

    # If using 'where' command on Windows
    if os.name == 'nt':
        try:
            result = subprocess.run(["where", "wkhtmltopdf"], capture_output=True, text=True, check=True, shell=True)
            # 'where' can return multiple paths, take the first one
            first_path = result.stdout.splitlines()[0].strip()
            if os.path.exists(first_path):
                return first_path
        except (FileNotFoundError, subprocess.CalledProcessError):
            pass

    # If using 'which' command on Linux/macOS
    elif os.name == 'posix':
         try:
            result = subprocess.run(["which", "wkhtmltopdf"], capture_output=True, text=True, check=True)
            first_path = result.stdout.strip()
            if os.path.exists(first_path):
                 return first_path
         except (FileNotFoundError, subprocess.CalledProcessError):
            pass


    print("Warning: wkhtmltopdf executable not found in common locations or system PATH.")
    return None


# --- Main execution block ---
if __name__ == '__main__':
    # --- Argument Parsing ---
    parser = argparse.ArgumentParser(description='Consolidate Moodle quiz attempts from MHTML files.')
    parser.add_argument(
        'mhtml_folder_arg',
        nargs='?',
        default='Files',
        help='Path to the folder containing MHTML files (defaults to "Files" subfolder)'
    )
    parser.add_argument(
        '-p', '--pdf',
        action='store_true',
        help='Also generate a PDF output (requires wkhtmltopdf)'
    )
    parser.add_argument(
        '-r', '--recursive',
        action='store_true',
        help='Search for MHTML files recursively in subfolders'
    )
    parser.add_argument(
        '-n', '--name',
        type=str,
        default=None, # Default is None, meaning we extract from header
        help='Specify a custom base name for the output files (overrides header extraction)'
    )

    args = parser.parse_args()

    # --- Validate Folder Path ---
    mhtml_folder = args.mhtml_folder_arg
    if not os.path.exists(mhtml_folder):
        print(f"Error: The specified folder does not exist: {mhtml_folder}")
        sys.exit(1)
    if not os.path.isdir(mhtml_folder):
        print(f"Error: The specified path is not a directory: {mhtml_folder}")
        sys.exit(1)

    print(f"Using MHTML folder: {mhtml_folder}")

    # --- List MHTML files (Conditional Recursive Search) ---
    mhtml_files = []
    try:
        if args.recursive:
            print("Searching recursively for MHTML files...")
            for root, dirs, files in os.walk(mhtml_folder):
                for filename in files:
                    if filename.lower().endswith('.mhtml'):
                        full_path = os.path.join(root, filename)
                        mhtml_files.append(full_path)
        else:
            print("Searching non-recursively for MHTML files...")
            for filename in os.listdir(mhtml_folder):
                 if filename.lower().endswith('.mhtml'):
                    full_path = os.path.join(mhtml_folder, filename)
                    if os.path.isfile(full_path):
                        mhtml_files.append(full_path)

        mhtml_files.sort()
        print(f"Found {len(mhtml_files)} MHTML file(s) to process.")

    except Exception as e:
        print(f"Error listing files in folder {mhtml_folder}: {e}")
        sys.exit(1)

    # --- Check if files were found and proceed ---
    if not mhtml_files:
        print(f"No .mhtml files found in '{mhtml_folder}'" + (" or its subfolders." if args.recursive else "."))
    else:
        # --- Extract Header String and potentially Title from the first file ---
        print(f"Extracting header structure from first file: {mhtml_files[0]}")
        _, _, first_header_str, extracted_title = extract_html_from_mhtml(mhtml_files[0]) # Assumes this function exists

        # --- Determine Base Filename (Custom or Extracted) ---
        if args.name:
            print(f"Using custom base name: '{args.name}'")
            base_filename = sanitize_filename(args.name) # Assumes this function exists
        else:
            print(f"Using extracted title for base name: '{extracted_title}'")
            base_filename = sanitize_filename(extracted_title) # Assumes this function exists

        # --- Determine Output Filenames ---
        output_file = f"{base_filename}.html"
        output_pdf = f"{base_filename}.pdf"
        print(f"Output HTML filename set to: {output_file}")
        if args.pdf:
            print(f"Output PDF filename set to: {output_pdf}")

        # --- Modify Header String with Determined Title ---
        modified_header_str = first_header_str # Start with the original
        if first_header_str and base_filename: # Only modify if we have a header and a name
            try:
                header_soup = BeautifulSoup(first_header_str, 'html.parser')
                # Find the first h1, h2, h3, or h4 tag within the header
                heading_tag = header_soup.find(['h1', 'h2', 'h3', 'h4'])
                if heading_tag:
                    # Create a display-friendly title from the base filename
                    display_title = base_filename.replace('_', ' ')
                    print(f"Updating header tag '{heading_tag.name}' to: '{display_title}'")
                    # Replace the content of the heading tag
                    heading_tag.string = display_title
                    # Get the modified header string
                    modified_header_str = str(header_soup)
                else:
                    print("Warning: Could not find a heading tag (h1-h4) in the extracted header to update.")
            except Exception as e:
                print(f"Warning: Error occurred while modifying header string: {e}")
                # Fallback to using the original header string
                modified_header_str = first_header_str
        # --- End of Header Modification ---


        # --- Consolidate the files ---
        # Pass the list of files, the dynamic output HTML name, and the MODIFIED header string
        consolidate_mhtml_files(mhtml_files, output_file, modified_header_str) # Assumes this function exists

        # --- Conditional PDF Conversion ---
        if args.pdf:
            print("\nAttempting PDF conversion...")
            try:
                convert_html_to_pdf(output_file, output_pdf) # Assumes this function exists
            except Exception as e:
                print(f"Failed to convert HTML to PDF: {e}")
        else:
            print("\nSkipping PDF generation (use -p or --pdf option to enable).")

