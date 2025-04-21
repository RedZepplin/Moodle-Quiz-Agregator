import os
import email
from email import policy
from bs4 import BeautifulSoup
import base64
import pdfkit
import subprocess
import time
import re # <-- Import the re module

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
    """Extracts HTML content and images from an MHTML file."""
    images = {}  # Dictionary to hold images as base64 data
    header_content = ""  # Variable to store the header content
    html_content = ""
    with open(mhtml_file, 'rb') as f:
        msg = email.message_from_binary_file(f, policy=policy.default)

    # Look for the HTML and image parts in the MHTML file
    for part in msg.iter_parts():
        content_type = part.get_content_type()
        if content_type == 'text/html':
            current_html_content = part.get_payload(decode=True).decode('utf-8', errors='ignore')
            # Extract header content (for title)
            soup = BeautifulSoup(current_html_content, 'html.parser')
            # Find the header content within the body
            body = soup.find('body', id='page-mod-quiz-review')
            if body:
                # Extract the header content
                print("Extracting Header")
                header_div = body.find('header', id ='page-header')
                
                if header_div:
                    print("Header Found")
                    header_content = str(header_div)

                    
                    
                # Remove the header from the html content
                if header_div:
                    print("Removing Header")
                    header_div.extract()
                html_content = str(body)

            else:
                print("Not Extracting")
                #html_content = current_html_content
                print(html_content)
        elif content_type.startswith('image'):
            image_data = part.get_payload(decode=True)
            content_disposition = part.get('Content-Disposition', '')
            # Get the filename of the image (usually used as the reference)
            filename = content_disposition.split('filename=')[-1].strip('"') if 'filename=' in content_disposition else 'image'
            # Convert the image to base64
            image_base64 = base64.b64encode(image_data).decode('utf-8')
            # Content-Location will be the URL we need to map from the HTML <img src=...>
            content_location = part.get('Content-Location', '')
            images[content_location] = (image_base64, content_type.split('/')[1])  # Store both base64 and MIME type
    return html_content, images, header_content

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
def consolidate_mhtml_files(mhtml_files, output_file):
    """Consolidates divs with class 'que' from multiple MHTML files into one document."""
    consolidated_html = '<html><head><title>Consolidated Document</title>'

    # Extract CSS and header from the first MHTML file
    if mhtml_files:
        first_mhtml_file = mhtml_files[0]
        css_content = extract_css_from_mhtml(first_mhtml_file)
        html_content, images, header_content = extract_html_from_mhtml(first_mhtml_file)

        if css_content:
            consolidated_html += f'<style>{css_content}</style>'

        # Add wrapper header content from the first MHTML file
        if header_content:
            consolidated_html += f'{header_content}'  # Ensure we insert header content here

    consolidated_html += '</head><body><section>' # Start the main content section

    question_number = 1  # Initialize the question number

    all_images = {}  # Dictionary to store all images (prioritize first file's images if conflicts)

    questions_to_process = []  # List to collect all 'que' divs from all files

    # Process all files to gather questions and images
    for mhtml_file in mhtml_files:
        print(f'Processing {mhtml_file}...')
        html_content, images, _ = extract_html_from_mhtml(mhtml_file)

        # Merge images, giving priority to existing ones (usually from the first file)
        for loc, data in images.items():
            if loc not in all_images:
                all_images[loc] = data
        if html_content:
            soup = BeautifulSoup(html_content, 'html.parser')
            # Extract all the 'que' divs
            for div in soup.find_all('div', class_='que'):
                
                questions_to_process.append(div)

    # Deduplicate and prioritize based on the grade using the modified function
    final_questions = deduplicate_and_replace_with_correct(questions_to_process)
    # Now, process the final list of unique, best-state questions
    for question in final_questions:
        # Find the span with class ending in 'qno' and update the number
        # Use re.compile to find span with class like 'qno', 'qnounderway', etc.
        qno_span = question.find('span', class_=re.compile(r'qno')) # More robust finding
        if qno_span:
            # Find the actual number element, often a sibling or child
            num_element = qno_span.find(string=re.compile(r'\d+')) # Find text node containing digits
            if num_element:
                 num_element.replace_with(str(question_number)) # Replace the number text
            else:
                 # Fallback if number is directly in the span
                 qno_span.string = str(question_number)
            question_number += 1  # Increment the question number
        else:
             # If no qno span found, maybe add one? Or just log it.
             print(f"Warning: Question number span ('qno') not found in a question div.")


        # Replace img src with base64 if it's a local image (from MHTML)
        for img in question.find_all('img'):
            img_src = img.get('src', '')
            if img_src in all_images:
                image_base64, mime_type = all_images[img_src]
                img['src'] = f"data:image/{mime_type};base64,{image_base64}"
            # Optional: Handle images not found in all_images (e.g., external URLs)
            # else:
            #    print(f"Warning: Image source '{img_src}' not found in extracted images.")

        # Append the processed question div to the HTML content
        consolidated_html += str(question)

    consolidated_html += '</section></body></html>' # Close the section and body

    # Save the consolidated HTML to a file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(consolidated_html)
    print(f'Consolidated document saved as {output_file}')


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
    # Directory containing the .mhtml files
    mhtml_folder = 'Files'  # Changed to relative path

    # List all the mhtml files in the folder
    mhtml_files = [os.path.join(mhtml_folder, f) for f in os.listdir(mhtml_folder) if f.endswith('.mhtml')]

    if not mhtml_files:
        print(f"No .mhtml files found in the '{mhtml_folder}' directory.")
    else:
        # Output file for the consolidated document
        output_file = 'Consolidated_Assessment.html'
        output_pdf = 'Consolidated_Assessment.pdf'

        # Consolidate the files
        consolidate_mhtml_files(mhtml_files, output_file)

        # Convert HTML to PDF with one question per page
        # try:
        #     convert_html_to_pdf(output_file, output_pdf)
        # except Exception as e: # Catch broader exceptions from PDF conversion
        #     print(f"Failed to convert HTML to PDF: {e}")

