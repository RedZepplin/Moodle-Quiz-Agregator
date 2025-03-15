import os
import email
from email import policy
from bs4 import BeautifulSoup
import base64
import pdfkit
import subprocess
import time

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
            body = soup.find('body')
            if body:
                header_div = body.find('div', class_='wrapper-header')
                if header_div:
                    header_content = str(header_div)
                
                # Remove the header from the html content
                if header_div:
                    header_div.extract()
                html_content = str(body)
            else:
                html_content = current_html_content
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

def deduplicate_and_replace_with_correct(questions):
    """Refine deduplication by replacing 'Partially Correct' and 'Incorrect' answers with 'Correct'."""
    question_map = {}  # Map to store unique questions with the best (Correct) answer
    
    # Loop through all the questions
    for question in questions:
        question_text = question.find('div', class_='qtext').get_text(strip=True)  # Get the question text
        
        # Get the answer state
        state_div = question.find('div', class_='state mx-2')
        state = state_div.get_text(strip=True) if state_div else ""
        
        if question_text in question_map:
            # If we already have a question, only replace it with 'Correct' answers
            existing_answer_state = question_map[question_text]['state']
            
            if state == 'Correct':
                question_map[question_text] = {'question': question, 'state': 'Correct'}  # Replace with 'Correct'
            elif state == 'Partially Correct' and existing_answer_state != 'Correct':
                # If the existing answer isn't 'Correct', replace 'Partially Correct' with 'Correct'
                question_map[question_text] = {'question': question, 'state': 'Partially Correct'}
            # If the state is 'Incorrect', do nothing, since we only replace with 'Correct'
        else:
            # If the question is not in the map, add it
            question_map[question_text] = {'question': question, 'state': state}
    
    # Return the list of questions with the best (Correct) answers prioritized
    return [value['question'] for value in question_map.values()]

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
    
    consolidated_html += '</head><body><section>'
    
    seen_divs = set()  # Set to track unique divs
    question_number = 1  # Initialize the question number
    
    all_images = {}  # Dictionary to store all images from the first MHTML file
    
    questions = []  # List to collect questions from all files
    
    # Merge all images from the first file
    if mhtml_files:
        first_mhtml_file = mhtml_files[0]
        _, images, _ = extract_html_from_mhtml(first_mhtml_file)
        all_images.update(images)
    
    for mhtml_file in mhtml_files:
        print(f'Processing {mhtml_file}...')
        html_content, images, _ = extract_html_from_mhtml(mhtml_file)
        all_images.update(images)  # Merge images from this file
        
        if html_content:
            soup = BeautifulSoup(html_content, 'html.parser')
            # Extract all the divs
            for div in soup.find_all('div', class_='que'):
                questions.append(div)
    
    # Deduplicate and prioritize 'Correct' answers, replacing others
    questions = deduplicate_and_replace_with_correct(questions)
    
    # Now, process the questions
    for question in questions:
        div_str = str(question)  # Convert div to string to check for uniqueness
        if div_str not in seen_divs:
            seen_divs.add(div_str)  # Add div to the set of seen divs
            
            # Find the span with class 'rui-qno' and update the number
            span = question.find('span', class_='rui-qno')
            if span:
                span.string = str(question_number)  # Replace with the consecutive number
                question_number += 1  # Increment the question number
            
            # Replace img src with base64 if it's a local image (from MHTML)
            for img in question.find_all('img'):
                img_src = img.get('src', '')
                if img_src in all_images:
                    image_base64, mime_type = all_images[img_src]
                    img['src'] = f"data:image/{mime_type};base64,{image_base64}"
            
            # Append the div with its updated question number and images to the HTML content
            consolidated_html += str(question)
    
    consolidated_html += '</section></body></html>'

    # Save the consolidated HTML to a file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(consolidated_html)
    print(f'Consolidated document saved as {output_file}')

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
    
    # Add a page break after each question in the HTML content
    html_content = html_content.replace('<div class="que">', '<div class="que" style="page-break-before:always;">')
    
    # Save the modified HTML as a temporary file and convert to PDF
    try:
        pdfkit.from_string(html_content, output_pdf, options=options, configuration=config)
        print(f'PDF saved as {output_pdf}')
    except OSError as e:
        print(f"Error during PDF conversion: {e}")
        print("Retrying in 5 seconds...")
        time.sleep(5)
        try:
            pdfkit.from_string(html_content, output_pdf, options=options, configuration=config)
            print(f'PDF saved as {output_pdf} after retry')
        except OSError as e:
            print(f"Error during PDF conversion after retry: {e}")
            print("Please check your internet connection and try again.")

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
        subprocess.run(["wkhtmltopdf", "--version"], capture_output=True, check=True)
        return "wkhtmltopdf"  # If it's in PATH, just return the command name
    except (FileNotFoundError, subprocess.CalledProcessError):
        pass

    return None

if __name__ == '__main__':
    # Directory containing the .mhtml files
    mhtml_folder = 'Files'  # Changed to relative path

    # List all the mhtml files in the folder
    mhtml_files = [os.path.join(mhtml_folder, f) for f in os.listdir(mhtml_folder) if f.endswith('.mhtml')]

    # Output file for the consolidated document
    output_file = 'Consolidated_Assessment.html'
    output_pdf = 'Consolidated_Assessment.pdf'

    # Consolidate the files
    consolidate_mhtml_files(mhtml_files, output_file)
    # Convert HTML to PDF with one question per page
    #
    # convert_html_to_pdf(output_file, output_pdf)
