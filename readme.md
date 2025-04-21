# Moodle Quiz Aggregator

## Overview

This Python script aggregates multiple Moodle quiz attempt review pages, saved as MHTML files, into a single, consolidated HTML document. It intelligently deduplicates questions based on their text content and correctness (determined by the "Mark X out of Y" grade), ensuring only the _best_ attempt for each unique question is included in the final output.

The script also embeds images directly into the HTML using Base64 encoding and optionally generates a PDF version of the consolidated document with one question per page. Output filenames and the main header title are automatically generated based on the quiz title found in the header of the first MHTML file processed, but can be overridden with a command-line argument.

## Features

- Processes multiple Moodle quiz review `.mhtml` files from a specified folder.
- **Optionally searches recursively** through subfolders for `.mhtml` files (`-r` flag).
- Extracts individual questions (`<div class="que">`).
- Identifies question correctness based on the `<div class="grade">` content (e.g., "Mark 1.00 out of 1.00").
- Deduplicates questions based on the question text (`<div class="qtext">`).
- Prioritizes and keeps the version of a question with the highest correctness state (Correct > Partially Correct > Incorrect).
- Embeds images referenced within the MHTML files directly into the output HTML using Base64.
- Renumbers questions sequentially in the final document.
- Generates a single, self-contained HTML output file.
- Optionally generates a PDF output file with page breaks before each question (`-p` flag, requires `wkhtmltopdf`).
- Automatically names output files and sets the main header based on the quiz title extracted from the first MHTML file's header (e.g., `Your_Quiz_Title.html`).
- **Allows specifying a custom base name** for output files and the main header title (`-n` flag), overriding automatic extraction.
- Accepts command-line arguments to specify the input folder.

## Prerequisites

1.  **Python 3.x:** Ensure you have Python 3 installed. You can download it from python.org.
2.  **pip:** Python's package installer (usually included with Python 3).
3.  **wkhtmltopdf:** (Required _only_ for PDF generation) This is an external command-line tool.
    - Download it from wkhtmltopdf.org.
    - **Important:** Install it and ensure the `wkhtmltopdf` executable is added to your system's PATH environment variable, or the script might not find it.

## Setup & Installation

1.  **Clone the repository:**

    ```bash
    git clone <your-repository-url>
    cd Moodle-Quiz-Agregator
    ```

2.  **Install Python dependencies:**

    ```bash
    pip install beautifulsoup4 pdfkit
    # Optional but recommended parser for BeautifulSoup:
    # pip install lxml
    ```

3.  **Install `wkhtmltopdf`:** Follow the instructions from the wkhtmltopdf website for your operating system. Remember to add it to your system's PATH if you plan to generate PDFs.

4.  **Create Input Folder (Optional):** By default, the script looks for MHTML files in a subfolder named `Files`. Create it if it doesn't exist:
    ```bash
    mkdir Files
    ```
    Alternatively, you can specify a different folder path when running the script.

## Usage

1.  **Save MHTML Files:** Save your Moodle quiz review pages as MHTML files (`.mhtml` or `.mht`) into the input folder (e.g., the `Files/` directory or another folder you specify). If using the recursive option, you can organize them into subfolders within the main input folder. Make sure you save the _review_ page that shows the questions, answers, and grades.

2.  **Run the Script:** Open your terminal or command prompt, navigate to the project directory (`Moodle-Quiz-Agregator`), and run the script using one of the following methods:

    - **Basic Usage (HTML only, uses default `./Files` folder, non-recursive):**

      ```bash
      python moodle_quiz_agregator.py
      ```

      _(Or use the `run_aggregator.bat` file on Windows)_

    - **Specify Input Folder (HTML only, non-recursive):**

      ```bash
      python moodle_quiz_agregator.py "C:\path\to\your\mhtml_files"
      ```

      _(Or `run_aggregator.bat "C:\path\to\your\mhtml_files"` on Windows)_

    - **Generate PDF (uses default `./Files` folder, non-recursive):**

      ```bash
      python moodle_quiz_agregator.py -p
      ```

      _(Or `run_aggregator.bat -p` on Windows)_

    - **Specify Input Folder and Generate PDF (non-recursive):**

      ```bash
      python moodle_quiz_agregator.py "C:\path\to\your\mhtml_files" --pdf
      ```

      _(Or `run_aggregator.bat "C:\path\to\your\mhtml_files" --pdf` on Windows)_

    - **Recursive Search (HTML only, uses default `./Files` folder):**

      ```bash
      python moodle_quiz_agregator.py -r
      ```

      _(Or `run_aggregator.bat -r` on Windows)_

    - **Specify Custom Name (HTML only, uses default `./Files` folder, non-recursive):**

      ```bash
      python moodle_quiz_agregator.py -n "My Custom Quiz Name"
      ```

      _(Or `run_aggregator.bat -n "My Custom Quiz Name"` on Windows)_

    - **Combine Options (Specify folder, recursive, PDF, custom name):**
      ```bash
      python moodle_quiz_agregator.py "D:\Quizzes" -r -p -n "Final Exam Consolidated"
      ```
      _(Or `run_aggregator.bat "D:\Quizzes" -r -p -n "Final Exam Consolidated"` on Windows)_

3.  **Output:** The script will generate:

    - An HTML file (e.g., `Your_Quiz_Title.html` or `My_Custom_Quiz_Name.html`) in the project's root directory.
    - Optionally, a PDF file (e.g., `Your_Quiz_Title.pdf` or `My_Custom_Quiz_Name.pdf`) in the project's root directory if the `-p` or `--pdf` flag was used.

    _Note: Filenames and the main header title are based on the title found in the first processed MHTML file's header unless overridden using the `-n` or `--name` argument. Names are sanitized to be valid filenames._

## File Structure

Moodle-Quiz-Agregator/
├── moodle_quiz_agregator.py # The main Python script
├── run_aggregator.bat # Example batch file for running on Windows
├── Files/ # Default directory for input .mhtml files
│ ├── attempt1.mhtml
│ └── attempt2.mhtml
├── .gitignore # Specifies files to ignore for Git
└── README.md # This file

## Contributing

Feel free to submit issues or pull requests. Please note that generated output files (`*.html`, `*.pdf`) should generally not be committed to the repository, as specified in the `.gitignore` file.
