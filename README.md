# Resume Converter Repository

This repository contains a Python application that tailors Markdown-based resumes to specific job listings and converts them into formatted Word documents. The process preserves the original resume structure and styling while tailoring the content to match job requirements. Right now you need to use the template provided and add details to it. You can't just use any old resume. 

## How It Works

1. **Tailoring the Resume**  
   The main script, tailor_and_convert.py, takes a job listing URL and a Markdown resume. It fetches the job details, applies tailoring rules (without fabricating new skills or altering employment dates), and generates a tailored Markdown resume.  
   - Moderate Mode: Applies minimal changes to keep the original resume largely intact.  
   - Unmoderated Mode: Allows for creative modifications while preserving key date and proficiency details.

2. **Converting to DOCX**  
   After tailoring, tailor_and_convert.py calls convert_resume.py internally to convert the tailored Markdown resume into a DOCX file. The conversion script handles:
   - Markdown heading conversion to corresponding Word styles.
   - Inline formatting (bold, italics, hyperlinks) using the Aptos font.
   - Custom table formatting for certifications and other sections.
   
   Note: convert_resume.py is not intended to be run directly.

## How to Use

1. **Prepare Your Resume**  
   Write or edit your resume in Markdown format. Ensure the formatting follows standard Markdown conventions. You can preview your Markdown at https://markdownlivepreview.com/.

2. **Tailor and Convert the Resume**  
   Open a terminal and execute the following command to tailor your resume and convert it to a DOCX file:

   python tailor_and_convert.py \
     --job_url "https://www.linkedin.com/jobs/view/JOB_ID" \
     --resume "path/to/your_resume.md" \
     --output_md "tailored_resume.md" \
     --output_docx "tailored_resume.docx" \
     --model "gpt-3.5-turbo" \
     --moderate "true" \
     --verbose
   
   **Explanation of Arguments:**
   - `--job_url`: The URL of the job listing you want to tailor your resume for.
   - `--resume`: Path to your original Markdown resume.
   - `--output_md`: Path where the tailored Markdown resume will be saved.
   - `--output_docx`: Path to save the final DOCX file.
   - `--model`: Specifies the AI model (e.g., "gpt-3.5-turbo") used for processing the resume content.
   - `--moderate`: A flag ("true" or "false") indicating whether minimal changes (moderate) should be applied.
   - `--verbose`: Enables detailed output for troubleshooting and confirmation of the process.

   - Replace "https://www.linkedin.com/jobs/view/JOB_ID" with the actual job listing URL.
   - Replace "path/to/your_resume.md" with the path to your Markdown resume.

3. **Review the Output**  
   - The tailored Markdown resume is saved in the markdown_resumes folder.
   - The converted DOCX file is saved in the exported_resumes folder (or in a custom path if provided).

## Requirements

- Python 3.x
- python-docx: Install with pip install python-docx
- Other dependencies listed in requirements.txt

## License

This project is provided as-is, without any warranties. Contributions and improvements are welcome.