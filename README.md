# Resume Converter Repository

This repository contains a Python application that tailors Markdown-based resumes to specific job listings and converts them into formatted Word documents. The process preserves the original resume structure and styling while tailoring the content to match job requirements. Right now you need to use the template provided and add details to it. You can't just use any old resume. 

There are currently two dummy resumes to choose from; edit them to your liking but keep the current layout. Modify the current layout and add your skills and proficiencies and work experience keeping the same format and style. Its designed this way to allow it to filter through without being flagged by recruitment programs or turning into a jumbled mess in their systems and ultimately just being ignored. 

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
   
   Note: convert_resume.py is not intended to be run directly; but can be if you just want to convert a markdown resume you already have. 

## How to Use

1. **Prepare Your Resume**  
   Write or edit your resume in Markdown format. Ensure the formatting follows standard Markdown conventions. You can preview your Markdown at https://markdownlivepreview.com/. The header information that is inserted into your document is in settings/header.txt. Here you can also edit the prompts sent to the AI model. 

2. **Setup Environment**
   *Step 1*
   Create a new file in the root of your directory, in the same folder as tailor_and_convert.py exists called .env 
   In this file called .env add the following line: (Save the file once you are done)
   ``` bash
   OPENAI_API_KEY=sk-proj-YOUROPENAIKEYGOESHEREWITHOUTANY""MARKS
   ```
   You will need some credits in your account for this to work. Head to https://platform.openai.com/ to create an account and keys if your do not have them already. 

   *Step 2*
   Install Python 3 - https://www.python.org/downloads/
   Open up a cli console and navigate to the folder where the tailor_and_convert.py file exists. Then run:
   ``` bash
   pip install requirements.txt
   ```
   *(Should you wish to your python environments that is your choice. And more details on how to set this up will be created later.)*

2. **Tailor and Convert the Resume**  
   Open a terminal and execute the following command to tailor your resume and convert it to a DOCX file:

      *Easy Run* - This will default to a resume that doesn't modify the resume too much and places it in exported_resumes folder using the o1-mini model.
   
   python tailor_and_convert.py \
     --job_url "https://www.linkedin.com/jobs/view/JOB_ID" \
     --resume "path/to/your_resume.md" 

   *Detailed Run*

   python tailor_and_convert.py \
     --job_url "https://www.linkedin.com/jobs/view/JOB_ID" \
     --resume "path/to/your_resume.md" \
     --output_md "tailored_resume.md" \
     --output_docx "tailored_resume.docx" \
     --model "o1-mini" \
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

   **Convert Resume**
   Should you wish to run the conversion without tailoring a resume use the following commands. 

   *Basic command which will export the file called resume.docx to the exported_resumes folder*
   ``` bash
   python convert_resume.py resume.md resume.docx
   ```

   *Extended command where you have control over where it saves the file*
   ``` bash
   python convert_resume.py resume.md resume.docx --path /home/user/my_resumes
   ```


| **Argument** | **Required** | **Description** |
|--------------|--------------|-----------------|
| `input`      | Yes          | Input Markdown file. |
| `output`     | Yes          | Output DOCX file name (without a folder path if using `--path`). |
| `--path`     | No           | Output folder path where the DOCX file will be saved. If not provided, the file is saved in the `exported_resumes` folder. |


3. **Review the Output**  
   - The tailored Markdown resume is saved in the markdown_resumes folder.
   - The converted DOCX file is saved in the exported_resumes folder (or in a custom path if provided).

## Requirements

- Python 3.x
- python-docx: Install with pip install python-docx
- Other dependencies listed in requirements.txt
- Built on a Windows Machine - Not tested on macOS yet

## License

This project is provided as-is, without any warranties. Contributions and improvements are welcome.