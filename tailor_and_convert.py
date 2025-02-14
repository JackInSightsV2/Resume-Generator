#!/usr/bin/env python3

"""
tailor_and_convert.py

This script:
  1. Takes a URL for a job listing and a Markdown resume file.
  2. Fetches the job details from the URL.
  3. Uses an OpenAI model (default "o1-mini") to tailor the resume to the job details,
     ensuring that it only uses content already present in the resume.
     - Do NOT add or invent any new skills, project names, or experience details.
     - Remove any skills, experiences, or sections that do not match the job description.
     - Dynamically filter out irrelevant cloud experience (e.g. if the job is AWS‑focused, remove or de‑emphasize Azure/GCP).
     - Expand on and clarify details that are relevant to the job requirements.
     - Process any '@Notes:' in the resume: update content (e.g. recalculate durations) and remove the note.
     - **Moderate mode:** 
         * If true, favor the baseline resume—make only minimal changes to align with the role while preserving original wording, job dates (including "Current" and duration details), and proficiency levels.
         * If false, allow creative modifications (except that employment dates and durations must remain exactly as in the baseline).
  4. Saves the tailored resume to a Markdown file with a user‑specified output name.
  5. Calls convert_resume.py to convert the tailored Markdown resume into a DOCX file with a user‑specified output name.
  6. **Verbose mode:** If set, print full job details and the prompt sent to OpenAI; if not set, only minimal status messages are shown.

  python tailor_and_convert.py \
  --job_url "https://www.linkedin.com/jobs/view/4149702217" \
  --resume "markdown_resumes/resume.dummy.md" \
  --output_md "tailored_resume.md" \
  --output_docx "tailored_resume.docx" \
  --model "gpt-3.5-turbo" \
  --moderate "false" \
  --verbose \
  --path "exported_resumes"

"""

import argparse
import os
import sys
import shutil
import subprocess
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
import datetime
import uuid

# Load environment variables from .env file.
load_dotenv()

# -------------------------------
# Helper functions
# -------------------------------

def fetch_job_details(url):
    """
    Fetch the job listing webpage at `url` and extract its text content.
    Returns a cleaned-up string with the job details.
    """
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/90.0.4430.93 Safari/537.36"
        )
    }
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        raise Exception(f"Failed to fetch URL {url}. Status code: {response.status_code}")
    
    html = response.text
    soup = BeautifulSoup(html, "html.parser")
    # Remove scripts and styles.
    for element in soup(["script", "style"]):
        element.decompose()
    # Get text and remove extra whitespace.
    text = soup.get_text(separator="\n")
    lines = [line.strip() for line in text.splitlines()]
    clean_text = "\n".join(line for line in lines if line)
    return clean_text

def tailor_resume(resume_md, job_details, moderate, model="o1-mini", verbose=False):
    """
    Use OpenAI's API to tailor the Markdown resume (`resume_md`)
    so that it better fits the job described in `job_details`.

    The tailoring instructions are read from external text files in the settings folder:
      - The header from 'settings/header.txt'
      - If moderate == True, instructions are read from 'settings/moderate.txt'
      - If moderate == False, instructions are read from 'settings/unmoderated.txt'

    These instructions, along with the current date, job listing, and original resume,
    form the prompt sent to the OpenAI model.

    If verbose is True, print the full prompt sent to OpenAI.
    """
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    
    # Read header from settings/header.txt
    header_path = os.path.join("settings", "header.txt")
    try:
        with open(header_path, "r", encoding="utf-8") as f:
            header = f.read()
    except Exception as e:
        print(f"Error reading header file at {header_path}:", e)
        sys.exit(1)
    
    # Read instructions from the appropriate prompt file.
    prompt_filename = "moderate.txt" if moderate else "unmoderated.txt"
    prompt_path = os.path.join("settings", prompt_filename)
    try:
        with open(prompt_path, "r", encoding="utf-8") as f:
            instructions = f.read()
    except Exception as e:
        print(f"Error reading prompt file at {prompt_path}:", e)
        sys.exit(1)
    
    prompt = (
        f"{header}\n\n"
        f"Current Date: {current_date}\n\n"
        f"{instructions}\n\n"
        "Job Listing:\n"
        "------------------\n"
        f"{job_details}\n\n"
        "Original Resume:\n"
        "------------------\n"
        f"{resume_md}\n\n"
        "Tailored Resume (Markdown):\n"
    )
    
    if verbose:
        print("\n--- FULL PROMPT TO OPENAI ---")
        print(prompt)
        print("-----------------------------\n")
    
    # For o1-mini, omit the system message (unsupported).
    if model.lower() == "o1-mini":
        messages = [{"role": "user", "content": prompt}]
    else:
        messages = [
            {"role": "system", "content": "You are a helpful resume tailoring assistant."},
            {"role": "user", "content": prompt},
        ]
    
    # Set temperature based on model compatibility.
    temp_value = 1 if model.lower() == "o1-mini" else 0.7
    
    try:
        from openai import OpenAI  # Import the new client class
        client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        response = client.chat.completions.create(
            model=model,
            messages=messages,
            temperature=temp_value,
        )
        tailored_resume = response.choices[0].message.content.strip()
        return tailored_resume
    except Exception as e:
        print("Error calling OpenAI API:", e)
        sys.exit(1)

# -------------------------------
# Main function
# -------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Tailor a Markdown resume to a job listing and convert it to DOCX."
    )
    parser.add_argument(
        "--job_url", required=True,
        help="URL of the job listing"
    )
    parser.add_argument(
        "--resume", required=True,
        help="Path to the original Markdown resume file"
    )
    parser.add_argument(
        "--output_md", default=None,
        help="Optional output Markdown file name for the tailored resume. If not provided, a random name is generated and the file is saved in the markdown_resumes folder."
    )
    parser.add_argument(
        "--output_docx", default=None,
        help="Optional output DOCX file name for the converted resume. If not provided, the file will be saved in the exported_resumes folder with a random name (prefix 'docx_resume_')."
    )
    parser.add_argument(
        "--model", default="o1-mini",
        help="OpenAI model to use for tailoring (e.g. o1-mini, gpt-3.5-turbo)"
    )
    parser.add_argument(
        "--moderate", default="true", choices=["true", "false"],
        help="Set moderate mode: true = favor baseline with minimal changes; false = allow creative modifications (but preserve dates/durations)"
    )
    parser.add_argument(
        "--verbose", action="store_true",
        help="Enable verbose mode to show full job details and prompt; default is false."
    )
    # New argument to pass the output folder for the DOCX file.
    parser.add_argument(
        "--path", type=str, default=None,
        help="Optional output folder path where the DOCX file will be saved. "
             "If not provided, the file is saved in the same folder as the input Markdown file."
    )
    
    args = parser.parse_args()
    moderate = (args.moderate.lower() == "true")
    verbose = args.verbose

    # -------------------------------
    # Determine output Markdown file name and location
    # -------------------------------
    if not args.output_md:
        args.output_md = f"resume_{uuid.uuid4().hex[:8]}.md"
    
    if not os.path.dirname(args.output_md):
        markdown_folder = "markdown_resumes"
        if not os.path.exists(markdown_folder):
            os.makedirs(markdown_folder)
            if verbose:
                print(f"Created markdown folder: {markdown_folder}")
        args.output_md = os.path.join(markdown_folder, args.output_md)
    
    # -------------------------------
    # Determine output DOCX file name and location
    # -------------------------------
    if not args.output_docx:
        args.output_docx = f"docx_resume_{uuid.uuid4().hex[:8]}.docx"
    
    if not os.path.dirname(args.output_docx):
        exported_folder = "exported_resumes"
        if not os.path.exists(exported_folder):
            os.makedirs(exported_folder)
            if verbose:
                print(f"Created exported folder: {exported_folder}")
        args.output_docx = os.path.join(exported_folder, args.output_docx)
    
    # -------------------------------
    # Backup the original resume file
    # -------------------------------
    backup_folder = os.path.join("markdown_resumes", "backups")
    if not os.path.exists(backup_folder):
        os.makedirs(backup_folder)
        if verbose:
            print(f"Created backup folder: {backup_folder}")
    
    original_filename = os.path.basename(args.resume)
    backup_filename = original_filename + ".bak"
    backup_path = os.path.join(backup_folder, backup_filename)
    
    # If a backup already exists, append the current date and random letters.
    if os.path.exists(backup_path):
        now = datetime.datetime.now()
        date_str = now.strftime("%d%m%y")  # DDMMYY format
        random_letters = uuid.uuid4().hex[:5]  # 5 random hexadecimal characters
        parts = original_filename.split(".")
        if len(parts) > 1:
            backup_filename = ".".join(parts[:-1]) + f".{date_str}_{random_letters}." + parts[-1] + ".bak"
        else:
            backup_filename = original_filename + f".{date_str}_{random_letters}.bak"
        backup_path = os.path.join(backup_folder, backup_filename)
    
    try:
        shutil.copy(args.resume, backup_path)
        if verbose:
            print(f"Backup of original resume saved as {backup_path}")
    except Exception as e:
        print("Error creating backup of the resume:", e)
        sys.exit(1)
    
    # Load the original resume content.
    try:
        with open(args.resume, "r", encoding="utf-8") as f:
            resume_md = f.read()
    except Exception as e:
        print("Error reading the resume file:", e)
        sys.exit(1)
    
    # Fetch job details from the provided URL.
    if verbose:
        print("Fetching job details from the URL...")
    try:
        job_details = fetch_job_details(args.job_url)
    except Exception as e:
        print("Error fetching job details:", e)
        sys.exit(1)
    if verbose:
        print("Job details fetched successfully.")
        print("\n--- JOB DETAILS ---")
        print(job_details)
        print("-------------------\n")
    else:
        print("Job details fetched successfully.")
    
    # Tailor the resume using OpenAI's API.
    print("Tailoring the resume to match the job listing...")
    tailored_resume = tailor_resume(resume_md, job_details, moderate, model=args.model, verbose=verbose)
    print("Resume tailored successfully.")
    
    # Save the tailored resume to the specified Markdown file.
    try:
        with open(args.output_md, "w", encoding="utf-8") as f:
            f.write(tailored_resume)
        if verbose:
            print(f"Tailored resume saved as {args.output_md}")
        else:
            print("Tailored resume saved.")
    except Exception as e:
        print("Error writing tailored resume:", e)
        sys.exit(1)
    
    # Convert the tailored Markdown resume to DOCX using convert_resume.py.
    convert_script = "convert_resume.py"
    if not os.path.exists(convert_script):
        print(f"Error: {convert_script} not found in the current directory.")
        sys.exit(1)
    
    print("Converting the tailored resume to DOCX...")
    try:
        # Build the conversion command with the --path argument if provided.
        convert_command = ["python", convert_script, args.output_md, args.output_docx]
        if args.path:
            convert_command.extend(["--path", args.path])
        subprocess.run(convert_command, check=True)
        if verbose:
            print(f"Conversion successful. DOCX saved as {args.output_docx}")
        else:
            print("Conversion successful.")
    except subprocess.CalledProcessError as e:
        print("Error during conversion:", e)
        sys.exit(1)

if __name__ == "__main__":
    main()