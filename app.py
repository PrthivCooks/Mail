import streamlit as st
import pandas as pd
import requests
import json
import openpyxl
import PyPDF2
import docx

# OpenAI API Key (Replace with your actual key)
OPENAI_API_KEY = "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"  # Update with your key
OPENAI_API_URL = "https://api.openai.com/v1/chat/completions"

# Function to extract text from a PDF
def extract_text_from_pdf(pdf_file):
    text = ""
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    for page in pdf_reader.pages:
        extracted_text = page.extract_text()
        if extracted_text:
            text += extracted_text + "\n"
    return text

# Function to extract text from a DOCX file
def extract_text_from_docx(docx_file):
    doc = docx.Document(docx_file)
    text = "\n".join([para.text for para in doc.paragraphs])
    return text

# Streamlit UI
st.title("📧 Automated Research Email Generator")

# Upload Excel
uploaded_file = st.file_uploader("📂 Upload Professor Details Excel", type=["xlsx"])

# Upload Resume File
resume_file = st.file_uploader("📄 Upload Your Resume (PDF/DOCX)", type=["pdf", "docx"])

# User Inputs
your_name = st.text_input("🧑‍🎓 Your Name")
your_university = st.text_input("🏫 Your University")
reason = st.text_area("📌 Why are you approaching the professor?")
resume_link = st.text_input("📎 Google Drive Link of Resume (Optional)")
sample_email = st.text_area("📄 Paste a Sample Email for Format Reference (Copy-Paste Exact Email Format Here)")

# Extract Resume Content
resume_text = ""
if resume_file:
    file_extension = resume_file.name.split(".")[-1].lower()
    if file_extension == "pdf":
        resume_text = extract_text_from_pdf(resume_file)
    elif file_extension == "docx":
        resume_text = extract_text_from_docx(resume_file)
    else:
        st.error("Unsupported file format. Please upload a PDF or DOCX file.")

# Generate Emails Button
if st.button("🚀 Generate Emails") and uploaded_file:
    df = pd.read_excel(uploaded_file)

    output_data = []

    for _, row in df.iterrows():
        professor_name = row["Professor Name"]
        professor_university = row["Professor University"]
        professor_email = row["Professor Email"]
        research_topic = row["Research Topic"]
        research_abstract = row["Research Abstract"]

        # Constructing the prompt with actual values
        prompt = f"""
        You are tasked with generating an email for a professor using the **exact structure** of the sample email below.
        
        **Replace placeholders with actual details** but do not modify the formatting, structure, or wording of the sample email.

        **Professor Details:**
        - Professor Name: {professor_name}
        - University: {professor_university}
        - Research Topic: {research_topic}
        - Research Abstract: {research_abstract}

        **My Details:**
        - Name: {your_name}
        - University: {your_university}
        - Reason for reaching out: {reason}
        - Resume Summary (Extracted from uploaded file):
          ```
          {resume_text[:1000]}  # Limiting to first 1000 characters for brevity
          ```
        - Resume Link: {resume_link}

        **Use the following email structure exactly as provided (DO NOT change format or style):**

        ```
        {sample_email}
        ```
        """

        # Prepare API request
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {OPENAI_API_KEY}"
        }

        data = {
            "model": "gpt-4o-mini",  # Change to "gpt-4o" or "gpt-3.5-turbo" if needed
            "messages": [
                {"role": "user", "content": prompt}
            ],
            "store": True
        }

        # Make API request
        response = requests.post(OPENAI_API_URL, headers=headers, json=data)

        if response.status_code == 200:
            email_body = response.json()["choices"][0]["message"]["content"]
        else:
            email_body = f"Error generating email (Status {response.status_code})"

        # Extract subject: Take the first meaningful line
        lines = email_body.split("\n")
        subject_line = "Research Collaboration Inquiry"  # Default subject
        for line in lines:
            clean_line = line.strip()
            if clean_line and len(clean_line.split()) > 2:  # Ensure it's a meaningful sentence
                subject_line = clean_line
                break

        # Ensure the resume link is clickable
        if resume_link:
            resume_button = f"[Click Here]({resume_link}) to view my resume."
        else:
            resume_button = "Resume link not provided."

        # Store data (Separate Subject & Body)
        email_body_cleaned = email_body.replace(subject_line, "").strip()  # Remove subject from body
        email_body_cleaned += f"\n\n{resume_button}"  # Append the clickable resume link
        output_data.append([professor_email, subject_line, email_body_cleaned])

    # Convert to DataFrame
    output_df = pd.DataFrame(output_data, columns=["Professor Email", "Email Subject", "Email Body"])

    # Save to Excel
    output_file = "Generated_Emails.xlsx"
    output_df.to_excel(output_file, index=False)

    # Display Download Link
    st.success("✅ Emails Generated Successfully!")
    st.download_button("📥 Download Emails Excel", data=open(output_file, "rb"), file_name=output_file)
