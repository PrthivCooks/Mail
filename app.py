import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader
import google.generativeai as genai
import os

# Configure Google Generative AI
genai.configure(api_key="AIzaSyAOw3Y-QYSQ1uq8XzcEdxdUS9tOHWcSRZw")
model = genai.GenerativeModel("gemini-1.5-flash")

# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
        return text
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
        return ""

# Generate the email content using prompts
def generate_email(prompt_template, **data):
    try:
        # Populate the prompt with data from the Excel row
        filled_prompt = prompt_template.format(**data)
        response = model.generate_content(filled_prompt).text.strip()
        return response
    except Exception as e:
        st.error(f"Error generating content: {e}")
        return None

# Streamlit App
st.title("Dynamic Email Generation App")

# Step 1: Ask the User about the Email Purpose
st.subheader("Step 1: Configure Email Context")
email_purpose = st.radio(
    "What is the purpose of this email?",
    ("Internship Application", "Research Collaboration", "Other")
)

specific_program = st.text_input("If applicable, specify the program or internship (e.g., ThinkSwiss Program).")

fine_tuning = st.text_area(
    "Optional: Add any specific instructions or customizations for the email (e.g., tone, specific points to include).",
    height=100
)

# Step 2: Base Prompt Template
st.subheader("Step 2: Review and Confirm the Prompt Template")
base_prompt = """
Generate a professional email to Prof. {professor_name} at {professor_university}. Begin by introducing yourself as {student_name}, a student from {student_university}. Mention that you are contacting them for {email_purpose}, specifically {specific_program}.

Discuss the professor's research: "{research_topic}" and abstract: "{research_abstract}". Highlight how impressed you are by their work and its relevance, and why it inspired you to reach out.

Summarize your skills and achievements based on your resume: "{resume_text}".

Conclude by expressing your enthusiasm to contribute under their guidance and mentioning that your CV is attached. End with gratitude and your contact information.
"""

if fine_tuning:
    base_prompt += "\n\nAdditional Instructions: " + fine_tuning

st.write(base_prompt)

# Step 3: Upload Files
st.subheader("Step 3: Upload Required Files")
uploaded_excel = st.file_uploader("Upload Excel File (Template Format)", type=["xlsx"])
uploaded_resume = st.file_uploader("Upload Resume PDF", type=["pdf"])

# Step 4: Generate Emails
if st.button("Generate Emails"):
    if uploaded_excel and uploaded_resume:
        # Save uploaded resume locally
        resume_path = "uploaded_resume.pdf"
        with open(resume_path, "wb") as f:
            f.write(uploaded_resume.read())

        # Extract text from the uploaded resume
        resume_text = extract_text_from_pdf(resume_path)

        # Read the uploaded Excel file
        try:
            df = pd.read_excel(uploaded_excel)
        except Exception as e:
            st.error(f"Error reading Excel file: {e}")
            st.stop()

        # Process the Excel data and generate emails
        output_data = []
        for _, row in df.iterrows():
            # Prepare data dictionary for prompt
            data = {
                "student_name": row["Student Name"],
                "student_university": row["Student University"],
                "student_email": row["Student Email"],
                "professor_name": row["Professor Name"],
                "professor_university": row["Professor University"],
                "professor_email": row["Professor Email"],
                "research_topic": row["Research Topic"],
                "research_abstract": row["Research Abstract"],
                "resume_text": resume_text,
                "email_purpose": email_purpose,
                "specific_program": specific_program,
            }

            # Generate email content
            email_content = generate_email(base_prompt, **data)

            if email_content:
                output_data.append({
                    "Student Name": data["student_name"],
                    "Professor Name": data["professor_name"],
                    "Email to Send": data["professor_email"],
                    "Generated Email": email_content,
                })

        # Save output data to an Excel file
        if output_data:
            output_df = pd.DataFrame(output_data)
            output_excel_path = "generated_emails.xlsx"
            output_df.to_excel(output_excel_path, index=False)
            with open(output_excel_path, "rb") as f:
                st.download_button("Download Generated Emails", f, "generated_emails.xlsx")
        else:
            st.error("No emails were generated. Please check your inputs.")
    else:
        st.error("Please upload all required files.")
