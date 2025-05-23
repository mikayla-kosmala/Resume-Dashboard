import pyresparser

# Load and parse the resume file
file = "Running Resume.docx"
data = ResumeParser(file).get_extracted_data()

# Print the structured data
print(data)


