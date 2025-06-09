from docx import Document
import string
import re
import pandas as pd
from datetime import datetime
from lxml import etree
from utils import get_hyperlinks_from_para, define_section, add_experience, add_interests, add_skills, add_education, add_projects, sqlite_db
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import Experience as ex
# Access the relationships

# Load your resume
doc = Document(r'C:\Users\mikay\OneDrive\Documents\Resume to Tableau\Running Resume.docx')
rels = doc.part.rels

""" Call Initalizer for Classes and pass it the raw text of the experience section

First for loop finds the section
Second for loop creates the list of experience objects
Third for loop extracts objects to dataframe for excel output

# resume_df = pd.DataFrame(columns=["Section", "Title", "Company", "Description", "Accomplishments", "Start Date", "End Date"])
"""
resume_df = pd.DataFrame(columns=["Section", "Title", "Company", "Description", "Accomplishments", "Start Date", "End Date"])
section = ''
section_list = ["Experience","Education","Projects"]
sections = {}
for i in section_list:
    sections[f"{i}"] = ""
for para in doc.paragraphs:
    if define_section(para)!='':
        section = define_section(para)
    if section in section_list and para.text != section: 
        sections[f"{section_list[section_list.index(section)]}"]+='\n'+para.text 
    
list_of_experience = []
for section in sections:
    if section == 'Education':
        sections[f"{section}"] = sections[f"{section}"].split('\n')[1:]
    else:
        sections[f"{section}"] = sections[f"{section}"].split('\n\n')
        
    for item in sections[f"{section}"]:
        experience = ex.Experience(section,item)
        experience = experience.parse()
        list_of_experience.append(experience)

print([experience.to_dict() for experience in list_of_experience])

# Need to fix the lists of dictionaries
resume_df = pd.DataFrame([experience.to_dict() for experience in list_of_experience])

resume_df.head(10)


# """
# Excel Version
# """
# personal_df.to_excel(r"C:\Users\mikay\OneDrive\Documents\Resume to Tableau\personal_data.xlsx", index=False)
# resume_df.to_excel(r"C:\Users\mikay\OneDrive\Documents\Resume to Tableau\resume_data.xlsx", index=False)
# skills_df.to_excel(r'C:\Users\mikay\OneDrive\Documents\Resume to Tableau\skills_data.xlsx', index=False)

# """
# DB version

# db_name = "resume_data"
# db_location = "\Documents\Github\Resume-Dashboard"

# sqlite_db(resume_df, db_name, "experience", db_location)
# sqlite_db(personal_df, db_name, "personal", db_location)
# sqlite_db(resume_df, db_name, "skills", db_location)

# """