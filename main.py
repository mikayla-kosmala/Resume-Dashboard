from docx import Document
import string
import re
import pandas as pd
from datetime import datetime
from lxml import etree
from utils import get_hyperlinks_from_para, define_section, add_experience, add_interests, add_skills
from docx.opc.constants import RELATIONSHIP_TYPE as RT
# Access the relationships

# Load your resume
doc = Document(r'C:\Users\mikay\OneDrive\Documents\Resume to Tableau\Running Resume.docx')
rels = doc.part.rels

# Print all the paragraphs
#for para in doc.paragraphs:
#  print(para.text)
job = []
job_info = []
job_title = []
company = []
start_end_date = []
job_desc = []
job_achievements = []
found_title = 0
found_company = 0
found_date = 0
found_desc = 0
found_achievements = 0
desc_found =0
section = ''
resume_df = pd.DataFrame(columns=["Section","Title", "Company", "Desc", "Accomplishments", "Start Date", "End Date"])
skills_df = pd.DataFrame(columns=["Skill","Level"])
personal_df = pd.DataFrame(columns=["Section", "Information","Interest Level","Links","Path"])
new_rows = [{"Section":'Personal', "Information":'Resume',"Interest Level":'',"Links":'https://docs.google.com/document/d/1ehoNrqLzMcSuB2BzAyix7t51HYPdDZ7a/edit?usp=sharing&ouid=100832385938879723557&rtpof=true&sd=true','Path':""}]
personal_df = pd.concat([personal_df, pd.DataFrame(new_rows)], ignore_index=True)
for para in doc.paragraphs:
    
    section = define_section(para)
    resume_df = add_experience(section=='Experience', para)
    resume_df = add_project(section=='Projects',para)
    resume_df = add_education(section=='Projects',para) # if section is 'Experience' then add rows to resume_df
    
    skills_df = add_skills(skills_df, section =="Skills", para)
    personal_df = add_interests(personal_df, section in ["Interests", ""], para)

personal_df.to_excel(r"C:\Users\mikay\OneDrive\Documents\Resume to Tableau\personal_data.xlsx", index=False)
resume_df.to_excel(r"C:\Users\mikay\OneDrive\Documents\Resume to Tableau\resume_data.xlsx", index=False)
skills_df.to_excel(r'C:\Users\mikay\OneDrive\Documents\Resume to Tableau\skills_data.xlsx', index=False)