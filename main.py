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

# resume_df = pd.DataFrame(columns=["Section","Title", "Company", "Description", "Accomplishments", "Start Date", "End Date"])
"""
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
        print(section,para.text)

print(sections['Education'])
list_of_experience = []
for i in sections:
    if i == 'Education':
        #print(sections[f"{i}"])
        sections[f"{i}"] = sections[f"{i}"].split('\n')[1:]
        #print(sections[f"{i}"])
    else:
        sections[f"{i}"] = sections[f"{i}"].split('\n\n')
    for item in sections[f"{i}"]:
        experience = ex.Experience(i,item)
        experience = experience.parse()
        list_of_experience.append(experience)

resume_df = pd.DataFrame([experience.to_dict for experience in list_of_experience])





# job_class = ex.Experience()

# for job in sections['Experience']:
#     ex.Experience(job)




# Print all the paragraphs
#for para in doc.paragraphs:
#  print(para.text)
# job = []
# job_info = []
# job_title = []
# company = []
# start_end_date = []
# job_desc = []
# job_achievements = []
# found_title = 0
# found_company = 0
# found_date = 0
# found_desc = 0
# found_achievements = 0
# desc_found =0
# section = ''
# resume_df = pd.DataFrame(columns=["Section","Title", "Company", "Desc", "Accomplishments", "Start Date", "End Date"])
# skills_df = pd.DataFrame(columns=["Skill","Level"])
# personal_df = pd.DataFrame(columns=["Section", "Information","Interest Level","Links","Path"])
# new_rows = [{"Section":'Personal', "Information":'Resume',"Interest Level":'',"Links":'https://docs.google.com/document/d/1ehoNrqLzMcSuB2BzAyix7t51HYPdDZ7a/edit?usp=sharing&ouid=100832385938879723557&rtpof=true&sd=true','Path':""}]
# personal_df = pd.concat([personal_df, pd.DataFrame(new_rows)], ignore_index=True)


# lines = [para.text.strip() for para in doc.paragraphs if para.text.strip() != '']

# # Recombine with real line breaks to preserve spacing
# full_text = '\n'.join(lines)

# # Now split by blank lines
# sections = full_text.strip().split('\n\n')



# for para in doc.paragraphs:
#     section = define_section(para)
#     #resume_df = add_experience(section=='Experience', para)
#     print(para.text)
#     #resume_df = add_projects(section=='Projects',para)
#     #resume_df = add_education(section=='Education',para)
#     #skills_df = add_skills(skills_df, section =="Skills", para)
#     #personal_df = add_interests(personal_df, section in ["Interests", ""], para)
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