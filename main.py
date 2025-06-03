from docx import Document
import string
import re
import pandas as pd
from datetime import datetime
from lxml import etree
from utils import get_hyperlinks_from_para, define_section, add_experience, add_interests
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

    if section == '' and 'Mikayla.Kosmala' in para.text:
        print(get_hyperlinks_from_para(para))
        for label, url in get_hyperlinks_from_para(para): 
            if ':' in label:
                information, interest_level = label.split(':')
            else:
                information = 'Email'
                interest_level = label
            new_rows = [{"Section":'Personal', "Information":information,"Interest Level":interest_level,"Links":url,'Path':""}]
            personal_df = pd.concat([personal_df, pd.DataFrame(new_rows)], ignore_index=True)
    add_experience(section=='Experience', para) # if section is 'Experience' then add rows to resume_df
    if section == "Experience":
        #print(para.text)
        if desc_found:
            desc = para.text
            desc_found = 0
        if "Bullet List" in para.style.name:
            # if end == " Current":
            #     end = datetime.today().strftime("%b %Y")
            new_row = {"Section": section, "Title": title, "Company": company, "Desc": desc, "Accomplishments":para.text, "Start Date":start, "End Date":end}
            # Append the row
            resume_df = pd.concat([resume_df, pd.DataFrame([new_row])], ignore_index=True)
        for run in para.runs:
            if run.bold:
                text = run.text.strip()
                if text:  # Ignore empty strings
                    title = text
                    list = re.sub(r'\t+','|',para.text[len(text)+1:]).split('|')
                    company=list[0]
                    start,end=list[1].split('–')
                    desc_found = 1
                    break
    if section == "Projects":
        if desc_found:
            start, end = para.text.split('|')[0].split(' – ')
            desc, url = get_hyperlinks_from_para(para)[0]
            desc_found = 0
        if "Bullet List" in para.style.name:
            new_row = {"Section": section, "Title": title, "Company": company, "Desc": desc, "Accomplishments":para.text, "Start Date":start, "End Date":end,"Link":url}
            print(new_row)
            # Append the row
            resume_df = pd.concat([resume_df, pd.DataFrame([new_row])], ignore_index=True)
        for run in para.runs:
            if run.bold:
                text = run.text.strip()
                if text:  # Ignore empty strings
                    title = text
                    list = re.sub(r'\t+','|',para.text[len(text)+2:]).split('|')
                    company=list[0]
                    #print(list)
                    desc_found = 1
                    break
    if section == "Education":
        for run in para.runs:
            text = run.text.strip()
            if run.bold and found_company==0:
                if text:  # Ignore empty strings
                    company = text
                    found_company = 1
            if run.italic and found_desc==0:
                text = run.text.strip()
                if text:
                    desc = text
                    found_desc = 1
            if  not run.bold and not run.italic and not ("Heading" in para.style.name) and text and found_company and found_desc: 
                list = re.sub(r'\t+','|',para.text[len(text)+2:]).split('|')
                start,end=list[1].split('–')
                new_row = {"Section": section, "Title": 'Student', "Company": company[:-1], "Desc": desc, "Accomplishments":'', "Start Date":start, "End Date":end}
                # Append the row
                resume_df = pd.concat([resume_df, pd.DataFrame([new_row])], ignore_index=True)
                found_company = 0
                found_desc = 0
    if section == "Skills" and para.text != 'Skills':
            list = para.text.split('\n')
            for item in list:
                if para.text:
                    group = item.split(':')
                    level = group[0]
                    skills = group[1].split('·')
                    #print(skills)
                    for skill in skills:
                        #print(skill)
                        new_row = {"Skill": skill, "Level": level}
                        skills_df = pd.concat([skills_df, pd.DataFrame([new_row])], ignore_index=True)
    personal_df = add_interests(personal_df, section=='Interests', para)

personal_df.to_excel(r"C:\Users\mikay\OneDrive\Documents\Resume to Tableau\personal_data.xlsx", index=False)
resume_df.to_excel(r"C:\Users\mikay\OneDrive\Documents\Resume to Tableau\resume_data.xlsx", index=False)
skills_df.to_excel(r'C:\Users\mikay\OneDrive\Documents\Resume to Tableau\skills_data.xlsx', index=False)