#Importing Libraries
from docx.oxml.ns import qn #qn for get_hyperlinks_from_para

# for add_experience
import re 
import pandas as pd 

def get_hyperlinks_from_para(para):
    #initalize hyperlinks variable output
    hyperlinks = []

    # Loop through all hyperlink elements in the paragraph
    for hyperlink in para._element.findall(".//w:hyperlink", para._element.nsmap):
        # Get the r:id (relationship ID) that links to the actual URL
        r_id = hyperlink.get(qn('r:id'))
        if r_id:
            # Get the actual target (URL) from the paragraph's part
            url = para.part.rels[r_id].target_ref

            # Extract the visible link text
            texts = [node.text for node in hyperlink.findall(".//w:t", para._element.nsmap) if node.text]
            print(texts)

            # Join items in texts to display the true link_text
            link_text = "".join(texts)

            # Append link_text and url for user to be able to extract the display text (link_text) and the website url (url)
            hyperlinks.append((link_text, url))

    return hyperlinks

def define_section(para):
    if "Heading" in para.style.name:
        section = para.text
    return section

def add_experience(resume_df, section, para, title, company, start_date, end_date, desc, desc_found):                  
    #initalizing desc_found to 
    if section == "Experience":
        #print(para.text)
        if desc_found:
            desc = para.text
            desc_found = 0
        if "Bullet List" in para.style.name:
            # if end == " Current":
            #     end = datetime.today().strftime("%b %Y")
            new_row = {"Section": section, "Title": title, "Company": company, "Desc": desc, "Accomplishments":para.text, "Start Date":start_date, "End Date":end_date}
            # Append the row
            resume_df = pd.concat([resume_df, pd.DataFrame([new_row])], ignore_index=True)
        for run in para.runs:
            if run.bold:
                text = run.text.strip()
                if text:  # Ignore empty strings
                    title = text
                    list = re.sub(r'\t+','|',para.text[len(text)+1:]).split('|')
                    company=list[0]
                    start_date, end_date =list[1].split('–')
                    desc_found = 1
                    break
    return resume_df, section, para, title, company, start_date, end_date, desc, desc_found

def add_interests(personal_df, section, para):
    if section and para.text != 'Interests':
        if para.text:
            interest_level=para.text.split(' ')[-1]
            information=' '.join(para.text.split(' ')[:-1])
            new_rows = [{"Section":section, "Information":information,"Interest Level":interest_level,"Links":"",'Path':1},{"Section":section, "Information":information,"Interest Level":interest_level,"Links":"",'Path':270}]
            personal_df = pd.concat([personal_df, pd.DataFrame(new_rows)], ignore_index=True)
    return personal_df