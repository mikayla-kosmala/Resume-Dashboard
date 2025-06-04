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

def add_projects():
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
    return

def add_education():
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
    return

def add_skills(skills_df, section, para):
    """
    Extracts skills and their levels from a formatted text block and appends them to a DataFrame.

    Parameters:
    - skills_df: DataFrame with columns ["Skill", "Level"]
    - section: Boolean or string flag ("1" or "0") indicating whether to process this section
    - para: Object with a .text attribute, expected to be a string like:
      "Expert: Alteryx · Excel · PowerPoint · Debugging · Troubleshooting\nAdvanced: Tableau · Regression Analysis · MySQL"

    Notes:
    - Skips processing if para.text is blank, equals "Skills", or section is 0
    - Each skill is added as a separate row with its corresponding level
    """

    if para.text != 'Skills' and para.text and section:
        # Split the paragraph into separate skill level lines
        list = para.text.split('\n')

        # Iterate through each group (e.g., "Expert: Alteryx · Excel")
        for item in list:

            if para.text:

                # Extract the skill level (e.g., "Expert") and skills list (e.g., ["Alteryx", "Excel"])
                skill_level = item.split(':')[0]
                skills = item.split(':')[1].split('·')

                for skill in skills:
                    # Construct new row to add to the DataFrame
                    new_row = {"Skill": skill, "Level": skill_level}

                    # Appending new row DataFrame
                    skills_df = pd.concat([skills_df, pd.DataFrame([new_row])], ignore_index=True)

    return skills_df


def add_interests(personal_df, section, para):
    """
    Extracts personal interests and their interest levels + Extracts personal information and the hyperlinks
    from a formatted text block and appends them to a DataFrame.

    Parameters:
    - personal_df: DataFrame with columns ["Section", "Information","Interest Level","Links","Path"]
    - section: Boolean or string flag ("1" or "0") indicating whether to process this section
    - para: Object with a .text attribute, expected to be a string like:
      "Board Games 20/100" or "· Mikayla.Kosmala@gmail.com · LinkedIn: Mikayla-Kosmala · Github: mikayla-kosmala"

    Notes:
    - Skips processing if para.text is blank, equals "Interests", or section is 0
    - Skips processing part 1 if 'Mikayla.Kosmala' isn't in para.text or section is not ''
    - Each skill is added as a separate row with its corresponding level
    """
    # Variable Creation
    new_row = []


    if section == '' and 'Mikayla.Kosmala' in para.text:
        # Iterate through hyperlinks in para to add to DataFrame
        for label, url in get_hyperlinks_from_para(para): 
            # Labels will be the text displayed 'LinkedIn: Mikayla-Kosmala' and url will be the link 'https://www.linkedin.com/in/mikayla-kosmala/'
            if ':' in label:
                # Extracting the information ("LinkedIn") and interest level "Mikayla-Kosmala"
                information, interest_level = label.split(':')

            else:
                # Special Case for my email
                information = 'Email'
                interest_level = label

            # Construct new row to add to the DataFrame
            new_row.append([{"Section":'Personal', "Information":information,"Interest Level":interest_level,"Links":url,'Path':""}])

    if para.text != 'Interests' and para.text and section:
        # Extracting the information ("Board Games") by taking all words except the last
        information=' '.join(para.text.split(' ')[:-1])

        # Extracting the interest level ("20/100") as the last word
        interest_level=para.text.split(' ')[-1]

        # Construct new row to add to the DataFrame
        new_row.append([{"Section":section, "Information":information,"Interest Level":interest_level,"Links":"",'Path':1},{"Section":section, "Information":information,"Interest Level":interest_level,"Links":"",'Path':270}])

    # Appending new row DataFrame
    personal_df = pd.concat([personal_df, pd.DataFrame(new_row)], ignore_index=True)
    return personal_df