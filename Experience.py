from utils import clean_leading_trailing_whitespace

class Experience:
    """
    title
    company
    Start Date
    End Date
    Desc
    Accomplishments (List)
    """

    def __init__(self, section, raw_text):
        self.raw_text = raw_text.strip()
        self.section = section
        self.title = ''
        self.company = ''
        self.start_date = ''
        self.end_date = ''
        self.description = ''
        self.accomplishments = []

    def clean(self):
        self.title = self.title.strip()
        self.company = self.company.strip()
        self.start_date = self.start_date.strip()
        self.end_date = self.end_date.strip()
        self.description = self.description.strip()
        self.accomplishments = clean_leading_trailing_whitespace(self.accomplishments)
        return self

    def parse(self):
        if self.section == 'Experience':
            lines = [line.strip() for line in self.raw_text.strip().split('\n') if line.strip()]
            title_company = lines[0].split('\t')[0].split(',')
            self.title = ','.join(title_company[:-1])
            self.company = title_company[-1]
            self.start_date, self.end_date = lines[0].split('\t')[-1].split(' – ')
            self.description = lines[1]
            for i in lines[2:]:
                if i:
                    self.accomplishments.append(i)
        if self.section == 'Projects':
            lines = [line.strip() for line in self.raw_text.strip().split('\n') if line.strip()]
            title_company = lines[0].split(', ')
            self.title = ','.join(title_company[:-1])
            self.company = title_company[-1]
            dates_links = lines[1].split('|')
            self.start_date, self.end_date = dates_links[0].split(' – ')
            for i in lines[2:]:
                if i:
                    self.accomplishments.append(i)
        if self.section == 'Education':
            lines = [line.strip() for line in self.raw_text.strip().split('\n') if line!='']
            title_company = lines[0].split('\t')[0].split(',')
            self.company = title_company[0]
            self.title = ','.join(title_company[1:])
            self.start_date, self.end_date = lines[0].split('\t')[-1].split(' – ')
        return self.clean()

    def to_dict(self):
        if self.section == 'Education':
            return[{
            "Section": self.section,
            "Title": self.title, 
            "Company": self.company,
            "Description": self.description,
            "Accomplishments": self.accomplishments, 
            "Start Date":self.start_date, 
            "End Date": self.end_date}]
        return [{
            "Section": self.section,
            "Title": self.title, 
            "Company": self.company,
            "Description": self.description,
            "Accomplishments": self.accomplishments[i], 
            "Start Date":self.start_date, 
            "End Date": self.end_date} for i in range(len(self.accomplishments))]
