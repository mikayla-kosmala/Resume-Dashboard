class Experience:
    """
    title
    company
    Start Date
    End Date
    Desc
    Accomplishments (List)
    """

    def __init__(self, raw_text):
        self.raw_text = raw_text.strip()
        self.title = ''
        self.company = ''
        self.start_date = ''
        self.end_date = ''
        self.description = ''
        self.accomplishments = ['']

    def _parse(self):
        