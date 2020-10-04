from pycroaktools.datapack import DataPack
from pycroaktools.word import Word
from datetime import date

class DataBridge:

    def __init__(self, filename: str, datapack: DataPack, references: dict):
        self.pack = datapack.pack
        self.references = references
        self.doc = Word(filename)
        self.switcher = {'table': self.replaceWithTable, 'string': self.replaceWithString,
            'date': self.replaceWithDate, 'images': self.replaceWithImages}

    def match(self, matchs: dict):
        for keyword in matchs:
            element_type = matchs[keyword]['replacewith']
            self.switcher[element_type](keyword, matchs[keyword]['parameters'])

    def replaceWithTable(self, keyword: str, parameters: dict):
        header = None
        replacement = parameters['replacement']
        if 'header' in parameters and parameters['header']:
            header = self.pack[replacement].columns

        table = self.doc.findTableByKeyword(keyword)
        if not table:
            raise ValueError('no table found with keyword {}'.format(keyword))
        
        df = self.pack[replacement]
        if header is not None:
            self.doc.addTableHeader(table, header)
            df = df.loc[1:]

        df.columns = range(0, df.shape[1])

        from_row = 1 if 'header' in parameters else 0

        self.doc.fillTableWithData(table, df, from_row)


    def replaceWithString(self, keyword: str, parameters: dict):
        self.doc.replaceKeyword(keyword, self.pack[parameters['replacement']])


    def replaceWithDate(self, keyword: str, parameters: dict):

        format = r'%d/%m/%Y'
        if 'format' in parameters:
            format = parameters['format']

        if parameters['replacement'] == 'today':
            self.doc.replaceKeyword(keyword, date.today().strftime(format))
        else:
            raise ValueError('replacement {} in {} is not recognised'.format(
                parameters['replacement'], keyword))


    def replaceWithImages(self, keyword: str, parameters: dict):
        width = parameters['width'] if 'width' in parameters else None
        height = parameters['height'] if 'height' in parameters else None
        self.doc.replaceKeywordByImages(
            keyword, self.references.fileset[parameters['replacement']], width, height)

    def save(self, filename: str):
        self.doc.save(filename)
