from pyDocxReport import DocxTemplate
from datetime import date


class DataBridge:
    '''DataBridge class manages resources and match them with keyword set in a template docx file.
    All keywords in the template so referenced ar replaced by the appropriate content.
    An example of use with a yml file as a matchs dictionary is given below:

    bridge = DataBridge(
        'path/to/template.docx',
        {'text1':'this is my replacement text', 'text2':'and another one'},
        {'table1': df1},
        {'imageset1': ['path/to/image1.jpg', 'path/to/image1.jpg'], 'imageset2':['path/to/image2.tiff']}
        )

    bridge.match(matchs)
    bridge.save('path/to/output.docx')

    where matchs is defined as a yml file like below:

        _keyword1_:
            replacewith: string
            parameters:
                replacement: text1
        _myimage1set_:
            replacewith: images
            parameters:
                replacement: imageset1
                width: 120
        _logo_:
            replacewith: images
            parameters:
                replacement: imageset2
                height: 10
        _keyword2_:
            replacewith: table
            parameters:
                replacement: table1
                header: false               # if header is true, the column names of the DataFrame are used as header. Otherwiser no header. Default is no header
        _text2_:
            replacewith: string
            parameters:
                replacement: text2
    '''

    def __init__(self, template_filename: str, texts: dict, tables: dict, images: dict):
        '''create object by setting resources
        Parameters:
        -----------
        template_filename: path to docx template
        texts: dictionary of key/string 
        tables: dictionary of key/pandas DataFrame
        images: dictionary of key/image paths
        '''
        self.texts = texts
        self.tables = tables
        self.images = images
        self.doc = DocxTemplate(template_filename)
        self.switcher = {'table': self._replaceWithTable, 'string': self._replaceWithString,
                         'images': self._replaceWithImages}

    def match(self, matchs: dict):
        '''look for keywords in docx template and replace them according to the set parameters.
        The replacement (either text, image or table) is called by its key in respectively one of the resource textx, images or tables
        set while creating this object
        Parameters:
        -----------
        matchs: dictionary go which key is the searched keyword in the docx template and value is a dictionary with the following keys:
        - replacewith: the asocciated value is the data type (either "string", "table", "images")
        - parameters : a dictionary of parameters dependent on the data type. The common and mandatory key/value pair is :
            - replacement:  the key pointing on one of the dict (tables, texts, images)
        Then for table data type :
            - header: [optional] boolean, true if the table has an header. Default is False
        For images data type:
            - width: [optional] expected width of images in mm in the docx. Default is original width
            - height: [optional] expected height of images in mm in the docx. Giving only width or height preserves aspsct ratio. Default is original height
        '''
        for keyword in matchs:
            element_type = matchs[keyword]['replacewith']
            self.switcher[element_type](keyword, matchs[keyword]['parameters'])

    def _replaceWithTable(self, keyword: str, parameters: dict):
        header = None
        replacement = parameters['replacement']
        if 'header' in parameters and parameters['header']:
            header = self.tables[replacement].columns

        table = self.doc.findTableByKeyword(keyword)
        if not table:
            raise ValueError('no table found with keyword {}'.format(keyword))

        df = self.tables[replacement]
        if header:
            self.doc.addTableHeader(table, header)
        from_row = 1 if header else 0

        df.columns = range(0, df.shape[1])

        self.doc.fillTableWithData(table, df, from_row)

    def _replaceWithString(self, keyword: str, parameters: dict):
        self.doc.replaceKeywordByString(
            keyword, self.texts[parameters['replacement']])

    def _replaceWithImages(self, keyword: str, parameters: dict):

        width = parameters['width'] if 'width' in parameters else None
        height = parameters['height'] if 'height' in parameters else None
        self.doc.replaceKeywordByImages(
            keyword, self.images[parameters['replacement']], width, height)

    def save(self, template_filename: str):
        self.doc.save(template_filename)
