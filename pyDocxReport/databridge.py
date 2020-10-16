from pyDocxReport import DocxTemplate
from datetime import date
from pandas import DataFrame


class DataBridge:
    '''DataBridge class manages resources and match them with keyword set in a template docx file.
    All keywords in the template so referenced ar replaced by the appropriate content.
    An example of use with a yml file as a matchs dictionary is given below:

    bridge = DataBridge('path/to/template.docx')

    bridge.match(matchs)
    bridge.save('path/to/output.docx')

    where matchs is defined as a yml file like below:

        _keyword1_:
            type: string
            replacement: text1
        _myimage1set_:
            type: images
            replacement: 
                - path/to/image1.png
                - path/to/image2.jpg
            width: 120
        _logo_:
            type: images
            replacement: path/to/logo.png
            height: 10
        _keyword2_:
            type: table
            replacement: table1
            header: false               # if header is true, the column names of the DataFrame are used as header. Otherwiser no header. Default is no header
        _text2_:
            type: string
            replacement: text2
    '''

    def __init__(self, template_filename: str):
        '''create object by setting resources
        Parameters:
        -----------
        template_filename: path to docx template
        texts: dictionary of key/string 
        tables: dictionary of key/pandas DataFrame
        images: dictionary of key/image paths
        '''
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
        - type: the asocciated value is the data type (either "string", "table", "images")
        - replacement:  the replacement value
        Then for table data type :
            - header: [optional] boolean, true if the table should display the datagrame columns as a header. Default is False
        For images data type:
            - width: [optional] expected width of images in mm in the docx. Default is original width
            - height: [optional] expected height of images in mm in the docx. Giving only width or height preserves aspsct ratio. Default is original height
        '''
        for keyword in matchs:
            element_type = matchs[keyword]['type']
            self.switcher[element_type](keyword, matchs[keyword])

    def _replaceWithTable(self, keyword: str, parameters: dict):
        header = None
        df: DataFrame = parameters['replacement']
        if 'header' in parameters and parameters['header']:
            header = df.columns

        table = self.doc.findTableByKeyword(keyword)
        if not table:
            raise ValueError('no table found with keyword {}'.format(keyword))

        if header:
            self.doc.addTableHeader(table, header)
        from_row = 1 if header else 0

        self.doc.fillTableWithData(table, df, from_row)

    def _replaceWithString(self, keyword: str, parameters: dict):
        self.doc.replaceKeywordByString(keyword, parameters['replacement'])

    def _replaceWithImages(self, keyword: str, parameters: dict):
        width = parameters['width'] if 'width' in parameters else None
        height = parameters['height'] if 'height' in parameters else None
        images = parameters['replacement'] if type(parameters['replacement']) is list else [parameters['replacement']]
        self.doc.replaceKeywordByImages(
            keyword, images, width, height)

    def save(self, template_filename: str):
        self.doc.save(template_filename)
