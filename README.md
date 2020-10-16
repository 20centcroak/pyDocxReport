# pyDocxReport
Ease the docx report generation using templates and importing features

## DataBridge
DataBridge class manages resources and match them with keyword set in a template docx file.
All keywords in the template so referenced ar replaced by the appropriate content.
An example of use with a yml file as a matchs dictionary is given below:

bridge = DataBridge('path/to/template.docx')
matchs['_keyword2_'] = {
                            'type': 'table',
                            'replacement': df1,
                            'header': False
                        }
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

See [tests](https://github.com/20centcroak/pyDocxReport/blob/main/tests/unit/test_databridge.py) to see this example implemented.

## DocxTemplate
The DocxTemplate class makes use of python-docx to modify a word document.
Use DataBridge for a standard operation and use DocxTemplate when you need to tune some replacements.