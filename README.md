# pyDocxReport
Ease the docx report generation using templates and importing features

## DataBridge
DataBridge class manages resources and match them with keyword set in a template docx file.
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

See [tests](https://github.com/20centcroak/pyDocxReport) to see this example implemented.

## DocxTemplate
The DocxTemplate class makes use of python-docx to modify a word document.
Use DataBridge for a standard operation and use DocxTemplate when you need to tune some replacements.