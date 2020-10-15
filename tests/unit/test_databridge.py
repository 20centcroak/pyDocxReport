import unittest
from pandas import DataFrame
from yaml import safe_load

from pyDocxReport import DataBridge


class TestDataBridge(unittest.TestCase):

    def test_bridge(self):

        d = {'col1': [1, 2], 'col2': [3, 4]}
        df1 = DataFrame(data=d)

        bridge = DataBridge(
            'tests/unit/resources/template.docx',
            {'text1': 'When animals such as frogs and crows croak, they make deep rough sounds.',
                'text2': 'A frog is any member of a diverse and largely carnivorous group of short-bodied, tailless amphibians composing the order Anura (literally without tail in Ancient Greek)'},
            {'table1': df1},
            {'imageset1': ['tests/unit/resources/image1.jpg', 'tests/unit/resources/image2.jpg'],
                'imageset2': ['tests/unit/resources/logo.jpg']}
        )

        with open('tests/unit/resources/matchs.yml', 'r') as file:
            matchs = safe_load(file.read())

        bridge.match(matchs)
        bridge.save('tests/unit/output/output.docx')
