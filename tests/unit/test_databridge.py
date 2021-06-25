import unittest
from pandas import DataFrame
from yaml import safe_load

from pyDocxReport import DataBridge


class TestDataBridge(unittest.TestCase):

    def test_bridge(self):

        d = {'col1': [1, 2], 'col2': [3, 4]}
        df1 = DataFrame(data=d)

        bridge = DataBridge('tests/unit/resources/template.docx')

        with open('tests/unit/resources/matchs.yml', 'r') as file:
            matchs = safe_load(file.read())

        matchs['_keyword2_'] = {
            'type': 'table',
            'replacement': df1,
            'header': False
        }

        bridge.match(matchs)
        bridge.save('output/output.docx')
