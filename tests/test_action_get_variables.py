# Licensed to the StackStorm, Inc ('StackStorm') under one or more
# contributor license agreements.  See the NOTICE file distributed with
# this work for additional information regarding copyright ownership.
# The ASF licenses this file to You under the Apache License, Version 2.0
# (the "License"); you may not use this file except in compliance with
# the License.  You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and

import mock
import yaml

from excel_base_action_test_case import ExcelBaseActionTestCase

from datetime import datetime

from get_variables import GetExcelVariablesAction


class GetVariablesTestCase(ExcelBaseActionTestCase):
    __test__ = True
    action_cls = GetExcelVariablesAction

    SHEET_1 = [ [ "Col1", "Col2", "Col3" ], [ "key1", "ro1_2", "rol1_3" ], [ "key2", "ro2_2", "ro2_3" ], ["key3", "ro3_2", "ro3_3"] ]
    _MOCK_SHEETS = {"sheet1": SHEET_1}

    def setUp(self):
        super(GetVariablesTestCase, self).setUp()
        self._full_config = self.load_yaml('full.yaml')

    def load_yaml(self, filename):
        return yaml.safe_load(self.get_fixture_content(filename))

    @property
    def full_config(self):
        return self._full_config

    def return_workbook(filename, data_only):
      GetVariablesTestCase.WB =  ExcelBaseActionTestCase.MockWorkbook(GetVariablesTestCase._MOCK_SHEETS, None)
      GetVariablesTestCase.WB.save = mock.MagicMock()
      return GetVariablesTestCase.WB

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_get_variables_row_exists(self):
        action = self.get_action_instance(self.full_config)
        result = action.run('key2', 'sheet1', excel_file="mock_excel.xlsx")

        self.assertIsNotNone(result)
        self.assertEquals("ro2_2", result["Col2"]) 
        self.assertEquals("ro2_3", result["Col3"]) 
        GetVariablesTestCase.WB.save.assert_not_called()

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_get_variables_sheet_not_exist(self):
        action = self.get_action_instance(self.full_config)
        with self.assertRaises(KeyError):
            action.run('key1', 'sheet3', excel_file="mock_excel.xlsx")

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', mock.Mock(return_value=False))
    def test_get_variables_spreadsheet_not_exist(self):
        action = self.get_action_instance(self.full_config)
        with self.assertRaises(ValueError):
            action.run('key1', 'sheet3', excel_file="mock_excel.xlsx")
