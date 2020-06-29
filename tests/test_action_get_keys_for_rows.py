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

from get_keys_for_rows import GetExcelSheetsAction


class GetKeysForRowsTestCase(ExcelBaseActionTestCase):
    __test__ = True
    action_cls = GetExcelSheetsAction

    SHEET_1 = [["Col1", "Col2", "Col3"], ["key1", "ro1_2", "ro1_3"],
               ["key2", "ro2_2", "ro2_3"],
               ["key3", "ro3_2", "ro3_3"]]
    SHEET_2 = [["Col1", "Col2"]]
    _MOCK_SHEETS = {"sheet1": SHEET_1,
                    "sheet2": SHEET_2}

    def setUp(self):
        super(GetKeysForRowsTestCase, self).setUp()
        self._full_config = self.load_yaml('full.yaml')

    def load_yaml(self, filename):
        return yaml.safe_load(self.get_fixture_content(filename))

    @property
    def full_config(self):
        return self._full_config

    def return_workbook(filename, data_only):
        GetKeysForRowsTestCase.WB = ExcelBaseActionTestCase.MockWorkbook(
            GetKeysForRowsTestCase._MOCK_SHEETS, None)
        GetKeysForRowsTestCase.WB.save = mock.MagicMock()
        return GetKeysForRowsTestCase.WB

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_get_keys_default_sheet_exists(self):
        action = self.get_action_instance(self.full_config)
        result = action.run('sheet1', "mock_excel.xlsx")

        self.assertIsNotNone(result)
        self.assertTrue("key1" in result)
        self.assertTrue("key2" in result)
        self.assertTrue("key3" in result)
        self.assertEquals(3, len(result))
        GetKeysForRowsTestCase.WB.save.assert_not_called()

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_get_keys_col_2_sheet_exists(self):
        action = self.get_action_instance(self.full_config)
        result = action.run('sheet1', "mock_excel.xlsx", 2)

        self.assertIsNotNone(result)
        self.assertTrue("ro1_2" in result)
        self.assertTrue("ro2_2" in result)
        self.assertTrue("ro3_2" in result)
        self.assertEquals(3, len(result))
        GetKeysForRowsTestCase.WB.save.assert_not_called()

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_get_keys_row_col_specify_exists(self):
        action = self.get_action_instance(self.full_config)
        result = action.run('sheet1', "mock_excel.xlsx", 3, 2)

        self.assertIsNotNone(result)
        self.assertTrue("ro2_3" in result)
        self.assertTrue("ro3_3" in result)
        self.assertEquals(2, len(result))
        GetKeysForRowsTestCase.WB.save.assert_not_called()

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_get_keys_sheet_not_exist(self):
        action = self.get_action_instance(self.full_config)
        with self.assertRaises(KeyError):
            action.run('sheet3', "mock_excel.xlsx")

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', mock.Mock(return_value=False))
    def test_get_keys_spreadsheet_not_exist(self):
        action = self.get_action_instance(self.full_config)
        with self.assertRaises(ValueError):
            action.run('sheet3', "mock_excel.xlsx")

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_get_keys_sheet_not_exist_strict(self):
        action = self.get_action_instance(self.full_config)
        with self.assertRaises(KeyError):
            action.run('sheet3', "mock_excel.xlsx", strict=True)

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_get_keys_sheet_not_exist_not_strict(self):
        action = self.get_action_instance(self.full_config)
        result = action.run('sheet3', "mock_excel.xlsx", strict=False)

        self.assertIsNotNone(result)
        self.assertEquals([], result)
        GetKeysForRowsTestCase.WB.save.assert_not_called()

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', mock.Mock(return_value=False))
    def test_get_keys_spreadsheet_not_exist_not_strict(self):
        action = self.get_action_instance(self.full_config)
        with self.assertRaises(ValueError):
            action.run('sheet3', "mock_excel.xlsx", strict=False)
