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

from delete_row import DeleteExcelRowAction


class DeleteRowsTestCase(ExcelBaseActionTestCase):
    __test__ = True
    action_cls = DeleteExcelRowAction

    SHEET_1 = [["Col1", "Col2"], ["key1", "ro1_2"], ["key2", "ro2_2"],
               ["key3", "ro3_2"]]
    SHEET_2 = [["Col1", "Col2"]]
    _MOCK_SHEETS = {"sheet1": SHEET_1,
                    "sheet2": SHEET_2}

    def setUp(self):
        super(DeleteRowsTestCase, self).setUp()
        self._full_config = self.load_yaml('full.yaml')

    def load_yaml(self, filename):
        return yaml.safe_load(self.get_fixture_content(filename))

    @property
    def full_config(self):
        return self._full_config

    def return_workbook(filename, data_only):
        DeleteRowsTestCase.WB = ExcelBaseActionTestCase.MockWorkbook(
            DeleteRowsTestCase._MOCK_SHEETS, None)
        DeleteRowsTestCase.WB.save = mock.MagicMock()
        return DeleteRowsTestCase.WB

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_delete_first_row_exists(self):
        action = self.get_action_instance(self.full_config)
        result = action.run('sheet1', 'key1', True, "mock_excel.xlsx")

        self.assertIsNotNone(result)
        self.assertTrue(result[0])
        self.assertEquals("Success", result[1])
        DeleteRowsTestCase.WB.save.assert_called()
        firstcol = ExcelBaseActionTestCase._util_get_column(
            DeleteRowsTestCase.WB.get_sheet_by_name("sheet1").mockrows,
            0)
        self.assertTrue("key1" not in firstcol)
        self.assertTrue("key2" in firstcol)
        self.assertTrue("key3" in firstcol)
        self.assertTrue("Col1" in firstcol)

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_middle_first_row_exists(self):
        action = self.get_action_instance(self.full_config)
        result = action.run('sheet1', 'key2', True, "mock_excel.xlsx")

        self.assertIsNotNone(result)
        self.assertTrue(result[0])
        self.assertEquals("Success", result[1])
        DeleteRowsTestCase.WB.save.assert_called()
        firstcol = ExcelBaseActionTestCase._util_get_column(
            DeleteRowsTestCase.WB.get_sheet_by_name("sheet1").mockrows,
            0)
        self.assertTrue("key2" not in firstcol)
        self.assertTrue("key1" in firstcol)
        self.assertTrue("key3" in firstcol)
        self.assertTrue("Col1" in firstcol)

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_last_first_row_exists(self):
        action = self.get_action_instance(self.full_config)
        result = action.run('sheet1', 'key3', True, "mock_excel.xlsx")

        self.assertIsNotNone(result)
        self.assertTrue(result[0])
        self.assertEquals("Success", result[1])
        DeleteRowsTestCase.WB.save.assert_called()
        firstcol = ExcelBaseActionTestCase._util_get_column(
            DeleteRowsTestCase.WB.get_sheet_by_name("sheet1").mockrows,
            0)
        self.assertTrue("key3" not in firstcol)
        self.assertTrue("key1" in firstcol)
        self.assertTrue("key2" in firstcol)
        self.assertTrue("Col1" in firstcol)

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_row_exists_specify_key_column(self):
        action = self.get_action_instance(self.full_config)
        result = action.run('sheet1', 'ro1_2', True, "mock_excel.xlsx", 2)

        self.assertIsNotNone(result)
        self.assertTrue(result[0])
        self.assertEquals("Success", result[1])
        DeleteRowsTestCase.WB.save.assert_called()
        firstcol = ExcelBaseActionTestCase._util_get_column(
            DeleteRowsTestCase.WB.get_sheet_by_name("sheet1").mockrows,
            0)
        self.assertTrue("key1" not in firstcol)
        self.assertTrue("key2" in firstcol)
        self.assertTrue("key3" in firstcol)
        self.assertTrue("Col1" in firstcol)

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_row_exists_specify_key_row(self):
        action = self.get_action_instance(self.full_config)
        result = action.run('sheet1', 'key2', True, "mock_excel.xlsx",
                            variable_name_row=2)

        self.assertIsNotNone(result)
        self.assertTrue(result[0])
        self.assertEquals("Success", result[1])
        DeleteRowsTestCase.WB.save.assert_called()
        firstcol = ExcelBaseActionTestCase._util_get_column(
            DeleteRowsTestCase.WB.get_sheet_by_name("sheet1").mockrows,
            0)
        self.assertTrue("key2" not in firstcol)
        self.assertTrue("key1" in firstcol)
        self.assertTrue("key3" in firstcol)
        self.assertTrue("Col1" in firstcol)

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_row_exists_specify_key_before_key_row(self):
        action = self.get_action_instance(self.full_config)
        with self.assertRaises(ValueError):
            action.run('sheet1', 'key1', True, "mock_excel.xlsx",
                       variable_name_row=2)
        DeleteRowsTestCase.WB.save.assert_not_called()

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_row_not_exist_and_strict(self):
        action = self.get_action_instance(self.full_config)
        with self.assertRaises(ValueError):
            action.run('sheet1', 'key4', True, "mock_excel.xlsx")
        DeleteRowsTestCase.WB.save.assert_not_called()

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_row_not_exist_and_not_strict(self):
        action = self.get_action_instance(self.full_config)
        result = action.run('sheet1', 'key4', False, "mock_excel.xlsx")

        self.assertIsNotNone(result)
        self.assertTrue(result[0])
        self.assertEquals("Success", result[1])
        DeleteRowsTestCase.WB.save.assert_called()
        firstcol = ExcelBaseActionTestCase._util_get_column(
            DeleteRowsTestCase.WB.get_sheet_by_name("sheet1").mockrows,
            0)
        self.assertTrue("key2" in firstcol)
        self.assertTrue("key1" in firstcol)
        self.assertTrue("key3" in firstcol)
        self.assertTrue("Col1" in firstcol)

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_sheet_not_exist_and_strict(self):
        action = self.get_action_instance(self.full_config)
        with self.assertRaises(KeyError):
            action.run('sheet3', 'key4', True, "mock_excel.xlsx")
        DeleteRowsTestCase.WB.save.assert_not_called()
        with self.assertRaises(KeyError):
            DeleteRowsTestCase.WB.get_sheet_by_name("sheet4")

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', ExcelBaseActionTestCase.mock_file_exists)
    def test_sheet_not_exist_and_not_strict(self):
        action = self.get_action_instance(self.full_config)
        result = action.run('sheet3', 'key4', False, "mock_excel.xlsx")

        self.assertIsNotNone(result)
        self.assertTrue(result[0])
        self.assertEquals("Success", result[1])
        DeleteRowsTestCase.WB.save.assert_called()
        firstcol = ExcelBaseActionTestCase._util_get_column(
            DeleteRowsTestCase.WB.get_sheet_by_name("sheet1").mockrows,
            0)
        self.assertTrue("key2" in firstcol)
        self.assertTrue("key1" in firstcol)
        self.assertTrue("key3" in firstcol)
        self.assertTrue("Col1" in firstcol)
        self.assertIsNotNone(DeleteRowsTestCase.WB.get_sheet_by_name("sheet3"))
