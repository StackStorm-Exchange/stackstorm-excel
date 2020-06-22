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

from delete_row import DeleteExcelRowAction


class DeleteRowsTestCase(ExcelBaseActionTestCase):
    __test__ = True
    action_cls = DeleteExcelRowAction

    SHEET_1 = [ [ "Col1", "Col2" ], [ "key1", "ro1_2" ], [ "key2", "ro2_2" ] ]
    SHEET_2 = [ [ "Col1", "Col2" ] ]
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

    def mock_is_file(filename):
        if "lock" in filename:
            return False
        else:
            return True

    def return_workbook(filename, data_only):
      DeleteRowsTestCase.WB =  ExcelBaseActionTestCase.MockWorkbook(DeleteRowsTestCase._MOCK_SHEETS, None)
      DeleteRowsTestCase.WB.save = mock.MagicMock()
      return DeleteRowsTestCase.WB

    def _get_column(rows, column):
        keys = []
        for i in range(len(rows)):
            keys.append(rows[i][column].value)
        return keys

    #@mock.patch('openpyxl.load_workbook',

    @mock.patch('openpyxl.load_workbook', return_workbook)
    @mock.patch('os.path.isfile', mock_is_file)
    def test_delete_row_exists(self):
        action = self.get_action_instance(self.full_config)
        result = action.run('sheet1', 'key1', True, "mock_excel.xlsx")

        self.assertIsNotNone(result)
        self.assertTrue(result[0])
        self.assertEquals("Success", result[1])
        DeleteRowsTestCase.WB.save.assert_called()
    #    self.assertEqual(2, len(DeleteRowsTestCase.WB.get_sheet_by_name("sheet1").mockrows))
        keys = DeleteRowsTestCase._get_column(DeleteRowsTestCase.WB.get_sheet_by_name("sheet1").mockrows, 0)
        self.assertTrue("key1" not in keys)
        self.assertTrue("key2" in keys)
        self.assertTrue("Col1" in keys)

        
