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

from st2tests.base import BaseActionTestCase
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from openpyxl import Workbook


class ExcelBaseActionTestCase(BaseActionTestCase):
    __test__ = False

    class MockSheet(Worksheet):
        def __init__(self, rows, **kwargs):
            ''' Mocks worksheet as an array of rows, each row is an array
                of mock cells taken from value in rows array '''
            super(ExcelBaseActionTestCase.MockSheet, self).__init__(**kwargs)
            self.mockrows = []
            for i in range(len(rows)-1):
                newrow = []
                for j in range(len(rows[i])-1):
                    newrow.append(mock.Mock(Cell(self,
                                  row=(i-1), column=(j-1),
                                  value=rows[i][j])))
                self.mockrows.append(newrow)

        def delete_rows(idx, amount=1):
            for i in range(idx - 1, idx + amount - 1):
                self.mockrows.pop(i)

        def cell(row, column):
            rowfound = self.mockrows[row-1]
            if not rowfound:
                self.mockrows[row-1] = []
                rowfound = self.mockrows[row-1]
            cellfound = rowfound[column-1]
            if not cellfound:
                cellfound = mock.Mock(Cell(
                                 row=(row-1), column=(column-1), value=None))
                rowfound[column-1] = cellfound
            return cellfound

    class MockWorkbook(object):
        def __init__(self, worksheets, new_sheet, **kwargs):
            """ Takes in dictionary of sheets """
            #super(ExcelBaseActionTestCase.MockWorkbook, self).__init__(**kwargs)
            # Convert dictionary of sheetname -> array of values to
            # dictionary of sheetname -> MockSheet
            self._mocksheets = {}
            for sheetname in worksheets.keys():
                self._mocksheets[sheetname] = ExcelBaseActionTestCase.MockSheet(worksheets[sheetname], parent=self)

        @property
        def sheetnames(self):
             return self._mocksheets.keys()

        def get_sheet_by_name(sheet_name):
            return self._mocksheets[sheet_name]


        def save():
            pass

        def create_sheet(sheet_name):
            return self.new_sheet

    def setUp(self):
        super(ExcelBaseActionTestCase, self).setUp()
