"""
Copyright 2016 Brocade Communications Systems, Inc.
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at
    http://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""

from lib import excel_action, excel_reader


class GetExcelSheetsAction(excel_action.ExcelAction):
    def run(self, sheet):
        excel = excel_reader.ExcelReader(self._excel_file)  # pylint: disable=no-member
        excel.set_sheet(sheet, key_column=self._key_column,  # pylint: disable=no-member
                        var_name_row=self._var_name_row,  # pylint: disable=no-member
                        strict=True)
        return excel.get_keys()
