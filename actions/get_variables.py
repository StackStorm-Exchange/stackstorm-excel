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
import json


class GetExcelVariablesAction(excel_action.ExcelAction):
    def run(self, key, sheet='Sheet', variables='[]', excel_file=None,
            key_column=None, variable_name_row=None):

        self.replace_defaults(excel_file, key_column, variable_name_row)

        excel = excel_reader.ExcelReader(self._excel_file)
        excel.set_sheet(sheet, key_column=self._key_column,
                        var_name_row=self._var_name_row,
                        strict=True)

        vfk = excel.get_variables_for_key(key)
        if variables == '[]':  # default
            return vfk

        variables = json.loads(variables)
        filtered = {}
        for var_name in variables:
            if var_name in vfk:
                filtered[var_name] = vfk[var_name]

        return filtered
