"""
Excel delete row action runner script
"""

from lib import excel_action, excel_reader, string_converter


class DeleteExcelRowAction(excel_action.ExcelAction):
    def run(self, sheet, key, strict, excel_file=None,
            key_column=None, variable_name_row=None):

        self.replace_defaults(excel_file, key_column, variable_name_row)

        excel = excel_reader.ExcelReader(self._excel_file, lock=True)
        excel.set_sheet(sheet, key_column=self._key_column,
                        var_name_row=self._var_name_row,
                        strict=strict)

        key = string_converter.convert_string_to_float_int(key)

        excel.delete_row(key)
        excel.save()

        return (True, 'Success')
