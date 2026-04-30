import os
import tempfile
import unittest

import pandas
import xlsxwriter

from docassemble.base.translation_xlsx import (
    INTERVIEW_TRANSLATION_XLSX_COLUMNS,
    WORD_TRANSLATION_XLSX_COLUMNS,
    read_xlsx_columns,
)


class TestTranslationXlsx(unittest.TestCase):

    def test_read_xlsx_columns_ignores_stray_far_right_header_cell(self):
        fd, path = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        try:
            workbook = xlsxwriter.Workbook(path)
            worksheet = workbook.add_worksheet()
            for column_number, header in enumerate(INTERVIEW_TRANSLATION_XLSX_COLUMNS):
                worksheet.write(0, column_number, header)
            worksheet.write(1, 0, 'docassemble.demo:data/questions/example.yml')
            worksheet.write(1, 1, 'q1')
            worksheet.write(1, 2, 0)
            worksheet.write(1, 3, 'abc123')
            worksheet.write(1, 4, 'en')
            worksheet.write(1, 5, 'es')
            worksheet.write(1, 6, 'Hello')
            worksheet.write(1, 7, 'Hola')
            worksheet.write(0, 16383, ' ')
            workbook.close()

            self.assertEqual(pandas.read_excel(path).shape[1], 16384)

            df = read_xlsx_columns(path, INTERVIEW_TRANSLATION_XLSX_COLUMNS)

            self.assertEqual(df.shape, (1, 8))
            self.assertEqual(list(df.columns), list(INTERVIEW_TRANSLATION_XLSX_COLUMNS))
        finally:
            os.remove(path)

    def test_read_xlsx_columns_supports_word_translation_layout(self):
        fd, path = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        try:
            workbook = xlsxwriter.Workbook(path)
            worksheet = workbook.add_worksheet()
            for column_number, header in enumerate(WORD_TRANSLATION_XLSX_COLUMNS):
                worksheet.write(0, column_number, header)
            worksheet.write(1, 0, 'en')
            worksheet.write(1, 1, 'es')
            worksheet.write(1, 2, 'Hello')
            worksheet.write(1, 3, 'Hola')
            worksheet.write(0, 16383, ' ')
            workbook.close()

            df = read_xlsx_columns(path, WORD_TRANSLATION_XLSX_COLUMNS)

            self.assertEqual(df.shape, (1, 4))
            self.assertEqual(list(df.columns), list(WORD_TRANSLATION_XLSX_COLUMNS))
        finally:
            os.remove(path)


if __name__ == '__main__':
    unittest.main()
