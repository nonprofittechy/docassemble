import ast
import re

INTERVIEW_TRANSLATION_XLSX_COLUMNS = (
    'interview',
    'question_id',
    'index_num',
    'hash',
    'orig_lang',
    'tr_lang',
    'orig_text',
    'tr_text',
)

WORD_TRANSLATION_XLSX_COLUMNS = (
    'orig_lang',
    'tr_lang',
    'orig_text',
    'tr_text',
)

missing_columns_re = re.compile(r"Usecols do not match columns, columns expected but not found: (\[.*?\])")


def read_xlsx_columns(path, expected_columns, **kwargs):
    import pandas  # pylint: disable=import-outside-toplevel
    read_kwargs = {'usecols': list(expected_columns)}
    read_kwargs.update(kwargs)
    return pandas.read_excel(path, **read_kwargs)


def extract_missing_usecols_columns(exc):
    match = missing_columns_re.search(str(exc))
    if not match:
        return None
    return list(ast.literal_eval(match.group(1)))
