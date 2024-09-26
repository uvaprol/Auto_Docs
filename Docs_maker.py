import os
import docx
import pandas as pd


def parse_docssheet(path:str = '.') -> list:
    docs_sheet = []
    with os.scandir(path) as it:
        for entry in it:
            if not entry.name.startswith('.') and entry.is_file() and entry.name[:7] == 'Шаблон_' and entry.name[-5:] == '.docx':
                docs_sheet.append(entry)
    return docs_sheet

def parse_excel_to_dict_list(filepath: str, sheet_name='Лист1') -> tuple:
    df = pd.read_excel(filepath, sheet_name=sheet_name)
    dict_list = df.to_dict(orient='records')
    names = {}
    for i in dict_list:
        f = i['ФИО обучающегося']
        if type(i['ФИО обучающегося']) is str:
            if names.get('ФИО обучающегося') == None:
                names[f] = {}
            i['Серия']             = i['Паспорт']
            i['Номер']             = i['Unnamed: 17']
            i['Дата выдачи']       = i['Unnamed: 18']
            i['Кем выдан']         = i['Unnamed: 19']
            i['Код подразделения'] = i['Unnamed: 20']
            i['Адрес регистрации'] = i['Unnamed: 21']
            del i['№ '], i['Паспорт'], i['Unnamed: 17'], i['Unnamed: 18'], i['Unnamed: 19'], i['Unnamed: 20'], i['Unnamed: 21']
            names[f] = i
    return names


def format_and_make(doc: str, change_text: dict = {}) -> None:
    print('Run')
    d = docx.Document(doc)
    for table in d.tables:
        for row in table.rows:
            for cell in row.cells:
                for mark in change_text.keys():
                    cell.text = cell.text.replace(f'<{mark}>', f'{change_text[mark]}')

    for i in d.paragraphs:
        for mark in change_text.keys():
            i.text = i.text.replace(f'<{mark}>', f'{change_text[mark]}')
    d.save(f'{change_text["ФИО обучающегося"]} {doc[7:]}')
    print("Done")

