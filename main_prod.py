import os, docx, customtkinter, pandas as pd
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
            i['Дата рождения'] = '.'.join(str(i['Дата рождения'])[:10].split('-')[-1::-1])
            i['Дата договора'] = '.'.join(str(i['Дата договора'])[:10].split('-')[-1::-1])
            i['Дата выдачи'] = '.'.join(str(i['Дата выдачи'])[:10].split('-')[-1::-1])
            del i['№ '], i['Паспорт'], i['Unnamed: 17'], i['Unnamed: 18'], i['Unnamed: 19'], i['Unnamed: 20'], i['Unnamed: 21']
            names[f] = i
    return names
def format_and_make(doc: str, save_path: str, change_text: dict = {}) -> None:
    print('Run')
    d = docx.Document(doc)
    styles = []
    for paragraph in d.paragraphs:
        styles.append(paragraph.style)
    for table in d.tables:
        for row in table.rows:
            for cell in row.cells:
                for mark in change_text.keys():
                    if mark in cell.text:
                        for paragraph in cell.paragraphs:
                            if '<' in paragraph.text and '>' in paragraph.text:
                                key = False
                                for run in paragraph.runs:
                                    if run.text == '<':
                                        run.text = run.text.replace('<', f'{change_text[mark]}')
                                        key = True
                                        continue
                                    elif key and run.text.find('>') == -1:
                                        run.text = ''
                                    elif run.text.find('>') != -1:
                                        run.text = run.text[:run.text.find('>')]
                                        key = False
    for paragraph in d.paragraphs:
        for mark in change_text.keys():
            if mark in paragraph.text:
                key = False
                for run in paragraph.runs:
                    if run.text == '<':
                        run.text = run.text.replace('<', f'{change_text[mark]}')
                        key = True
                        continue
                    elif key and run.text.find('>') == -1:
                        run.text = ''
                    elif run.text.find('>') != -1:
                        run.text = run.text[:run.text.find('>')]
                        key = False
    d.save(f'{save_path}{change_text["ФИО обучающегося"]} {doc[7:]}')
    print("Done")
class SelectWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
class SetingsWindow(customtkinter.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("400x350")
        self.focus()
        self.title('Настройки')
        self.label = customtkinter.CTkLabel(self, text="Путь к шаблонам")
        self.label.grid(row=0, column=0, padx=20, pady=20)
        self.out_entry = customtkinter.CTkEntry(self, placeholder_text="По умолчанию")
        self.out_entry.grid(row=0, column=1, padx=20, pady=20)
        self.label = customtkinter.CTkLabel(self, text="Путь до Exel")
        self.label.grid(row=1, column=0, padx=20, pady=20)
        self.exel_entry = customtkinter.CTkEntry(self, placeholder_text="По умолчанию")
        self.exel_entry.grid(row=1, column=1, padx=20, pady=20)
        self.label = customtkinter.CTkLabel(self, text="Какой лист использовать")
        self.label.grid(row=2, column=0, padx=20, pady=20)
        self.sheet_entry = customtkinter.CTkEntry(self, placeholder_text="По умолчанию")
        self.sheet_entry.grid(row=2, column=1, padx=20, pady=20)
        self.label = customtkinter.CTkLabel(self, text="Куда сохранять")
        self.label.grid(row=3, column=0, padx=20, pady=20)
        self.in_entry = customtkinter.CTkEntry(self, placeholder_text="По умолчанию")
        self.in_entry.grid(row=3, column=1, padx=20, pady=20)
        self.button = customtkinter.CTkButton(self, text="Принять настройки", command=self.set_setings)
        self.button.grid(row=4, column=1, padx=20, pady=20)
        self.button = customtkinter.CTkButton(self, text="Сбросить настройки", command=self.reset_setings)
        self.button.grid(row=4, column=0, padx=20, pady=20)
        if seting[0] != '' and seting[0] != '.':
            self.out_entry.insert(0, seting[0])
        if seting[1] != '' and seting[1] != 'РЕЕСТР ДЛЯ ОТЧЕТА.xlsx':
            self.exel_entry.insert(0, seting[1])
        if seting[2] != '' and seting[2] != 'Профессиональная переподготовка':
            self.sheet_entry.insert(0, seting[2])
        if seting[3] != '':
            self.in_entry.insert(0, seting[3])
    def set_setings(self):
        global app, seting
        with open('SETINGS.txt', 'w') as f:
            if self.out_entry.get() == '':
                self.out_entry.insert(0, seting[0])
            if self.exel_entry.get() == '':
                self.exel_entry.insert(0, seting[1])
            if self.sheet_entry.get() == '':
                self.sheet_entry.insert(0, seting[2])
            if self.in_entry.get() == '':
                self.in_entry.insert(0, seting[3])
            f.write(f'{self.out_entry.get()}|{self.exel_entry.get()}|{self.sheet_entry.get()}|{self.in_entry.get()}')
            seting = [self.out_entry.get(), self.exel_entry.get(), self.sheet_entry.get(), self.in_entry.get()]
            self.reload_data()
    def reset_setings(self):
        global app, seting
        with open('SETINGS.txt', 'w') as f:
            f.write('.|РЕЕСТР ДЛЯ ОТЧЕТА.xlsx|Профессиональная переподготовка|')
            seting = ('.|РЕЕСТР ДЛЯ ОТЧЕТА.xlsx|Профессиональная переподготовка|').split('|')
            self.reload_data()
    def reload_data(self):
        global docs_name, name, app
        try:
            docs_name = parse_docssheet(seting[0])
        except:
            docs_name = []
        try:
            name = parse_excel_to_dict_list(seting[1], seting[2])
        except:
            name = {}
        app.state('normal')
        app.update_frames()
        self.destroy()
class MyFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        global name, docs_name
        self.geometry("1520x600")
        self.title("Auto docx")
        customtkinter.set_appearance_mode("dark")
        self.Name_frame_label = customtkinter.CTkLabel(self, text="Выберите записи", fg_color="transparent")
        self.Name_frame_label.grid(row=0, column=0, padx=20, pady=20)
        self.Name_frame = MyFrame(master=self, width=700, height=400)
        self.Name_frame.grid(row=1, column=0, padx=20, pady=20)
        self.Docs_frame_label = customtkinter.CTkLabel(self, text="Выберите документы", fg_color="transparent")
        self.Docs_frame_label.grid(row=0, column=1, padx=20, pady=20)
        self.Docs_frame = MyFrame(master=self, width=700, height=400)
        self.Docs_frame.grid(row=1, column=1, padx=20, pady=20)
        self.name_checkboxes = []
        self.docs_checkboxes = []
        self.select_all_name_button = customtkinter.CTkButton(self.Name_frame, text="Выбрать всё", command=self.select_all_name)
        self.select_all_name_button.pack(padx=20, pady=20, anchor="n")
        self.select_all_docs_button = customtkinter.CTkButton(self.Docs_frame, text="Выбрать всё", command=self.select_all_docs)
        self.select_all_docs_button.pack(padx=20, pady=20, anchor="n")
        self.update_frames()
        self.create_button = customtkinter.CTkButton(self, text="Создать", command=self.create_button_event)
        self.create_button.grid(row=2, column=1, padx=20, pady=20)
        self.seting_button = customtkinter.CTkButton(self, text="Настройки", command=self.open_setings)
        self.seting_button.grid(row=2, column=0, padx=20, pady=20)
    def select_all_name(self):
        if self.select_all_name_button.cget('text') == 'Выбрать всё':
            self.select_all_name_button.configure(text="Отменить выбор")
        else:
            self.select_all_name_button.configure(text="Выбрать всё")
        for n in self.name_checkboxes:
            if n._variable.get() == 'on' and self.select_all_name_button.cget('text') == 'Выбрать всё':
                n._variable.set('off')
            elif self.select_all_name_button.cget('text') != 'Выбрать всё':
                n._variable.set('on')
    def select_all_docs(self):
        if self.select_all_docs_button.cget('text') == 'Выбрать всё':
            self.select_all_docs_button.configure(text="Отменить выбор")
        else:
            self.select_all_docs_button.configure(text="Выбрать всё")
        for d in self.docs_checkboxes:
            if d._variable.get() == 'on' and self.select_all_docs_button.cget('text') == 'Выбрать всё':
                d._variable.set('off')
            elif self.select_all_docs_button.cget('text') != 'Выбрать всё':
                d._variable.set('on')
    def create_button_event(self):
        global name
        for d in self.docs_checkboxes:
            if d._variable.get() == 'on':
                for n in self.name_checkboxes:
                    if n._variable.get() == 'on':
                        print(seting)
                        format_and_make(d.cget('text'), seting[3], name[n.cget('text')])
        for d in self.docs_checkboxes:
            d._variable.set('off')
        for n in self.name_checkboxes:
            n._variable.set('off')
    def open_setings(self):
        print(seting)
        try:
            self.seting_window.focus()
        except:
            self.seting_window = SetingsWindow(self)
            self.withdraw()
    def update_frames(self):
        self.focus()
        for n in self.name_checkboxes:
            n.destroy()
        for d in self.docs_checkboxes:
            d.destroy()
        self.name_checkboxes = []
        self.docs_checkboxes = []
        self.row_index = 0
        for i in name.keys():
            self.name_checkboxes.append(customtkinter.CTkCheckBox(self.Name_frame, text=f'{i}', variable=customtkinter.StringVar(value="off"), onvalue="on", offvalue="off"))
            self.name_checkboxes[self.row_index].pack(padx=20, pady=20, anchor="w")
            self.row_index += 1
        self.row_index = 0
        for i in docs_name:
            self.docs_checkboxes.append(
                customtkinter.CTkCheckBox(self.Docs_frame, text=f'{i.name}', variable=customtkinter.StringVar(value="off"), onvalue="on", offvalue="off"))
            self.docs_checkboxes[self.row_index].pack(padx=20, pady=20, anchor="w")
            self.row_index += 1
with open('SETINGS.txt', 'r+') as f:
    seting = f.read()
    if len(seting) > 0:
        seting = seting.split('|')
    else:
        f.write('.|РЕЕСТР ДЛЯ ОТЧЕТА.xlsx|Профессиональная переподготовка|')
        seting = ('.|РЕЕСТР ДЛЯ ОТЧЕТА.xlsx|Профессиональная переподготовка|').split('|')
try:
    docs_name = parse_docssheet(seting[0])
except:
    docs_name = []
try:
    name = parse_excel_to_dict_list(seting[1], seting[2])
except:
    name = {}
app = App()
app.mainloop()