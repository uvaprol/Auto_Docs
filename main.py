import customtkinter
import Docs_maker as DM


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
            docs_name = DM.parse_docssheet(seting[0])
        except:
            docs_name = []
        try:
            name = DM.parse_excel_to_dict_list(seting[1], seting[2])
        except:
            name = {}


        app.state('normal')
        app.focus()
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



        # add widgets to app
        self.Name_frame_label = customtkinter.CTkLabel(self, text="Выберите записи", fg_color="transparent")
        self.Name_frame_label.grid(row=0, column=0, padx=20, pady=20)
        self.Name_frame = MyFrame(master=self, width=700, height=400)
        self.Name_frame.grid(row=1, column=0, padx=20, pady=20)

        self.Docs_frame_label = customtkinter.CTkLabel(self, text="Выберите документы", fg_color="transparent")
        self.Docs_frame_label.grid(row=0, column=1, padx=20, pady=20)
        self.Docs_frame = MyFrame(master=self, width=700, height=400)
        self.Docs_frame.grid(row=1, column=1, padx=20, pady=20)

        self.name_checkboxes = []
        self.row_index = 0
        for i in name.keys():
            self.name_checkboxes.append(customtkinter.CTkCheckBox(self.Name_frame, text=f'{i}',
                                                             variable=customtkinter.StringVar(value="off"),
                                                             onvalue="on", offvalue="off"))
            self.name_checkboxes[self.row_index].pack(padx=20, pady=20, anchor="w")
            self.row_index += 1

        self.docs_checkboxes = []
        self.row_index = 0

        for i in docs_name:
            self.docs_checkboxes.append(
                customtkinter.CTkCheckBox(self.Docs_frame, text=f'{i.name}',
                                          variable=customtkinter.StringVar(value="off"),
                                          onvalue="on", offvalue="off"))
            self.docs_checkboxes[self.row_index].pack(padx=20, pady=20, anchor="w")
            self.row_index += 1


        self.create_button = customtkinter.CTkButton(self, text="Создать", command=self.create_button_event)
        self.create_button.grid(row=2, column=1, padx=20, pady=20)

        self.seting_button = customtkinter.CTkButton(self, text="Настройки", command=self.open_setings)
        self.seting_button.grid(row=2, column=0, padx=20, pady=20)

    # add methods to app
    def create_button_event(self):
        global name
        for d in self.docs_checkboxes:
            if d._variable.get() == 'on':
                for n in self.name_checkboxes:
                    if n._variable.get() == 'on':
                        print(seting)
                        DM.format_and_make(d.cget('text'), seting[3], name[n.cget('text')])
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



with open('SETINGS.txt', 'r+') as f:
    seting = f.read()
    if len(seting) > 0:
        seting = seting.split('|')
    else:
        f.write('.|РЕЕСТР ДЛЯ ОТЧЕТА.xlsx|Профессиональная переподготовка|')
        seting = ('.|РЕЕСТР ДЛЯ ОТЧЕТА.xlsx|Профессиональная переподготовка|').split('|')

try:
    docs_name = DM.parse_docssheet(seting[0])
except:
    docs_name = []

try:
    name = DM.parse_excel_to_dict_list(seting[1], seting[2])
except:
    name = {}

app = App()
app.mainloop()
