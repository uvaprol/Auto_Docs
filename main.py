import customtkinter
import Docs_maker as DM

class MyFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("1600x1000")
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
        self.name = DM.parse_excel_to_dict_list('РЕЕСТР ДЛЯ ОТЧЕТА.xlsx', 'Профессиональная переподготовка')
        for i in self.name.keys():
            self.name_checkboxes.append(customtkinter.CTkCheckBox(self.Name_frame, text=f'{i}',
                                                             variable=customtkinter.StringVar(value="off"),
                                                             onvalue="on", offvalue="off"))
            self.name_checkboxes[self.row_index].pack(padx=20, pady=20, anchor="w")
            self.row_index += 1

        self.docs_checkboxes = []
        self.row_index = 0
        self.docs_name = DM.parse_docssheet()
        for i in self.docs_name:
            self.docs_checkboxes.append(
                customtkinter.CTkCheckBox(self.Docs_frame, text=f'{i.name}',
                                          variable=customtkinter.StringVar(value="off"),
                                          onvalue="on", offvalue="off"))
            self.docs_checkboxes[self.row_index].pack(padx=20, pady=20, anchor="w")
            self.row_index += 1


        self.create_button = customtkinter.CTkButton(self, text="Создать", command=self.button_event)
        self.create_button.grid(row=2, column=1, padx=20, pady=20)
    # add methods to app
    def button_event(self):
        for d in self.docs_checkboxes:
            if d._variable.get() == 'on':
                for n in self.name_checkboxes:
                    if n._variable.get() == 'on':
                        DM.format_and_make(d.cget('text'), self.name[n.cget('text')])
        for d in self.docs_checkboxes:
            d._variable.set('off')
        for n in self.name_checkboxes:
            n._variable.set('off')





app = App()
app.mainloop()
