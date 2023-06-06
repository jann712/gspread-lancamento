import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from sheets_new import Sheets
# data, relatorio, cliente, consultor, abrangencia, produto, servico, horario, problema, solucao, observacao


class InterfaceEntry(ttk.Frame):

    def __init__(self, master):
        super().__init__(master, padding=(20, 10))
        self.pack(fill=BOTH, expand=YES)
        self.sheet = Sheets("MACRO", "RELATORIO", "Planilha5")
        self.variaveis_box = []


        texto_header = "Por favor insira as informacoes necessarias."
        header = ttk.Label(master=self, text=texto_header, width=50)
        header.pack(fill=X, pady=10)


    def create_form_entry(self, label, variable):
        """Create a single form entry"""
        container = ttk.Frame(self)
        container.pack(fill=X, expand=YES, pady=5)

        lbl = ttk.Label(master=container, text=label.title(), width=15)
        lbl.pack(side=LEFT, padx=5)

        ent = ttk.Entry(master=container, textvariable=variable)
        ent.pack(side=LEFT, padx=5, fill=X, expand=YES)

        return ent


    def create_combobox(self, label, variable):
        """Create a single form entry"""
        container = ttk.Frame(self)
        container.pack(fill=X, expand=YES, pady=5)

        lbl = ttk.Label(master=container, text=label.title(), width=15)
        lbl.pack(side=LEFT, padx=5)

        ent = ttk.Combobox(master=container, values=variable)
        ent.pack(side=LEFT, padx=5, fill=X, expand=YES)

        return ent


    def create_buttonbox(self):
        """Create the application buttonbox"""
        container = ttk.Frame(self)
        container.pack(fill=X, expand=YES, pady=(15, 10))

        sub_btn = ttk.Button(
            master=container,
            text="Submit",
            command=self.on_submit,
            bootstyle=SUCCESS,
            width=6,
        )
        sub_btn.pack(side=RIGHT, padx=5)
        sub_btn.focus_set()

        cnl_btn = ttk.Button(
            master=container,
            text="Cancel",
            command=self.on_cancel,
            bootstyle=DANGER,
            width=6,
        )
        cnl_btn.pack(side=RIGHT, padx=5)

    def create_box_variable(self):
        i = 0
        (self.sheet.cabecalhos)
        for cabecalho in self.sheet.cabecalhos:
            if len(self.sheet.get_col_values(i)) >= 1:
                setattr(self, f"entry_box{i}",
                        self.create_combobox(cabecalho,
                        self.sheet.get_col_values(i)))

            elif len(self.sheet.get_col_values(i)) == 0:
                setattr(self, f"entry_box{i}",
                        self.create_form_entry(cabecalho, ttk.StringVar(value="")))

            self.variaveis_box.append(getattr(self, f"entry_box{i}"))

            i += 1
        self.create_buttonbox()


    def on_submit(self):
        valores_lancamento = [attr.get() for attr in self.variaveis_box]
        self.sheet.relatorio_sheet.append_row(valores_lancamento)


    def on_cancel(self):
        """Cancel and close the application."""
        self.quit()


if __name__ == "__main__":
    app = ttk.Window("Relatorio", "lumen", resizable=(True, True))
    test = InterfaceEntry(app)
    test.create_box_variable()
    app.mainloop()

