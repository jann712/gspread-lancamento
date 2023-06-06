import gspread

# gc = gspread.service_account(filename='resonant-feat-386212-de29c2104176.json')
# sheet = gc.open("MACRO").worksheet("RELATORIO")
# data, relatorio, cliente, consultor, abrangencia, produto, servico, horario, problema, solucao, observacao

class Sheets():
    def __init__(self, spreadsheet_name, relatorio_sheet_name, valores_sheet_name):
        self.gsa = gspread.service_account(filename='')
        self.spreadsheet = self.gsa.open(spreadsheet_name)
        self.relatorio_sheet = self.spreadsheet.worksheet(relatorio_sheet_name)
        self.valores_sheet = self.spreadsheet.worksheet(valores_sheet_name)
        self.todos_valores = self.valores_sheet.get_all_values()
        self.valores = self.get_somente_valores()
        self.cabecalhos = self.get_headers()



    # Retorna todos os valores da aba atual menos o cabecalho
    def get_somente_valores(self):
        return self.todos_valores[1:]

    # Retorna somente o cabecalho da planilha atual
    def get_headers(self):
        return self.todos_valores[0]


    # Retorna os valores de determinada coluna
    def get_col_values(self, col):
        return [lista[col] for lista in self.valores
         if lista[col] != ""]

