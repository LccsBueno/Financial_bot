
class MenuOptions:
    
    @staticmethod
    def menuOpcoes():
        return int(input("""
    *-------------------------*
    |      CONFIGURAÇÕES      |
    *-------------------------*
    | 1-Caminho da planilha   |
    |                         |
    |                         |
    |                         |
    | 5-Voltar                |
    *-------------------------*
    |Escolha uma: """))

    @staticmethod
    def volteSempre():
        return print("""
    *-------------------------*
    |      Volte sempre!      |
    *-------------------------*
        """)

    @staticmethod
    def siglaAtivo():
        return input("""
    |Sigla do ativo: """)