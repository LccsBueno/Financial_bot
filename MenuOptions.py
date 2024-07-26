
class MenuOptions:
    
    @staticmethod
    def menuPrincipal():
        return int(input("""
    *-------------------------*
    |       I N V E S T       |
    *-------------------------*
    | 1-Procurar ação         |
    | 2-Baixar todas ações    |
    | 3-Baixar todos FIIs     |
    | 4-Config                |
    | 5-Sair                  |
    *-------------------------*
    |Escolha uma: """))
    
    @staticmethod
    def configuracoes():
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
    def opcoesFiis():
        return int(input("""
    *-------------------------*
    |  CONFIGURAÇÕES - FIIS   |
    *-------------------------*
    | 1-Exibir Colunas        |
    | 2-Adicionar Coluna      |                        
    | 3-Editar Coluna         |
    | 4-Prosseguir            |
    *-------------------------*
    |Escolha uma: """))
        

    @staticmethod
    def siglaAtivo():
        return input("""
    |Sigla do ativo: """)