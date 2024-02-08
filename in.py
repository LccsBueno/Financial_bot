import importlib
import subprocess

import scrappers
from scrappers import stocksCollector

import scrappers
from scrappers import fiisCollector

import scrappers
from scrappers import stocksShower


sair = False

libsUsadas = ["selenium", "argparse", "xlsxwriter"]

def checarBiblitecas(bibliotecas):
    for lib in bibliotecas:
        try:
            importlib.import_module(lib)
        except ImportError:
            print(f"Instalando {lib} ...")
            try:
                subprocess.check_call(["pip", "install", lib])
                print(f"{lib} instalado com sucesso!!")
            except subprocess.CalledProcessError:
                print(f"Não foi possível instalar {lib}, por favor instale manualmente")
                return
            
checarBiblitecas(libsUsadas)

def configs():
    
    while True:
        
        userInput = int(input("""
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
            
        if userInput == 5:
            
            return
        
        elif userInput == 1:
            caminhoConfigurado = input("""
    |Caminho absoluto: """)
                       
def main():

    while not sair:

        userInput = int(input("""
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
        
        if userInput == 5:
            print("""
    *-------------------------*
    |      Volte sempre!      |
    *-------------------------*
        """)
            break
        
        elif userInput == 1:
            acaoInput = input("""
    |Sigla do ativo: """)
            comando = "-n" 
            script = ["python", "./invest.py", comando, acaoInput]
            subprocess.call(script)
            exit()
            
        elif userInput == 2:
            stocksCollector.StocksCollector()
                 
        elif userInput == 3:
            script = ["python", "./invest.py", "-fiis"]
            subprocess.call(script)
            
        elif userInput == 4:
            configs()
        
main()    

    
    

    
        
