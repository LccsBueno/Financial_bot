import importlib
import subprocess

from scrappers import stocksCollector

from scrappers import fiisCollector

from formatters import dataFormatter
from formatters import sheetFormatter

from scrappers.stocksCollector import StocksCollector as stock
import MenuOptions 


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
        
        userInput = MenuOptions.menuOpcoes()
            
        if userInput == 5:
            
            return
        
        elif userInput == 1:
            caminhoConfigurado = input("""
    |Caminho absoluto: """)
                       
def main():

    while not sair:

        userInput = MenuOptions.menuOpcoes()
        
        if userInput == 5:
            MenuOptions.volteSempre()
            break
        
        elif userInput == 1:
            acaoInput = MenuOptions.siglaAtivo()
            comando = "-n" 
            script = ["python", "./invest.py", comando, acaoInput]
            subprocess.call(script)
            exit()
            
        elif userInput == 2:
            stock.scrap()
            
                 
        elif userInput == 3:
            pass
            
        elif userInput == 4:
            configs()
        
main()    

    
    

    
        
