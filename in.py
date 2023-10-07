import importlib
import subprocess

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
            

def configs():
    input("""
*-------------------------*
|      CONFIGURAÇÕES      |
*-------------------------*
| 1-Caminho da planilha   |
| 2-                      |
| 3-                      |
| 4-                      |
| 5-Sair                  |
*-------------------------*
|Escolha uma: """)
        
    
checarBiblitecas(libsUsadas)

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
        break
    elif userInput == 2:
        comando = "-acoes"
    elif userInput == 3:
        comando = "-fiis"
    elif userInput == 4:
        comando = "./invest.py"
    
    script = ["python", "./invest.py", comando]
    subprocess.call(script)
    
    

    
    

    
        
