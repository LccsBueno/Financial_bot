import argparse

import subprocess
import xlsxwriter 

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument('headless')

navegador = webdriver.Chrome(options=options)

parser = argparse.ArgumentParser(
    description="Procura o preço da ação hoje no mercado"    
)

acao_name = parser.add_argument(
    '-n',
    type=str,
    help="Nome da ação",
)

parser.add_argument(
    '-p',
    type=str,
    help="Pesquisa uma empresa",
)

parser.add_argument(
    '-acoes',
    help="Baixa todas as ações da bolsa"
)

parser.add_argument(
    '-fiis',
    help="Baixa todos os fiis da bolsa"
)

args = parser.parse_args()

if args.n:

    acao_sigla = args.n

    navegador.get(f"https://www.fundamentus.com.br/detalhes.php?papel={acao_sigla}")

    print(f'\nprocurando {acao_sigla} ...')
    
    try:
        web_acao_valor = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, ('/html/body/div[1]/div[2]/table[1]/tbody/tr[1]/td[4]/span') )))
        web_acao_sigla = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, ('/html/body/div[1]/div[2]/table[1]/tbody/tr[1]/td[2]/span'))))
        web_acao_nome = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, ('/html/body/div[1]/div[2]/table[1]/tbody/tr[3]/td[2]/span') )))
        web_acao_dividendYield = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, ('/html/body/div[1]/div[2]/table[3]/tbody/tr[9]/td[4]/span') )))
        web_acao_pl = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, ('/html/body/div[1]/div[2]/table[3]/tbody/tr[2]/td[4]/span') )))
        web_acao_pvp = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, ('/html/body/div[1]/div[2]/table[3]/tbody/tr[3]/td[4]/span') )))
        web_acao_roe = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, ('/html/body/div[1]/div[2]/table[3]/tbody/tr[9]/td[6]/span') )))
        web_acao_roic = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, ('/html/body/div[1]/div[2]/table[3]/tbody/tr[8]/td[6]/span') )))
        web_acao_ebit = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, ('/html/body/div[1]/div[2]/table[5]/tbody/tr[4]/td[2]/span') )))
        web_acao_lucroliquido = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, ('/html/body/div[1]/div[2]/table[5]/tbody/tr[5]/td[4]/span') )))

        print(f'\n Ação: {web_acao_sigla.text} \n Preço: R$ {web_acao_valor.text} \n Empresa: {web_acao_nome.text} \n Dividend Yield: {web_acao_dividendYield.text} \n P/L: {web_acao_pl.text} \n P/VP: {web_acao_pvp.text} \n ROE: {web_acao_roe.text} \n ROIC: {web_acao_roic.text} \n EBIT: R$ {web_acao_ebit.text} \n Lucro Líquido: R$ {web_acao_lucroliquido.text} \n')

    except:
        print("\nNada foi encontrado")        

elif args.p:
    print("funcionou patrão!!")

elif args.acoes:
    print("\n Baixando as informações ... \n")
    
    navegador.get(f"https://www.fundamentus.com.br/resultado.php")
    
    tabela = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, ('//*[@id="resultado"]/tbody') )))
    
    row = 0
    column = 0
    array_auxiliar = []
    
    no_column = [4, 6, 7, 8, 9, 11, 12, 13, 14, 18]
    
    table_header = ["Papel", "Cotação", "P/L", "P/VP", "PSR", "Div.Yield", "P/Ativo", "P/Cap.Giro", "P/EBIT", "P/Ativ Circ.Liq", "EV/EBIT", "EV/BIT", "Mrg EBIT", "Mrg Liq", "Liq Corr", "ROIC", "ROE", "Liq 2 meses", "Patrim. Liq", "Div Brut/Patrim", "Cresc Rec.5a"]
    
    workbook =xlsxwriter.Workbook('C:/Users/lucca/OneDrive - SPTech School/Documents/Planilhas/Investimento/Investimento_acoes.xlsx')
    worksheet = workbook.add_worksheet("Todas Ações")
    
    array_tabela = tabela.text.split("\n")
    
    for cada_acao in array_tabela:
        array_auxiliar.append(cada_acao.split(" "))
        
    array_tabela = array_auxiliar
    array_tabela.insert(0, table_header)
    
    for table_row in array_tabela:
        table_row.pop(4)
        table_row.pop(5)
        table_row.pop(5)
        table_row.pop(5)
        table_row.pop(5)
        table_row.pop(6)
        table_row.pop(6)
        table_row.pop(6)
        table_row.pop(6)
        table_row.pop(9)
            
    for linha_tabela in array_tabela:    
        
        for coluna_tabela in linha_tabela:
            
            if column == 11:
                
                row+= 1
                column = 0   
                        
            try: 
                
                worksheet.write_number(row, column, float(coluna_tabela))
                column+= 1
                
            except: 
                
                worksheet.write(row, column, str(coluna_tabela))
                column+= 1

    # worksheet.filter_column('B:B', 'x >= 5 and x <= 20')
    
    workbook.close()
    
    excel_path = "C:/Program Files/Microsoft Office/root\Office16/EXCEL.EXE"
    subprocess.run([excel_path, "C:/Users/lucca/OneDrive - SPTech School/Documents/Planilhas/Investimento/Investimento_acoes.xlsx"])

elif args.fiis: 
    print("\n Baixando as informações ... \n")
    
    navegador.get(f"https://www.fundamentus.com.br/fii_resultado.php")
     
    workbook =xlsxwriter.Workbook('C:/Users/lucca/OneDrive - SPTech School/Documents/Planilhas/Investimento/Investimento_fiis.xlsx')
    worksheet = workbook.add_worksheet("Todos FIIS")
     
    cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
    
    table_data = navegador.execute_script("""
    var table = document.getElementById('tabelaResultado'); 
    var data = [];
    for (var i = 0, row; row = table.rows[i]; i++) {
        var rowData = [];
 
            for (var j = 0, cell; j < 14; j++) {

                cell = row.cells[j]
                rowData.push(cell.textContent);
            }
        
        data.push(rowData);
    }
    return data;""")

    qtd_column = 1
    qtd_row = 1

    for table_row in table_data:
        
        for table_data in table_row:
            
            if qtd_column == 14:
                qtd_column = 1
                qtd_row+=1
                
            else:
                
                if qtd_column == 3 and qtd_row >= 2:
                    
                    formatoDinheiro = workbook.add_format({'num_format': '$#,##'})
                                
                    # worksheet.write(qtd_row, qtd_column, cell_format)
                    worksheet.write_number(3, qtd_column, table_data, formatoDinheiro)    
                    qtd_column+=1
                    
                    
                else:
                    worksheet.write(qtd_row, qtd_column, str(table_data))    
                    # worksheet.write(qtd_row, qtd_column, cell_format)
                    qtd_column+=1    
                        
                
        
  
    workbook.close()
        
    excel_path = "C:/Program Files/Microsoft Office/root\Office16/EXCEL.EXE"
    subprocess.run([excel_path, "C:/Users/lucca/OneDrive - SPTech School/Documents/Planilhas/Investimento/Investimento_fiis.xlsx"])
