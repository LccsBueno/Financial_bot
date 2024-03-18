import argparse

import subprocess
import xlsxwriter 
import logging

import selenium
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

selenium_logger = logging.getLogger('selenium')
selenium_logger.setLevel(logging.WARNING)

caminhoDefault = "."
excel_path = r"C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE"

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
    action="store_true",
    help="Baixa todas as ações da bolsa"
)

parser.add_argument(
    '-fiis',
    action="store_true",
    help="Baixa todos os fiis da bolsa"
)


args = parser.parse_args()

def formatarNumeros(number, type):
    
    if type == "float":
        number = number.replace(".", "")
        number = round(float(number.replace(",", ".")), 2)
        return number
    
    elif type == "percentage":
        number = number.replace("%", "")
        number = number.replace(".", "")
        number = round(float(number.replace(",", ".")), 2) / 100
        return number
    
    elif type == "integer":
        number = int(number.replace(".", ""))
        return number
        
if args.n:

    acao_sigla = args.n

    navegador.get(f"https://www.fundamentus.com.br/detalhes.php?papel={acao_sigla}")

    print(f'\n    |Procurando {acao_sigla} ...')
    
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

        print(f"""
    |Sigla: {web_acao_sigla.text}
    |Preço: R$ {web_acao_valor.text} 
    |Empresa: {web_acao_nome.text} 
    |Dividend Yield: {web_acao_dividendYield.text} 
    |P/L: {web_acao_pl.text} 
    |P/VP: {web_acao_pvp.text} 
    |ROE: {web_acao_roe.text} 
    |ROIC: {web_acao_roic.text} 
    |EBIT: R$ {web_acao_ebit.text} 
    |Lucro Líquido: R$ {web_acao_lucroliquido.text} \n
    """)

    except:
        print("\n    |Nada foi encontrado")        

elif args.acoes:
    print("\n    |Baixando as informações ...")
    
    navegador.get(f"https://www.fundamentus.com.br/resultado.php")
    
    arquivoCaminho = caminhoDefault + "/invest_acoes.xlsx"
        
    workbook =xlsxwriter.Workbook(arquivoCaminho)
    worksheet = workbook.add_worksheet("Todas Ações")
    
    cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
    
    formatoDinheiro = workbook.add_format({'num_format': '\\R$ #,##0.00'})        
    formatoDecimal = workbook.add_format({'num_format': '#,##0.00'})        
    formatoPorcentagem = workbook.add_format({'num_format': '0.00,##%'})
    
    table = navegador.execute_script("""
    var table = document.getElementById('resultado'); 
    var data = [];
    for (var i = 0, row; row = table.rows[i]; i++) {
        var rowData = [];
 
            for (var j = 0, cell; j < 21; j++) {

                cell = row.cells[j]
                rowData.push(cell.textContent);
            }
        
        data.push(rowData);
    }
    return data;""")
    
    qtd_column = 1
    qtd_row = 1

    for table_row in table:
        
        for table_data in table_row:
            
            if qtd_column == 21:
                qtd_column = 1
                qtd_row+=1
                
            else:

                if qtd_row == 1:
                    
                    negrito = workbook.add_format()
                    negrito.set_bold()
                    negrito.set_font_size(13)
                    
                    worksheet.autofilter('B2:N2')
                    
                    worksheet.write(qtd_row, qtd_column, table_data, negrito)
                
                elif qtd_column in [2, 19] and qtd_row >= 2:
                    
                    worksheet.write(qtd_row, qtd_column, formatarNumeros(table_data, "float"), formatoDinheiro)    
                
                elif qtd_column in [3, 4, 5, 7, 8, 9, 10, 11, 12, 15, 18, 20] and qtd_row >=2:

                    worksheet.write(qtd_row, qtd_column, formatarNumeros(table_data, "float"), formatoDecimal)    
                    
                
                elif qtd_column in [6, 13, 14, 16, 17, 21] and qtd_row >= 2:
                
                    worksheet.write(qtd_row, qtd_column, formatarNumeros(table_data, "percentage"), formatoPorcentagem)    
             
                    
                else:
                    worksheet.write(qtd_row, qtd_column, str(table_data))    
                    # worksheet.write(qtd_row, qtd_column, cell_format)

                qtd_column+=1

    
    workbook.close()
    
    subprocess.run([excel_path, arquivoCaminho])
    


elif args.fiis: 
    print("\n    |Baixando as informações ...")
    
    navegador.get(f"https://www.fundamentus.com.br/fii_resultado.php")
     
    arquivoCaminho = caminhoDefault + "/invest_fiis.xlsx" 
     
    workbook =xlsxwriter.Workbook(arquivoCaminho)
    worksheet = workbook.add_worksheet("Todos FIIS")
     
    cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
    
    formatoDinheiro = workbook.add_format({'num_format': '\\R$ #,##0.00'})        
    formatoDecimal = workbook.add_format({'num_format': '#,##0.00'})        
    formatoPorcentagem = workbook.add_format({'num_format': '0.00,##%'})
    formatoInteiro = workbook.add_format({'num_format': '0'})
    
    table = navegador.execute_script("""
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

    for table_row in table:
        
        for table_data in table_row:
            
            if qtd_column == 14:
                qtd_column = 1
                qtd_row+=1
                
            else:
                
                if qtd_row == 1:
                    
                    negrito = workbook.add_format()
                    negrito.set_bold()
                    negrito.set_font_size(13)
                    
                    worksheet.autofilter('B2:N2')
                    
                    worksheet.write(qtd_row, qtd_column, table_data, negrito)
                
                elif qtd_column in [3, 7, 10, 11] and qtd_row >= 2:
                    
                    worksheet.write(qtd_row, qtd_column, formatarNumeros(table_data, "float"), formatoDinheiro)    
                
                elif qtd_column in [6] and qtd_row >=2:

                    worksheet.write(qtd_row, qtd_column, formatarNumeros(table_data, "float"), formatoDecimal)    
                    
                
                elif qtd_column in [4, 5, 12, 13] and qtd_row >= 2:
                
                    worksheet.write(qtd_row, qtd_column, formatarNumeros(table_data, "percentage"), formatoPorcentagem)    
                    
                elif qtd_column in [8, 9] and qtd_row >= 2:
                
                    worksheet.write(qtd_row, qtd_column, formatarNumeros(table_data, "integer"), formatoInteiro)                    
                    
                else:
                    worksheet.write(qtd_row, qtd_column, str(table_data))    
                    # worksheet.write(qtd_row, qtd_column, cell_format)

                qtd_column+=1
                        
                
    workbook.close()
        
    subprocess.run([excel_path, arquivoCaminho])
