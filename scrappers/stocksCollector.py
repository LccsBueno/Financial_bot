# from webAcess import WebAcess;

# import formatters 

# from formatters import sheetFormatter
# import sys
# sys.path.insert(1, '../formatters/sheetFormatter.py')

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from formatters import sheetFormatter

import logging

class WebAcess:

    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument('headless')
    navegador = webdriver.Chrome(options=options)

    selenium_logger = logging.getLogger('selenium')
    selenium_logger.setLevel(logging.WARNING)

class StocksCollector(WebAcess):
    
    @staticmethod
    def scrap():
    
        WebAcess.navegador.get(f"https://www.fundamentus.com.br/resultado.php")

        data = WebAcess.navegador.execute_script("""
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

        
        sht = sheetFormatter.SheetFormatter("./stocks", 
                                            [2, 19],
                                            [3, 4, 5, 7, 8, 9, 10, 11, 12, 15, 18, 20], 
                                            [6, 13, 14, 16, 17, 21])

        sht.sheetGenerator(data)
    

