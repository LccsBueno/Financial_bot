from formatters import sheetFormatter
from . import webAcess 
import datetime
import pandas as pd
import numpy as np
from formatters import dataFormatter as dt

import sys

class StocksCollector(webAcess.WebAcess):
    
    @staticmethod
    def scrap():
    
        StocksCollector.navegador.get(f"https://www.fundamentus.com.br/resultado.php")
        
        data = StocksCollector.navegador.execute_script("""
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

                    
        currencyColumns = [1, 18]
        decimalColumns = [2, 3, 4, 6, 7, 8, 9, 10, 11, 14, 17, 19]
        percentageColumns = [5, 12, 13, 15, 16, 20]
        integerColumn = []
        
        dataLength = len(data[0])
        
        cabecalho = data[0]
        data.pop(0)
        
        df = pd.DataFrame(data)

        for dataFrameColumn in currencyColumns :
            df[dataFrameColumn] = pd.to_numeric(df[dataFrameColumn].str.replace('.', '').str.replace(',', '.'), errors='coerce')
        
        for dataFrameColumn in decimalColumns:
            df[dataFrameColumn] = pd.to_numeric(df[dataFrameColumn].str.replace('.', '').str.replace(',', '.'), errors='coerce')
        
        for dataFrameColumn in percentageColumns:
            df[dataFrameColumn] = pd.to_numeric(df[dataFrameColumn].str.replace('%', '').str.replace('.', '').str.replace(',', '.'), errors='coerce')
            df[dataFrameColumn] = df[dataFrameColumn] / 100 
        
        #ADICIONANDO COLUNA
        df[dataLength] = df[1] * (df[5] * 100) / 12
        currencyColumns.append(dataLength)
        cabecalho.append("Dividendo/Mês")
        dataLength+=1
        
        df[dataLength] = "Resultados"
        cabecalho.append("Resultados")
        dataLength+=1
                
        array = df.to_numpy()
     
        array = np.insert(array, 0, cabecalho, axis=0)

        sht = sheetFormatter.SheetFormatter("./Stocks.xlsx", 
                                            dataLength,
                                            currencyColumns = currencyColumns,
                                            decimalColumns = decimalColumns,
                                            percentageColumns = percentageColumns,
                                            )
        
        
        
        sht.sheetGenerator(array)
        
    @staticmethod    
    def getResultados(acao):
        url = f"https://www.fundamentus.com.br/resultados_trimestrais.php?papel={acao}&tipo=1"
        return url
    
    

