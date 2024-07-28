from formatters import sheetFormatter
from . import webAcess
import datetime
import pandas as pd
import numpy as np

class FiisCollector(webAcess.WebAcess):
    
    @staticmethod
    def scrap():
    
        FiisCollector.navegador.get(f"https://www.fundamentus.com.br/fii_resultado.php")
        
        data = FiisCollector.navegador.execute_script("""
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
        
        # ESPERAR UMA ENTRADA DO USUARIO PARA MOSTRAR AS COLUNAS E PEDIR CONFIRMACAO NO TIPO DELAS
        
        currencyColumns = [3, 7, 10, 11]
        decimalColumns = [6]
        percentageColumns = [4, 5, 12, 13]
        integerColumns = [8, 9]
                        
        cabecalho = data[0]
        data.pop(0)
        df = pd.DataFrame(data)  
        
        print(df) 
        
        array = df.to_numpy()
        array = np.insert(array, 0, cabecalho, axis=0)

        now = datetime.datetime.now()

        sht = sheetFormatter.SheetFormatter("./Fiis-"+str(now.month)+"-"+str(now.day)+"-"+str(now.year)+".xlsx",
                                            len(data[0]), 
                                            currencyColumns = currencyColumns,
                                            decimalColumns = decimalColumns,
                                            percentageColumns = percentageColumns,
                                            integerColumn = integerColumns
                                            )

        # cont=1
        # string = " {posicao:6} | {coluna:18} | {tipoColuna:12}"
        
        # print(string.format(posicao="Nº Col", coluna="Nome Col", tipoColuna="Tipo Col"))
        # print("---------------------------------------")
            
        # for i in data[0]:
            
        #     if cont in currencyColumns:
        #         print(string.format(posicao=cont, coluna=i, tipoColuna="Monetário"))
            
        #     elif cont in decimalColumns:
        #         print(string.format(posicao=cont, coluna=i, tipoColuna="Decimal"))
                
        #     elif cont in percentageColumns:
        #         print(string.format(posicao=cont, coluna=i, tipoColuna="Porcentagem"))
                
        #     elif cont in integerColumns:
        #         print(string.format(posicao=cont, coluna=i, tipoColuna="Inteiro"))
            
        #     else:
        #         print(string.format(posicao=cont, coluna=i, tipoColuna="Texto"))
            
        #     cont+=1
                    
            
        sht.sheetGenerator(array)
    