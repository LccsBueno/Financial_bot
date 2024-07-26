from formatters import sheetFormatter
from . import webAcess 
import datetime

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
               
        currencyColumns = [2, 19]
        decimalColumns = [3, 4, 5, 7, 8, 9, 10, 11, 12, 15, 18, 20]
        percentageColumns = [6, 13, 14, 16, 17, 21]
        integerColumn = [0]
               
        print(currencyColumns)
        now = datetime.datetime.now()
               
        sht = sheetFormatter.SheetFormatter("./Stocks-"+str(now.month)+"-"+str(now.day)+"-"+str(now.year)+".xlsx", 
                                            len(data[0]),
                                            currencyColumns = currencyColumns,
                                            decimalColumns = decimalColumns,
                                            percentageColumns = percentageColumns,
                                            )
        sht.sheetGenerator(data)
    

