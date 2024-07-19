from formatters import sheetFormatter
from . import webAcess
import datetime

class FiisCollector(webAcess.WebAcess):
    
    @staticmethod
    def scrap():
    
        FiisCollector.navegador.get(f"https://www.fundamentus.com.br/fii_resultado.php")

        qtd_columns = 14
        
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
        
        now = datetime.datetime.now()

        sht = sheetFormatter.SheetFormatter("./Fiis-"+str(now.month)+"-"+str(now.day)+"-"+str(now.year)+".xlsx",
                                            qtd_columns, 
                                            [3, 7, 10, 11],
                                            [6],
                                            [4, 5, 12, 13],
                                            integerColumn = [8, 9])

        sht.sheetGenerator(data)
    