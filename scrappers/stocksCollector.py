import webAcess
from webAcess import WebAcess

import formatters 
from formatters import sheetFormatter
sheetFormatter.SheetFormatter 



class StocksCollector(WebAcess):

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
    

