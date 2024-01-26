import xlsxwriter
import subprocess
import dataFormatter
from dataFormatter import DataFormatter


class SheetFormatter:

    def __init__(self,
                 generatedSheetPath,
                 currencyColumns,
                 decimalColumns,
                 percentageColumns):
                
        self.excel_path = r"C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE"
        self.generatedSheetPath = generatedSheetPath
        self.currencyColumns = currencyColumns
        self.decimalColumns = decimalColumns
        self.percentageColumns = percentageColumns
    
    def sheetGenerator(self, data):

        workbook =xlsxwriter.Workbook(self.generatedSheetPath)
        worksheet = workbook.add_worksheet("Todas Ações")

        qtd_column = 1
        qtd_row = 1

        for table_row in data:
            
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
                    
                    elif qtd_column in self.currencyColumns and qtd_row >= 2:
                        
                        worksheet.write(qtd_row, qtd_column, formatarNumeros(table_data, "float"), formatoDinheiro)    
                    
                    elif qtd_column in self.decimalColumns and qtd_row >=2:

                        worksheet.write(qtd_row, qtd_column, formatarNumeros(table_data, "float"), formatoDecimal)    
                        
                    
                    elif qtd_column in self.percentageColumns and qtd_row >= 2:
                    
                        worksheet.write(qtd_row, qtd_column, formatarNumeros(table_data, "percentage"), formatoPorcentagem)    
                
                        
                    else:
                        worksheet.write(qtd_row, qtd_column, str(table_data))    
                        # worksheet.write(qtd_row, qtd_column, cell_format)

                    qtd_column+=1

        
        workbook.close()
        
        subprocess.run([self.excel_path, self.generatedSheetPath])