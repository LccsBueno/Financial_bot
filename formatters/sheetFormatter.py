import xlsxwriter
import subprocess
import datetime

from . import dataFormatter as dt

class SheetFormatter: 
    
    def __init__(
            self,
            generatedSheetPath,
            qtd_columns,
            currencyColumns,
            decimalColumns,
            percentageColumns,
            **kwargs):
            
        self.excel_path = r"C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE"
        self.generatedSheetPath = generatedSheetPath
        self.qtd_columns = qtd_columns
        self.currencyColumns = currencyColumns
        self.decimalColumns = decimalColumns
        self.percentageColumns = percentageColumns 
        self.integerColumns = kwargs.get("integerColumn")
        
        self.workbook = xlsxwriter.Workbook(self.generatedSheetPath)
        
        now = datetime.datetime.now()
        self.worksheet = self.workbook.add_worksheet("FIIS") 
    
    def sheetGenerator(self, data):
        
        currency = self.workbook.add_format({'num_format': '\\R$ #,##0.00'})        
        decimal = self.workbook.add_format({'num_format': '#,##0.00'})        
        percentage = self.workbook.add_format({'num_format': '0.00,##%'})
        integer = self.workbook.add_format({'num_format': '0'})
        
        qtd_column = 1
        qtd_row = 1

        for table_row in data:
            
            for table_data in table_row:
                
                if qtd_column == self.qtd_columns:
                    qtd_column = 1
                    qtd_row+=1
                    
                else:

                    if qtd_row == 1:
                        
                        bold = self.workbook.add_format()
                        bold.set_bold()
                        bold.set_font_size(13)
                        
                        self.worksheet.autofilter('B2:N2')
                        
                        self.worksheet.write(qtd_row, qtd_column, table_data, bold)

                    
                    elif qtd_column in self.currencyColumns and qtd_row >= 2:
                        
                        self.worksheet.write(qtd_row, qtd_column, dt.DataFormatter.formatDataToFloat(table_data), currency)    
                    
                    
                    elif qtd_column in self.decimalColumns and qtd_row >=2:
                        self.worksheet.write(qtd_row, qtd_column, dt.DataFormatter.formatDataToFloat(table_data), decimal)    
                        
                    
                    elif qtd_column in self.percentageColumns and qtd_row >= 2:
                    
                        self.worksheet.write(qtd_row, qtd_column, dt.DataFormatter.formatDataToPercentage(table_data), percentage)    
                

                    elif (qtd_column in self.integerColumns and qtd_row >= 2) and not type(self.integerColumns) == "NoneType": 
                        
                        self.worksheet.write(qtd_row, qtd_column, dt.DataFormatter.formatDataToInteger(table_data), integer)    
                        
                    else:
                        self.worksheet.write(qtd_row, qtd_column, str(table_data))    
                        # worksheet.write(qtd_row, qtd_column, cell_format)

                    qtd_column+=1


        try:
            self.workbook.close()
        except Exception as e:
            print("Permissão negada, não pode fechar o workbook: ",e)
        
        
        subprocess.run([self.excel_path, self.generatedSheetPath])
        