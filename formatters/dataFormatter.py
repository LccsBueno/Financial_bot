import xlsxwriter

class DataFormatter:
    
    @staticmethod
    def formatDataToFloat(data):
        number = data.replace(".", "")
        number = round(float(number.replace(",", ".")), 2)
        return number
    
    @staticmethod
    def formatDataToPercentage(data):
        number = data.replace("%", "")
        number = number.replace(".", "")
        number = round(float(number.replace(",", ".")), 2) / 100
        return number
    
    @staticmethod
    def formatDataToInteger(data):
        number = int(data.replace(".", ""))
        return number
    
    @staticmethod
    def formatDataToString(data):
        string = str(data)
        return string
    