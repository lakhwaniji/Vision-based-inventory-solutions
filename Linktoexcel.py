import openpyxl


def createfile(filename):
    workbook = openpyxl.Workbook()
    work = workbook.active
    work.column_dimensions['A'].width = 20
    work.column_dimensions['B'].width = 20
    work.column_dimensions['C'].width = 20
    workbook.save(filename=filename)


def writedata(filename, name, regno, sid):
    try:
        wb = openpyxl.load_workbook(filename)
        sheet = wb.active
        data = (name, regno, sid)
        sheet.append(data)
        wb.save(filename)
        return "Success"
    except:
        return "Error"


def readcardno(filename):
    try:
        wb2 = openpyxl.load_workbook("IOT_SEC_A.xlsx")
        sheet = wb2.active
        values = [sheet.cell(row=i, column=3).value for i in range(1, sheet.max_row + 1)]
        return values
    except:
        return "Error"

def readstudname(filename):
    try:
        wb2 = openpyxl.load_workbook("IOT_SEC_A.xlsx")
        sheet = wb2.active
        values = [sheet.cell(row=i, column=1).value for i in range(1, sheet.max_row + 1)]
        return values
    except:
        return "Error"

def readregno(filename):
    try:
        wb2 = openpyxl.load_workbook("IOT_SEC_A.xlsx")
        sheet = wb2.active
        values = [sheet.cell(row=i, column=2).value for i in range(1, sheet.max_row + 1)]
        return values
    except:
        return "Error"
