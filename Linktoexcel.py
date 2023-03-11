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

#createfile("IOT_SEC_A.xlsx")
writedata("IOT_SEC_A.xlsx", "LOVE", "219311046", "43 TR IO TD")
