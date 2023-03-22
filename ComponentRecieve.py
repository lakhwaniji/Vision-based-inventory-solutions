import os.path
import PySimpleGUI as sg
import Linktoexcel
import time
import ard_data

student_name = []
reg_number = []
card_number = []
components = ""
clock = sg.Text(time.strftime("%d %b %Y"), key="clock")
sg.theme('BluePurple')
label_name = sg.Text("Name--:")
label_regname = sg.Text("Registration--:")
studname = sg.Text("", key="studname", text_color="Green")
regno = sg.Text("", key="regno", text_color="Green")
IssuedComponents = sg.Text("", key="issuedcomponents", text_color="Black")
label_1 = sg.Text("Choose Your Section")
label_2 = sg.Text("Scan your Card")
label_3 = sg.Radio("IOT SEC A", "sec", default=False, key="IOT_SEC_A")
label_4 = sg.Radio("IOT SEC B", "sec", default=False, key="IOT_SEC_B")
label_5 = sg.Radio("IOT SEC C", "sec", default=False, key="IOT_SEC_C")
label_6 = sg.Radio("IOT SEC D", "sec", default=False, key="IOT_SEC_D")
button_2 = sg.Button("Scan Your Card", key="scan", size=15)
button_3 = sg.Button("Exit", key="exit", size=15)
submit = sg.Button("Submit", size=15, key="submit")
card_no = sg.Text("", text_color="Red", key="card_no")
result = sg.Text("", text_color="Green", key="result")
window = sg.Window("Recieve", layout=[[clock],
                                      [label_name,studname],
                                      [label_regname,regno],
                                      [label_1],
                                      [label_3, label_4, label_5, label_6],
                                      [IssuedComponents],
                                      [button_2,card_no],
                                      [button_3],
                                      [submit],
                                      [result]],
                   font=('Helvetica',15))
while True:
    event, value = window.read(timeout=1000)
    match event:
        case "scan":
            data = ard_data.scanning_card()
            window["card_no"].update(value=data)
            if value["IOT_SEC_A"]:
                card_number = Linktoexcel.readcardno("IOT_SEC_A.xlsx")
                reg_number = Linktoexcel.readregno("IOT_SEC_A.xlsx")
                student_name = Linktoexcel.readstudname("IOT_SEC_A.xlsx")
            elif value["IOT_SEC_B"]:
                card_number = Linktoexcel.readcardno("IOT_SEC_B.xlsx")
                reg_number = Linktoexcel.readregno("IOT_SEC_B.xlsx")
                student_name = Linktoexcel.readstudname("IOT_SEC_B.xlsx")
            elif value["IOT_SEC_C"]:
                card_number = Linktoexcel.readcardno("IOT_SEC_C.xlsx")
                reg_number = Linktoexcel.readregno("IOT_SEC_C.xlsx")
                student_name = Linktoexcel.readstudname("IOT_SEC_C.xlsx")
            elif value["IOT_SEC_D"]:
                card_number = Linktoexcel.readcardno("IOT_SEC_D.xlsx")
                reg_number = Linktoexcel.readregno("IOT_SEC_D.xlsx")
                student_name = Linktoexcel.readstudname("IOT_SEC_D.xlsx")
            idx = card_number.index(data)
            window["studname"].update(value=student_name[idx])
            window["regno"].update(value=reg_number[idx])
            components=Linktoexcel.getcomponents(data)
            window["issuedcomponents"].update(value=components)
            if(len(components)==0):
                window["issuedcomponents"].update(value="No Component Issued")
        case "exit":
            break
        case "submit":
            status=Linktoexcel.delete(data)
            if (status == "Success"):
                window["result"].update(text_color="Green", value="Data Deleted Successfully")
            else:
                window["result"].update(text_color="Red", value="Error In Deleting Data")
        case sg.WIN_CLOSED:
            break
window.close()
