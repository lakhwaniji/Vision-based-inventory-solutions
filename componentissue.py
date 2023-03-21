import os.path
import PySimpleGUI as sg
import Linktoexcel
import time
import ard_data

student_name = []
reg_number = []
card_number = []
clock = sg.Text(time.strftime("%d %b %Y"), key="clock")
sg.theme('BluePurple')
label_name = sg.Text("Name--:")
label_regname = sg.Text("Registration--:")
studname = sg.Text("", key="studname", text_color="Green")
regno = sg.Text("", key="regno", text_color="Green")
label_1 = sg.Text("Choose Your Section")
label_2 = sg.Text("Scan your Card")
label_3 = sg.Radio("IOT SEC A", "sec", default=False, key="IOT_SEC_A")
label_4 = sg.Radio("IOT SEC B", "sec", default=False, key="IOT_SEC_B")
label_5 = sg.Radio("IOT SEC C", "sec", default=False, key="IOT_SEC_C")
label_6 = sg.Radio("IOT SEC D", "sec", default=False, key="IOT_SEC_D")
button_1 = sg.Button("Open Camera", size=15, key="camera")
button_2 = sg.Button("Scan Your Card", key="scan", size=15)
button_3 = sg.Button("Exit", key="exit", size=15)
submit = sg.Button("Submit", size=15, key="submit")
card_no = sg.Text("", text_color="Red", key="card_no")
result = sg.Text("", text_color="Green", key="result")
window = sg.Window("Issue", layout=[[clock],
                                    [label_name, studname],
                                    [label_regname, regno],
                                    [label_1],
                                    [label_3, label_4, label_5, label_6],
                                    [button_2, card_no],
                                    [button_1],
                                    [button_3],
                                    [submit],
                                    [result]],
                   font=('Helvetica', 15))
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

        case "exit":
            break
        case "submit":
            try:
                if value["IOT_SEC_A"]:
                    if not os.path.exists("IOT_SEC_A_data.xlsx"):
                        Linktoexcel.createfile("IOT_SEC_A_data.xlsx")
                    if result == 'Success':
                        window["result"].update(value="Data Registered Successfully")
                    else:
                        window["result"].update(text_color="Red", value="Error in registering Data")
                elif value["IOT_SEC_B"]:
                    if not os.path.exists("IOT_SEC_B_data.xlsx"):
                        Linktoexcel.createfile("IOT_SEC_B_data.xlsx")
                    if result == 'Success':
                        window["result"].update(value="Data Registered Successfully")
                    else:
                        window["result"].update(text_color="Red", value="Error in registering Data")
                elif value["IOT_SEC_C"]:
                    if not os.path.exists("IOT_SEC_C_data.xlsx"):
                        Linktoexcel.createfile("IOT_SEC_C_data.xlsx")
                    if result == 'Success':
                        window["result"].update(value="Data Registered Successfully")
                    else:
                        window["result"].update(text_color="Red", value="Error in registering Data")
                elif value["IOT_SEC_D"]:
                    if not os.path.exists("IOT_SEC_D_data.xlsx"):
                        Linktoexcel.createfile("IOT_SEC_D_data.xlsx")
                    if result == 'Success':
                        window["result"].update(value="Data Registered Successfully")
                    else:
                        window["result"].update(text_color="Red", value="Error in registering Data")
            except:
                window["result"].update(text_color="Red", value="Error in registering Data")
        case sg.WIN_CLOSED:
            break
window.close()
