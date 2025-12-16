from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import main as Main

root = Tk()
root.geometry("700x475")
root.title('Expo LTE KPI Report Generation Tool')
root.iconbitmap('Images\\huawei_icon.ico')
root.iconify()
root.configure(bg = 'white')
logo = PhotoImage(file = "Images\\auditor3.png")
logo_lbl = Label(root, image = logo)
logo_lbl.pack()




def file_report():
    global NE_report
    NE_report = filedialog.askopenfilename(title='Select "Input Report"')
    NE_report_lbl = Label(root, text=NE_report, bg='azure2', font=("Times New Roman", 10, 'bold'))
    NE_report_lbl.place(x=120, y=265)

def output():
    global out_path
    out_path = filedialog.askdirectory(title='Select Output directory')
    out_path_lable = Label(root, text=out_path, bg='azure2', font=("Times New Roman", 10, 'bold'))
    out_path_lable.place(x=120, y=315)

def close():
    gh = messagebox.askquestion('Warning', 'Are you sure you want to quit?')
    if gh == 'yes':
        root.quit()


btn_rep = Button(root, text = 'INPUT File', bg = 'white smoke', font = ("Times New Roman", 10, "bold"), width=10, command=lambda: file_report())
btn_rep.place(x=30, y=260)
btn_out = Button(root, text = 'OUTPUT Dir', bg = 'white smoke', font=("Times New Roman", 10, "bold"), width=10, command=lambda: output())
btn_out.place(x=30, y=310)
btn_start = Button(root, text = 'Start', bg = 'lavender', font=("Times New Roman", 12, "bold"), width=6, command=lambda:Main.OD_IBS_SLA_STATUS(NE_report, out_path))
btn_start.place(x=220, y=370)
btn_quit = Button(root, text = 'Quit', bg = 'lavender', font=("Times New Roman", 12, "bold"), width=6, command=lambda: close())
btn_quit.place(x=310, y=370)
lbl_signature1 = Label(root, text = '               For Support :  Danish Ali (dwx854280)\n             Contact :  00971508552942', bg='white', font=("Times New Roman", 8, 'bold'))
lbl_signature1.place(x=150, y=430)
root.mainloop()