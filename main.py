from tkinter import *
from tkinter import filedialog
import xlrd


def browse():
    global filename
    filename = filedialog.askopenfilename(initialdir="/Users/Sinan/Desktop", title="Select file",
                                          filetypes=[("Excel files", ".xlsx .xls"), ("all files",
                                                                                     "*.*")])


def calculate():
    global text, OUTPUT
    selected_file = xlrd.open_workbook(filename).sheet_by_index(0)
    total_attendance = 10
    failed_student_count = 0
    for cell in selected_file.col(1, 1):
        if float(cell.value) < total_attendance * float(threshold.get()) / 100:
            failed_student_count += 1
    OUTPUT = 'Total Number of Students Failed due to inadequate attendance: {}'.format(failed_student_count)
    text.insert(END, OUTPUT)


def clear():
    global OUTPUT
    text.delete('1.0', END)
    threshold.set('')


root = Tk()
root.title("Attendance Calculator")
root.geometry('640x480-8-200')

label_main = Label(root, text="Attendance Calculator")
label_main.grid(row=0, column=0, columnspan=2, pady=10)


label_browse = Label(root, text="Enter Input File Name")
label_browse.grid(row=1, column=0, sticky=E, pady=5)

threshold = StringVar()
threshold.set('')
entry_threshold = Entry(root, textvariable=threshold, width=5)
entry_threshold.grid(row=2, column=1, sticky=W, pady=5)

pass_treshold = Label(root, text="Pass Threshold(%)")
pass_treshold.grid(row=2, column=0, sticky=E, pady=5)

browse_button = Button(root, text='Browse', command=browse)
browse_button.grid(row=1, column=1, sticky=W, pady=5)

button_calculate = Button(root, text="Calculate", width=10, command=calculate)
button_calculate.grid(row=3, column=0, pady=5, padx=20)

button_clear = Button(root, text="Clear", width=10, command=clear)
button_clear.grid(row=3, column=1, pady=5, padx=20)

text = Text(width=50, height=10)
text.grid(row=4, column=0, columnspan=2)


if __name__ == '__main__':
    root.mainloop()
