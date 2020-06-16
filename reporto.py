import tkinter as tk
from fpdf import FPDF
from tkinter.messagebox import showinfo


def makeform():
    rp_save = rpnum_text.get('1.0', 'end').split('\n')[0].strip()
    serial_save = serial_text.get('1.0', 'end').split('\n')[0].strip()
    pdf = FPDF()
    pdf.add_page()
    sum_out = ''
    count = 0

    if rpnum_text.get('1.0', 'end').split('\n')[0] == '':
        if key_text.get('1.0', 'end').split('\n')[0] == '':
            if rp.get() == False and nrp.get() == False and rt.get() == False and kj.get() == False:
                norp('RP#\nKeycode\nResult\n')
                return None
            else:
                norp('RP#\nKeycode\n')
                return None
        else:
            norp('RP#\n')
            return None
    elif key_text.get('1.0', 'end').split('\n')[0] == '':
        if rp.get() == False and nrp.get() == False and rt.get() == False and kj.get() == False:
            norp('Keycode\nResult\n')
            return None
        else:
            norp('Keycode\n')
            return None
    elif rp.get() == False and nrp.get() == False and rt.get() == False and kj.get() == False:
        norp('Result\n')
        return None
    else:
        pass

    pdf.set_font('Arial', 'I', 18)
    pdf.cell(200, 15, 'FA RSS', ln=1, align='C')
    pdf.set_font('Arial', '', 12)
    pdf.cell(40, 10, 'RP#:' + ' ' + rpnum_text.get('1.0', 'end'), ln=1)
    pdf.cell(40, 10, 'Serial:' + serial_text.get('1.0', 'end'), ln=1)
    pdf.cell(40, 10, 'Keycode: ' + key_text.get('1.0', 'end'), ln=1)
    pdf.cell(40, 10, 'PIC: ' + pic_text.get('1.0', 'end'), ln=1)
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(40, 10, '< Apperance >', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.cell(40, 10, apperance_text.get('1.0', 'end'), ln=1)
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(40, 10, '', ln=1)
    pdf.cell(40, 10, '< Symptom Information >', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.cell(40, 10, symptom_text.get('1.0', 'end'), ln=1)
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(40, 10, '', ln=1)
    pdf.cell(40, 10, '< Tests Performed >', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.cell(40, 10, tests_text.get('1.0', 'end'), ln=1)
    pdf.cell(40, 10, '', ln=1)
    pdf.cell(40, 10, 'Failed Board: ' + fb_text.get('1.0', 'end'), ln=1)
    pdf.cell(40, 10, 'Location of Component: ' + comp_text.get('1.0', 'end'), ln=1)
    pdf.cell(40, 10, 'Failed Structure Part: ' + structure_text.get('1.0', 'end'), ln=1)
    pdf.cell(40, 10, '', ln=1)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(40, 10, '< Analysis Steps (list) > ', ln=1)
    pdf.set_font('Arial', '', 12)

    for a in range(len(analysis_text.get('1.0', 'end').split('\n'))):
        pdf.cell(40, 10, analysis_text.get('1.0', 'end').split('\n')[count], ln=1)
        count += 1

    pdf.set_font('Arial', 'B', 14)
    pdf.cell(40, 10, '', ln=1)
    pdf.cell(40, 10, '< Failure Cause >', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.cell(40, 10, failure_text.get('1.0', 'end'), ln=1)
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(40, 10, '', ln=1)
    pdf.cell(40, 10, '< Summary >', ln=1)
    pdf.set_font('Arial', '', 12)

    if rp.get() == True:
        sum_out = 'This unit was repaired and sent back to the customer'
    elif nrp.get() == True:
        sum_out = 'This unit was not repaired and sent back to the customer'
    elif rt.get() == True:
        sum_out = 'This unit was sent back to the customer in the condition it was received'
    elif kj.get() == True:
        sum_out = 'This unit was sent to KJ'
    pdf.cell(40, 10, sum_out, ln=1)

    pdf.output(rp_save + ' - ' + serial_save + '.pdf', 'F')
    rpnum_text.delete('1.0', 'end')
    serial_text.delete('1.0', 'end')
    key_text.delete('1.0', 'end')
    apperance_text.delete('1.0', 'end')
    symptom_text.delete('1.0', 'end')
    tests_text.delete('1.0', 'end')
    failure_text.delete('1.0', 'end')
    comp_text.delete('1.0', 'end')
    structure_text.delete('1.0', 'end')
    analysis_text.delete('1.0', 'end')
    fb_text.delete('1.0', 'end')
    item_text.delete('1.0', 'end')


def focus_next_window(event):
    event.widget.tk_focusNext().focus()
    return "break"


def norp(wrds):
    showinfo('Window', 'You are missing:\n' + wrds)


root = tk.Tk()
w, h = root.winfo_screenwidth(), root.winfo_screenheight()
# root.geometry("%dx%d+0+0" % (w, h))
root.geometry('900x600')
root.title('RSS Sheet Generator')

rpnum_label = tk.Label(text='RP#')
rpnum_label.grid(row=1, column=1)
rpnum_text = tk.Text(root, height=1, width=10)
rpnum_text.grid(row=2, column=1)
rpnum_text.bind("<Tab>", focus_next_window)

item_label = tk.Label(text='Item Name')
item_label.grid(row=3, column=1)
item_text = tk.Text(root, height=1, width=10)
item_text.grid(row=4, column=1)
item_text.bind("<Tab>", focus_next_window)

serial_label = tk.Label(text='Serial#')
serial_label.grid(row=5, column=1)
serial_text = tk.Text(root, height=1, width=10)
serial_text.grid(row=6, column=1)
serial_text.bind("<Tab>", focus_next_window)

key_label = tk.Label(text='Keycode')
key_label.grid(row=1, column=2)
key_text = tk.Text(root, height=1, width=10)
key_text.grid(row=2, column=2)
key_text.bind("<Tab>", focus_next_window)

fb_label = tk.Label(text='Failed Board')
fb_label.grid(row=3, column=2)
fb_text = tk.Text(root, height=1, width=10)
fb_text.grid(row=4, column=2)
fb_text.bind("<Tab>", focus_next_window)

comp_label = tk.Label(text='Location of Component')
comp_label.grid(row=5, column=2)
comp_text = tk.Text(root, height=1, width=10)
comp_text.grid(row=6, column=2)
comp_text.bind("<Tab>", focus_next_window)

structure_label = tk.Label(text='Failed Structure Part')
structure_label.grid(row=1, column=3)
structure_text = tk.Text(root, height=1, width=10)
structure_text.grid(row=2, column=3)
structure_text.bind("<Tab>", focus_next_window)

pic_label = tk.Label(text='PIC')
pic_label.grid(row=3, column=3)
pic_text = tk.Text(root, height=1, width=10)
pic_text.grid(row=4, column=3)
pic_text.bind("<Tab>", focus_next_window)

''' Check Boxes'''

rp = tk.BooleanVar()
nrp = tk.BooleanVar()
rt = tk.BooleanVar()
kj = tk.BooleanVar()

key_label = tk.Label(text='Result')
key_label.grid(row=1, column=4)

rpchk_label = tk.Checkbutton(root, text="Repaired",
                             onvalue=1, offvalue=0, height=1,
                             width=10, variable=rp)
rpchk_label.grid(row=2, column=4)

nrpchk_label = tk.Checkbutton(root, text="Not repaired",
                              onvalue=1, offvalue=0, height=1,
                              width=10, variable=nrp)
nrpchk_label.grid(row=3, column=4)

rtchk_label = tk.Checkbutton(root, text="Returned (NFF)",
                             onvalue=1, offvalue=0, height=1,
                             width=10, variable=rt)
rtchk_label.grid(row=4, column=4)

kjchk_label = tk.Checkbutton(root, text="Sent to KJ",
                             onvalue=1, offvalue=0, height=1,
                             width=10, variable=kj)
kjchk_label.grid(row=5, column=4)

'''Form'''

apperance = tk.Label(text='Apperance Inspection')
apperance.grid(row=8, column=1)
apperance_text = tk.Text(root, height=10, width=30)
apperance_text.grid(row=9, column=1)
apperance_text.bind("<Tab>", focus_next_window)

symptom = tk.Label(text='Symptom Information')
symptom.grid(row=8, column=2)
symptom_text = tk.Text(root, height=10, width=30)
symptom_text.grid(row=9, column=2)
symptom_text.bind("<Tab>", focus_next_window)

tests = tk.Label(text='Tests Performed')
tests.grid(row=8, column=3)
tests_text = tk.Text(root, height=10, width=30)
tests_text.grid(row=9, column=3)
tests_text.bind("<Tab>", focus_next_window)

analysis = tk.Label(text='Analysis Steps')
analysis.grid(row=10, column=1)
analysis_text = tk.Text(root, height=10, width=30)
analysis_text.grid(row=11, column=1)
analysis_text.bind("<Tab>", focus_next_window)

failure = tk.Label(text='Cause of Failure')
failure.grid(row=10, column=2)
failure_text = tk.Text(root, height=10, width=30)
failure_text.grid(row=11, column=2)
failure_text.bind("<Tab>", focus_next_window)

# summary = tk.Label(text='Summary')
# summary.grid(row=10, column=3)
# summary_text = tk.Text(root, height=10, width=30)
# summary_text.grid(row=11, column=3)
# summary_text.bind("<Tab>", focus_next_window)

btnRead = tk.Button(root, height=1, width=10, text="Save", command=makeform)
btnRead.grid(row=20, column=1)

root.mainloop()
