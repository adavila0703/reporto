import tkinter as tk
from fpdf import FPDF
from tkinter.messagebox import showinfo

def checkerror():
    error = ''
    if rpnum_text.get('1.0', 'end').split('\n')[0] == '':
        error += 'RP#\n'
    if rp.get() == False or rp.get() == False or nrp.get() == False or rt.get() == False or kj.get() == False:
        error += 'Result\n'
    if bhotswap.get() == False and bvibration.get() == False and bnoise.get() == False and bheat.get() == False and \
        bnff.get() == False:
        error += 'Prevenative Measure\n'
    if c_appearance_text.get('1.0', 'end').split('\n')[0] == '' and bnofaults.get() == False:
        error += 'Appearance Check\n'
    warning(error)
    return None

def suspected_cause():
    str_out = ''
    if bhotswap.get() == True:
        str_out += ', electrical component damage due to hot swap'
    if bvibration.get() == True:
        str_out += ', internal component damage due to vibration or impact damage'
    if bnoise.get() == True:
        str_out += ', electrical component damage due to inductive noise'
    if bheat.get() == True:
        str_out += ', internal component damage due to heat'
    if bnff.get() == True:
        str_out += 'No faults were found.'
    else:
        pass
    return str_out

def preventative_measure():
    str_out = ''
    if bhotswap.get() == True:
        str_out += 'How swap damage could occur when the units power is not fully cycled off. Please make' \
                   ' sure to complete a unit shutdown before unplugging the unit. '
    if bvibration.get() == True:
        str_out += 'To prevent damage from vibration, please make sure to reduce the amount of vibration ' \
                   'the unit is subjected to by introducing vibration dampeners. '
    if bnoise.get() == True:
        str_out += 'To prevent damage from inductive noise, please make sure to keep this unit away from ' \
                   'machinery that can cause electrical noise. '
    if bheat.get() == True:
        str_out += 'To prevent damage from heat, please make sure to provide proper air flow to keep the unit cool. '
    if bnff.get() == True:
        str_out += 'No faults were found.'
    else:
        pass
    return str_out

def makeform():
    rp_save = rpnum_text.get('1.0', 'end').split('\n')[0].strip()
    serial_save = serial_text.get('1.0', 'end').split('\n')[0].strip()
    pdf = FPDF()
    pdf.add_page()
    sum_out = ''
    count = 0
    error = ''

    if rpnum_text.get('1.0', 'end').split('\n')[0] == '' or rp.get() == False or \
            pic_text.get('1.0', 'end').split('\n')[0] == '' or rp.get() == False and nrp.get() == False and \
            rt.get() == False and kj.get() == False or bhotswap.get() == False and bvibration.get() == False and \
            bnoise.get() == False and bheat.get() == False and bnff.get() == False or \
            c_appearance_text.get('1.0', 'end').split('\n')[0] == '' and bnofaults.get() == False:
        checkerror()
        return None
    else:
        pass

    pdf.set_font('Arial', 'I', 18)
    pdf.cell(200, 15, 'FA RSS', ln=1, align='C')
    pdf.set_font('Arial', '', 12)
    pdf.cell(40, 10, 'RP#:' + ' ' + rpnum_text.get('1.0', 'end'))
    pdf.cell(40, 10, 'Serial:' + serial_text.get('1.0', 'end'))
    pdf.cell(40, 10, 'Keycode: ' + key_text.get('1.0', 'end'))
    pdf.cell(40, 10, 'PIC: ' + pic_text.get('1.0', 'end'))
    pdf.cell(40, 10, '', ln=1)
    pdf.cell(40, 10, '', ln=1)
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(40, 10, '< Appearance >', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.cell(40, 10, appearance_text.get('1.0', 'end'), ln=1)
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(40, 10, '', ln=1)
    pdf.cell(40, 10, '< Symptom Information >', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.cell(40, 10, symptom_text.get('1.0', 'end'), ln=1)
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(40, 10, '', ln=1)
    pdf.cell(40, 10, '< Tests Conducted >', ln=1)
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
    pdf.cell(40, 10, '< What caused the failure? >', ln=1)
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

    pdf.set_font('Arial', 'I', 18)
    pdf.cell(175, 15, 'Customer Inspection Report', ln=1, align='C')
    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Appearance Check > ', ln=1)
    pdf.set_font('Arial', '', 12)
    if bnofaults.get() == True:
        pdf.cell(40, 10, 'With a visual inspection, there were no major damage to the exterior of the unit.', ln=1)
    else:
        pdf.cell(40, 10, 'With a visual inspection, ' + c_appearance_text.get('1.0', 'end').lower(), ln=1)
    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Operation Check > ', ln=1)
    pdf.set_font('Arial', '', 12)
    if c_tests_text.get('1.0', 'end') == '':
        if breplicated.get() == True:
            pdf.multi_cell(175, 10, 'With an operation check, the symptom was replicated.')
        else:
            pdf.multi_cell(175, 10, 'With an operation check, '
                                    'the symptom was not replicated.' + c_symptom_text.get('1.0', 'end').capitalize())
    else:
        if breplicated.get() == True:
            pdf.multi_cell(175, 10, 'With an operation check, the symptom was replicated. '
                                    '' + c_symptom_text.get('1.0', 'end').capitalize() + c_tests_text.get('1.0', 'end').capitalize())
        else:
            pdf.multi_cell(175, 10, 'With an operation check, the symptom was not replicated.' + c_symptom_text.get('1.0', 'end').capitalize() + c_tests_text.get('1.0', 'end').capitalize())
    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Internal Check > ', ln=1)
    pdf.set_font('Arial', '', 12)
    if bdisassembled == True:
        pdf.multi_cell(175, 10, 'This unit was disassembled and an internal inspection was conducted.'
                                '' + c_disassembly_text.get('1.0', 'end'))
    else:
        pdf.multi_cell(175, 10, 'There was no reason to disassemble the unit.')

    pdf.set_font('Arial', '', 14)

    pdf.cell(40, 10, '< Suspected Cause > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, 'The suspected cause of failure is likely due to' + suspected_cause() + '.')

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Conclusion > ', ln=1)
    pdf.set_font('Arial', '', 12)
    if rp.get() == True:
        pdf.multi_cell(175, 10, 'In order to bring this unit back to full functionality, the unit will be repaired and'
                                ' tested to meet manufacturing specifications. A full operation and function check'
                                ' will be conducted to confirm normal operation.')
    elif nrp.get() == True:
        pdf.multi_cell(175, 10, 'This unit cannot be repaired as either it is damaged beyond repairing capabilities or '
                                'the repair will approach the cost of a new unit. The unit will be sent back in the '
                                'condition it was received.')
    elif rt.get() == True:
        pdf.multi_cell(175, 10, 'During testing there was no fault found with the unit. The unit will be returned in'
                                ' the condition it was received.')
    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Preventative Measure > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, preventative_measure())

    # pdf.output('I:/PRD/Check Sheets/RSS/' + rp_save + ' - ' + serial_save + '.pdf', 'F')
    pdf.output(rp_save + ' - ' + serial_save + '.pdf', 'F')
    rpnum_text.delete('1.0', 'end')
    serial_text.delete('1.0', 'end')
    key_text.delete('1.0', 'end')
    comp_text.delete('1.0', 'end')
    structure_text.delete('1.0', 'end')
    fb_text.delete('1.0', 'end')


def focus_next_window(event):
    event.widget.tk_focusNext().focus()
    return "break"


def warning(wrds):
    showinfo('Window', 'You are missing:\n' + wrds)


root = tk.Tk()
w, h = root.winfo_screenwidth(), root.winfo_screenheight()
# root.geometry("%dx%d+0+0" % (w, h))
root.geometry('1225x500')
root.title('RSS Generator')

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
bhotswap = tk.BooleanVar()
bvibration = tk.BooleanVar()
bnoise = tk.BooleanVar()
bheat = tk.BooleanVar()
bnofaults = tk.BooleanVar()
breplicated = tk.BooleanVar()
bdisassembled = tk.BooleanVar()
bnff = tk.BooleanVar()

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

key_label = tk.Label(text=' Preventative Measures ')
key_label.grid(row=1, column=5)

hot_swap = tk.Checkbutton(root, text="Hot Swap",
                          onvalue=1, offvalue=0, height=1,
                          width=10, variable=bhotswap)
hot_swap.grid(row=2, column=5)

vibration = tk.Checkbutton(root, text="Vibration",
                           onvalue=1, offvalue=0, height=1,
                           width=10, variable=bvibration)
vibration.grid(row=3, column=5)

noise = tk.Checkbutton(root, text="Electrical Noise",
                       onvalue=1, offvalue=0, height=1,
                       width=10, variable=bnoise)
noise.grid(row=4, column=5)

heat = tk.Checkbutton(root, text="Heat",
                      onvalue=1, offvalue=0, height=1,
                      width=10, variable=bheat)
heat.grid(row=5, column=5)

nff = tk.Checkbutton(root, text="NFF",
                      onvalue=1, offvalue=0, height=1,
                      width=10, variable=bnff)
nff.grid(row=6, column=5)



'''Form'''

appearance = tk.Label(text='Appearance Inspection (Internal)', fg="blue")
appearance.grid(row=9, column=1)
appearance_text = tk.Text(root, height=5, width=30)
appearance_text.grid(row=10, column=1)
appearance_text.bind("<Tab>", focus_next_window)

symptom = tk.Label(text='Symptom Information (Internal)', fg="blue")
symptom.grid(row=9, column=2)
symptom_text = tk.Text(root, height=5, width=30)
symptom_text.grid(row=10, column=2)
symptom_text.bind("<Tab>", focus_next_window)

tests = tk.Label(text='Tests Performed (Internal)', fg="blue")
tests.grid(row=9, column=3)
tests_text = tk.Text(root, height=5, width=30)
tests_text.grid(row=10, column=3)
tests_text.bind("<Tab>", focus_next_window)

analysis = tk.Label(text='Analysis Steps (Internal)', fg="blue")
analysis.grid(row=9, column=4)
analysis_text = tk.Text(root, height=5, width=30)
analysis_text.grid(row=10, column=4)
analysis_text.bind("<Tab>", focus_next_window)

failure = tk.Label(text='Cause of Failure (Internal)', fg="blue")
failure.grid(row=9, column=5)
failure_text = tk.Text(root, height=5, width=30)
failure_text.grid(row=10, column=5)
failure_text.bind("<Tab>", focus_next_window)

# quck_label = tk.Label(text=' Quick Answer ')
# quck_label.grid(row=7, column=1)

no_faults = tk.Checkbutton(root, text="Appearance NFF",
                           onvalue=1, offvalue=0, height=1,
                           width=25, variable=bnofaults)
no_faults.grid(row=11, column=1)

replicated = tk.Checkbutton(root, text="Was the symptom replicated?",
                            onvalue=1, offvalue=0, height=1,
                            width=25, variable=breplicated)
replicated.grid(row=11, column=2)

disassembled = tk.Checkbutton(root, text="Was the unit disassembled?",
                              onvalue=1, offvalue=0, height=1,
                              width=25, variable=bdisassembled)
disassembled.grid(row=11, column=4)

c_appearance = tk.Label(text='Appearance Inspection (for customer)', fg="red")
c_appearance.grid(row=12, column=1)
c_appearance_text = tk.Text(root, height=5, width=30)
c_appearance_text.grid(row=13, column=1)
c_appearance_text.bind("<Tab>", focus_next_window)

c_symptom = tk.Label(text='Symptom Information (for customer)', fg="red")
c_symptom.grid(row=12, column=2)
c_symptom_text = tk.Text(root, height=5, width=30)
c_symptom_text.grid(row=13, column=2)
c_symptom_text.bind("<Tab>", focus_next_window)

c_tests = tk.Label(text='Tests Performed (for customer)', fg="red")
c_tests.grid(row=12, column=3)
c_tests_text = tk.Text(root, height=5, width=30)
c_tests_text.grid(row=13, column=3)
c_tests_text.bind("<Tab>", focus_next_window)

c_disassembly = tk.Label(text='Disassembly (for customer)', fg="red")
c_disassembly.grid(row=12, column=4)
c_disassembly_text = tk.Text(root, height=5, width=30)
c_disassembly_text.grid(row=13, column=4)
c_disassembly_text.bind("<Tab>", focus_next_window)

# c_failure = tk.Label(text='Cause of Failure (for customer)', fg="red")
# c_failure.grid(row=12, column=5)
# c_failure_text = tk.Text(root, height=5, width=30)
# c_failure_text.grid(row=13, column=5)
# c_failure_text.bind("<Tab>", focus_next_window)

# summary = tk.Label(text='Summary')
# summary.grid(row=10, column=3)
# summary_text = tk.Text(root, height=10, width=30)
# summary_text.grid(row=11, column=3)
# summary_text.bind("<Tab>", focus_next_window)

btnRead = tk.Button(root, height=1, width=10, text="Save", command=makeform)
btnRead.grid(row=20, column=1)

root.mainloop()
