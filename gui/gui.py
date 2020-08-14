import tkinter as tk
from tkinter.messagebox import showinfo
from presets.clear_example import clearexample
from presets.examples import example
from presets.preset_save import save_exe_preset1, save_exe_preset2, save_exe_preset3
from reportgen.gen_report import form
from openpyxl import load_workbook

serial_list = []
count = 7

"""test"""
def entries():
    """
    Summary: Stores SN entries

    Parameters: None

    Return: A string of multiple serials
    """
    sout = ''
    for s in serial_list:
        sout += s.get('1.0', 'end')
    return sout


def add_box():
    """
    Summary: Adds another box to the list for SN inputs

    Parameters: None

    Return: None
    """
    global count
    if count < 11:
        serial_box = tk.Text(root, height=1, width=20)
        serial_box.grid(row=count, column=1)
        serial_box.bind("<Tab>", focus_next_window)
        count += 1
        serial_list.append(serial_box)
    else:
        pass


def focus_next_window(event):
    """
    Summary: Allows you to use TAB to switch to the next window.

    Parameters: (event)

    Return: None
    """
    event.widget.tk_focusNext().focus()
    return "break"


def fileopen():
    """
    Summary: A warning to alert user that the file trying modify is open.

    Parameters: None

    Return: Warning information
    """
    showinfo('Warning!', 'You have the file open, close it before save.')


def warning(wrds):
    """
    Summary: A warning to alert user that the file trying modify is open.

    Parameters: None

    Return: Warning information
    """
    showinfo('Window', 'You are missing:\n' + wrds)


def copy_message():
    showinfo('Success!', 'Your inspection report has been copied to your clipboard.\n(CTL+V to paste in AS2)')


def too_many_char(incoming):
    showinfo('Warning!', f'Preset character limit is 18,\nyou have {incoming}')


root = tk.Tk()
w, h = root.winfo_screenwidth(), root.winfo_screenheight()
root.geometry("%dx%d+0+0" % (w, h))
# root.geometry('1300x900')
root.title('RSS Generator')

rpnum_label = tk.Label(text='RP#')
rpnum_label.grid(row=1, column=1)
rpnum_text = tk.Text(root, height=1, width=20)
rpnum_text.grid(row=2, column=1)
rpnum_text.bind("<Tab>", focus_next_window)

item_label = tk.Label(text='Item Name')
item_label.grid(row=3, column=1)
item_text = tk.Text(root, height=1, width=20)
item_text.grid(row=4, column=1)
item_text.bind("<Tab>", focus_next_window)

serial_label = tk.Label(text='Serial#')
serial_label.grid(row=5, column=1)
# serial_text = tk.Text(root, height=1, width=20)
# serial_text.grid(row=6, column=1)
# serial_text.bind("<Tab>", focus_next_window)
add_box()
serial_text = tk.Button(root, text='Add Serial Box', fg="Red", command=add_box)
serial_text.grid(row=6, column=1)

key_label = tk.Label(text='Keycode')
key_label.grid(row=1, column=2)
key_text = tk.Text(root, height=1, width=20)
key_text.grid(row=2, column=2)
key_text.bind("<Tab>", focus_next_window)

fb_label = tk.Label(text='Failed Board')
fb_label.grid(row=3, column=2)
fb_text = tk.Text(root, height=1, width=20)
fb_text.grid(row=4, column=2)
fb_text.bind("<Tab>", focus_next_window)

comp_label = tk.Label(text='Location of Component')
comp_label.grid(row=5, column=2)
comp_text = tk.Text(root, height=1, width=20)
comp_text.grid(row=6, column=2)
comp_text.bind("<Tab>", focus_next_window)

structure_label = tk.Label(text='Failed Structure Part')
structure_label.grid(row=1, column=3)
structure_text = tk.Text(root, height=1, width=20)
structure_text.grid(row=2, column=3)
structure_text.bind("<Tab>", focus_next_window)

parts_label = tk.Label(text='Parts Replaced')
parts_label.grid(row=3, column=3)
parts_text = tk.Text(root, height=1, width=20)
parts_text.grid(row=4, column=3)
parts_text.bind("<Tab>", focus_next_window)

pic_label = tk.Label(text='PIC')
pic_label.grid(row=5, column=3)
pic_text = tk.Text(root, height=1, width=20)
pic_text.grid(row=6, column=3)
pic_text.bind("<Tab>", focus_next_window)

bench_label = tk.Label(text='Bench Time')
bench_label.grid(row=1, column=4)
bench_text = tk.Text(root, height=1, width=20)
bench_text.grid(row=2, column=4)
bench_text.bind("<Tab>", focus_next_window)

bench_label = tk.Label(text='Error')
bench_label.grid(row=3, column=4)
bench_text = tk.Text(root, height=1, width=20)
bench_text.grid(row=4, column=4)
bench_text.bind("<Tab>", focus_next_window)

his_label = tk.Label(text='RP History')
his_label.grid(row=1, column=5)

his_rp = tk.Label(text='RP#')
his_rp.grid(row=2, column=5)
his_rp_text = tk.Text(root, height=1, width=20)
his_rp_text.grid(row=3, column=5)
his_rp_text.bind("<Tab>", focus_next_window)

his_comp_date = tk.Label(text='Completion Date')
his_comp_date.grid(row=4, column=5)
his_comp_date_text = tk.Text(root, height=1, width=20)
his_comp_date_text.grid(row=5, column=5)
his_comp_date_text.bind("<Tab>", focus_next_window)

his_symptom = tk.Label(text='Symptom')
his_symptom.grid(row=6, column=5)
his_symptom_text = tk.Text(root, height=1, width=20)
his_symptom_text.grid(row=7, column=5)
his_symptom_text.bind("<Tab>", focus_next_window)

his_result = tk.Label(text='Result')
his_result.grid(row=8, column=5)
his_result_text = tk.Text(root, height=1, width=20)
his_result_text.grid(row=9, column=5)
his_result_text.bind("<Tab>", focus_next_window)

rp = tk.BooleanVar()
nrp = tk.BooleanVar()
rt = tk.BooleanVar()
kj = tk.BooleanVar()
bhotswap = tk.BooleanVar()
bvibration = tk.BooleanVar()
bnoise = tk.BooleanVar()
bheat = tk.BooleanVar()
bsurge = tk.BooleanVar()
bnff = tk.BooleanVar()
bcable = tk.BooleanVar()
bimpact = tk.BooleanVar()

bapp_nff = tk.BooleanVar()
bapp_cable = tk.BooleanVar()
bapp_impact = tk.BooleanVar()
bapp_oil = tk.BooleanVar()
bapp_filterglass = tk.BooleanVar()

bsym_nopower = tk.BooleanVar()
bsym_noint = tk.BooleanVar()
bsym_nocom = tk.BooleanVar()
bsym_nolaser = tk.BooleanVar()
bsym_nffstruct = tk.BooleanVar()
bsym_symnff = tk.BooleanVar()
bsym_cable = tk.BooleanVar()
bsym_inout = tk.BooleanVar()

c_addition = tk.Label(text='Inspection Report', fg="black")
c_addition.grid(row=10, column=3)

key_label = tk.Label(text='Appearance')
key_label.grid(row=11, column=1)

app_cable = tk.Checkbutton(root, text="Cable Damage",
                           onvalue=1, offvalue=0, height=1,
                           width=25, variable=bapp_cable)
app_cable.grid(row=12, column=1)

app_impact = tk.Checkbutton(root, text="Impact Damage",
                            onvalue=1, offvalue=0, height=1,
                            width=25, variable=bapp_impact)
app_impact.grid(row=13, column=1)

app_oil = tk.Checkbutton(root, text="Oil",
                         onvalue=1, offvalue=0, height=1,
                         width=25, variable=bapp_oil)
app_oil.grid(row=14, column=1)

app_filter = tk.Checkbutton(root, text="Filter Glass",
                            onvalue=1, offvalue=0, height=1,
                            width=25, variable=bapp_filterglass)
app_filter.grid(row=15, column=1)

app_nff = tk.Checkbutton(root, text="NFF",
                         onvalue=1, offvalue=0, height=1,
                         width=25, variable=bapp_nff)
app_nff.grid(row=16, column=1)

key_label = tk.Label(text='Symptom')
key_label.grid(row=10, column=2)

nopower = tk.Checkbutton(root, text="No Power",
                         onvalue=1, offvalue=0, height=1,
                         width=25, variable=bsym_nopower)
nopower.grid(row=12, column=2)

nolaser = tk.Checkbutton(root, text="No Laser",
                         onvalue=1, offvalue=0, height=1,
                         width=25, variable=bsym_nolaser)
nolaser.grid(row=13, column=2)

noint = tk.Checkbutton(root, text="No Initialization",
                       onvalue=1, offvalue=0, height=1,
                       width=25, variable=bsym_noint)
noint.grid(row=14, column=2)

nocom = tk.Checkbutton(root, text="No Communication",
                       onvalue=1, offvalue=0, height=1,
                       width=25, variable=bsym_nocom)
nocom.grid(row=15, column=2)

inout = tk.Checkbutton(root, text="Input/Output Error",
                       onvalue=1, offvalue=0, height=1,
                       width=25, variable=bsym_inout)
inout.grid(row=16, column=2)

sym_cable = tk.Checkbutton(root, text="Cable Damage",
                           onvalue=1, offvalue=0, height=1,
                           width=25, variable=bsym_cable)
sym_cable.grid(row=17, column=2)

sym_nffstruct = tk.Checkbutton(root, text="NFF(Structural Damage)",
                               onvalue=1, offvalue=0, height=1,
                               width=25, variable=bsym_nffstruct)
sym_nffstruct.grid(row=10, column=2)

symnff = tk.Checkbutton(root, text="NFF",
                        onvalue=1, offvalue=0, height=1,
                        width=25, variable=bsym_symnff)
symnff.grid(row=11, column=2)

key_label = tk.Label(text='Result')
key_label.grid(row=12, column=3)

rpchk_label = tk.Checkbutton(root, text="Repaired",
                             onvalue=1, offvalue=0, height=1,
                             width=25, variable=rp)
rpchk_label.grid(row=13, column=3)

nrpchk_label = tk.Checkbutton(root, text="Not repaired",
                              onvalue=1, offvalue=0, height=1,
                              width=25, variable=nrp)
nrpchk_label.grid(row=14, column=3)

rtchk_label = tk.Checkbutton(root, text="Returned (NFF)",
                             onvalue=1, offvalue=0, height=1,
                             width=25, variable=rt)
rtchk_label.grid(row=15, column=3)

kjchk_label = tk.Checkbutton(root, text="Sent to KJ(RRC)",
                             onvalue=1, offvalue=0, height=1,
                             width=25, variable=kj)
kjchk_label.grid(row=16, column=3)

key_label = tk.Label(text=' Suspected Cause/Preventative Measures ')
key_label.grid(row=10, column=4)

key_label = tk.Label(text=' Suspected Cause/Preventative Measures ')
key_label.grid(row=10, column=5)

hot_swap = tk.Checkbutton(root, text="Hot Swap",
                          onvalue=1, offvalue=0, height=1,
                          width=25, variable=bhotswap)
hot_swap.grid(row=11, column=4)

noise = tk.Checkbutton(root, text="Electrical Noise",
                       onvalue=1, offvalue=0, height=1,
                       width=25, variable=bnoise)
noise.grid(row=12, column=4)

vibration = tk.Checkbutton(root, text="Vibration",
                           onvalue=1, offvalue=0, height=1,
                           width=25, variable=bvibration)
vibration.grid(row=13, column=4)

surge = tk.Checkbutton(root, text="Electrical Surge",
                       onvalue=1, offvalue=0, height=1,
                       width=25, variable=bsurge)
surge.grid(row=14, column=4)

heat = tk.Checkbutton(root, text="Heat",
                      onvalue=1, offvalue=0, height=1,
                      width=25, variable=bheat)
heat.grid(row=15, column=4)

cable = tk.Checkbutton(root, text="Cable Damage",
                       onvalue=1, offvalue=0, height=1,
                       width=25, variable=bcable)
cable.grid(row=11, column=5)

impact = tk.Checkbutton(root, text="Impact",
                        onvalue=1, offvalue=0, height=1,
                        width=25, variable=bimpact)
impact.grid(row=12, column=5)

nff = tk.Checkbutton(root, text="NFF",
                     onvalue=1, offvalue=0, height=1,
                     width=25, variable=bnff)
nff.grid(row=13, column=5)

c_rss = tk.Label(text='RSS Comments', fg="blue")
c_rss.grid(row=18, column=3)

appearance = tk.Label(text='Appearance Inspection (RSS)', fg="blue")
appearance.grid(row=19, column=1)
appearance_text = tk.Text(root, height=5, width=30)
appearance_text.grid(row=20, column=1)
appearance_text.bind("<Tab>", focus_next_window)

symptom = tk.Label(text='Symptom Information (RSS)', fg="blue")
symptom.grid(row=19, column=2)
symptom_text = tk.Text(root, height=5, width=30)
symptom_text.grid(row=20, column=2)
symptom_text.bind("<Tab>", focus_next_window)

tests = tk.Label(text='Tests Performed (RSS)', fg="blue")
tests.grid(row=19, column=3)
tests_text = tk.Text(root, height=5, width=30)
tests_text.grid(row=20, column=3)
tests_text.bind("<Tab>", focus_next_window)

analysis = tk.Label(text='Analysis Steps (RSS)', fg="blue")
analysis.grid(row=19, column=4)
analysis_text = tk.Text(root, height=5, width=30)
analysis_text.grid(row=20, column=4)
analysis_text.bind("<Tab>", focus_next_window)

failure = tk.Label(text='Cause of Failure (RSS)', fg="blue")
failure.grid(row=19, column=5)
failure_text = tk.Text(root, height=5, width=30)
failure_text.grid(row=20, column=5)
failure_text.bind("<Tab>", focus_next_window)

c_addition = tk.Label(text='Additional Comments for Inspection Report', fg="red")
c_addition.grid(row=24, column=3)

c_appearance = tk.Label(text='Appearance Check', fg="red")
c_appearance.grid(row=25, column=1)
c_appearance_text = tk.Text(root, height=5, width=30)
c_appearance_text.grid(row=26, column=1)
c_appearance_text.bind("<Tab>", focus_next_window)

c_operation = tk.Label(text='Operation Check', fg="red")
c_operation.grid(row=25, column=2)
c_operation_text = tk.Text(root, height=5, width=30)
c_operation_text.grid(row=26, column=2)
c_operation_text.bind("<Tab>", focus_next_window)

c_internal = tk.Label(text='Internal Check', fg="red")
c_internal.grid(row=25, column=3)
c_internal_text = tk.Text(root, height=5, width=30)
c_internal_text.grid(row=26, column=3)
c_internal_text.bind("<Tab>", focus_next_window)

c_suspected = tk.Label(text='Suspected Cause', fg="red")
c_suspected.grid(row=25, column=4)
c_suspected_text = tk.Text(root, height=5, width=30)
c_suspected_text.grid(row=26, column=4)
c_suspected_text.bind("<Tab>", focus_next_window)

c_conclusion = tk.Label(text='Conclusion', fg="red")
c_conclusion.grid(row=25, column=5)
c_conclusion_text = tk.Text(root, height=5, width=30)
c_conclusion_text.grid(row=26, column=5)
c_conclusion_text.bind("<Tab>", focus_next_window)

preset_label = tk.Label(text='Presets ("clear" for reset)')
preset_label.grid(row=33, column=4)
preset_text = tk.Text(root, height=1, width=20)
preset_text.grid(row=34, column=4)
preset_text.bind("<Tab>", focus_next_window)

lb = load_workbook(filename='presets.xlsx')
ws = lb.active
sheet_ranges = lb['Sheet']

btn_preset_text = tk.StringVar()
btn_preset_text.set(sheet_ranges['A1'].value)
btn_preset1 = tk.Button(root, height=1, width=15, textvariable=btn_preset_text, command=save_exe_preset1)
btn_preset1.grid(row=33, column=5)

btn_preset2_text = tk.StringVar()
btn_preset2_text.set(sheet_ranges['B1'].value)
btn_preset2 = tk.Button(root, height=1, width=15, textvariable=btn_preset2_text, command=save_exe_preset2)
btn_preset2.grid(row=34, column=5)

btn_preset3_text = tk.StringVar()
btn_preset3_text.set(sheet_ranges['C1'].value)
btn_preset3 = tk.Button(root, height=1, width=15, textvariable=btn_preset3_text, command=save_exe_preset3)
btn_preset3.grid(row=35, column=5)

btnRead = tk.Button(root, height=1, width=15, text="Save", command=form)
btnRead.grid(row=33, column=1)

btn_example = tk.Button(root, height=1, width=15, text="Example", command=example)
btn_example.grid(row=34, column=1)

btn_example = tk.Button(root, height=1, width=15, text="Clear All Values", command=clearexample)
btn_example.grid(row=33, column=2)

btn_example = tk.Button(root, height=1, width=15, text="Help", command=clearexample)
btn_example.grid(row=34, column=2)

angel_label = tk.Label(text='Created by Angel Davila')
angel_label.grid(row=35, column=1)
