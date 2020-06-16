import tkinter as tk
import reportlab


def getTextInput():
    result = rpnum_text.get("1.0", "end")
    print(result)


root = tk.Tk()
w, h = root.winfo_screenwidth(), root.winfo_screenheight()
# root.geometry("%dx%d+0+0" % (w, h))
root.geometry('900x600')
root.title('RSS Sheet Generator')

rpnum_label = tk.Label(text='RP#')
rpnum_label.grid(row=1, column=1)
rpnum_text = tk.Text(root, height=1, width=10)
rpnum_text.grid(row=2, column=1)

item_label = tk.Label(text='Item Name')
item_label.grid(row=3, column=1)
item_text = tk.Text(root, height=1, width=10)
item_text.grid(row=4, column=1)

serial_label = tk.Label(text='Serial#')
serial_label.grid(row=5, column=1)
serial_text = tk.Text(root, height=1, width=10)
serial_text.grid(row=6, column=1)

key_label = tk.Label(text='Keycode')
key_label.grid(row=1, column=2)
key_text = tk.Text(root, height=1, width=10)
key_text.grid(row=2, column=2)

fb_label = tk.Label(text='Failed Board')
fb_label.grid(row=3, column=2)
fb_text = tk.Text(root, height=1, width=10)
fb_text.grid(row=4, column=2)

comp_label = tk.Label(text='Location of Component')
comp_label.grid(row=5, column=2)
comp_text = tk.Text(root, height=1, width=10)
comp_text.grid(row=6, column=2)

structure_label = tk.Label(text='Failed Structure Part')
structure_label.grid(row=1, column=3)
structure_text = tk.Text(root, height=1, width=10)
structure_text.grid(row=2, column=3)

pic_label = tk.Label(text='PIC')
pic_label.grid(row=3, column=3)
pic_text = tk.Text(root, height=1, width=10)
pic_text.grid(row=4, column=3)

''' Check Boxes'''

key_label = tk.Label(text='Result')
key_label.grid(row=1, column=4)

rpchk_label = tk.Checkbutton(root, text="Repaired",
                             onvalue=1, offvalue=0, height=1,
                             width=10)
rpchk_label.grid(row=2, column=4)

nrpchk_label = tk.Checkbutton(root, text="Not repaired",
                              onvalue=1, offvalue=0, height=1,
                              width=10)
nrpchk_label.grid(row=3, column=4)

rtchk_label = tk.Checkbutton(root, text="Returned (NFF)",
                             onvalue=1, offvalue=0, height=1,
                             width=10)
rtchk_label.grid(row=4, column=4)

kjchk_label = tk.Checkbutton(root, text="Sent to KJ",
                             onvalue=1, offvalue=0, height=1,
                             width=10)
kjchk_label.grid(row=5, column=4)

'''Form'''

apperance = tk.Label(text='Apperance Inspection')
apperance.grid(row=8, column=1)
apperance_text = tk.Text(root, height=10, width=30)
apperance_text.grid(row=9, column=1)

symptom = tk.Label(text='Symptom Information')
symptom.grid(row=8, column=2)
symptom_text = tk.Text(root, height=10, width=30)
symptom_text.grid(row=9, column=2)

tests = tk.Label(text='Tests Performed')
tests.grid(row=8, column=3)
tests_text = tk.Text(root, height=10, width=30)
tests_text.grid(row=9, column=3)

analysis = tk.Label(text='Analysis Steps')
analysis.grid(row=10, column=1)
analysis_text = tk.Text(root, height=10, width=30)
analysis_text.grid(row=11, column=1)

failure = tk.Label(text='Cause of Failure')
failure.grid(row=10, column=2)
failure_text = tk.Text(root, height=10, width=30)
failure_text.grid(row=11, column=2)

failure = tk.Label(text='Summary')
failure.grid(row=10, column=3)
failure_text = tk.Text(root, height=10, width=30)
failure_text.grid(row=11, column=3)

btnRead = tk.Button(root, height=1, width=10, text="Read", command=getTextInput)
btnRead.grid(row=20, column=1)


root.mainloop()
