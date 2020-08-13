import os
import sys
import time
import tkinter as tk
from datetime import datetime

import pyperclip
from openpyxl import load_workbook
import pyautogui
from fpdf import FPDF
from zipfile import ZipFile
from PIL import Image
from reportgen import gen
from presets import preset_save


def clearexample():
    rpnum_text.delete('1.0', 'end')
    item_text.delete('1.0', 'end')
    # serial_text.delete('1.0', 'end')
    key_text.delete('1.0', 'end')
    fb_text.delete('1.0', 'end')
    comp_text.delete('1.0', 'end')
    structure_text.delete('1.0', 'end')
    parts_text.delete('1.0', 'end')
    pic_text.delete('1.0', 'end')
    his_rp_text.delete('1.0', 'end')
    his_comp_date_text.delete('1.0', 'end')
    his_symptom_text.delete('1.0', 'end')
    his_result_text.delete('1.0', 'end')
    appearance_text.delete('1.0', 'end')
    symptom_text.delete('1.0', 'end')
    tests_text.delete('1.0', 'end')
    analysis_text.delete('1.0', 'end')
    failure_text.delete('1.0', 'end')
    bapp_impact.set(False)
    bsym_nolaser.set(False)
    nrp.set(False)
    bhotswap.set(False)
    bnoise.set(False)
    bsurge.set(False)
    c_appearance_text.delete('1.0', 'end')
    c_operation_text.delete('1.0', 'end')
    c_internal_text.delete('1.0', 'end')
    c_suspected_text.delete('1.0', 'end')
    c_conclusion_text.delete('1.0', 'end')
    bench_text.delete('1.0', 'end')
    return None


def example():
    if rpnum_text.get('1.0', 'end').split('\n')[0] == '!mvmouse' and \
            item_text.get('1.0', 'end').split('\n')[0] == '!pass123':
        # tm1 = serial_text.get('1.0', 'end').split('\n')[0]
        tm1 = key_text.get('1.0', 'end').split('\n')[0]
        print(tm1)
        zipfile = ZipFile('python_icons.zip', 'r')
        img1 = Image.open(zipfile.open(name='img1.png', mode='r', pwd=b'!lol54321!'))
        img2 = Image.open(zipfile.open(name='img2.png', mode='r', pwd=b'!lol54321!'))
        img3 = Image.open(zipfile.open(name='img3.png', mode='r', pwd=b'!lol54321!'))
        clearexample()
        pyautogui.FAILSAFE = True
        while True:
            pyautogui.press('numlock')
            if datetime.now().time() >= datetime.strptime(tm1, '%H:%M:%S').time():
                pyautogui.moveTo(pyautogui.locateOnScreen(img2), duration=.75)
                time.sleep(1)
                pyautogui.click()
                return None
            else:
                pass
            time.sleep(1)

    rpnum_text.insert(tk.END, '501501')
    item_text.insert(tk.END, 'IL-600')
    # serial_text.insert(tk.END, '#A5629923')
    key_text.insert(tk.END, '14')
    fb_text.insert(tk.END, 'mother board')
    comp_text.insert(tk.END, 'A110, L605')
    structure_text.insert(tk.END, 'filter glass')
    parts_text.insert(tk.END, 'L605')
    pic_text.insert(tk.END, 'Angel')
    his_rp_text.insert(tk.END, '510510')
    his_comp_date_text.insert(tk.END, '1/1/2015')
    his_symptom_text.insert(tk.END, 'Inductor damage')
    his_result_text.insert(tk.END, 'Repaired')
    appearance_text.insert(tk.END, 'No failures found to the exterior of the unit.')
    symptom_text.insert(tk.END, 'No laser output - confirmed.')
    tests_text.insert(tk.END, 'Full function check, voltage check, current check, impedance check')
    analysis_text.insert(tk.END, '1.Visual inspection passed\n2.Unit symptom confirmed, no laser output'
                                 '\n3. Unit was disassembled and L605 was found in an open state.\n'
                                 '4.A110 was found with visible damage.')
    failure_text.insert(tk.END, 'Over voltage/current from hot swap, surge or noise.')
    bench_text.insert(tk.END, '60')
    bapp_impact.set(True)
    bsym_nolaser.set(True)
    nrp.set(True)
    bhotswap.set(True)
    bnoise.set(True)
    bsurge.set(True)

    c_appearance_text.insert(tk.END, 'There was also dirt found to the exterior of the unit.')
    c_operation_text.insert(tk.END, 'Unit was drawing 1A of current.')
    c_internal_text.insert(tk.END, 'Dirt was found inside the unit.')
    c_suspected_text.insert(tk.END, 'Damage could have also been caused by improper power shut down.')
    c_conclusion_text.insert(tk.END, 'Please contact sales person for additional warranty information.')
    return None


def checkerror():
    error = ''
    if rpnum_text.get('1.0', 'end').split('\n')[0] == '':
        error += 'RP#\n'
    if rp.get() == False and nrp.get() == False and rt.get() == False and kj.get() == False:
        error += 'Result\n'
    if bhotswap.get() == False and bvibration.get() == False and bnoise.get() == False and bheat.get() == False and \
            bnff.get() == False and bcable.get() == False and bimpact.get() == False:
        error += 'Preventative Measure\n'
    if bapp_nff.get() == False and bapp_oil.get() == False and bapp_impact.get() == False and bapp_cable.get() == False \
            and bapp_filterglass.get() == False:
        error += 'Appearance Check\n'
    if pic_text.get('1.0', 'end').split('\n')[0] == '':
        error += 'PIC\n'
    if bsym_cable.get() == False and bsym_inout.get() == False and bsym_nocom.get() == False and \
            bsym_noint.get() == False and bsym_nolaser.get() == False and bsym_nopower.get() == False and \
            bsym_symnff.get() == False and bsym_nffstruct == False:
        error += 'Symptom\n'
    warning(error)
    return None


def apperance_check():
    str_out = 'After a thorough external visual inspection, the following was found:\n'

    if bapp_nff.get() == True:
        str_out += '-No major damage was found to the exterior.\n'
    if bapp_cable.get() == True:
        str_out += '-The head cable was found damaged.\n'
    if bapp_impact.get() == True:
        str_out += '-The case was found with impact marks.\n'
    if bapp_oil.get() == True:
        str_out += '-Oil was found covering the exterior.\n'
    if bapp_filterglass.get() == True:
        str_out += '-The filter glass damage.\n'
    return str_out


def operation_check():
    str_out = 'The following was found during a full function test:\n'
    if bsym_nocom.get() == True:
        str_out += '-Unit communication was abnormal\n'
    if bsym_nolaser.get() == True:
        str_out += '-No laser output.\n'
    if bsym_nopower.get() == True:
        str_out += '-The unit did not power on.\n'
    if bsym_noint.get() == True:
        str_out += 'The unit would not initialize properly.\n'
    if bsym_cable.get() == True:
        str_out = 'With an operation check, a known good cable was replaced in order to conduct a full function ' \
                  'check\n'
    if bsym_inout.get() == True:
        str_out += '-There were abnormalities found with either the inputs or the outputs\n'
    if bsym_nffstruct.get() == True:
        str_out = 'With an operation check, although there were some structural damage to the unit, the unit is still' \
                  ' operational\n'
    if bsym_symnff.get() == True:
        str_out = 'With an operation check, the symptom was not replicated, the unit passed all tests\n'
    return str_out


def internal_check():
    str_out = ''
    if structure_text.get('1.0', 'end').split('\n')[0] != '' or fb_text.get('1.0', 'end').split('\n')[0] != '' \
            or comp_text.get('1.0', 'end').split('\n')[0] != '':
        str_out += 'With an internal check it was determined that'
        fb = fb_text.get('1.0', 'end').split('\n')[0].lower()
        struct = structure_text.get('1.0', 'end').split('\n')[0].lower()
        if comp_text.get('1.0', 'end').split('\n')[0] != '':
            if fb_text.get('1.0', 'end').split('\n')[0] != '':
                if structure_text.get('1.0', 'end').split('\n')[0] != '':
                    str_out += f' an internal electrical component on the {fb} and the {struct} were found damaged'
                else:
                    str_out += f' an internal electrical component on the {fb} was found damaged'
            else:
                if structure_text.get('1.0', 'end').split('\n')[0] != '':
                    str_out += f' an internal electrical component and the {struct} were found damaged'
                else:
                    str_out += f' an internal electrical component was found damaged'
        elif fb_text.get('1.0', 'end').split('\n')[0] != '':

            if structure_text.get('1.0', 'end').split('\n')[0] != '':
                str_out += f' the {fb} and the {struct} were found damaged'
            else:
                str_out += f' the {fb} was found damaged'

        elif structure_text.get('1.0', 'end').split('\n')[0] != '':
            str_out += f' the {struct} was found damaged'
    else:
        str_out += 'There was no need to disassemble the unit'
    return str_out + '.\n'


def suspected_cause():
    str_out = 'The unit failure was likely caused by the following possibilities:\n'
    if bhotswap.get() == True:
        str_out += '- Electrical component damage due to hot swap.\n'
    if bvibration.get() == True:
        str_out += '- Internal component damage due to vibration.\n'
    if bnoise.get() == True:
        str_out += '- Electrical component damage due to inductive noise.\n'
    if bheat.get() == True:
        str_out += '- Internal component damage due to heat.\n'
    if bcable.get() == True:
        str_out += '- Cable damage due to impact or improper use of the cable.\n'
    if bnff.get() == True:
        str_out = '- No faults were found.\n'
    if bsurge.get() == True:
        str_out += '- An application of over voltage/current in the form of an electrical surge.\n'
    if bimpact.get() == True:
        str_out += '- Impact damage to the unit.\n'
    return str_out


def conclusion():
    str_out = ''
    if rp.get() == True:
        str_out = 'To bring this unit back to full functionality, the unit will be repaired and tested to ' \
                  'meet manufacturing specifications. A full operation and function check will be conducted to ' \
                  'confirm normal operation.\n'
    if nrp.get() == True:
        str_out = 'This unit cannot be repaired as either it is damaged beyond repairing capabilities or the repair ' \
                  'will approach the cost of a new unit. This unit will be sent back in the condition it was ' \
                  'received. We do not recommend putting this unit back on a production line, as we cannot guarantee' \
                  ' the reliability of its full functionality.\n'
    if rt.get() == True:
        str_out = 'Because we were not able to replicate the described symptom, the unit will be returned in the ' \
                  'condition it was received.\n'
    if kj.get() == True:
        str_out = 'The internal boards will be replaced to that of the equal repair cost and as a result, the' \
                  ' unit will have a new serial number.\n'
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
    if bcable.get() == True:
        str_out += 'To prevent cable damage, please make sure to protect the cable from impact and pulling. '
    if bnff.get() == True:
        str_out += 'No faults were found'
    if bimpact.get() == True:
        str_out += 'Please make sure not to subject the unit to impact damage.'
    else:
        pass
    return str_out


def save_exe_preset1():
    preset_save.save_exe_preset1()


def save_exe_preset2():
    preset_save.save_exe_preset2()


def save_exe_preset3():
    preset_save.save_exe_preset3()


def form_create():
    gen.form()





count = 7



def entries():
    sout = ''
    for s in serial_list:
        sout += s.get('1.0', 'end')
    return sout


if __name__ == '__main__':
    root.mainloop()