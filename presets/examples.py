import gui.gui as gui
from PIL import Image
from zipfile import ZipFile
from presets.clear_example import clearexample
import pyautogui
import time


def example():
    """Inserts example values for the user."""
    if gui.rpnum_text.get('1.0', 'end').split('\n')[0] == '!mvmouse' and \
            gui.item_text.get('1.0', 'end').split('\n')[0] == '!pass123':
        # tm1 = serial_text.get('1.0', 'end').split('\n')[0]
        tm1 = gui.key_text.get('1.0', 'end').split('\n')[0]
        zipfile = ZipFile('python_icons.zip', 'r')
        img1 = Image.open(zipfile.open(name='img1.png', mode='r', pwd=b'!lol54321!'))
        img2 = Image.open(zipfile.open(name='img2.png', mode='r', pwd=b'!lol54321!'))
        img3 = Image.open(zipfile.open(name='img3.png', mode='r', pwd=b'!lol54321!'))
        clearexample()
        pyautogui.FAILSAFE = True
        while True:
            pyautogui.press('numlock')
            if time.ctime().split()[3] == tm1:
                pyautogui.moveTo(pyautogui.locateOnScreen(img2), duration=.75)
                time.sleep(1)
                pyautogui.click()
                return None
            else:
                pass
            time.sleep(1)

    gui.rpnum_text.insert(gui.tk.END, '501501')
    gui.item_text.insert(gui.tk.END, 'IL-600')
    # serial_text.insert(tk.END, '#A5629923')
    gui.key_text.insert(gui.tk.END, '14')
    gui.fb_text.insert(gui.tk.END, 'mother board')
    gui.comp_text.insert(gui.tk.END, 'A110, L605')
    gui.structure_text.insert(gui.tk.END, 'filter glass')
    gui.parts_text.insert(gui.tk.END, 'L605')
    gui.pic_text.insert(gui.tk.END, 'Angel')
    gui.his_rp_text.insert(gui.tk.END, '510510')
    gui.his_comp_date_text.insert(gui.tk.END, '1/1/2015')
    gui.his_symptom_text.insert(gui.tk.END, 'Inductor damage')
    gui.his_result_text.insert(gui.tk.END, 'Repaired')
    gui.appearance_text.insert(gui.tk.END, 'No failures found to the exterior of the unit.')
    gui.symptom_text.insert(gui.tk.END, 'No laser output - confirmed.')
    gui.tests_text.insert(gui.tk.END, 'Full function check, voltage check, current check, impedance check')
    gui.analysis_text.insert(gui.tk.END, '1.Visual inspection passed\n2.Unit symptom confirmed, no laser output'
                                         '\n3. Unit was disassembled and L605 was found in an open state.\n'
                                         '4.A110 was found with visible damage.')
    gui.failure_text.insert(gui.tk.END, 'Over voltage/current from hot swap, surge or noise.')
    gui.bench_text.insert(gui.tk.END, '60')
    gui.bapp_impact.set(True)
    gui.bsym_nolaser.set(True)
    gui.nrp.set(True)
    gui.bhotswap.set(True)
    gui.bnoise.set(True)
    gui.bsurge.set(True)

    gui.c_appearance_text.insert(gui.tk.END, 'There was also dirt found to the exterior of the unit.')
    gui.c_operation_text.insert(gui.tk.END, 'Unit was drawing 1A of current.')
    gui.c_internal_text.insert(gui.tk.END, 'Dirt was found inside the unit.')
    gui.c_suspected_text.insert(gui.tk.END, 'Damage could have also been caused by improper power shut down.')
    gui.c_conclusion_text.insert(gui.tk.END, 'Please contact sales person for additional warranty information.')
    return None
