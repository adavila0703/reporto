# import app as app


def save_exe_preset1():
    if len(app.preset_text.get('1.0', 'end').split('\n')[0]) > 18:
        app.too_many_char(len(app.preset_text.get('1.0', 'end').split('\n')[0]))
        return None
    else:
        pass

    if app.preset_text.get('1.0', 'end').split('\n')[0] == '':
        lb = app.load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'A'
        app.rpnum_text.insert(app.tk.END, sheet_ranges[f'{letter}2'].value)
        app.item_text.insert(app.tk.END, sheet_ranges[f'{letter}3'].value)
        # serial_text.insert(tk.END, sheet_ranges[f'{letter}4'].value)
        app.key_text.insert(app.tk.END, sheet_ranges[f'{letter}5'].value)
        app.fb_text.insert(app.tk.END, sheet_ranges[f'{letter}6'].value)
        app.comp_text.insert(app.tk.END, sheet_ranges[f'{letter}7'].value)
        app.structure_text.insert(app.tk.END, sheet_ranges[f'{letter}8'].value)
        app.parts_text.insert(app.tk.END, sheet_ranges[f'{letter}9'].value)
        app.pic_text.insert(app.tk.END, sheet_ranges[f'{letter}10'].value)
        app.his_rp_text.insert(app.tk.END, sheet_ranges[f'{letter}11'].value)
        app.his_comp_date_text.insert(app.tk.END, sheet_ranges[f'{letter}12'].value)
        app.his_symptom_text.insert(app.tk.END, sheet_ranges[f'{letter}13'].value)
        app.his_result_text.insert(app.tk.END, sheet_ranges[f'{letter}14'].value)
        app.appearance_text.insert(app.tk.END, sheet_ranges[f'{letter}15'].value)
        app.symptom_text.insert(app.tk.END, sheet_ranges[f'{letter}16'].value)
        app.tests_text.insert(app.tk.END, sheet_ranges[f'{letter}17'].value)
        app.analysis_text.insert(app.tk.END, sheet_ranges[f'{letter}18'].value)
        app.failure_text.insert(app.tk.END, sheet_ranges[f'{letter}19'].value)
        if str(sheet_ranges[f'{letter}20'].value) == 'True':
            app.bapp_impact.set(True)
        else:
            app.bapp_impact.set(False)
        if str(sheet_ranges[f'{letter}21'].value) == 'True':
            app.bapp_cable.set(True)
        else:
            app.bapp_cable.set(False)
        if str(sheet_ranges[f'{letter}22'].value) == 'True':
            app.bapp_filterglass.set(True)
        else:
            app.bapp_filterglass.set(False)
        if str(sheet_ranges[f'{letter}23'].value) == 'True':
            app.bapp_nff.set(True)
        else:
            app.bapp_nff.set(False)
        if str(sheet_ranges[f'{letter}24'].value) == 'True':
            app.bapp_oil.set(True)
        else:
            app.bapp_oil.set(False)
        if str(sheet_ranges[f'{letter}25'].value) == 'True':
            app.bsym_cable.set(True)
        else:
            app.bsym_cable.set(False)
        if str(sheet_ranges[f'{letter}26'].value) == 'True':
            app.bsym_inout.set(True)
        else:
            app.bsym_inout.set(False)
        if str(sheet_ranges[f'{letter}27'].value) == 'True':
            app.bsym_nffstruct.set(True)
        else:
            app.bsym_nffstruct.set(False)
        if str(sheet_ranges[f'{letter}28'].value) == 'True':
            app.bsym_nocom.set(True)
        else:
            app.bsym_nocom.set(False)
        if str(sheet_ranges[f'{letter}29'].value) == 'True':
            app.bsym_noint.set(True)
        else:
            app.bsym_noint.set(False)
        if str(sheet_ranges[f'{letter}30'].value) == 'True':
            app.bsym_nolaser.set(True)
        else:
            app.bsym_nolaser.set(False)
        if str(sheet_ranges[f'{letter}31'].value) == 'True':
            app.bsym_nopower.set(True)
        else:
            app.bsym_nopower.set(False)
        if str(sheet_ranges[f'{letter}32'].value) == 'True':
            app.bsym_symnff.set(True)
        else:
            app.bsym_symnff.set(False)
        if str(sheet_ranges[f'{letter}33'].value) == 'True':
            app.rp.set(True)
        else:
            app.rp.set(False)
        if str(sheet_ranges[f'{letter}34'].value) == 'True':
            app.nrp.set(True)
        else:
            app.nrp.set(False)
        if str(sheet_ranges[f'{letter}35'].value) == 'True':
            app.rt.set(True)
        else:
            app.rt.set(False)
        if str(sheet_ranges[f'{letter}36'].value) == 'True':
            app.kj.set(True)
        else:
            app.kj.set(False)
        if str(sheet_ranges[f'{letter}37'].value) == 'True':
            app.bhotswap.set(True)
        else:
            app.bhotswap.set(False)
        if str(sheet_ranges[f'{letter}38'].value) == 'True':
            app.bvibration.set(True)
        else:
            app.bvibration.set(False)
        if str(sheet_ranges[f'{letter}39'].value) == 'True':
            app.bnoise.set(True)
        else:
            app.bnoise.set(False)
        if str(sheet_ranges[f'{letter}40'].value) == 'True':
            app.bheat.set(True)
        else:
            app.bheat.set(False)
        if str(sheet_ranges[f'{letter}41'].value) == 'True':
            app.bnff.set(True)
        else:
            app.bnff.set(False)
        if str(sheet_ranges[f'{letter}42'].value) == 'True':
            app.bcable.set(True)
        else:
            app.bcable.set(False)
        if str(sheet_ranges[f'{letter}43'].value) == 'True':
            app.bsym_symnff.set(True)
        else:
            app.bsym_symnff.set(False)
        if str(sheet_ranges[f'{letter}44'].value) == 'True':
            app.bsym_cable.set(True)
        else:
            app.bsym_cable.set(False)
        if str(sheet_ranges[f'{letter}45'].value) == 'True':
            app.bsurge.set(True)
        else:
            app.bsurge.set(False)
        app.c_appearance_text.insert(app.tk.END, sheet_ranges[f'{letter}46'].value)
        app.c_operation_text.insert(app.tk.END, sheet_ranges[f'{letter}47'].value)
        app.c_internal_text.insert(app.tk.END, sheet_ranges[f'{letter}48'].value)
        app.c_suspected_text.insert(app.tk.END, sheet_ranges[f'{letter}49'].value)
        app.c_conclusion_text.insert(app.tk.END, sheet_ranges[f'{letter}50'].value)
        app.bench_text.insert(app.tk.END, sheet_ranges[f'{letter}51'].value)
    elif app.preset_text.get('1.0', 'end').split('\n')[0] == 'clear':
        lb = app.load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'A'
        app.btn_preset_text.set('Preset 1')
        app.preset_text.delete('1.0', 'end')

        sheet_ranges[f'{letter}1'].value = 'Preset 1'
        for a in range(2, 51):
            sheet_ranges[f'{letter}{a}'].value = ''
        lb.save('presets.xlsx')
    else:
        lb = app.load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'A'

        sheet_ranges[f'{letter}1'].value = app.preset_text.get('1.0', 'end').split('\n')[0]
        app.btn_preset_text.set(sheet_ranges[f'{letter}1'].value)
        app.preset_text.delete('1.0', 'end')

        sheet_ranges[f'{letter}2'].value = app.rpnum_text.get('1.0', 'end')
        sheet_ranges[f'{letter}3'].value = app.item_text.get('1.0', 'end')
        # sheet_ranges[f'{letter}4'].value = serial_text.get('1.0', 'end')
        sheet_ranges[f'{letter}5'].value = app.key_text.get('1.0', 'end')
        sheet_ranges[f'{letter}6'].value = app.fb_text.get('1.0', 'end')
        sheet_ranges[f'{letter}7'].value = app.comp_text.get('1.0', 'end')
        sheet_ranges[f'{letter}8'].value = app.structure_text.get('1.0', 'end')
        sheet_ranges[f'{letter}9'].value = app.parts_text.get('1.0', 'end')
        sheet_ranges[f'{letter}10'].value = app.pic_text.get('1.0', 'end')
        sheet_ranges[f'{letter}11'].value = app.his_rp_text.get('1.0', 'end')
        sheet_ranges[f'{letter}12'].value = app.his_comp_date_text.get('1.0', 'end')
        sheet_ranges[f'{letter}13'].value = app.his_symptom_text.get('1.0', 'end')
        sheet_ranges[f'{letter}14'].value = app.his_result_text.get('1.0', 'end')
        sheet_ranges[f'{letter}15'].value = app.appearance_text.get('1.0', 'end')
        sheet_ranges[f'{letter}16'].value = app.symptom_text.get('1.0', 'end')
        sheet_ranges[f'{letter}17'].value = app.tests_text.get('1.0', 'end')
        sheet_ranges[f'{letter}18'].value = app.analysis_text.get('1.0', 'end')
        sheet_ranges[f'{letter}19'].value = app.failure_text.get('1.0', 'end')
        sheet_ranges[f'{letter}20'].value = app.bapp_impact.get()
        sheet_ranges[f'{letter}21'].value = app.bapp_cable.get()
        sheet_ranges[f'{letter}22'].value = app.bapp_filterglass.get()
        sheet_ranges[f'{letter}23'].value = app.bapp_nff.get()
        sheet_ranges[f'{letter}24'].value = app.bapp_oil.get()
        sheet_ranges[f'{letter}25'].value = app.bsym_cable.get()
        sheet_ranges[f'{letter}26'].value = app.bsym_inout.get()
        sheet_ranges[f'{letter}27'].value = app.bsym_nffstruct.get()
        sheet_ranges[f'{letter}28'].value = app.bsym_nocom.get()
        sheet_ranges[f'{letter}29'].value = app.bsym_noint.get()
        sheet_ranges[f'{letter}30'].value = app.bsym_nolaser.get()
        sheet_ranges[f'{letter}31'].value = app.bsym_nopower.get()
        sheet_ranges[f'{letter}32'].value = app.bsym_symnff.get()
        sheet_ranges[f'{letter}33'].value = app.rp.get()
        sheet_ranges[f'{letter}34'].value = app.nrp.get()
        sheet_ranges[f'{letter}35'].value = app.rt.get()
        sheet_ranges[f'{letter}36'].value = app.kj.get()
        sheet_ranges[f'{letter}37'].value = app.bhotswap.get()
        sheet_ranges[f'{letter}38'].value = app.bvibration.get()
        sheet_ranges[f'{letter}39'].value = app.bnoise.get()
        sheet_ranges[f'{letter}40'].value = app.bheat.get()
        sheet_ranges[f'{letter}41'].value = app.bnff.get()
        sheet_ranges[f'{letter}42'].value = app.bcable.get()
        sheet_ranges[f'{letter}43'].value = app.bsym_symnff.get()
        sheet_ranges[f'{letter}44'].value = app.bsym_cable.get()
        sheet_ranges[f'{letter}45'].value = app.bsurge.get()
        sheet_ranges[f'{letter}46'].value = app.c_appearance_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}47'].value = app.c_operation_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}48'].value = app.c_internal_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}49'].value = app.c_suspected_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}50'].value = app.c_conclusion_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}51'].value = app.bench_text.get('1.0', 'end').split('\n')[0]
        lb.save('presets.xlsx')
    return None


def save_exe_preset2():
    if app.preset_text.get('1.0', 'end').split('\n')[0] == '':
        lb = app.load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'B'
        app.rpnum_text.insert(app.tk.END, sheet_ranges[f'{letter}2'].value)
        app.item_text.insert(app.tk.END, sheet_ranges[f'{letter}3'].value)
        # serial_text.insert(tk.END, sheet_ranges[f'{letter}4'].value)
        app.key_text.insert(app.tk.END, sheet_ranges[f'{letter}5'].value)
        app.fb_text.insert(app.tk.END, sheet_ranges[f'{letter}6'].value)
        app.comp_text.insert(app.tk.END, sheet_ranges[f'{letter}7'].value)
        app.structure_text.insert(app.tk.END, sheet_ranges[f'{letter}8'].value)
        app.parts_text.insert(app.tk.END, sheet_ranges[f'{letter}9'].value)
        app.pic_text.insert(app.tk.END, sheet_ranges[f'{letter}10'].value)
        app.his_rp_text.insert(app.tk.END, sheet_ranges[f'{letter}11'].value)
        app.his_comp_date_text.insert(app.tk.END, sheet_ranges[f'{letter}12'].value)
        app.his_symptom_text.insert(app.tk.END, sheet_ranges[f'{letter}13'].value)
        app.his_result_text.insert(app.tk.END, sheet_ranges[f'{letter}14'].value)
        app.appearance_text.insert(app.tk.END, sheet_ranges[f'{letter}15'].value)
        app.symptom_text.insert(app.tk.END, sheet_ranges[f'{letter}16'].value)
        app.tests_text.insert(app.tk.END, sheet_ranges[f'{letter}17'].value)
        app.analysis_text.insert(app.tk.END, sheet_ranges[f'{letter}18'].value)
        app.failure_text.insert(app.tk.END, sheet_ranges[f'{letter}19'].value)
        if str(sheet_ranges[f'{letter}20'].value) == 'True':
            app.bapp_impact.set(True)
        else:
            app.bapp_impact.set(False)
        if str(sheet_ranges[f'{letter}21'].value) == 'True':
            app.bapp_cable.set(True)
        else:
            app.bapp_cable.set(False)
        if str(sheet_ranges[f'{letter}22'].value) == 'True':
            app.bapp_filterglass.set(True)
        else:
            app.bapp_filterglass.set(False)
        if str(sheet_ranges[f'{letter}23'].value) == 'True':
            app.bapp_nff.set(True)
        else:
            app.bapp_nff.set(False)
        if str(sheet_ranges[f'{letter}24'].value) == 'True':
            app.bapp_oil.set(True)
        else:
            app.bapp_oil.set(False)
        if str(sheet_ranges[f'{letter}25'].value) == 'True':
            app.bsym_cable.set(True)
        else:
            app.bsym_cable.set(False)
        if str(sheet_ranges[f'{letter}26'].value) == 'True':
            app.bsym_inout.set(True)
        else:
            app.bsym_inout.set(False)
        if str(sheet_ranges[f'{letter}27'].value) == 'True':
            app.bsym_nffstruct.set(True)
        else:
            app.bsym_nffstruct.set(False)
        if str(sheet_ranges[f'{letter}28'].value) == 'True':
            app.bsym_nocom.set(True)
        else:
            app.bsym_nocom.set(False)
        if str(sheet_ranges[f'{letter}29'].value) == 'True':
            app.bsym_noint.set(True)
        else:
            app.bsym_noint.set(False)
        if str(sheet_ranges[f'{letter}30'].value) == 'True':
            app.bsym_nolaser.set(True)
        else:
            app.bsym_nolaser.set(False)
        if str(sheet_ranges[f'{letter}31'].value) == 'True':
            app.bsym_nopower.set(True)
        else:
            app.bsym_nopower.set(False)
        if str(sheet_ranges[f'{letter}32'].value) == 'True':
            app.bsym_symnff.set(True)
        else:
            app.bsym_symnff.set(False)
        if str(sheet_ranges[f'{letter}33'].value) == 'True':
            app.rp.set(True)
        else:
            app.rp.set(False)
        if str(sheet_ranges[f'{letter}34'].value) == 'True':
            app.nrp.set(True)
        else:
            app.nrp.set(False)
        if str(sheet_ranges[f'{letter}35'].value) == 'True':
            app.rt.set(True)
        else:
            app.rt.set(False)
        if str(sheet_ranges[f'{letter}36'].value) == 'True':
            app.kj.set(True)
        else:
            app.kj.set(False)
        if str(sheet_ranges[f'{letter}37'].value) == 'True':
            app.bhotswap.set(True)
        else:
            app.bhotswap.set(False)
        if str(sheet_ranges[f'{letter}38'].value) == 'True':
            app.bvibration.set(True)
        else:
            app.bvibration.set(False)
        if str(sheet_ranges[f'{letter}39'].value) == 'True':
            app.bnoise.set(True)
        else:
            app.bnoise.set(False)
        if str(sheet_ranges[f'{letter}40'].value) == 'True':
            app.bheat.set(True)
        else:
            app.bheat.set(False)
        if str(sheet_ranges[f'{letter}41'].value) == 'True':
            app.bnff.set(True)
        else:
            app.bnff.set(False)
        if str(sheet_ranges[f'{letter}42'].value) == 'True':
            app.bcable.set(True)
        else:
            app.bcable.set(False)
        if str(sheet_ranges[f'{letter}43'].value) == 'True':
            app.bsym_symnff.set(True)
        else:
            app.bsym_symnff.set(False)
        if str(sheet_ranges[f'{letter}44'].value) == 'True':
            app.bsym_cable.set(True)
        else:
            app.bsym_cable.set(False)
        if str(sheet_ranges[f'{letter}45'].value) == 'True':
            app.bsurge.set(True)
        else:
            app.bsurge.set(False)
        app.c_appearance_text.insert(app.tk.END, sheet_ranges[f'{letter}46'].value)
        app.c_operation_text.insert(app.tk.END, sheet_ranges[f'{letter}47'].value)
        app.c_internal_text.insert(app.tk.END, sheet_ranges[f'{letter}48'].value)
        app.c_suspected_text.insert(app.tk.END, sheet_ranges[f'{letter}49'].value)
        app.c_conclusion_text.insert(app.tk.END, sheet_ranges[f'{letter}50'].value)
        app.bench_text.insert(app.tk.END, sheet_ranges[f'{letter}51'].value)
    elif app.preset_text.get('1.0', 'end').split('\n')[0] == 'clear':
        lb = app.load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'B'
        app.btn_preset2_text.set('Preset 2')
        app.preset_text.delete('1.0', 'end')

        sheet_ranges[f'{letter}1'].value = 'Preset 2'
        for a in range(2, 51):
            sheet_ranges[f'{letter}{a}'].value = ''
        lb.save('presets.xlsx')
    else:
        lb = app.load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'B'

        sheet_ranges[f'{letter}1'].value = app.preset_text.get('1.0', 'end').split('\n')[0]
        app.btn_preset2_text.set(sheet_ranges[f'{letter}1'].value)
        app.preset_text.delete('1.0', 'end')

        sheet_ranges[f'{letter}2'].value = app.rpnum_text.get('1.0', 'end')
        sheet_ranges[f'{letter}3'].value = app.item_text.get('1.0', 'end')
        # sheet_ranges[f'{letter}4'].value = serial_text.get('1.0', 'end')
        sheet_ranges[f'{letter}5'].value = app.key_text.get('1.0', 'end')
        sheet_ranges[f'{letter}6'].value = app.fb_text.get('1.0', 'end')
        sheet_ranges[f'{letter}7'].value = app.comp_text.get('1.0', 'end')
        sheet_ranges[f'{letter}8'].value = app.structure_text.get('1.0', 'end')
        sheet_ranges[f'{letter}9'].value = app.parts_text.get('1.0', 'end')
        sheet_ranges[f'{letter}10'].value = app.pic_text.get('1.0', 'end')
        sheet_ranges[f'{letter}11'].value = app.his_rp_text.get('1.0', 'end')
        sheet_ranges[f'{letter}12'].value = app.his_comp_date_text.get('1.0', 'end')
        sheet_ranges[f'{letter}13'].value = app.his_symptom_text.get('1.0', 'end')
        sheet_ranges[f'{letter}14'].value = app.his_result_text.get('1.0', 'end')
        sheet_ranges[f'{letter}15'].value = app.appearance_text.get('1.0', 'end')
        sheet_ranges[f'{letter}16'].value = app.symptom_text.get('1.0', 'end')
        sheet_ranges[f'{letter}17'].value = app.tests_text.get('1.0', 'end')
        sheet_ranges[f'{letter}18'].value = app.analysis_text.get('1.0', 'end')
        sheet_ranges[f'{letter}19'].value = app.failure_text.get('1.0', 'end')
        sheet_ranges[f'{letter}20'].value = app.bapp_impact.get()
        sheet_ranges[f'{letter}21'].value = app.bapp_cable.get()
        sheet_ranges[f'{letter}22'].value = app.bapp_filterglass.get()
        sheet_ranges[f'{letter}23'].value = app.bapp_nff.get()
        sheet_ranges[f'{letter}24'].value = app.bapp_oil.get()
        sheet_ranges[f'{letter}25'].value = app.bsym_cable.get()
        sheet_ranges[f'{letter}26'].value = app.bsym_inout.get()
        sheet_ranges[f'{letter}27'].value = app.bsym_nffstruct.get()
        sheet_ranges[f'{letter}28'].value = app.bsym_nocom.get()
        sheet_ranges[f'{letter}29'].value = app.bsym_noint.get()
        sheet_ranges[f'{letter}30'].value = app.bsym_nolaser.get()
        sheet_ranges[f'{letter}31'].value = app.bsym_nopower.get()
        sheet_ranges[f'{letter}32'].value = app.bsym_symnff.get()
        sheet_ranges[f'{letter}33'].value = app.rp.get()
        sheet_ranges[f'{letter}34'].value = app.nrp.get()
        sheet_ranges[f'{letter}35'].value = app.rt.get()
        sheet_ranges[f'{letter}36'].value = app.kj.get()
        sheet_ranges[f'{letter}37'].value = app.bhotswap.get()
        sheet_ranges[f'{letter}38'].value = app.bvibration.get()
        sheet_ranges[f'{letter}39'].value = app.bnoise.get()
        sheet_ranges[f'{letter}40'].value = app.bheat.get()
        sheet_ranges[f'{letter}41'].value = app.bnff.get()
        sheet_ranges[f'{letter}42'].value = app.bcable.get()
        sheet_ranges[f'{letter}43'].value = app.bsym_symnff.get()
        sheet_ranges[f'{letter}44'].value = app.bsym_cable.get()
        sheet_ranges[f'{letter}45'].value = app.bsurge.get()
        sheet_ranges[f'{letter}46'].value = app.c_appearance_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}47'].value = app.c_operation_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}48'].value = app.c_internal_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}49'].value = app.c_suspected_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}50'].value = app.c_conclusion_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}51'].value = app.bench_text.get('1.0', 'end').split('\n')[0]
        lb.save('presets.xlsx')
    return None


def save_exe_preset3():
    if app.preset_text.get('1.0', 'end').split('\n')[0] == '':
        lb = app.load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'C'
        app.rpnum_text.insert(app.tk.END, sheet_ranges[f'{letter}2'].value)
        app.item_text.insert(app.tk.END, sheet_ranges[f'{letter}3'].value)
        # serial_text.insert(tk.END, sheet_ranges[f'{letter}4'].value)
        app.key_text.insert(app.tk.END, sheet_ranges[f'{letter}5'].value)
        app.fb_text.insert(app.tk.END, sheet_ranges[f'{letter}6'].value)
        app.comp_text.insert(app.tk.END, sheet_ranges[f'{letter}7'].value)
        app.structure_text.insert(app.tk.END, sheet_ranges[f'{letter}8'].value)
        app.parts_text.insert(app.tk.END, sheet_ranges[f'{letter}9'].value)
        app.pic_text.insert(app.tk.END, sheet_ranges[f'{letter}10'].value)
        app.his_rp_text.insert(app.tk.END, sheet_ranges[f'{letter}11'].value)
        app.his_comp_date_text.insert(app.tk.END, sheet_ranges[f'{letter}12'].value)
        app.his_symptom_text.insert(app.tk.END, sheet_ranges[f'{letter}13'].value)
        app.his_result_text.insert(app.tk.END, sheet_ranges[f'{letter}14'].value)
        app.appearance_text.insert(app.tk.END, sheet_ranges[f'{letter}15'].value)
        app.symptom_text.insert(app.tk.END, sheet_ranges[f'{letter}16'].value)
        app.tests_text.insert(app.tk.END, sheet_ranges[f'{letter}17'].value)
        app.analysis_text.insert(app.tk.END, sheet_ranges[f'{letter}18'].value)
        app.failure_text.insert(app.tk.END, sheet_ranges[f'{letter}19'].value)
        if str(sheet_ranges[f'{letter}20'].value) == 'True':
            app.bapp_impact.set(True)
        else:
            app.bapp_impact.set(False)
        if str(sheet_ranges[f'{letter}21'].value) == 'True':
            app.bapp_cable.set(True)
        else:
            app.bapp_cable.set(False)
        if str(sheet_ranges[f'{letter}22'].value) == 'True':
            app.bapp_filterglass.set(True)
        else:
            app.bapp_filterglass.set(False)
        if str(sheet_ranges[f'{letter}23'].value) == 'True':
            app.bapp_nff.set(True)
        else:
            app.bapp_nff.set(False)
        if str(sheet_ranges[f'{letter}24'].value) == 'True':
            app.bapp_oil.set(True)
        else:
            app.bapp_oil.set(False)
        if str(sheet_ranges[f'{letter}25'].value) == 'True':
            app.bsym_cable.set(True)
        else:
            app.bsym_cable.set(False)
        if str(sheet_ranges[f'{letter}26'].value) == 'True':
            app.bsym_inout.set(True)
        else:
            app.bsym_inout.set(False)
        if str(sheet_ranges[f'{letter}27'].value) == 'True':
            app.bsym_nffstruct.set(True)
        else:
            app.bsym_nffstruct.set(False)
        if str(sheet_ranges[f'{letter}28'].value) == 'True':
            app.bsym_nocom.set(True)
        else:
            app.bsym_nocom.set(False)
        if str(sheet_ranges[f'{letter}29'].value) == 'True':
            app.bsym_noint.set(True)
        else:
            app.bsym_noint.set(False)
        if str(sheet_ranges[f'{letter}30'].value) == 'True':
            app.bsym_nolaser.set(True)
        else:
            app.bsym_nolaser.set(False)
        if str(sheet_ranges[f'{letter}31'].value) == 'True':
            app.bsym_nopower.set(True)
        else:
            app.bsym_nopower.set(False)
        if str(sheet_ranges[f'{letter}32'].value) == 'True':
            app.bsym_symnff.set(True)
        else:
            app.bsym_symnff.set(False)
        if str(sheet_ranges[f'{letter}33'].value) == 'True':
            app.rp.set(True)
        else:
            app.rp.set(False)
        if str(sheet_ranges[f'{letter}34'].value) == 'True':
            app.nrp.set(True)
        else:
            app.nrp.set(False)
        if str(sheet_ranges[f'{letter}35'].value) == 'True':
            app.rt.set(True)
        else:
            app.rt.set(False)
        if str(sheet_ranges[f'{letter}36'].value) == 'True':
            app.kj.set(True)
        else:
            app.kj.set(False)
        if str(sheet_ranges[f'{letter}37'].value) == 'True':
            app.bhotswap.set(True)
        else:
            app.bhotswap.set(False)
        if str(sheet_ranges[f'{letter}38'].value) == 'True':
            app.bvibration.set(True)
        else:
            app.bvibration.set(False)
        if str(sheet_ranges[f'{letter}39'].value) == 'True':
            app.bnoise.set(True)
        else:
            app.bnoise.set(False)
        if str(sheet_ranges[f'{letter}40'].value) == 'True':
            app.bheat.set(True)
        else:
            app.bheat.set(False)
        if str(sheet_ranges[f'{letter}41'].value) == 'True':
            app.bnff.set(True)
        else:
            app.bnff.set(False)
        if str(sheet_ranges[f'{letter}42'].value) == 'True':
            app.bcable.set(True)
        else:
            app.bcable.set(False)
        if str(sheet_ranges[f'{letter}43'].value) == 'True':
            app.bsym_symnff.set(True)
        else:
            app.bsym_symnff.set(False)
        if str(sheet_ranges[f'{letter}44'].value) == 'True':
            app.bsym_cable.set(True)
        else:
            app.bsym_cable.set(False)
        if str(sheet_ranges[f'{letter}45'].value) == 'True':
            app.bsurge.set(True)
        else:
            app.bsurge.set(False)
        app.c_appearance_text.insert(app.tk.END, sheet_ranges[f'{letter}46'].value)
        app.c_operation_text.insert(app.tk.END, sheet_ranges[f'{letter}47'].value)
        app.c_internal_text.insert(app.tk.END, sheet_ranges[f'{letter}48'].value)
        app.c_suspected_text.insert(app.tk.END, sheet_ranges[f'{letter}49'].value)
        app.c_conclusion_text.insert(app.tk.END, sheet_ranges[f'{letter}50'].value)
        app.bench_text.insert(app.tk.END, sheet_ranges[f'{letter}51'].value)
    elif app.preset_text.get('1.0', 'end').split('\n')[0] == 'clear':
        lb = app.load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'C'
        app.btn_preset3_text.set('Preset 3')
        app.preset_text.delete('1.0', 'end')

        sheet_ranges[f'{letter}1'].value = 'Preset 3'
        for a in range(2, 51):
            sheet_ranges[f'{letter}{a}'].value = ''
        lb.save('presets.xlsx')
    else:
        lb = app.load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'C'

        sheet_ranges[f'{letter}1'].value = app.preset_text.get('1.0', 'end').split('\n')[0]
        app.btn_preset3_text.set(sheet_ranges[f'{letter}1'].value)
        app.preset_text.delete('1.0', 'end')

        sheet_ranges[f'{letter}2'].value = app.rpnum_text.get('1.0', 'end')
        sheet_ranges[f'{letter}3'].value = app.item_text.get('1.0', 'end')
        # sheet_ranges[f'{letter}4'].value = serial_text.get('1.0', 'end')
        sheet_ranges[f'{letter}5'].value = app.key_text.get('1.0', 'end')
        sheet_ranges[f'{letter}6'].value = app.fb_text.get('1.0', 'end')
        sheet_ranges[f'{letter}7'].value = app.comp_text.get('1.0', 'end')
        sheet_ranges[f'{letter}8'].value = app.structure_text.get('1.0', 'end')
        sheet_ranges[f'{letter}9'].value = app.parts_text.get('1.0', 'end')
        sheet_ranges[f'{letter}10'].value = app.pic_text.get('1.0', 'end')
        sheet_ranges[f'{letter}11'].value = app.his_rp_text.get('1.0', 'end')
        sheet_ranges[f'{letter}12'].value = app.his_comp_date_text.get('1.0', 'end')
        sheet_ranges[f'{letter}13'].value = app.his_symptom_text.get('1.0', 'end')
        sheet_ranges[f'{letter}14'].value = app.his_result_text.get('1.0', 'end')
        sheet_ranges[f'{letter}15'].value = app.appearance_text.get('1.0', 'end')
        sheet_ranges[f'{letter}16'].value = app.symptom_text.get('1.0', 'end')
        sheet_ranges[f'{letter}17'].value = app.tests_text.get('1.0', 'end')
        sheet_ranges[f'{letter}18'].value = app.analysis_text.get('1.0', 'end')
        sheet_ranges[f'{letter}19'].value = app.failure_text.get('1.0', 'end')
        sheet_ranges[f'{letter}20'].value = app.bapp_impact.get()
        sheet_ranges[f'{letter}21'].value = app.bapp_cable.get()
        sheet_ranges[f'{letter}22'].value = app.bapp_filterglass.get()
        sheet_ranges[f'{letter}23'].value = app.bapp_nff.get()
        sheet_ranges[f'{letter}24'].value = app.bapp_oil.get()
        sheet_ranges[f'{letter}25'].value = app.bsym_cable.get()
        sheet_ranges[f'{letter}26'].value = app.bsym_inout.get()
        sheet_ranges[f'{letter}27'].value = app.bsym_nffstruct.get()
        sheet_ranges[f'{letter}28'].value = app.bsym_nocom.get()
        sheet_ranges[f'{letter}29'].value = app.bsym_noint.get()
        sheet_ranges[f'{letter}30'].value = app.bsym_nolaser.get()
        sheet_ranges[f'{letter}31'].value = app.bsym_nopower.get()
        sheet_ranges[f'{letter}32'].value = app.bsym_symnff.get()
        sheet_ranges[f'{letter}33'].value = app.rp.get()
        sheet_ranges[f'{letter}34'].value = app.nrp.get()
        sheet_ranges[f'{letter}35'].value = app.rt.get()
        sheet_ranges[f'{letter}36'].value = app.kj.get()
        sheet_ranges[f'{letter}37'].value = app.bhotswap.get()
        sheet_ranges[f'{letter}38'].value = app.bvibration.get()
        sheet_ranges[f'{letter}39'].value = app.bnoise.get()
        sheet_ranges[f'{letter}40'].value = app.bheat.get()
        sheet_ranges[f'{letter}41'].value = app.bnff.get()
        sheet_ranges[f'{letter}42'].value = app.bcable.get()
        sheet_ranges[f'{letter}43'].value = app.bsym_symnff.get()
        sheet_ranges[f'{letter}44'].value = app.bsym_cable.get()
        sheet_ranges[f'{letter}45'].value = app.bsurge.get()
        sheet_ranges[f'{letter}46'].value = app.c_appearance_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}47'].value = app.c_operation_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}48'].value = app.c_internal_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}49'].value = app.c_suspected_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}50'].value = app.c_conclusion_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}51'].value = app.bench_text.get('1.0', 'end').split('\n')[0]
        lb.save('presets.xlsx')
    return None
