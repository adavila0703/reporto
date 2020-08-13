def save_exe_preset1():
    if len(preset_text.get('1.0', 'end').split('\n')[0]) > 18:
        too_many_char(len(preset_text.get('1.0', 'end').split('\n')[0]))
        return None
    else:
        pass

    if preset_text.get('1.0', 'end').split('\n')[0] == '':
        lb = load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'A'
        rpnum_text.insert(tk.END, sheet_ranges[f'{letter}2'].value)
        item_text.insert(tk.END, sheet_ranges[f'{letter}3'].value)
        # serial_text.insert(tk.END, sheet_ranges[f'{letter}4'].value)
        key_text.insert(tk.END, sheet_ranges[f'{letter}5'].value)
        fb_text.insert(tk.END, sheet_ranges[f'{letter}6'].value)
        comp_text.insert(tk.END, sheet_ranges[f'{letter}7'].value)
        structure_text.insert(tk.END, sheet_ranges[f'{letter}8'].value)
        parts_text.insert(tk.END, sheet_ranges[f'{letter}9'].value)
        pic_text.insert(tk.END, sheet_ranges[f'{letter}10'].value)
        his_rp_text.insert(tk.END, sheet_ranges[f'{letter}11'].value)
        his_comp_date_text.insert(tk.END, sheet_ranges[f'{letter}12'].value)
        his_symptom_text.insert(tk.END, sheet_ranges[f'{letter}13'].value)
        his_result_text.insert(tk.END, sheet_ranges[f'{letter}14'].value)
        appearance_text.insert(tk.END, sheet_ranges[f'{letter}15'].value)
        symptom_text.insert(tk.END, sheet_ranges[f'{letter}16'].value)
        tests_text.insert(tk.END, sheet_ranges[f'{letter}17'].value)
        analysis_text.insert(tk.END, sheet_ranges[f'{letter}18'].value)
        failure_text.insert(tk.END, sheet_ranges[f'{letter}19'].value)
        if str(sheet_ranges[f'{letter}20'].value) == 'True':
            bapp_impact.set(True)
        else:
            bapp_impact.set(False)
        if str(sheet_ranges[f'{letter}21'].value) == 'True':
            bapp_cable.set(True)
        else:
            bapp_cable.set(False)
        if str(sheet_ranges[f'{letter}22'].value) == 'True':
            bapp_filterglass.set(True)
        else:
            bapp_filterglass.set(False)
        if str(sheet_ranges[f'{letter}23'].value) == 'True':
            bapp_nff.set(True)
        else:
            bapp_nff.set(False)
        if str(sheet_ranges[f'{letter}24'].value) == 'True':
            bapp_oil.set(True)
        else:
            bapp_oil.set(False)
        if str(sheet_ranges[f'{letter}25'].value) == 'True':
            bsym_cable.set(True)
        else:
            bsym_cable.set(False)
        if str(sheet_ranges[f'{letter}26'].value) == 'True':
            bsym_inout.set(True)
        else:
            bsym_inout.set(False)
        if str(sheet_ranges[f'{letter}27'].value) == 'True':
            bsym_nffstruct.set(True)
        else:
            bsym_nffstruct.set(False)
        if str(sheet_ranges[f'{letter}28'].value) == 'True':
            bsym_nocom.set(True)
        else:
            bsym_nocom.set(False)
        if str(sheet_ranges[f'{letter}29'].value) == 'True':
            bsym_noint.set(True)
        else:
            bsym_noint.set(False)
        if str(sheet_ranges[f'{letter}30'].value) == 'True':
            bsym_nolaser.set(True)
        else:
            bsym_nolaser.set(False)
        if str(sheet_ranges[f'{letter}31'].value) == 'True':
            bsym_nopower.set(True)
        else:
            bsym_nopower.set(False)
        if str(sheet_ranges[f'{letter}32'].value) == 'True':
            bsym_symnff.set(True)
        else:
            bsym_symnff.set(False)
        if str(sheet_ranges[f'{letter}33'].value) == 'True':
            rp.set(True)
        else:
            rp.set(False)
        if str(sheet_ranges[f'{letter}34'].value) == 'True':
            nrp.set(True)
        else:
            nrp.set(False)
        if str(sheet_ranges[f'{letter}35'].value) == 'True':
            rt.set(True)
        else:
            rt.set(False)
        if str(sheet_ranges[f'{letter}36'].value) == 'True':
            kj.set(True)
        else:
            kj.set(False)
        if str(sheet_ranges[f'{letter}37'].value) == 'True':
            bhotswap.set(True)
        else:
            bhotswap.set(False)
        if str(sheet_ranges[f'{letter}38'].value) == 'True':
            bvibration.set(True)
        else:
            bvibration.set(False)
        if str(sheet_ranges[f'{letter}39'].value) == 'True':
            bnoise.set(True)
        else:
            bnoise.set(False)
        if str(sheet_ranges[f'{letter}40'].value) == 'True':
            bheat.set(True)
        else:
            bheat.set(False)
        if str(sheet_ranges[f'{letter}41'].value) == 'True':
            bnff.set(True)
        else:
            bnff.set(False)
        if str(sheet_ranges[f'{letter}42'].value) == 'True':
            bcable.set(True)
        else:
            bcable.set(False)
        if str(sheet_ranges[f'{letter}43'].value) == 'True':
            bsym_symnff.set(True)
        else:
            bsym_symnff.set(False)
        if str(sheet_ranges[f'{letter}44'].value) == 'True':
            bsym_cable.set(True)
        else:
            bsym_cable.set(False)
        if str(sheet_ranges[f'{letter}45'].value) == 'True':
            bsurge.set(True)
        else:
            bsurge.set(False)
        c_appearance_text.insert(tk.END, sheet_ranges[f'{letter}46'].value)
        c_operation_text.insert(tk.END, sheet_ranges[f'{letter}47'].value)
        c_internal_text.insert(tk.END, sheet_ranges[f'{letter}48'].value)
        c_suspected_text.insert(tk.END, sheet_ranges[f'{letter}49'].value)
        c_conclusion_text.insert(tk.END, sheet_ranges[f'{letter}50'].value)
        bench_text.insert(tk.END, sheet_ranges[f'{letter}51'].value)
    elif preset_text.get('1.0', 'end').split('\n')[0] == 'clear':
        lb = load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'A'
        btn_preset_text.set('Preset 1')
        preset_text.delete('1.0', 'end')

        sheet_ranges[f'{letter}1'].value = 'Preset 1'
        for a in range(2, 51):
            sheet_ranges[f'{letter}{a}'].value = ''
        lb.save('presets.xlsx')
    else:
        lb = load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'A'

        sheet_ranges[f'{letter}1'].value = preset_text.get('1.0', 'end').split('\n')[0]
        btn_preset_text.set(sheet_ranges[f'{letter}1'].value)
        preset_text.delete('1.0', 'end')

        sheet_ranges[f'{letter}2'].value = rpnum_text.get('1.0', 'end')
        sheet_ranges[f'{letter}3'].value = item_text.get('1.0', 'end')
        # sheet_ranges[f'{letter}4'].value = serial_text.get('1.0', 'end')
        sheet_ranges[f'{letter}5'].value = key_text.get('1.0', 'end')
        sheet_ranges[f'{letter}6'].value = fb_text.get('1.0', 'end')
        sheet_ranges[f'{letter}7'].value = comp_text.get('1.0', 'end')
        sheet_ranges[f'{letter}8'].value = structure_text.get('1.0', 'end')
        sheet_ranges[f'{letter}9'].value = parts_text.get('1.0', 'end')
        sheet_ranges[f'{letter}10'].value = pic_text.get('1.0', 'end')
        sheet_ranges[f'{letter}11'].value = his_rp_text.get('1.0', 'end')
        sheet_ranges[f'{letter}12'].value = his_comp_date_text.get('1.0', 'end')
        sheet_ranges[f'{letter}13'].value = his_symptom_text.get('1.0', 'end')
        sheet_ranges[f'{letter}14'].value = his_result_text.get('1.0', 'end')
        sheet_ranges[f'{letter}15'].value = appearance_text.get('1.0', 'end')
        sheet_ranges[f'{letter}16'].value = symptom_text.get('1.0', 'end')
        sheet_ranges[f'{letter}17'].value = tests_text.get('1.0', 'end')
        sheet_ranges[f'{letter}18'].value = analysis_text.get('1.0', 'end')
        sheet_ranges[f'{letter}19'].value = failure_text.get('1.0', 'end')
        sheet_ranges[f'{letter}20'].value = bapp_impact.get()
        sheet_ranges[f'{letter}21'].value = bapp_cable.get()
        sheet_ranges[f'{letter}22'].value = bapp_filterglass.get()
        sheet_ranges[f'{letter}23'].value = bapp_nff.get()
        sheet_ranges[f'{letter}24'].value = bapp_oil.get()
        sheet_ranges[f'{letter}25'].value = bsym_cable.get()
        sheet_ranges[f'{letter}26'].value = bsym_inout.get()
        sheet_ranges[f'{letter}27'].value = bsym_nffstruct.get()
        sheet_ranges[f'{letter}28'].value = bsym_nocom.get()
        sheet_ranges[f'{letter}29'].value = bsym_noint.get()
        sheet_ranges[f'{letter}30'].value = bsym_nolaser.get()
        sheet_ranges[f'{letter}31'].value = bsym_nopower.get()
        sheet_ranges[f'{letter}32'].value = bsym_symnff.get()
        sheet_ranges[f'{letter}33'].value = rp.get()
        sheet_ranges[f'{letter}34'].value = nrp.get()
        sheet_ranges[f'{letter}35'].value = rt.get()
        sheet_ranges[f'{letter}36'].value = kj.get()
        sheet_ranges[f'{letter}37'].value = bhotswap.get()
        sheet_ranges[f'{letter}38'].value = bvibration.get()
        sheet_ranges[f'{letter}39'].value = bnoise.get()
        sheet_ranges[f'{letter}40'].value = bheat.get()
        sheet_ranges[f'{letter}41'].value = bnff.get()
        sheet_ranges[f'{letter}42'].value = bcable.get()
        sheet_ranges[f'{letter}43'].value = bsym_symnff.get()
        sheet_ranges[f'{letter}44'].value = bsym_cable.get()
        sheet_ranges[f'{letter}45'].value = bsurge.get()
        sheet_ranges[f'{letter}46'].value = c_appearance_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}47'].value = c_operation_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}48'].value = c_internal_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}49'].value = c_suspected_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}50'].value = c_conclusion_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}51'].value = bench_text.get('1.0', 'end').split('\n')[0]
        lb.save('presets.xlsx')
    return None


def save_exe_preset2():
    if preset_text.get('1.0', 'end').split('\n')[0] == '':
        lb = load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'B'
        rpnum_text.insert(tk.END, sheet_ranges[f'{letter}2'].value)
        item_text.insert(tk.END, sheet_ranges[f'{letter}3'].value)
        # serial_text.insert(tk.END, sheet_ranges[f'{letter}4'].value)
        key_text.insert(tk.END, sheet_ranges[f'{letter}5'].value)
        fb_text.insert(tk.END, sheet_ranges[f'{letter}6'].value)
        comp_text.insert(tk.END, sheet_ranges[f'{letter}7'].value)
        structure_text.insert(tk.END, sheet_ranges[f'{letter}8'].value)
        parts_text.insert(tk.END, sheet_ranges[f'{letter}9'].value)
        pic_text.insert(tk.END, sheet_ranges[f'{letter}10'].value)
        his_rp_text.insert(tk.END, sheet_ranges[f'{letter}11'].value)
        his_comp_date_text.insert(tk.END, sheet_ranges[f'{letter}12'].value)
        his_symptom_text.insert(tk.END, sheet_ranges[f'{letter}13'].value)
        his_result_text.insert(tk.END, sheet_ranges[f'{letter}14'].value)
        appearance_text.insert(tk.END, sheet_ranges[f'{letter}15'].value)
        symptom_text.insert(tk.END, sheet_ranges[f'{letter}16'].value)
        tests_text.insert(tk.END, sheet_ranges[f'{letter}17'].value)
        analysis_text.insert(tk.END, sheet_ranges[f'{letter}18'].value)
        failure_text.insert(tk.END, sheet_ranges[f'{letter}19'].value)
        if str(sheet_ranges[f'{letter}20'].value) == 'True':
            bapp_impact.set(True)
        else:
            bapp_impact.set(False)
        if str(sheet_ranges[f'{letter}21'].value) == 'True':
            bapp_cable.set(True)
        else:
            bapp_cable.set(False)
        if str(sheet_ranges[f'{letter}22'].value) == 'True':
            bapp_filterglass.set(True)
        else:
            bapp_filterglass.set(False)
        if str(sheet_ranges[f'{letter}23'].value) == 'True':
            bapp_nff.set(True)
        else:
            bapp_nff.set(False)
        if str(sheet_ranges[f'{letter}24'].value) == 'True':
            bapp_oil.set(True)
        else:
            bapp_oil.set(False)
        if str(sheet_ranges[f'{letter}25'].value) == 'True':
            bsym_cable.set(True)
        else:
            bsym_cable.set(False)
        if str(sheet_ranges[f'{letter}26'].value) == 'True':
            bsym_inout.set(True)
        else:
            bsym_inout.set(False)
        if str(sheet_ranges[f'{letter}27'].value) == 'True':
            bsym_nffstruct.set(True)
        else:
            bsym_nffstruct.set(False)
        if str(sheet_ranges[f'{letter}28'].value) == 'True':
            bsym_nocom.set(True)
        else:
            bsym_nocom.set(False)
        if str(sheet_ranges[f'{letter}29'].value) == 'True':
            bsym_noint.set(True)
        else:
            bsym_noint.set(False)
        if str(sheet_ranges[f'{letter}30'].value) == 'True':
            bsym_nolaser.set(True)
        else:
            bsym_nolaser.set(False)
        if str(sheet_ranges[f'{letter}31'].value) == 'True':
            bsym_nopower.set(True)
        else:
            bsym_nopower.set(False)
        if str(sheet_ranges[f'{letter}32'].value) == 'True':
            bsym_symnff.set(True)
        else:
            bsym_symnff.set(False)
        if str(sheet_ranges[f'{letter}33'].value) == 'True':
            rp.set(True)
        else:
            rp.set(False)
        if str(sheet_ranges[f'{letter}34'].value) == 'True':
            nrp.set(True)
        else:
            nrp.set(False)
        if str(sheet_ranges[f'{letter}35'].value) == 'True':
            rt.set(True)
        else:
            rt.set(False)
        if str(sheet_ranges[f'{letter}36'].value) == 'True':
            kj.set(True)
        else:
            kj.set(False)
        if str(sheet_ranges[f'{letter}37'].value) == 'True':
            bhotswap.set(True)
        else:
            bhotswap.set(False)
        if str(sheet_ranges[f'{letter}38'].value) == 'True':
            bvibration.set(True)
        else:
            bvibration.set(False)
        if str(sheet_ranges[f'{letter}39'].value) == 'True':
            bnoise.set(True)
        else:
            bnoise.set(False)
        if str(sheet_ranges[f'{letter}40'].value) == 'True':
            bheat.set(True)
        else:
            bheat.set(False)
        if str(sheet_ranges[f'{letter}41'].value) == 'True':
            bnff.set(True)
        else:
            bnff.set(False)
        if str(sheet_ranges[f'{letter}42'].value) == 'True':
            bcable.set(True)
        else:
            bcable.set(False)
        if str(sheet_ranges[f'{letter}43'].value) == 'True':
            bsym_symnff.set(True)
        else:
            bsym_symnff.set(False)
        if str(sheet_ranges[f'{letter}44'].value) == 'True':
            bsym_cable.set(True)
        else:
            bsym_cable.set(False)
        if str(sheet_ranges[f'{letter}45'].value) == 'True':
            bsurge.set(True)
        else:
            bsurge.set(False)
        c_appearance_text.insert(tk.END, sheet_ranges[f'{letter}46'].value)
        c_operation_text.insert(tk.END, sheet_ranges[f'{letter}47'].value)
        c_internal_text.insert(tk.END, sheet_ranges[f'{letter}48'].value)
        c_suspected_text.insert(tk.END, sheet_ranges[f'{letter}49'].value)
        c_conclusion_text.insert(tk.END, sheet_ranges[f'{letter}50'].value)
        bench_text.insert(tk.END, sheet_ranges[f'{letter}51'].value)
    elif preset_text.get('1.0', 'end').split('\n')[0] == 'clear':
        lb = load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'B'
        btn_preset2_text.set('Preset 2')
        preset_text.delete('1.0', 'end')

        sheet_ranges[f'{letter}1'].value = 'Preset 2'
        for a in range(2, 51):
            sheet_ranges[f'{letter}{a}'].value = ''
        lb.save('presets.xlsx')
    else:
        lb = load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'B'

        sheet_ranges[f'{letter}1'].value = preset_text.get('1.0', 'end').split('\n')[0]
        btn_preset2_text.set(sheet_ranges[f'{letter}1'].value)
        preset_text.delete('1.0', 'end')

        sheet_ranges[f'{letter}2'].value = rpnum_text.get('1.0', 'end')
        sheet_ranges[f'{letter}3'].value = item_text.get('1.0', 'end')
        # sheet_ranges[f'{letter}4'].value = serial_text.get('1.0', 'end')
        sheet_ranges[f'{letter}5'].value = key_text.get('1.0', 'end')
        sheet_ranges[f'{letter}6'].value = fb_text.get('1.0', 'end')
        sheet_ranges[f'{letter}7'].value = comp_text.get('1.0', 'end')
        sheet_ranges[f'{letter}8'].value = structure_text.get('1.0', 'end')
        sheet_ranges[f'{letter}9'].value = parts_text.get('1.0', 'end')
        sheet_ranges[f'{letter}10'].value = pic_text.get('1.0', 'end')
        sheet_ranges[f'{letter}11'].value = his_rp_text.get('1.0', 'end')
        sheet_ranges[f'{letter}12'].value = his_comp_date_text.get('1.0', 'end')
        sheet_ranges[f'{letter}13'].value = his_symptom_text.get('1.0', 'end')
        sheet_ranges[f'{letter}14'].value = his_result_text.get('1.0', 'end')
        sheet_ranges[f'{letter}15'].value = appearance_text.get('1.0', 'end')
        sheet_ranges[f'{letter}16'].value = symptom_text.get('1.0', 'end')
        sheet_ranges[f'{letter}17'].value = tests_text.get('1.0', 'end')
        sheet_ranges[f'{letter}18'].value = analysis_text.get('1.0', 'end')
        sheet_ranges[f'{letter}19'].value = failure_text.get('1.0', 'end')
        sheet_ranges[f'{letter}20'].value = bapp_impact.get()
        sheet_ranges[f'{letter}21'].value = bapp_cable.get()
        sheet_ranges[f'{letter}22'].value = bapp_filterglass.get()
        sheet_ranges[f'{letter}23'].value = bapp_nff.get()
        sheet_ranges[f'{letter}24'].value = bapp_oil.get()
        sheet_ranges[f'{letter}25'].value = bsym_cable.get()
        sheet_ranges[f'{letter}26'].value = bsym_inout.get()
        sheet_ranges[f'{letter}27'].value = bsym_nffstruct.get()
        sheet_ranges[f'{letter}28'].value = bsym_nocom.get()
        sheet_ranges[f'{letter}29'].value = bsym_noint.get()
        sheet_ranges[f'{letter}30'].value = bsym_nolaser.get()
        sheet_ranges[f'{letter}31'].value = bsym_nopower.get()
        sheet_ranges[f'{letter}32'].value = bsym_symnff.get()
        sheet_ranges[f'{letter}33'].value = rp.get()
        sheet_ranges[f'{letter}34'].value = nrp.get()
        sheet_ranges[f'{letter}35'].value = rt.get()
        sheet_ranges[f'{letter}36'].value = kj.get()
        sheet_ranges[f'{letter}37'].value = bhotswap.get()
        sheet_ranges[f'{letter}38'].value = bvibration.get()
        sheet_ranges[f'{letter}39'].value = bnoise.get()
        sheet_ranges[f'{letter}40'].value = bheat.get()
        sheet_ranges[f'{letter}41'].value = bnff.get()
        sheet_ranges[f'{letter}42'].value = bcable.get()
        sheet_ranges[f'{letter}43'].value = bsym_symnff.get()
        sheet_ranges[f'{letter}44'].value = bsym_cable.get()
        sheet_ranges[f'{letter}45'].value = bsurge.get()
        sheet_ranges[f'{letter}46'].value = c_appearance_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}47'].value = c_operation_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}48'].value = c_internal_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}49'].value = c_suspected_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}50'].value = c_conclusion_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}51'].value = bench_text.get('1.0', 'end').split('\n')[0]
        lb.save('presets.xlsx')
    return None


def save_exe_preset3():
    if preset_text.get('1.0', 'end').split('\n')[0] == '':
        lb = load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'C'
        rpnum_text.insert(tk.END, sheet_ranges[f'{letter}2'].value)
        item_text.insert(tk.END, sheet_ranges[f'{letter}3'].value)
        # serial_text.insert(tk.END, sheet_ranges[f'{letter}4'].value)
        key_text.insert(tk.END, sheet_ranges[f'{letter}5'].value)
        fb_text.insert(tk.END, sheet_ranges[f'{letter}6'].value)
        comp_text.insert(tk.END, sheet_ranges[f'{letter}7'].value)
        structure_text.insert(tk.END, sheet_ranges[f'{letter}8'].value)
        parts_text.insert(tk.END, sheet_ranges[f'{letter}9'].value)
        pic_text.insert(tk.END, sheet_ranges[f'{letter}10'].value)
        his_rp_text.insert(tk.END, sheet_ranges[f'{letter}11'].value)
        his_comp_date_text.insert(tk.END, sheet_ranges[f'{letter}12'].value)
        his_symptom_text.insert(tk.END, sheet_ranges[f'{letter}13'].value)
        his_result_text.insert(tk.END, sheet_ranges[f'{letter}14'].value)
        appearance_text.insert(tk.END, sheet_ranges[f'{letter}15'].value)
        symptom_text.insert(tk.END, sheet_ranges[f'{letter}16'].value)
        tests_text.insert(tk.END, sheet_ranges[f'{letter}17'].value)
        analysis_text.insert(tk.END, sheet_ranges[f'{letter}18'].value)
        failure_text.insert(tk.END, sheet_ranges[f'{letter}19'].value)
        if str(sheet_ranges[f'{letter}20'].value) == 'True':
            bapp_impact.set(True)
        else:
            bapp_impact.set(False)
        if str(sheet_ranges[f'{letter}21'].value) == 'True':
            bapp_cable.set(True)
        else:
            bapp_cable.set(False)
        if str(sheet_ranges[f'{letter}22'].value) == 'True':
            bapp_filterglass.set(True)
        else:
            bapp_filterglass.set(False)
        if str(sheet_ranges[f'{letter}23'].value) == 'True':
            bapp_nff.set(True)
        else:
            bapp_nff.set(False)
        if str(sheet_ranges[f'{letter}24'].value) == 'True':
            bapp_oil.set(True)
        else:
            bapp_oil.set(False)
        if str(sheet_ranges[f'{letter}25'].value) == 'True':
            bsym_cable.set(True)
        else:
            bsym_cable.set(False)
        if str(sheet_ranges[f'{letter}26'].value) == 'True':
            bsym_inout.set(True)
        else:
            bsym_inout.set(False)
        if str(sheet_ranges[f'{letter}27'].value) == 'True':
            bsym_nffstruct.set(True)
        else:
            bsym_nffstruct.set(False)
        if str(sheet_ranges[f'{letter}28'].value) == 'True':
            bsym_nocom.set(True)
        else:
            bsym_nocom.set(False)
        if str(sheet_ranges[f'{letter}29'].value) == 'True':
            bsym_noint.set(True)
        else:
            bsym_noint.set(False)
        if str(sheet_ranges[f'{letter}30'].value) == 'True':
            bsym_nolaser.set(True)
        else:
            bsym_nolaser.set(False)
        if str(sheet_ranges[f'{letter}31'].value) == 'True':
            bsym_nopower.set(True)
        else:
            bsym_nopower.set(False)
        if str(sheet_ranges[f'{letter}32'].value) == 'True':
            bsym_symnff.set(True)
        else:
            bsym_symnff.set(False)
        if str(sheet_ranges[f'{letter}33'].value) == 'True':
            rp.set(True)
        else:
            rp.set(False)
        if str(sheet_ranges[f'{letter}34'].value) == 'True':
            nrp.set(True)
        else:
            nrp.set(False)
        if str(sheet_ranges[f'{letter}35'].value) == 'True':
            rt.set(True)
        else:
            rt.set(False)
        if str(sheet_ranges[f'{letter}36'].value) == 'True':
            kj.set(True)
        else:
            kj.set(False)
        if str(sheet_ranges[f'{letter}37'].value) == 'True':
            bhotswap.set(True)
        else:
            bhotswap.set(False)
        if str(sheet_ranges[f'{letter}38'].value) == 'True':
            bvibration.set(True)
        else:
            bvibration.set(False)
        if str(sheet_ranges[f'{letter}39'].value) == 'True':
            bnoise.set(True)
        else:
            bnoise.set(False)
        if str(sheet_ranges[f'{letter}40'].value) == 'True':
            bheat.set(True)
        else:
            bheat.set(False)
        if str(sheet_ranges[f'{letter}41'].value) == 'True':
            bnff.set(True)
        else:
            bnff.set(False)
        if str(sheet_ranges[f'{letter}42'].value) == 'True':
            bcable.set(True)
        else:
            bcable.set(False)
        if str(sheet_ranges[f'{letter}43'].value) == 'True':
            bsym_symnff.set(True)
        else:
            bsym_symnff.set(False)
        if str(sheet_ranges[f'{letter}44'].value) == 'True':
            bsym_cable.set(True)
        else:
            bsym_cable.set(False)
        if str(sheet_ranges[f'{letter}45'].value) == 'True':
            bsurge.set(True)
        else:
            bsurge.set(False)
        c_appearance_text.insert(tk.END, sheet_ranges[f'{letter}46'].value)
        c_operation_text.insert(tk.END, sheet_ranges[f'{letter}47'].value)
        c_internal_text.insert(tk.END, sheet_ranges[f'{letter}48'].value)
        c_suspected_text.insert(tk.END, sheet_ranges[f'{letter}49'].value)
        c_conclusion_text.insert(tk.END, sheet_ranges[f'{letter}50'].value)
        bench_text.insert(tk.END, sheet_ranges[f'{letter}51'].value)
    elif preset_text.get('1.0', 'end').split('\n')[0] == 'clear':
        lb = load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'C'
        btn_preset3_text.set('Preset 3')
        preset_text.delete('1.0', 'end')

        sheet_ranges[f'{letter}1'].value = 'Preset 3'
        for a in range(2, 51):
            sheet_ranges[f'{letter}{a}'].value = ''
        lb.save('presets.xlsx')
    else:
        lb = load_workbook(filename='presets.xlsx')
        sheet_ranges = lb['Sheet']
        letter = 'C'

        sheet_ranges[f'{letter}1'].value = preset_text.get('1.0', 'end').split('\n')[0]
        btn_preset3_text.set(sheet_ranges[f'{letter}1'].value)
        preset_text.delete('1.0', 'end')

        sheet_ranges[f'{letter}2'].value = rpnum_text.get('1.0', 'end')
        sheet_ranges[f'{letter}3'].value = item_text.get('1.0', 'end')
        # sheet_ranges[f'{letter}4'].value = serial_text.get('1.0', 'end')
        sheet_ranges[f'{letter}5'].value = key_text.get('1.0', 'end')
        sheet_ranges[f'{letter}6'].value = fb_text.get('1.0', 'end')
        sheet_ranges[f'{letter}7'].value = comp_text.get('1.0', 'end')
        sheet_ranges[f'{letter}8'].value = structure_text.get('1.0', 'end')
        sheet_ranges[f'{letter}9'].value = parts_text.get('1.0', 'end')
        sheet_ranges[f'{letter}10'].value = pic_text.get('1.0', 'end')
        sheet_ranges[f'{letter}11'].value = his_rp_text.get('1.0', 'end')
        sheet_ranges[f'{letter}12'].value = his_comp_date_text.get('1.0', 'end')
        sheet_ranges[f'{letter}13'].value = his_symptom_text.get('1.0', 'end')
        sheet_ranges[f'{letter}14'].value = his_result_text.get('1.0', 'end')
        sheet_ranges[f'{letter}15'].value = appearance_text.get('1.0', 'end')
        sheet_ranges[f'{letter}16'].value = symptom_text.get('1.0', 'end')
        sheet_ranges[f'{letter}17'].value = tests_text.get('1.0', 'end')
        sheet_ranges[f'{letter}18'].value = analysis_text.get('1.0', 'end')
        sheet_ranges[f'{letter}19'].value = failure_text.get('1.0', 'end')
        sheet_ranges[f'{letter}20'].value = bapp_impact.get()
        sheet_ranges[f'{letter}21'].value = bapp_cable.get()
        sheet_ranges[f'{letter}22'].value = bapp_filterglass.get()
        sheet_ranges[f'{letter}23'].value = bapp_nff.get()
        sheet_ranges[f'{letter}24'].value = bapp_oil.get()
        sheet_ranges[f'{letter}25'].value = bsym_cable.get()
        sheet_ranges[f'{letter}26'].value = bsym_inout.get()
        sheet_ranges[f'{letter}27'].value = bsym_nffstruct.get()
        sheet_ranges[f'{letter}28'].value = bsym_nocom.get()
        sheet_ranges[f'{letter}29'].value = bsym_noint.get()
        sheet_ranges[f'{letter}30'].value = bsym_nolaser.get()
        sheet_ranges[f'{letter}31'].value = bsym_nopower.get()
        sheet_ranges[f'{letter}32'].value = bsym_symnff.get()
        sheet_ranges[f'{letter}33'].value = rp.get()
        sheet_ranges[f'{letter}34'].value = nrp.get()
        sheet_ranges[f'{letter}35'].value = rt.get()
        sheet_ranges[f'{letter}36'].value = kj.get()
        sheet_ranges[f'{letter}37'].value = bhotswap.get()
        sheet_ranges[f'{letter}38'].value = bvibration.get()
        sheet_ranges[f'{letter}39'].value = bnoise.get()
        sheet_ranges[f'{letter}40'].value = bheat.get()
        sheet_ranges[f'{letter}41'].value = bnff.get()
        sheet_ranges[f'{letter}42'].value = bcable.get()
        sheet_ranges[f'{letter}43'].value = bsym_symnff.get()
        sheet_ranges[f'{letter}44'].value = bsym_cable.get()
        sheet_ranges[f'{letter}45'].value = bsurge.get()
        sheet_ranges[f'{letter}46'].value = c_appearance_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}47'].value = c_operation_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}48'].value = c_internal_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}49'].value = c_suspected_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}50'].value = c_conclusion_text.get('1.0', 'end').split('\n')[0]
        sheet_ranges[f'{letter}51'].value = bench_text.get('1.0', 'end').split('\n')[0]
        lb.save('presets.xlsx')
    return None
