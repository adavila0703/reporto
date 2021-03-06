import gui.gui as gui


def clearexample():
    """Clears the example values that are stored to help user understand UI."""
    gui.rpnum_text.delete('1.0', 'end')
    gui.item_text.delete('1.0', 'end')
    # serial_text.delete('1.0', 'end')
    gui.key_text.delete('1.0', 'end')
    gui.fb_text.delete('1.0', 'end')
    gui.comp_text.delete('1.0', 'end')
    gui.structure_text.delete('1.0', 'end')
    gui.parts_text.delete('1.0', 'end')
    gui.pic_text.delete('1.0', 'end')
    gui.his_rp_text.delete('1.0', 'end')
    gui.his_comp_date_text.delete('1.0', 'end')
    gui.his_symptom_text.delete('1.0', 'end')
    gui.his_result_text.delete('1.0', 'end')
    gui.appearance_text.delete('1.0', 'end')
    gui.symptom_text.delete('1.0', 'end')
    gui.tests_text.delete('1.0', 'end')
    gui.analysis_text.delete('1.0', 'end')
    gui.failure_text.delete('1.0', 'end')
    gui.bapp_impact.set(False)
    gui.bsym_nolaser.set(False)
    gui.nrp.set(False)
    gui.bhotswap.set(False)
    gui.bnoise.set(False)
    gui.bsurge.set(False)
    gui.c_appearance_text.delete('1.0', 'end')
    gui.c_operation_text.delete('1.0', 'end')
    gui.c_internal_text.delete('1.0', 'end')
    gui.c_suspected_text.delete('1.0', 'end')
    gui.c_conclusion_text.delete('1.0', 'end')
    gui.bench_text.delete('1.0', 'end')
    return None
