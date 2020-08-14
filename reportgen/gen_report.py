import gui.gui as gui
import pyperclip
from fpdf import FPDF
import reportgen.rep_categories as rep


def checkerror():
    """Checks to make sure all required values are entered.

    Notes:
        calls function warning() from gui.gui

    """
    error = ''
    if gui.rpnum_text.get('1.0', 'end').split('\n')[0] == '':
        error += 'RP#\n'
    if gui.rp.get() == False and gui.nrp.get() == False and gui.rt.get() == False and gui.kj.get() == False:
        error += 'Result\n'
    if gui.bhotswap.get() == False and gui.bvibration.get() == False and gui.bnoise.get() == False and gui.bheat.get() == False and \
            gui.bnff.get() == False and gui.bcable.get() == False and gui.bimpact.get() == False:
        error += 'Preventative Measure\n'
    if gui.bapp_nff.get() == False and gui.bapp_oil.get() == False and gui.bapp_impact.get() == False and gui.bapp_cable.get() == False \
            and gui.bapp_filterglass.get() == False:
        error += 'Appearance Check\n'
    if gui.pic_text.get('1.0', 'end').split('\n')[0] == '':
        error += 'PIC\n'
    if gui.bsym_cable.get() == False and gui.bsym_inout.get() == False and gui.bsym_nocom.get() == False and \
            gui.bsym_noint.get() == False and gui.bsym_nolaser.get() == False and gui.bsym_nopower.get() == False and \
            gui.bsym_symnff.get() == False and gui.bsym_nffstruct == False:
        error += 'Symptom\n'
    gui.warning(error)
    return None


def create_form():
    """Generates report and saves to a locaiton.

    """
    rp_save = gui.rpnum_text.get('1.0', 'end').split('\n')[0].strip()
    pdf = FPDF()
    pdf.add_page()
    sum_out = ''
    count = 0
    cliping = ''

    if gui.rpnum_text.get('1.0', 'end').split('\n')[0] == '' or gui.pic_text.get('1.0', 'end').split('\n')[0] == '' or \
            gui.rp.get() == False and gui.nrp.get() == False and gui.rt.get() == False and gui.kj.get() == False or \
            gui.bhotswap.get() == False and gui.bvibration.get() == False and \
            gui.bnoise.get() == False and gui.bheat.get() == False and gui.bnff.get() == False and gui.bcable.get() == False and \
            gui.bsurge.get() == False and gui.bimpact.get() == False or gui.bapp_nff.get() == False and gui.bapp_oil.get() == False and \
            gui.bapp_impact.get() == False and gui.bapp_cable.get() == False and gui.bapp_filterglass.get() == False or \
            gui.bsym_cable.get() == False and \
            gui.bsym_inout.get() == False and gui.bsym_nocom.get() == False and gui.bsym_noint.get() == False and \
            gui.bsym_nolaser.get() == False and gui.bsym_nopower.get() == False and gui.bsym_symnff.get() == False and \
            gui.bsym_nffstruct.get() == False:
        checkerror()
        return None
    else:
        pass
    pdf.image('report_img/logo.jpg', w=190, h=55, type='jpg')
    # pdf.set_font('Arial', 'IB', 22)
    # pdf.cell(190, 15, 'Failure Analysis Report', ln=1, align='C')
    pdf.ln(10)
    pdf.set_font('Arial', 'B', 14)
    if gui.rpnum_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'RP#: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'RP#: ' + gui.rpnum_text.get('1.0', 'end') + '          ')

    if gui.item_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Item: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'Item: ' + gui.item_text.get('1.0', 'end') + '          ')

    # if serial_text.get('1.0', 'end').strip() == '':
    #     pdf.cell(40, 10, 'Serial: ' + 'None          ')
    # else:
    pdf.cell(40, 10, 'Serial: ' + gui.entries())

    pdf.ln(h=15)

    if gui.key_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Keyword: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'Keyword: ' + gui.key_text.get('1.0', 'end'))

    if gui.pic_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'PIC: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'PIC: ' + gui.pic_text.get('1.0', 'end'))

    if gui.bench_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Time: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'Time: ' + gui.bench_text.get('1.0', 'end'))

    if gui.pic_text.get('1.0', 'end').strip() == 'Angel' or gui.pic_text.get('1.0', 'end').strip() == 'angel':
        pdf.image('report_img/angel.jpg', y=85, x=125, w=20, h=20, type='jpg')
    elif gui.pic_text.get('1.0', 'end').strip() == 'Grant' or gui.pic_text.get('1.0', 'end').strip() == 'grant':
        pdf.image('report_img/grant.jpg', y=85, x=125, w=20, h=20, type='jpg')
    elif gui.pic_text.get('1.0', 'end').strip() == 'Adolfo' or gui.pic_text.get('1.0', 'end').strip() == 'adolfo':
        pdf.image('report_img/adolfo.jpg', y=85, x=125, w=20, h=20, type='jpg')

    pdf.ln(h=15)
    pdf.image('report_img/rp-history.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)

    if gui.his_rp_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'RP#: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'RP#: ' + gui.his_rp_text.get('1.0', 'end'))

    if gui.his_comp_date_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Date: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'Date:' + gui.his_comp_date_text.get('1.0', 'end'))

    if gui.his_symptom_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Symptom: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'Symptom: ' + gui.his_symptom_text.get('1.0', 'end'), ln=1)

    if gui.his_result_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Result: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'Result: ' + gui.his_result_text.get('1.0', 'end'), ln=1)

    pdf.ln(h=15)
    pdf.image('report_img/appearance.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(40, 10, gui.appearance_text.get('1.0', 'end'), ln=1)
    pdf.ln(h=15)
    pdf.image('report_img/symptom.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(40, 10, gui.symptom_text.get('1.0', 'end'), ln=1)
    pdf.ln(h=15)
    pdf.image('report_img/tests.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(40, 10, gui.tests_text.get('1.0', 'end'), ln=1)
    pdf.ln(h=15)
    pdf.image('report_img/failure.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.ln(h=15)

    if gui.fb_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Failed Board: ' + ' None', ln=1)
    else:
        pdf.cell(40, 10, 'Failed Board: ' + gui.fb_text.get('1.0', 'end'), ln=1)

    if gui.comp_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Location of Component: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'Location of Component: ' + gui.comp_text.get('1.0', 'end'), ln=1)
    pdf.cell(40, 10, 'Failed Structure Part: ' + gui.structure_text.get('1.0', 'end'), ln=1)

    if gui.parts_text.get('1.0', 'end').split('\n')[0] == '':
        pdf.cell(40, 10, 'Were parts replaced? No')
    else:
        pdf.cell(40, 10, 'Were parts replaced? Yes. ' + gui.parts_text.get('1.0', 'end') + '.')

    pdf.ln(h=15)
    pdf.image('report_img/analysis.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)

    for a in range(len(gui.analysis_text.get('1.0', 'end').split('\n'))):
        pdf.cell(40, 10, gui.analysis_text.get('1.0', 'end').split('\n')[count], ln=1)
        count += 1

    pdf.ln(h=15)
    pdf.image('report_img/causeof.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(40, 10, gui.failure_text.get('1.0', 'end'), ln=1)
    pdf.ln(h=15)
    pdf.image('report_img/summary.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)

    if gui.rp.get():
        sum_out = 'This unit was repaired and sent back to the customer'
    elif gui.nrp.get():
        sum_out = 'This unit was not repaired and sent back to the customer in the condition it was received'
    elif gui.rt.get():
        sum_out = 'This unit was sent back to the customer in the condition it was received'
    elif gui.kj.get():
        sum_out = 'This unit was sent to KJ'

    pdf.cell(40, 10, sum_out, ln=1)
    pdf.ln(h=15)
    pdf.add_page()
    pdf.cell(175, 15, 'Customer Inspection Report', ln=1, align='C')
    pdf.ln(h=15)

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Appearance Check > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, rep.apperance_check())

    if gui.c_appearance_text.get('1.0', 'end').strip() == '':
        cliping += '< Appearance Check >\n' + rep.apperance_check() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + gui.c_appearance_text.get('1.0', 'end') + '\n')
        cliping += '< Appearance Check >\n' + rep.apperance_check() + '\nAdditional information:\n' + \
                   gui.c_appearance_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Operation Check > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, rep.operation_check())

    if gui.c_operation_text.get('1.0', 'end').strip() == '':
        cliping += '< Operation Check >\n' + rep.operation_check() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + gui.c_operation_text.get('1.0', 'end') + '\n')
        cliping += '< Operation Check >\n' + rep.operation_check() + '\nAdditional information:\n' + \
                   gui.c_operation_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Internal Check > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, rep.internal_check())

    if gui.c_internal_text.get('1.0', 'end').strip() == '':
        cliping += '< Internal Check >\n' + rep.internal_check() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + gui.c_internal_text.get('1.0', 'end') + '\n')
        cliping += '< Internal Check >\n' + rep.internal_check() + '\nAdditional information:\n' + \
                   gui.c_internal_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Suspected Cause > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, rep.suspected_cause())

    if gui.c_suspected_text.get('1.0', 'end').strip() == '':
        cliping += '< Suspected Cause >\n' + rep.suspected_cause() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + gui.c_suspected_text.get('1.0', 'end') + '\n')
        cliping += '< Suspected Cause >\n' + rep.suspected_cause() + '\nAdditional information:\n' + \
                   gui.c_suspected_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Conclusion > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, rep.conclusion())

    if gui.c_conclusion_text.get('1.0', 'end').strip() == '':
        cliping += '< Conclusion >\n' + rep.conclusion() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + gui.c_conclusion_text.get('1.0', 'end') + '\n')
        cliping += '< Conclusion >\n' + rep.conclusion() + '\nAdditional information:\n' + \
                   gui.c_conclusion_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Preventative Measure > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, rep.preventative_measure())
    cliping += '< Preventative Measure >\n' + rep.preventative_measure()

    try:
        pdf.output('I:/PRD/Check Sheets/RSS/' + rp_save + ' - ' + gui.serial_list[0].get('1.0', 'end').strip() + '.pdf',
                   'F')
        # pdf.output(rp_save + ' - ' + gui.serial_list[0].get('1.0', 'end').strip() + '.pdf', 'F')
    except OSError:
        gui.fileopen()
        return None

    pyperclip.copy(cliping)
    gui.copy_message()
    gui.rpnum_text.delete('1.0', 'end')
    # serial_text.delete('1.0', 'end')
    gui.key_text.delete('1.0', 'end')
    gui.comp_text.delete('1.0', 'end')
    gui.structure_text.delete('1.0', 'end')
    gui.fb_text.delete('1.0', 'end')
