import app as r


def makeform():
    rp_save = r.rpnum_text.get('1.0', 'end').split('\n')[0].strip()
    # serial_save = serial_text.get('1.0', 'end').split('\n')[0].strip()
    pdf = r.FPDF()
    pdf.add_page()
    sum_out = ''
    count = 0
    cliping = ''

    if r.rpnum_text.get('1.0', 'end').split('\n')[0] == '' or r.pic_text.get('1.0', 'end').split('\n')[0] == '' or \
            r.rp.get() == False and r.nrp.get() == False and r.rt.get() == False and r.kj.get() == False or \
            r.bhotswap.get() == False and r.bvibration.get() == False and \
            r.bnoise.get() == False and r.bheat.get() == False and r.bnff.get() == False and r.bcable.get() == False and \
            r.bsurge.get() == False and r.bimpact.get() == False or r.bapp_nff.get() == False and r.bapp_oil.get() == False and \
            r.bapp_impact.get() == False and r.bapp_cable.get() == False and r.bapp_filterglass.get() == False or \
            r.bsym_cable.get() == False and \
            r.bsym_inout.get() == False and r.bsym_nocom.get() == False and r.bsym_noint.get() == False and \
            r.bsym_nolaser.get() == False and r.bsym_nopower.get() == False and r.bsym_symnff.get() == False and \
            r.bsym_nffstruct.get() == False:
        r.checkerror()
        return None
    else:
        pass
    pdf.image('report_img/logo.jpg', w=190, h=55, type='jpg')
    # pdf.set_font('Arial', 'IB', 22)
    # pdf.cell(190, 15, 'Failure Analysis Report', ln=1, align='C')
    pdf.ln(10)
    pdf.set_font('Arial', 'B', 14)
    if r.rpnum_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'RP#: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'RP#: ' + r.rpnum_text.get('1.0', 'end') + '          ')

    if r.item_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Item: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'Item: ' + r.item_text.get('1.0', 'end') + '          ')

    # if serial_text.get('1.0', 'end').strip() == '':
    #     pdf.cell(40, 10, 'Serial: ' + 'None          ')
    # else:
    pdf.cell(40, 10, 'Serial: ' + r.entries())

    pdf.ln(h=15)

    if r.key_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Keyword: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'Keyword: ' + r.key_text.get('1.0', 'end'))

    if r.pic_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'PIC: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'PIC: ' + r.pic_text.get('1.0', 'end'))

    if r.bench_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Time: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'Time: ' + r.bench_text.get('1.0', 'end'))

    if r.pic_text.get('1.0', 'end').strip() == 'Angel' or r.pic_text.get('1.0', 'end').strip() == 'angel':
        pdf.image('report_img/angel.jpg', y=85, x=125, w=20, h=20, type='jpg')
    elif r.pic_text.get('1.0', 'end').strip() == 'Grant' or r.pic_text.get('1.0', 'end').strip() == 'grant':
        pdf.image('report_img/grant.jpg', y=85, x=125, w=20, h=20, type='jpg')
    elif r.pic_text.get('1.0', 'end').strip() == 'Adolfo' or r.pic_text.get('1.0', 'end').strip() == 'adolfo':
        pdf.image('report_img/adolfo.jpg', y=85, x=125, w=20, h=20, type='jpg')

    pdf.ln(h=15)
    pdf.image('report_img/rp-history.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)

    if r.his_rp_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'RP#: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'RP#: ' + r.his_rp_text.get('1.0', 'end'))

    if r.his_comp_date_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Date: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'Date:' + r.his_comp_date_text.get('1.0', 'end'))

    if r.his_symptom_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Symptom: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'Symptom: ' + r.his_symptom_text.get('1.0', 'end'), ln=1)

    if r.his_result_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Result: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'Result: ' + r.his_result_text.get('1.0', 'end'), ln=1)

    pdf.ln(h=15)
    pdf.image('report_img/apperance.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(40, 10, r.appearance_text.get('1.0', 'end'), ln=1)
    pdf.ln(h=15)
    pdf.image('report_img/symptom.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(40, 10, r.symptom_text.get('1.0', 'end'), ln=1)
    pdf.ln(h=15)
    pdf.image('report_img/tests.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(40, 10, r.tests_text.get('1.0', 'end'), ln=1)
    pdf.ln(h=15)
    pdf.image('report_img/failure.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.ln(h=15)

    if r.fb_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Failed Board: ' + ' None', ln=1)
    else:
        pdf.cell(40, 10, 'Failed Board: ' + r.fb_text.get('1.0', 'end'), ln=1)

    if r.comp_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Location of Component: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'Location of Component: ' + r.comp_text.get('1.0', 'end'), ln=1)
    pdf.cell(40, 10, 'Failed Structure Part: ' + r.structure_text.get('1.0', 'end'), ln=1)

    if r.parts_text.get('1.0', 'end').split('\n')[0] == '':
        pdf.cell(40, 10, 'Were parts replaced? No')
    else:
        pdf.cell(40, 10, 'Were parts replaced? Yes. ' + r.parts_text.get('1.0', 'end') + '.')

    pdf.ln(h=15)
    pdf.image('report_img/analysis.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)

    for a in range(len(r.analysis_text.get('1.0', 'end').split('\n'))):
        pdf.cell(40, 10, r.analysis_text.get('1.0', 'end').split('\n')[count], ln=1)
        count += 1

    pdf.ln(h=15)
    pdf.image('report_img/causeof.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(40, 10, r.failure_text.get('1.0', 'end'), ln=1)
    pdf.ln(h=15)
    pdf.image('report_img/summary.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)

    if r.rp.get() == True:
        sum_out = 'This unit was repaired and sent back to the customer'
    elif r.nrp.get() == True:
        sum_out = 'This unit was not repaired and sent back to the customer in the condition it was received'
    elif r.rt.get() == True:
        sum_out = 'This unit was sent back to the customer in the condition it was received'
    elif r.kj.get() == True:
        sum_out = 'This unit was sent to KJ'

    pdf.cell(40, 10, sum_out, ln=1)
    pdf.ln(h=15)
    pdf.add_page()
    pdf.cell(175, 15, 'Customer Inspection Report', ln=1, align='C')
    pdf.ln(h=15)

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Appearance Check > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, r.apperance_check())

    if r.c_appearance_text.get('1.0', 'end').strip() == '':
        cliping += '< Appearance Check >\n' + r.apperance_check() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + r.c_appearance_text.get('1.0', 'end') + '\n')
        cliping += '< Appearance Check >\n' + r.apperance_check() + '\nAdditional information:\n' + \
                   r.c_appearance_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Operation Check > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, r.operation_check())

    if r.c_operation_text.get('1.0', 'end').strip() == '':
        cliping += '< Operation Check >\n' + r.operation_check() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + r.c_operation_text.get('1.0', 'end') + '\n')
        cliping += '< Operation Check >\n' + r.operation_check() + '\nAdditional information:\n' + \
                   r.c_operation_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Internal Check > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, r.internal_check())

    if r.c_internal_text.get('1.0', 'end').strip() == '':
        cliping += '< Internal Check >\n' + r.internal_check() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + r.c_internal_text.get('1.0', 'end') + '\n')
        cliping += '< Internal Check >\n' + r.internal_check() + '\nAdditional information:\n' + \
                   r.c_internal_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Suspected Cause > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, r.suspected_cause())

    if r.c_suspected_text.get('1.0', 'end').strip() == '':
        cliping += '< Suspected Cause >\n' + r.suspected_cause() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + r.c_suspected_text.get('1.0', 'end') + '\n')
        cliping += '< Suspected Cause >\n' + r.suspected_cause() + '\nAdditional information:\n' + \
                   r.c_suspected_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Conclusion > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, r.conclusion())

    if r.c_conclusion_text.get('1.0', 'end').strip() == '':
        cliping += '< Conclusion >\n' + r.conclusion() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + r.c_conclusion_text.get('1.0', 'end') + '\n')
        cliping += '< Conclusion >\n' + r.conclusion() + '\nAdditional information:\n' + \
                   r.c_conclusion_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Preventative Measure > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, r.preventative_measure())
    cliping += '< Preventative Measure >\n' + r.preventative_measure()

    try:
        # pdf.output('I:/PRD/Check Sheets/RSS/' + rp_save + ' - ' + serial_save + '.pdf', 'F')
        pdf.output(rp_save + ' - ' + '.pdf', 'F')
    except OSError:
        r.fileopen()
        return None

    r.pyperclip.copy(cliping)
    r.copy_message()
    r.rpnum_text.delete('1.0', 'end')
    # serial_text.delete('1.0', 'end')
    r.key_text.delete('1.0', 'end')
    r.comp_text.delete('1.0', 'end')
    r.structure_text.delete('1.0', 'end')
    r.fb_text.delete('1.0', 'end')