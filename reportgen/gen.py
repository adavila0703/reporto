import app


def form():
    """
    This function is in charge of generating report.
    """
    rp_save = app.rpnum_text.get('1.0', 'end').split('\n')[0].strip()
    # serial_save = serial_text.get('1.0', 'end').split('\n')[0].strip()
    pdf = app.FPDF()
    pdf.add_page()
    sum_out = ''
    count = 0
    cliping = ''

    if app.rpnum_text.get('1.0', 'end').split('\n')[0] == '' or app.pic_text.get('1.0', 'end').split('\n')[0] == '' or \
            app.rp.get() == False and app.nrp.get() == False and app.rt.get() == False and app.kj.get() == False or \
            app.bhotswap.get() == False and app.bvibration.get() == False and \
            app.bnoise.get() == False and app.bheat.get() == False and app.bnff.get() == False and app.bcable.get() == False and \
            app.bsurge.get() == False and app.bimpact.get() == False or app.bapp_nff.get() == False and app.bapp_oil.get() == False and \
            app.bapp_impact.get() == False and app.bapp_cable.get() == False and app.bapp_filterglass.get() == False or \
            app.bsym_cable.get() == False and \
            app.bsym_inout.get() == False and app.bsym_nocom.get() == False and app.bsym_noint.get() == False and \
            app.bsym_nolaser.get() == False and app.bsym_nopower.get() == False and app.bsym_symnff.get() == False and \
            app.bsym_nffstruct.get() == False:
        app.checkerror()
        return None
    else:
        pass
    pdf.image('report_img/logo.jpg', w=190, h=55, type='jpg')
    # pdf.set_font('Arial', 'IB', 22)
    # pdf.cell(190, 15, 'Failure Analysis Report', ln=1, align='C')
    pdf.ln(10)
    pdf.set_font('Arial', 'B', 14)
    if app.rpnum_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'RP#: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'RP#: ' + app.rpnum_text.get('1.0', 'end') + '          ')

    if app.item_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Item: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'Item: ' + app.item_text.get('1.0', 'end') + '          ')

    # if serial_text.get('1.0', 'end').strip() == '':
    #     pdf.cell(40, 10, 'Serial: ' + 'None          ')
    # else:
    pdf.cell(40, 10, 'Serial: ' + app.entries())

    pdf.ln(h=15)

    if app.key_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Keyword: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'Keyword: ' + app.key_text.get('1.0', 'end'))

    if app.pic_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'PIC: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'PIC: ' + app.pic_text.get('1.0', 'end'))

    if app.bench_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Time: ' + 'None          ')
    else:
        pdf.cell(40, 10, 'Time: ' + app.bench_text.get('1.0', 'end'))

    if app.pic_text.get('1.0', 'end').strip() == 'Angel' or app.pic_text.get('1.0', 'end').strip() == 'angel':
        pdf.image('report_img/angel.jpg', y=85, x=125, w=20, h=20, type='jpg')
    elif app.pic_text.get('1.0', 'end').strip() == 'Grant' or app.pic_text.get('1.0', 'end').strip() == 'grant':
        pdf.image('report_img/grant.jpg', y=85, x=125, w=20, h=20, type='jpg')
    elif app.pic_text.get('1.0', 'end').strip() == 'Adolfo' or app.pic_text.get('1.0', 'end').strip() == 'adolfo':
        pdf.image('report_img/adolfo.jpg', y=85, x=125, w=20, h=20, type='jpg')

    pdf.ln(h=15)
    pdf.image('report_img/rp-history.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)

    if app.his_rp_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'RP#: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'RP#: ' + app.his_rp_text.get('1.0', 'end'))

    if app.his_comp_date_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Date: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'Date:' + app.his_comp_date_text.get('1.0', 'end'))

    if app.his_symptom_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Symptom: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'Symptom: ' + app.his_symptom_text.get('1.0', 'end'), ln=1)

    if app.his_result_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Result: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'Result: ' + app.his_result_text.get('1.0', 'end'), ln=1)

    pdf.ln(h=15)
    pdf.image('report_img/apperance.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(40, 10, app.appearance_text.get('1.0', 'end'), ln=1)
    pdf.ln(h=15)
    pdf.image('report_img/symptom.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(40, 10, app.symptom_text.get('1.0', 'end'), ln=1)
    pdf.ln(h=15)
    pdf.image('report_img/tests.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(40, 10, app.tests_text.get('1.0', 'end'), ln=1)
    pdf.ln(h=15)
    pdf.image('report_img/failure.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.ln(h=15)

    if app.fb_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Failed Board: ' + ' None', ln=1)
    else:
        pdf.cell(40, 10, 'Failed Board: ' + app.fb_text.get('1.0', 'end'), ln=1)

    if app.comp_text.get('1.0', 'end').strip() == '':
        pdf.cell(40, 10, 'Location of Component: ' + 'None', ln=1)
    else:
        pdf.cell(40, 10, 'Location of Component: ' + app.comp_text.get('1.0', 'end'), ln=1)
    pdf.cell(40, 10, 'Failed Structure Part: ' + app.structure_text.get('1.0', 'end'), ln=1)

    if app.parts_text.get('1.0', 'end').split('\n')[0] == '':
        pdf.cell(40, 10, 'Were parts replaced? No')
    else:
        pdf.cell(40, 10, 'Were parts replaced? Yes. ' + app.parts_text.get('1.0', 'end') + '.')

    pdf.ln(h=15)
    pdf.image('report_img/analysis.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)

    for a in range(len(app.analysis_text.get('1.0', 'end').split('\n'))):
        pdf.cell(40, 10, app.analysis_text.get('1.0', 'end').split('\n')[count], ln=1)
        count += 1

    pdf.ln(h=15)
    pdf.image('report_img/causeof.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(40, 10, app.failure_text.get('1.0', 'end'), ln=1)
    pdf.ln(h=15)
    pdf.image('report_img/summary.jpg', w=190, h=12, type='jpg')
    pdf.set_font('Arial', 'B', 12)

    if app.rp.get() == True:
        sum_out = 'This unit was repaired and sent back to the customer'
    elif app.nrp.get() == True:
        sum_out = 'This unit was not repaired and sent back to the customer in the condition it was received'
    elif app.rt.get() == True:
        sum_out = 'This unit was sent back to the customer in the condition it was received'
    elif app.kj.get() == True:
        sum_out = 'This unit was sent to KJ'

    pdf.cell(40, 10, sum_out, ln=1)
    pdf.ln(h=15)
    pdf.add_page()
    pdf.cell(175, 15, 'Customer Inspection Report', ln=1, align='C')
    pdf.ln(h=15)

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Appearance Check > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, app.apperance_check())

    if app.c_appearance_text.get('1.0', 'end').strip() == '':
        cliping += '< Appearance Check >\n' + app.apperance_check() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + app.c_appearance_text.get('1.0', 'end') + '\n')
        cliping += '< Appearance Check >\n' + app.apperance_check() + '\nAdditional information:\n' + \
                   app.c_appearance_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Operation Check > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, app.operation_check())

    if app.c_operation_text.get('1.0', 'end').strip() == '':
        cliping += '< Operation Check >\n' + app.operation_check() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + app.c_operation_text.get('1.0', 'end') + '\n')
        cliping += '< Operation Check >\n' + app.operation_check() + '\nAdditional information:\n' + \
                   app.c_operation_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Internal Check > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, app.internal_check())

    if app.c_internal_text.get('1.0', 'end').strip() == '':
        cliping += '< Internal Check >\n' + app.internal_check() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + app.c_internal_text.get('1.0', 'end') + '\n')
        cliping += '< Internal Check >\n' + app.internal_check() + '\nAdditional information:\n' + \
                   app.c_internal_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Suspected Cause > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, app.suspected_cause())

    if app.c_suspected_text.get('1.0', 'end').strip() == '':
        cliping += '< Suspected Cause >\n' + app.suspected_cause() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + app.c_suspected_text.get('1.0', 'end') + '\n')
        cliping += '< Suspected Cause >\n' + app.suspected_cause() + '\nAdditional information:\n' + \
                   app.c_suspected_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Conclusion > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, app.conclusion())

    if app.c_conclusion_text.get('1.0', 'end').strip() == '':
        cliping += '< Conclusion >\n' + app.conclusion() + '\n'
    else:
        pdf.multi_cell(175, 10, 'Additional information:\n' + app.c_conclusion_text.get('1.0', 'end') + '\n')
        cliping += '< Conclusion >\n' + app.conclusion() + '\nAdditional information:\n' + \
                   app.c_conclusion_text.get('1.0', 'end') + '\n'

    pdf.set_font('Arial', '', 14)
    pdf.cell(40, 10, '< Preventative Measure > ', ln=1)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(175, 10, app.preventative_measure())
    cliping += '< Preventative Measure >\n' + app.preventative_measure()

    try:
        # pdf.output('I:/PRD/Check Sheets/RSS/' + rp_save + ' - ' + serial_save + '.pdf', 'F')
        pdf.output(rp_save + ' - ' + '.pdf', 'F')
    except OSError:
        app.fileopen()
        return None

    app.pyperclip.copy(cliping)
    app.copy_message()
    app.rpnum_text.delete('1.0', 'end')
    # serial_text.delete('1.0', 'end')
    app.key_text.delete('1.0', 'end')
    app.comp_text.delete('1.0', 'end')
    app.structure_text.delete('1.0', 'end')
    app.fb_text.delete('1.0', 'end')

