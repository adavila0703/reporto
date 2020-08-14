import gui.gui as gui


def apperance_check():
    """Checks which values of appearance is relevant and stores it in report/pyperclip

    Return:
        combined string value
    """
    str_out = 'After a thorough external visual inspection, the following was found:\n'

    if gui.bapp_nff.get():
        str_out += '-No major damage was found to the exterior.\n'
    if gui.bapp_cable.get():
        str_out += '-The head cable was found damaged.\n'
    if gui.bapp_impact.get():
        str_out += '-The case was found with impact marks.\n'
    if gui.bapp_oil.get():
        str_out += '-Oil was found covering the exterior.\n'
    if gui.bapp_filterglass.get():
        str_out += '-The filter glass damage.\n'
    return str_out


def operation_check():
    """Checks which values of operation is relevant and stores it in report/pyperclip

        Return:
            combined string value
    """
    str_out = 'The following was found during a full function test:\n'
    if gui.bsym_nocom.get():
        str_out += '-Unit communication was abnormal\n'
    if gui.bsym_nolaser.get():
        str_out += '-No laser output.\n'
    if gui.bsym_nopower.get():
        str_out += '-The unit did not power on.\n'
    if gui.bsym_noint.get():
        str_out += 'The unit would not initialize properly.\n'
    if gui.bsym_cable.get():
        str_out = 'With an operation check, a known good cable was replaced in order to conduct a full function ' \
                  'check\n'
    if gui.bsym_inout.get():
        str_out += '-There were abnormalities found with either the inputs or the outputs\n'
    if gui.bsym_nffstruct.get():
        str_out = 'With an operation check, although there were some structural damage to the unit, the unit is still' \
                  ' operational\n'
    if gui.bsym_symnff.get():
        str_out = 'With an operation check, the symptom was not replicated, the unit passed all tests\n'
    return str_out


def internal_check():
    """Checks which values of internal is relevant and stores it in report/pyperclip

        Return:
            combined string value
    """
    str_out = ''
    if gui.structure_text.get('1.0', 'end').split('\n')[0] != '' or gui.fb_text.get('1.0', 'end').split('\n')[0] != '' \
            or gui.comp_text.get('1.0', 'end').split('\n')[0] != '':
        str_out += 'With an internal check it was determined that'
        fb = gui.fb_text.get('1.0', 'end').split('\n')[0].lower()
        struct = gui.structure_text.get('1.0', 'end').split('\n')[0].lower()
        if gui.comp_text.get('1.0', 'end').split('\n')[0] != '':
            if gui.fb_text.get('1.0', 'end').split('\n')[0] != '':
                if gui.structure_text.get('1.0', 'end').split('\n')[0] != '':
                    str_out += f' an internal electrical component on the {fb} and the {struct} were found damaged'
                else:
                    str_out += f' an internal electrical component on the {fb} was found damaged'
            else:
                if gui.structure_text.get('1.0', 'end').split('\n')[0] != '':
                    str_out += f' an internal electrical component and the {struct} were found damaged'
                else:
                    str_out += f' an internal electrical component was found damaged'
        elif gui.fb_text.get('1.0', 'end').split('\n')[0] != '':

            if gui.structure_text.get('1.0', 'end').split('\n')[0] != '':
                str_out += f' the {fb} and the {struct} were found damaged'
            else:
                str_out += f' the {fb} was found damaged'

        elif gui.structure_text.get('1.0', 'end').split('\n')[0] != '':
            str_out += f' the {struct} was found damaged'
    else:
        str_out += 'There was no need to disassemble the unit'
    return str_out + '.\n'


def suspected_cause():
    """Checks which values of cause is relevant and stores it in report/pyperclip

        Return:
            combined string value
    """
    str_out = 'The unit failure was likely caused by the following possibilities:\n'
    if gui.bhotswap.get():
        str_out += '- Electrical component damage due to hot swap.\n'
    if gui.bvibration.get():
        str_out += '- Internal component damage due to vibration.\n'
    if gui.bnoise.get():
        str_out += '- Electrical component damage due to inductive noise.\n'
    if gui.bheat.get():
        str_out += '- Internal component damage due to heat.\n'
    if gui.bcable.get():
        str_out += '- Cable damage due to impact or improper use of the cable.\n'
    if gui.bnff.get():
        str_out = '- No faults were found.\n'
    if gui.bsurge.get():
        str_out += '- An application of over voltage/current in the form of an electrical surge.\n'
    if gui.bimpact.get():
        str_out += '- Impact damage to the unit.\n'
    return str_out


def conclusion():
    """Checks which values of conclusion is relevant and stores it in report/pyperclip

        Return:
            combined string value
    """
    str_out = ''
    if gui.rp.get():
        str_out = 'To bring this unit back to full functionality, the unit will be repaired and tested to ' \
                  'meet manufacturing specifications. A full operation and function check will be conducted to ' \
                  'confirm normal operation.\n'
    if gui.nrp.get():
        str_out = 'This unit cannot be repaired as either it is damaged beyond repairing capabilities or the repair ' \
                  'will approach the cost of a new unit. This unit will be sent back in the condition it was ' \
                  'received. We do not recommend putting this unit back on a production line, as we cannot guarantee' \
                  ' the reliability of its full functionality.\n'
    if gui.rt.get():
        str_out = 'Because we were not able to replicate the described symptom, the unit will be returned in the ' \
                  'condition it was received.\n'
    if gui.kj.get():
        str_out = 'The internal boards will be replaced to that of the equal repair cost and as a result, the' \
                  ' unit will have a new serial number.\n'
    return str_out


def preventative_measure():
    """Checks which values of preventative measures is relevant and stores it in report/pyperclip

        Return:
            combined string value
    """
    str_out = ''
    if gui.bhotswap.get():
        str_out += 'How swap damage could occur when the units power is not fully cycled off. Please make' \
                   ' sure to complete a unit shutdown before unplugging the unit. '
    if gui.bvibration.get():
        str_out += 'To prevent damage from vibration, please make sure to reduce the amount of vibration ' \
                   'the unit is subjected to by introducing vibration dampeners. '
    if gui.bnoise.get():
        str_out += 'To prevent damage from inductive noise, please make sure to keep this unit away from ' \
                   'machinery that can cause electrical noise. '
    if gui.bheat.get():
        str_out += 'To prevent damage from heat, please make sure to provide proper air flow to keep the unit cool. '
    if gui.bcable.get():
        str_out += 'To prevent cable damage, please make sure to protect the cable from impact and pulling. '
    if gui.bnff.get():
        str_out += 'No faults were found'
    if gui.bimpact.get():
        str_out += 'Please make sure not to subject the unit to impact damage.'
    else:
        pass
    return str_out
