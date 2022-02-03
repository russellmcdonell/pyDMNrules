import sys
import pyDMNrules

if __name__ == '__main__':
    dmnRules = pyDMNrules.DMN()
    status = dmnRules.load('../pyDMNrules/AN-SNAP V4 rules (DMN).xlsx')
    if 'errors' in status:
        print('AN-SNAP rules (DMN).xlsx has errors', status['errors'])
        sys.exit(0)
    else:
        print('AN-SNAP rules (DMN).xlsx loaded')

    data = {}
    data['Multidisciplinary'] = False
    data['Admitted Flag'] = False
    print('testing:', data, 'Expect:4999')
    (status, newData) = dmnRules.decide(data)
    if status != {}:
        print(status)
    print('AN-SNAP V4 code', newData[-1]['Result']['AN-SNAP V4 code'])
    (decision, table, rule) = newData[-1]['Executed Rule']
    print('Rule:', rule)
    print()

    data['Multidisciplinary'] = True
    data['Care Type'] = 'GEM'
    data['Single Day of Care'] = False
    data['Ongoing Pain'] = False
    data['Clinic'] = 'Memory'
    print('testing:', data, 'Expect:4UC3')
    (status, newData) = dmnRules.decide(data)
    if status != {}:
        print(status)
    print('AN-SNAP V4 code:', newData[-1]['Result']['AN-SNAP V4 code'])
    (decision, table, rule) = newData[-1]['Executed Rule']
    print('Rule:', rule)
    print()

    data['Care Type'] = 'Rehabilitation'
    del data['Single Day of Care']
    del data['Ongoing Pain']
    del data['Clinic']
    data['Patient Age'] = 19
    data['Assessment Only'] = True
    print('testing:', data, 'Expect:adult)(2),4SY1')
    (status, newData) = dmnRules.decide(data)
    if status != {}:
        print(status)
    print('Computed Age Type:', newData[1]['Result']['Computed Age Type'])
    print('AN-SNAP V4 code:', newData[-1]['Result']['AN-SNAP V4 code'])
    (decision, table, rule) = newData[-1]['Executed Rule']
    print('Rule:', rule)
    print()

    data['Assessment Only'] = False
    data['AROC code'] = '7'
    print('testing:', data, 'Expect:adult(2),4SG1')
    (status, newData) = dmnRules.decide(data)
    if status != {}:
        print(status)
    print('Computed Age Type:', newData[1]['Result']['Computed Age Type'])
    print('AN-SNAP V4 code:', newData[-1]['Result']['AN-SNAP V4 code'])
    (decision, table, rule) = newData[-1]['Executed Rule']
    print('Rule:', rule)
    print()

    data['Patient Age Type'] = '1'
    print('testing:', data, 'Expect:paed(1),4X05')
    (status, newData) = dmnRules.decide(data)
    if status != {}:
        print(status)
    print('Computed Age Type:', newData[1]['Result']['Computed Age Type'])
    print('Patient Age:', newData[-1]['Result']['Patient Age'])
    print('AN-SNAP V4 code:', newData[-1]['Result']['AN-SNAP V4 code'])
    (decision, table, rule) = newData[-1]['Executed Rule']
    print('Rule:', rule)
