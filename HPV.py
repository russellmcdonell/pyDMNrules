import sys
import pyDMNrules

if __name__ == '__main__':

    dmnRules = pyDMNrules.DMN()
    status = dmnRules.load('ExampleHPV.xlsx')
    if 'errors' in status:
        print('ExampleHPV.xlsx has errors', status['errors'])
        sys.exit(0)
    else:
        print('ExampleHPV.xlsx loaded')

    data = {}
    data['Participant Age'] = 36
    data['In Test of Cure'] = True
    data['Hysterectomy Flag'] = False
    data['Cancer Flag'] = False
    data['HPV-V'] = 'V0'
    data['Current Participant Risk Category'] = 'low'
    print('Testing',repr(data))
    (status, newData) = dmnRules.decide(data)
    print('Decision(newData)',repr(newData))
    print()
    print(newData.keys())
    for key in newData.keys():
        if isinstance(newData[key], dict):
            print('\t', key, newData[key].keys())
    if 'errors' in status:
        print('With errors', status['errors'])
    print()

    (testStatus, results) = dmnRules.test()
    for test in range(len(results)):
        if 'Mismatches' not in results[test]:
            print('Test ID', results[test]['Test ID'], 'passed')
        else:
            print('Test ID', results[test]['Test ID'], 'failed')
            for failure in range(len(results[test]['Mismatches'])):
                print(results[test]['Mismatches'][failure])
            print(results[test]['DataAnnotations'])
            print(results[test]['TestAnnotations'])
            print(results[test]['data'])
            print(results[test]['newData'])
    if len(testStatus) > 0:
        print('with errors', repr(testStatus))
        sys.exit(0)

    # Report the structure of results
    print()
    print(results[0].keys())
    for key in results[0].keys():
        if isinstance(results[0][key], dict):
            print('\t', key, results[0][key].keys())
            for subKey in results[0][key].keys():
                if isinstance(results[0][key][subKey], dict):
                    print('\t', '\t', subKey, results[0][key][subKey].keys())
    sys.exit(1)
