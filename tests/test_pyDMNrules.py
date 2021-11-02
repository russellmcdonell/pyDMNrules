import pyDMNrules
from openpyxl import load_workbook
import csv
import pandas as pd

class TestClass:
    def test_HPV1(self):
        '''
        Check that the supplied ExampleHPv.xlsx workbook works
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/ExampleHPV.xlsx')
        assert 'errors' not in status
        data = {}
        data['Participant Age'] = 36
        data['In Test of Cure'] = True
        data['Hysterectomy Flag'] = False
        data['Cancer Flag'] = False
        data['HPV-V'] = 'V0'
        data['Current Participant Risk Category'] = 'low'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Test Risk Code' in newData['Result']
        assert newData['Result']['Test Risk Code'] == 'L'
        assert 'New Participant Risk Category' in newData['Result']
        assert newData['Result']['New Participant Risk Category'] == 'low'
        assert 'Participant Care Pathway' in newData['Result']
        assert newData['Result']['Participant Care Pathway'] == 'toBeDetermined'
        assert 'Next Rule' in newData['Result']
        assert newData['Result']['Next Rule'] == 'CervicalRisk2'

    def test_HPV2(self):
        '''
        Check that the supplied ExampleHPv.xlsx workbook works when loaded and passed as a workbook
        '''
        dmnRules = pyDMNrules.DMN()
        wb = load_workbook(filename='../pyDMNrules/ExampleHPV.xlsx')
        status = dmnRules.use(wb)
        assert 'errors' not in status
        data = {}
        data['Participant Age'] = 36
        data['In Test of Cure'] = True
        data['Hysterectomy Flag'] = False
        data['Cancer Flag'] = False
        data['HPV-V'] = 'V0'
        data['Current Participant Risk Category'] = 'low'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Test Risk Code' in newData['Result']
        assert newData['Result']['Test Risk Code'] == 'L'
        assert 'New Participant Risk Category' in newData['Result']
        assert newData['Result']['New Participant Risk Category'] == 'low'
        assert 'Participant Care Pathway' in newData['Result']
        assert newData['Result']['Participant Care Pathway'] == 'toBeDetermined'
        assert 'Next Rule' in newData['Result']
        assert newData['Result']['Next Rule'] == 'CervicalRisk2'


    def test_Therapy(self):
        '''
        Check that the supplied Therapy.xlsx workbook works
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/Therapy.xlsx')
        assert 'errors' not in status
        (testStatus, results) = dmnRules.test()
        assert 'errors' not in testStatus
        assert len(results) == 3
        assert 'Mismatches' not in results[0]
        assert 'newData' in results[0]
        assert 'Result' in results[0]['newData']
        assert 'Encounter Diagnosis' in results[0]['newData']['Result']
        assert results[0]['newData']['Result']['Encounter Diagnosis'] == 'Acute Sinusitis'
        assert 'Recommended Medication' in results[0]['newData']['Result']
        assert results[0]['newData']['Result']['Recommended Medication'] == 'Levofloxacin'
        assert 'Recommended Dose' in results[0]['newData']['Result']
        assert results[0]['newData']['Result']['Recommended Dose'] == '500mg every 24 hours for 14 days'
        assert 'Warning' in results[0]['newData']['Result']
        assert results[0]['newData']['Result']['Warning'] == 'Coumadin and Levofloxacin can result in reduced effectiveness of Coumadin.'
        assert 'Error Message' in results[0]['newData']['Result']
        assert results[0]['newData']['Result']['Error Message'] is None
        assert 'Result' in results[1]['newData']
        assert 'Encounter Diagnosis' in results[1]['newData']['Result']
        assert results[1]['newData']['Result']['Encounter Diagnosis'] == 'Acute Sinusitis'
        assert 'Recommended Medication' in results[1]['newData']['Result']
        assert results[1]['newData']['Result']['Recommended Medication'] == 'Amoxicillin'
        assert 'Recommended Dose' in results[1]['newData']['Result']
        assert results[1]['newData']['Result']['Recommended Dose'] == '250mg every 24 hours for 14 days'
        assert 'Warning' in results[1]['newData']['Result']
        assert results[1]['newData']['Result']['Warning'] is None
        assert 'Error Message' in results[0]['newData']['Result']
        assert results[0]['newData']['Result']['Error Message'] is None
        assert 'Result' in results[2]['newData']
        assert 'Encounter Diagnosis' in results[2]['newData']['Result']
        assert results[2]['newData']['Result']['Encounter Diagnosis'] == 'Diabetes'
        assert 'Recommended Medication' in results[2]['newData']['Result']
        assert results[2]['newData']['Result']['Recommended Medication'] is None
        assert 'Recommended Dose' in results[2]['newData']['Result']
        assert results[2]['newData']['Result']['Recommended Dose'] is None
        assert 'Warning' in results[2]['newData']['Result']
        assert results[2]['newData']['Result']['Warning'] is None
        assert 'Error Message' in results[2]['newData']['Result']
        assert results[2]['newData']['Result']['Error Message'] == 'Sorry, this decision service can handle only Acute Sinusitis'
 
    def test_Example1(self):
        '''
        Check that the supplied Example1.xlsx workbook works
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/Example1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Applicant Age'] = 61
        data['Medical History'] = 'good'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Applicant Risk Rating' in newData['Result']
        assert newData['Result']['Applicant Risk Rating'] == 'Medium'
        data['Medical History'] = 'bad'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert newData['Result']['Applicant Risk Rating'] == 'High'
        data['Applicant Age'] = 60
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert newData['Result']['Applicant Risk Rating'] == 'Medium'
        data['Applicant Age'] = 24
        data['Medical History'] = 'good'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert newData['Result']['Applicant Risk Rating'] == 'Low'
        data['Medical History'] = 'bad'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert newData['Result']['Applicant Risk Rating'] == 'Medium'

    def test_ExampleRows(self):
        '''
        Check that the supplied ExampleRows.xlsx workbook works
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/ExampleRows.xlsx')
        assert 'errors' not in status
        data = {}
        data['Customer'] = 'Business'
        data['OrderSize'] = 9
        data['Delivery'] = 'sameday'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0.05
        data['OrderSize'] = 10
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0.1
        data['Customer'] = 'Private'
        data['Delivery'] = 'sameday'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0
        data['Delivery'] = 'slow'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0.05
        data['Customer'] = 'Government'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0.15

    def test_ExampleColumns(self):
        '''
        Check that the supplied ExampleColumns.xlsx workbook works
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/ExampleColumns.xlsx')
        assert 'errors' not in status
        data = {}
        data['Customer'] = 'Business'
        data['OrderSize'] = 9
        data['Delivery'] = 'sameday'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0.05
        data['OrderSize'] = 10
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0.1
        data['Customer'] = 'Private'
        data['Delivery'] = 'sameday'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0
        data['Delivery'] = 'slow'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0.05
        data['Customer'] = 'Government'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0.15

    def test_ExampleCrosstab(self):
        '''
        Check that the supplied ExampleCrosstab.xlsx workbook works
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/ExampleCrosstab.xlsx')
        assert 'errors' not in status
        data = {}
        data['Customer'] = 'Business'
        data['OrderSize'] = 9
        data['Delivery'] = 'sameday'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0.05
        data['OrderSize'] = 10
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0.1
        data['Customer'] = 'Private'
        data['Delivery'] = 'sameday'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0
        data['Delivery'] = 'slow'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0.05
        data['Customer'] = 'Government'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert 'Discount' in newData['Result']
        assert newData['Result']['Discount'] == 0.15

    def test_AN_SNAP(self):
        '''
        Check that the supplied AN-SNAP rules (DMN).xlsx workbook works
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/AN-SNAP rules (DMN).xlsx')
        assert 'errors' not in status
        data = {}
        data['Multidisciplinary'] = False
        data['Admitted Flag'] = False
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list) == True
        assert 'Result' in newData[-1]
        assert 'Executed Rule' in newData[-1]
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert rule == 'General Multidisciplinary Flag Error Rule 1'
        assert 'AN-SNAP V4 code' in newData[-1]['Result']
        assert newData[-1]['Result']['AN-SNAP V4 code'] == '4999'
        data['Multidisciplinary'] = True
        data['Care Type'] = 'GEM'
        data['Single Day of Care'] = False
        data['Ongoing Pain'] = False
        data['Clinic'] = 'Memory'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list) == True
        assert 'Result' in newData[-1]
        assert 'Executed Rule' in newData[-1]
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert rule == 'Non-AdmittedGEM Rule 3'
        assert 'AN-SNAP V4 code' in newData[-1]['Result']
        assert newData[-1]['Result']['AN-SNAP V4 code'] == '4UC3'
        data['Care Type'] = 'Rehabilitation'
        del data['Single Day of Care']
        del data['Ongoing Pain']
        del data['Clinic']
        data['Patient Age'] = 19
        data['Assessment Only'] = True
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list) == True
        assert 'Result' in newData[-1]
        assert 'Executed Rule' in newData[-1]
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert rule == 'Non-Admitted Adult Rehab Rule 1'
        assert 'AN-SNAP V4 code' in newData[-1]['Result']
        assert newData[-1]['Result']['Computed Age Type'] == '2'
        assert newData[-1]['Result']['AN-SNAP V4 code'] == '4SY1'
        data['Assessment Only'] = False
        data['AROC code'] = '7'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list) == True
        assert 'Result' in newData[-1]
        assert 'Executed Rule' in newData[-1]
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert rule == 'Non-Admitted Adult Rehab Rule 5'
        assert 'AN-SNAP V4 code' in newData[-1]['Result']
        assert newData[-1]['Result']['Computed Age Type'] == '2'
        assert newData[-1]['Result']['AN-SNAP V4 code'] == '4SG1'
        data['Patient Age Type'] = '1'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list) == True
        assert 'Result' in newData[-1]
        assert 'Executed Rule' in newData[-1]
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert rule == 'Non-Admitted Paed Rehab Rule 5'
        assert 'AN-SNAP V4 code' in newData[-1]['Result']
        assert newData[-1]['Result']['Patient Age'] == 19.0
        assert newData[-1]['Result']['Computed Age Type'] == '1'
        assert newData[-1]['Result']['AN-SNAP V4 code'] == '4X05'


    def test_testMatchNumber(self):
        '''
        Check matching a simple number (no operator)
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/MatchNumber.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value'] = 5
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Executed Rule' in newData
        assert newData['Executed Rule'] == ('Do DMNtest', 'dmnTest', 'pass')
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        assert 'RuleAnnotations' in newData
        assert newData['RuleAnnotations'] == [('Description', 'Input Value == 5'), ('Reference', 'Rule Pass')]
        data['Input Value'] = 6
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_testEqualsNumber(self):
        '''
        Check matching a simple number ( = operator)
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/EqualsNumber.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value'] = 5
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value'] = 6
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_testNotEqualsNumber(self):
        '''
        Check not matching a simple number ( != operator)
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/NotEqualsNumber.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value'] = 6
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value'] = 5
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_testLessThanNumber(self):
        '''
        Check less than a simple number ( < operator)
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/LessThanNumber.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value'] = 4
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value'] = 5
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_GreaterThanNumber(self):
        '''
        Check greater than a simple number ( > operator)
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/GreaterThanNumber.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value'] = 6
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value'] = 5
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_LessThanOrEqualsNumber(self):
        '''
        Check less than or equals a simple number ( <= operator)
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/LessThanOrEqualsNumber.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value'] = 4
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value'] = 5
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value'] = 6
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_GreaterThanOrEqualsNumber(self):
        '''
        Check greater than or equals a simple number ( >= operator)
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/GreaterThanOrEqualsNumber.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value'] = 6
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value'] = 5
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value'] = 4
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_testMatchString1(self):
        '''
        Check matches a simple string (not enclosed in double quotes - no spaces)
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/MatchString1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value'] = 'abc'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value'] = 'def'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_testMatchString2(self):
        '''
        Check matches a complex string (enclosed in double quotes with a space)
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/MatchString2.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value'] = 'a c'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value'] = 'd f'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False


    def test_testEqualsString1(self):
        '''
        Check matches a simple string with equals operator (not enclosed in double quotes - no spaces)
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/EqualsString1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value'] = 'abc'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value'] = 'def'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_testEqualsString2(self):
        '''
        Check matches a complex string with equals operator (enclosed in double quotes with a space)
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/EqualsString2.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value'] = 'a c'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value'] = 'd f'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_testLessThanString(self):
        '''
        Check matches a complex string with equals operator (enclosed in double quotes with a space)
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/LessThanString.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value'] = ' bc'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value'] = 'abc'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_testIn1(self):
        '''
        Check string in list
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/In1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value'] = 'a'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value'] = 'd'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_testNotIn1(self):
        '''
        Check string not in list
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/NotIn1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value'] = 'a'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Value'] = 'd'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True

    def test_testVariableIn1(self):
        '''
        Check variable in list
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/InVariable.xlsx')
        assert 'errors' not in status
        data = {}
        data['Container'] = ['AAA','BBB','XXX']
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Container'] = ['AAA','BBB','CCC']
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_testANSNAP1(self):
        '''
        Check AN-SNAP decision
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/AN-SNAP rules (DMN).xlsx')
        assert 'errors' not in status
        data = {}
        data['Multidisciplinary'] = False
        data['Admitted Flag'] = False
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list)
        assert 'Result' in newData[-1]
        assert 'AN-SNAP V4 code' in newData[-1]['Result']
        assert newData[-1]['Result']['AN-SNAP V4 code'] == '4999'
        assert 'Executed Rule' in newData[-1]
        assert len(newData[-1]['Executed Rule']) == 3
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert rule == 'General Multidisciplinary Flag Error Rule 1'
        data['Multidisciplinary'] = True
        data['Care Type'] = 'GEM'
        data['Single Day of Care'] = False
        data['Ongoing Pain'] = False
        data['Clinic'] = 'Memory'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list)
        assert 'Result' in newData[-1]
        assert 'AN-SNAP V4 code' in newData[-1]['Result']
        assert newData[-1]['Result']['AN-SNAP V4 code'] == '4UC3'
        assert 'Executed Rule' in newData[-1]
        assert len(newData[-1]['Executed Rule']) == 3
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert rule == 'Non-AdmittedGEM Rule 3'
        data['Care Type'] = 'Rehabilitation'
        del data['Single Day of Care']
        del data['Ongoing Pain']
        del data['Clinic']
        data['Patient Age'] = 19
        data['Assessment Only'] = True
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list)
        assert 'Result' in newData[-1]
        assert 'AN-SNAP V4 code' in newData[-1]['Result']
        assert newData[-1]['Result']['AN-SNAP V4 code'] == '4SY1'
        assert 'Executed Rule' in newData[-1]
        assert len(newData[-1]['Executed Rule']) == 3
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert rule == 'Non-Admitted Adult Rehab Rule 1'
        data['Assessment Only'] = False
        data['AROC code'] = '7'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list)
        assert 'Result' in newData[-1]
        assert 'AN-SNAP V4 code' in newData[-1]['Result']
        assert newData[-1]['Result']['AN-SNAP V4 code'] == '4SG1'
        assert 'Executed Rule' in newData[-1]
        assert len(newData[-1]['Executed Rule']) == 3
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert rule == 'Non-Admitted Adult Rehab Rule 5'
        data['Patient Age Type'] = '1'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list)
        assert 'Result' in newData[-1]
        assert 'AN-SNAP V4 code' in newData[-1]['Result']
        assert newData[-1]['Result']['AN-SNAP V4 code'] == '4X05'
        assert 'Executed Rule' in newData[-1]
        assert len(newData[-1]['Executed Rule']) == 3
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert rule == 'Non-Admitted Paed Rehab Rule 5'

    def test_testANSNAP2(self):
        '''
        Check AN-SNAP decision
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/AN-SNAP rules (DMN).xlsx')
        assert 'errors' not in status
        thisPatient = thisAdmission = None
        with open('../pyDMNrules/subAcuteExtract.csv', 'r', newline='') as csvInFile:
            csvReader = csv.DictReader(csvInFile, dialect=csv.excel)
            for row in csvReader:
                ''' Let's go for SNAP '''
                for col in row:                 # Start by mapping the booleans
                    if row[col] == 'TRUE':
                        row[col] = True
                    elif row[col] == 'True':
                        row[col] = True
                    elif row[col] == 'true':
                        row[col] = True
                    elif row[col] == 'FALSE':
                        row[col] = False
                    elif row[col] == 'False':
                        row[col] = False
                    elif row[col] == 'false':
                        row[col] = False
                    elif row[col] == 'NULL':
                        row[col] = None
                    elif row[col] == 'Null':
                        row[col] = None
                    elif row[col] == 'null':
                        row[col] = None
                data = {}
                data['Multidisciplinary'] = True
                data['Admitted Flag'] = True
                data['Care Type'] = row['Care_Type']
                LOS = int(float(row['Phase_Length_of_Stay']))
                data['Length of Stay'] = LOS
                if LOS >= 92:
                    data['Long term care'] = True
                else:
                    data['Long term care'] = False
                data['Same-day admitted care'] = row['Same_day_admitted_care']
                data['GEM clinic'] = None                       # Required for non-admitted GEM AN-SNAP. No non-admitted GEM records
                data['Patient Age'] = int(float(row['Patient_Age']))
                data['Patient Age Type'] = None
                data['AROC code'] = row['AROC_code']
                data['FIM Motor score'] = int(float(row['Phase_FIM_Motor_Score']))
                data['FIM Cognition score'] = int(float(row['Phase_FIM_Cognition_Score']))
                data['Delirium or Dimentia'] = row['Delirium_or_Dimentia']
                data['Phase Type'] = row['Phase_Type']
                if ((row['Patient_UR'] != thisPatient) or (row['Admission_Date'] != thisAdmission)):
                    thisPatient = row['Patient_UR']
                    thisAdmission = row['Admission_Date']
                    if row['Phase_Type'] == 'Unstable':
                        data['First Phase'] = True
                    else:
                        data['First Phase'] = False
                else:
                    data['First Phase'] = False
                data['RUG-ADL'] = int(float(row['Phase_RUG_ADL_Score']))
                (status, newData) = dmnRules.decide(data)        
                assert isinstance(newData, list)
                assert 'Result' in newData[-1]
                assert 'AN-SNAP V4 code' in newData[-1]['Result']
                assert newData[-1]['Result']['AN-SNAP V4 code'] == row['Phase_AN_SNAP_V4_0_code']

    def test_testANWU21(self):
        '''
        Check ANWU21 decision
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/Subacute NWAU calculator (DMN).xlsx')
        assert 'errors' not in status
        thisPatient = thisAdmission = None
        with open('../pyDMNrules/subAcuteExtract.csv', 'r', newline='') as csvInFile:
            csvReader = csv.DictReader(csvInFile, dialect=csv.excel)
            for row in csvReader:
                ''' Let's go for SNAP '''
                for col in row:                 # Start by mapping the booleans
                    if row[col] == 'TRUE':
                        row[col] = True
                    elif row[col] == 'True':
                        row[col] = True
                    elif row[col] == 'true':
                        row[col] = True
                    elif row[col] == 'FALSE':
                        row[col] = False
                    elif row[col] == 'False':
                        row[col] = False
                    elif row[col] == 'false':
                        row[col] = False
                    elif row[col] == 'NULL':
                        row[col] = None
                    elif row[col] == 'Null':
                        row[col] = None
                    elif row[col] == 'null':
                        row[col] = None
                data = {}
                data['Care Type'] = row['Care_Type']
                data['LOS'] = int(float(row['Phase_Length_of_Stay']))
                data['Same Day Admission'] = row['Same_day_admitted_care']
                data['Patient Age'] = row['Patient_Age']
                data['AN-SNAP V4.0'] = row['Phase_AN_SNAP_V4_0_code']
                data['Care Type'] = row['Care_Type']
                data['Hospital Remoteness'] = '0'           # Hospital is not remote
                data['Postcode'] = row['Postcode']
                data['SA2'] = None                          # Addresses are not geocoded
                data['Dialysis Flag'] = row['Dialysis_Flag']
                data['RadioTherapy Flag'] = row['RadioTherapy_Flag']
                data['Funding Source'] = row['Funding_Source']
                data['Indigenous Status'] = row['Indigenous_Status']
                data['State'] = '2'                         # Funded by DHHS Victoria
                (status, newData) = dmnRules.decide(data)
                assert isinstance(newData, list)
                assert 'Result' in newData[-1]
                assert 'NWAU21' in newData[-1]['Result']
                assert (int(newData[-1]['Result']['NWAU21'] * 10000.0 + 0.5) / 10000.0) == (int(float(row['Phase_NWAU21']) * 10000.0 + 0.5) / 10000.0)

    def test_testANSNAPpandas(self):
        '''
        Check AN-SNAP decision using Pandas DataFrames
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/AN-SNAP rules (DMN).xlsx')
        assert 'errors' not in status
        dataTypes = {'Patient_Age':int, 'Episode_Length_of_stay':int, 'Phase_Length_of_stay':int,
        'Phase_FIM_Motor_Score':int, 'Phase_FIM_Cognition_Score':int, 'Phase_RUG_ADL_Score':int,
        'Phase_w01':float, 'Phase_NWAU21':float,
        'Indigenous_Status':str, 'Care_Type':str, 'Funding_Source':str, 'Phase_Impairment_Code':str,
        'AROC_code':str, 'Phase_AN_SNAP_V4_0_code':str,
        'Same_day_addmitted_care':bool, 'Delerium_or_Dimentia':bool, 'RadioTherapy_Flag':bool, 'Dialysis_Flag':bool}
        dates = ['BirthDate', 'Admission_Date', 'Discharge_Date', 'Phase_Start_Date', 'Phase_End_Date']
        dfInput = pd.read_csv('../pyDMNrules/subAcuteExtract.csv', dtype=dataTypes, parse_dates=dates)
        dfInput['Multidisciplinary'] = True
        dfInput['Admitted_Flag'] = True
        dfInput['Length_of_Stay'] = dfInput['Phase_Length_of_Stay']
        dfInput['Long_term_care'] = False
        dfInput.loc[dfInput['Length_of_Stay'] > 92, 'Long_term_care'] = True
        dfInput['Same_day_admitted_care'] = False
        dfInput['GEM_clinic'] = None
        dfInput['Patient_Age_Type'] = None
        dfInput['First_Phase'] = False
        grouped = dfInput.groupby(['Patient_UR', 'Admission_Date'])
        for index in grouped.head(1).index:
            if dfInput.loc[index]['Phase_Type'] == 'Unstable':
                dfInput.loc[index, 'First_Phase'] = True
        columns = {'Admitted_Flag':'Admitted Flag', 'Care_Type':'Care Type', 'Length_of_Stay':'Length of Stay', 'Long_term_care':'Long term care',
        'Same_day_admitted_care':'Same-day admitted care', 'GEM_clinic':'GEM clinic', 'Patient_Age':'Patient Age', 'Patient_Age_Type':'Patient Age Type',
        'AROC_code':'AROC code', 'Delirium_of_Dimentia':'Delirium or Dimentia', 'Phase_Type':'Phase Type', 'First_Phase':'First Phase', 'Phase_FIM_Motor_Score':'FIM Motor score',
        'Phase_FIM_Cognition_Score':'FIM Cognition score', 'Phase_RUG_ADL_Score':'RUG-ADL', 'Delirium_or_Dimentia':'Delirium or Dimentia',
        'Problem_Severity_Scrore':'Problem Severity Score'}
        (dfStatus, dfResults, dfDecision) = dmnRules.decidePandas(dfInput, headings=columns)
        assert dfStatus.where(dfStatus == 'no errors').count() == dfStatus.count()
        assert dfResults['AN_SNAP_V4_code'].count() == dfInput['Phase_AN_SNAP_V4_0_code'].count()
        for index in dfResults.index:
            assert dfResults.loc[index, 'AN_SNAP_V4_code'] == dfInput.loc[index, 'Phase_AN_SNAP_V4_0_code']

