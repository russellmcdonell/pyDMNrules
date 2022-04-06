import pyDMNrules
from openpyxl import load_workbook
import csv
import pandas as pd
import datetime

class TestClass:
    def test_HPV1(self):
        '''
        Check that the supplied ExampleHPV.xlsx workbook works
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
        Check that the supplied ExampleHPV.xlsx workbook works when loaded and passed as a workbook
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
        assert results[1]['newData']['Result']['Error Message'] is None
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
        Check that the supplied AN-SNAP V4 grouper (DMN).xlsx workbook works
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/AN-SNAP V4 grouper (DMN).xlsx')
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
        status = dmnRules.load('../pyDMNrules/AN-SNAP V4 grouper (DMN).xlsx')
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
        status = dmnRules.load('../pyDMNrules/AN-SNAP V4 grouper (DMN).xlsx')
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
                for col in ['Multidisciplinary', 'Admitted Flag', 'Care Type', 'Same-day admitted care', 'GEM clinic',
                            'Patient Age Type', 'AROC code', 'Delirium or Dimentia', 'Phase Type']:
                    data[col] = row[col]
                for col in ['Length of Stay', 'Patient Age', 'FIM Motor score', 'FIM Cognition score', 'RUG-ADL']:
                    data[col] = int(float(row[col]))
                if data['Length of Stay'] >= 92:
                    data['Long term care'] = True
                else:
                    data['Long term care'] = False
                if ((row['Patient UR'] != thisPatient) or (row['Episode Start Date'] != thisAdmission)):
                    thisPatient = row['Patient UR']
                    thisAdmission = row['Episode Start Date']
                    if row['Phase Type'] == 'Unstable':
                        data['First Phase'] = True
                    else:
                        data['First Phase'] = False
                else:
                    data['First Phase'] = False
                (status, newData) = dmnRules.decide(data)        
                assert isinstance(newData, list)
                for i in range(len(status)):
                    assert 'errors' not in status[i]
                assert 'Result' in newData[-1]
                assert 'AN-SNAP V4 code' in newData[-1]['Result']
                assert newData[-1]['Result']['AN-SNAP V4 code'] == row['Expected AN-SNAP V4 code']

    def test_testANWU21(self):
        '''
        Check ANWU21 decision
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/Subacute NWAU21 calculator (DMN).xlsx')
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
                for col in ['Care Type', 'Hospital Remoteness', 'Postcode', 'SA2',
                            'Dialysis Flag', 'RadioTherapy Flag', 'Funding Source', 'Indigenous Status', 'State']:
                    data[col] = row[col]
                data['Same Day Admission'] = row['Same-day admitted care']
                data['AN-SNAP V4.0'] = row['Expected AN-SNAP V4 code']
                for col in ['Length of Stay', 'Patient Age']:
                    data[col] = int(float(row[col]))
                (status, newData) = dmnRules.decide(data)
                assert isinstance(newData, list)
                for i in range(len(status)):
                    assert 'errors' not in status[i]
                assert 'Result' in newData[-1]
                assert 'NWAU21' in newData[-1]['Result']
                assert (int(newData[-1]['Result']['NWAU21'] * 10000.0 + 0.5) / 10000.0) == (int(float(row['Expected NWAU21']) * 10000.0 + 0.5) / 10000.0)

    def test_testANSNAPpandas(self):
        '''
        Check AN-SNAP decision using Pandas DataFrames
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/AN-SNAP V4 grouper (DMN).xlsx')
        assert 'errors' not in status
        dfInput = pd.read_excel('../pyDMNrules/SubacuteExtract.xlsx')
        dfInput['Long term care'] = False
        dfInput.loc[dfInput['Length of Stay'] > 92, 'Long term care'] = True
        dfInput['First Phase'] = False
        thisPatient = thisEpisode = None
        for index, row in dfInput.iterrows():
            if ((row['Patient UR'] != thisPatient) or (row['Episode Start Date'] != thisEpisode)):
                thisPatient = row['Patient UR']
                thisEpisode = row['Episode Start Date']
                if row['Phase Type'] == 'Unstable':
                    dfInput.loc[index, 'First Phase'] = True
        (dfStatus, dfResults, dfDecision) = dmnRules.decidePandas(dfInput)
        for index, value in dfStatus.items():
            assert value == 'no errors'
        assert dfResults['AN-SNAP V4 code'].count() == dfInput['Expected AN-SNAP V4 code'].count()
        for index in dfResults.index:
            assert dfResults.loc[index, 'AN-SNAP V4 code'] == dfInput.loc[index, 'Expected AN-SNAP V4 code']

    def test_is(self):
        '''
        Check is() function
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/Is1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Value1'] = 7
        data['Input Value2'] = 9
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value2'] = 'AAA'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Value1'] = 'abc'
        data['Input Value2'] = 'DEF'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value2'] = 7
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Value1'] = True
        data['Input Value2'] = False
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value2'] = 7
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Value1'] = '@"P2Y3M"'
        data['Input Value2'] = '@"P5Y0M"'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value2'] = 7
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Value1'] = datetime.timedelta(days=1, seconds=3000)
        data['Input Value2'] = datetime.timedelta(days=3)
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value2'] = 7
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Value1'] = datetime.time(hour=11, minute=3, second=19)
        data['Input Value2'] = datetime.time(hour=1, minute=13, second=11)
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value2'] = 7
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Value1'] = datetime.date(year=2021, month=9, day=15)
        data['Input Value2'] = datetime.date(year=2020, month=3, day=7)
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value2'] = 7
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Value1'] = datetime.datetime(year=2021, month=9, day=15, hour=15, minute=15, second=10)
        data['Input Value2'] = datetime.datetime(year=2020, month=5, day=7, hour=11, minute=5, second=0)
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Value2'] = 7
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_beforeRange(self):
        '''
        Check before() function
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/Before1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Range1'] = 1
        data['Input Range2'] = 10
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = 10
        data['Input Range2'] = 1
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = 1
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = 1
        data['Input Range2'] = ('(', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = 10
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 10, ')')
        data['Input Range2'] = 10
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = 15
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('[', 15, 20, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('[', 10, 20, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 10, ')')
        data['Input Range2'] = ('[', 10, 20, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('(', 10, 20, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True

    def test_afterRange(self):
        '''
        Check before() function
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/After1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Range1'] = 10
        data['Input Range2'] = 5
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = 5
        data['Input Range2'] = 10
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = 12
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = 10
        data['Input Range2'] = ('(', 1, 10, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = 10
        data['Input Range2'] = ('(', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 11, 20, ')')
        data['Input Range2'] = 12
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 11, 20, ']')
        data['Input Range2'] = 10
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 11, 20, ']')
        data['Input Range2'] = 11
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 11, 20, ']')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('[', 11, 20, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 11, 20, ']')
        data['Input Range2'] = ('[', 1, 11, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 11, 20, ']')
        data['Input Range2'] = ('[', 1, 11, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True

    def test_meetsRange(self):
        '''
        Check before() function
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/Meets1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('[', 5, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 5, ')')
        data['Input Range2'] = ('[', 5, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('(', 5, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('[', 6, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_metByRange(self):
        '''
        Check before() function
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/MetBy1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Range1'] = ('[', 5, 10, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 5, 10, ']')
        data['Input Range2'] = ('[', 1, 5, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('(', 5, 10, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 6, 10, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_overlapsRange(self):
        '''
        Check before() function
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/Overlaps1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('[', 3, 8, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 3, 8, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 8, ']')
        data['Input Range2'] = ('[', 3, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 3, 5, ']')
        data['Input Range2'] = ('[', 1, 8, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('[', 6, 8, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 6, 8, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('[', 5, 8, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('(', 5, 8, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ')')
        data['Input Range2'] = ('[', 5, 8, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ')')
        data['Input Range2'] = ('(', 5, 8, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 5, 8, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 5, 8, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 5, 8, ']')
        data['Input Range2'] = ('[', 1, 5, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('(', 5, 8, ']')
        data['Input Range2'] = ('[', 1, 5, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_overlapsBeforeRange(self):
        '''
        Check before() function
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/OverlapsBefore1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('[', 3, 8, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('[', 6, 8, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('[', 5, 8, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('(', 5, 8, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ')')
        data['Input Range2'] = ('(', 5, 8, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ')')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ')')
        data['Input Range2'] = ('(', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('(', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 5, ')')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_overlapsAfterRange(self):
        '''
        Check before() function
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/OverlapsAfter1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Range1'] = ('[', 3, 8, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 6, 8, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 5, 8, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 5, 8, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 5, 8, ']')
        data['Input Range2'] = ('[', 1, 5, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('(', 1, 5, ']')
        data['Input Range2'] = ('[', 1, 5, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 1, 5, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('[', 1, 5, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_finishesRange(self):
        '''
        Check before() function
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/Finishes1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Range1'] = 10
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = 10
        data['Input Range2'] = ('[', 1, 10, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 5, 10, ']')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 5, 10, ')')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 5, 10, ')')
        data['Input Range2'] = ('[', 1, 10, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 1, 10, ']')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('(', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True

    def test_includesRange(self):
        '''
        Check before() function
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/Includes1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = 5
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = 12
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = 1
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = 10
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 1, 10, ']')
        data['Input Range2'] = 1
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 10, ')')
        data['Input Range2'] = 10
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('[', 4, 6, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 1, 10, ']')
        data['Input Range2'] = ('(', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('(', 1, 10, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ')')
        data['Input Range2'] = ('[', 5, 10, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('[', 1, 10, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('(', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True

    def test_duringRange(self):
        '''
        Check before() function
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/During1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Range1'] = 5
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = 12
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = 1
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = 10
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = 1
        data['Input Range2'] = ('(', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = 10
        data['Input Range2'] = ('[', 1, 10, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 4, 6, ']')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 1, 5, ']')
        data['Input Range2'] = ('(', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 1, 10, ')')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 5, 10, ')')
        data['Input Range2'] = ('[', 1, 10, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ')')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 1, 10, ']')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True

    def test_startsRange(self):
        '''
        Check before() function
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/Starts1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Range1'] = 1
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = 1
        data['Input Range2'] = ('(', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = 2
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 1, 5, ']')
        data['Input Range2'] = ('(', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 1, 5, ']')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('(', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 10, ')')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ')')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 1, 10, ')')
        data['Input Range2'] = ('(', 1, 10, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True

    def test_startedByRange(self):
        '''
        Check before() function
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/StartedBy1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = 1
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 1, 10, ']')
        data['Input Range2'] = 1
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = 2
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 1,10, ']')
        data['Input Range2'] = ('(', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('(', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('(', 1, 10, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('[', 1, 10, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('[', 1, 10, ']')
        data['Input Range2'] = ('[', 1, 10, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 1, 10, ')')
        data['Input Range2'] = ('(', 1, 10, ')')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True

    def test_coincidesRange(self):
        '''
        Check before() function
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/tests/Coincides1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Input Range1'] = 5
        data['Input Range2'] = 5
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = 3
        data['Input Range2'] = 4
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('[', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == True
        data['Input Range1'] = ('(', 1, 5, ')')
        data['Input Range2'] = ('(', 1, 5, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False
        data['Input Range1'] = ('[', 1, 5, ']')
        data['Input Range2'] = ('(', 2, 6, ']')
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert 'Result' in newData
        assert 'Output Value' in newData['Result']
        assert newData['Result']['Output Value'] == False

    def test_HPV3(self):
        '''
        Check that the supplied ExampleHPVnoGlossary.xlsx workbook works
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/ExampleHPVnoGlossary.xlsx')
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
        (testStatus, results) = dmnRules.test()
        assert 'errors' not in testStatus
        assert len(results) == 26
        for i in range(len(results)):
            assert 'Mismatches' not in results[i]

    def test_ExecuteRows(self):
        '''
        Check that the supplied ExampleExecuteRows1.xlsx workbook works
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/ExampleExecuteByRows1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Patient Birthdate'] = datetime.date(year=1952, month=11, day=11)
        data['Admission Date'] = datetime.date(year=2021, month=1, day=16)
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert decision == 'Determine Patient Age Group'
        assert table == 'Patient Age Group'
        assert rule == '1'
        assert newData[-1]['Result']['Age Group'] == 13
        assert newData[-1]['Result']['Computed Patient Age'] == 69

    def test_ExecuteColumns(self):
        '''
        Check that the supplied ExampleExecuteColumns1.xlsx workbook works
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/ExampleExecuteByColumns1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Patient Birthdate'] = datetime.date(year=1952, month=11, day=11)
        data['Admission Date'] = datetime.date(year=2021, month=1, day=16)
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert decision == 'Determine Patient Age Group'
        assert table == 'Patient Age Group'
        assert rule == '1'
        assert newData[-1]['Result']['Age Group'] == 13
        assert newData[-1]['Result']['Computed Patient Age'] == 69

    def test_ExecuteCrossTabs(self):
        '''
        Check that the supplied ExampleExecuteCrosstab1.xlsx workbook works
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/ExampleExecuteCrosstab1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Patient Birthdate'] = datetime.date(year=1952, month=11, day=11)
        data['Admission Date'] = datetime.date(year=2021, month=1, day=16)
        data['Admission Weight'] = 125
        data['Discharge Weight'] = 115.5
        data['want Age'] = True
        data['want Weight Loss'] = False
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list)
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert decision == 'Determine Answers'
        assert table == 'Do As You Are Told'
        assert rule == '2:1'
        assert len(newData) == 2
        assert 'Result' in newData[-1]
        assert 'Computed Patient Age' in newData[-1]['Result']
        assert newData[-1]['Result']['Computed Patient Age'] == 69
        assert 'Computed Weight Loss' in newData[-1]['Result']
        assert newData[-1]['Result']['Computed Weight Loss'] == None
        data['want Age'] = False
        data['want Weight Loss'] = True
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list)
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert decision == 'Determine Answers'
        assert table == 'Do As You Are Told'
        assert rule == '1:2'
        assert len(newData) == 2
        assert 'Result' in newData[-1]
        assert 'Computed Patient Age' in newData[-1]['Result']
        assert newData[-1]['Result']['Computed Patient Age'] == None
        assert 'Computed Weight Loss' in newData[-1]['Result']
        assert newData[-1]['Result']['Computed Weight Loss'] == 9.5
        data['want Age'] = True
        data['want Weight Loss'] = True
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list)
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert decision == 'Determine Answers'
        assert table == 'Do As You Are Told'
        assert rule == '2:2'
        assert len(newData) == 4
        assert 'Result' in newData[-1]
        assert 'Computed Patient Age' in newData[-1]['Result']
        assert newData[-1]['Result']['Computed Patient Age'] == 69
        assert 'Computed Weight Loss' in newData[-1]['Result']
        assert newData[-1]['Result']['Computed Weight Loss'] == 9.5
        data['want Age'] = False
        data['want Weight Loss'] = False
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list) == False
        (decision, table, rule) = newData['Executed Rule']
        assert decision == 'Determine Answers'
        assert table == 'Do As You Are Told'
        assert rule == '1:1'
        assert 'Result' in newData
        assert 'Computed Patient Age' in newData['Result']
        assert newData['Result']['Computed Patient Age'] == None
        assert 'Computed Weight Loss' in newData['Result']
        assert newData['Result']['Computed Weight Loss'] == None

    def test_ExecuteInterest(self):
        '''
        Check that the supplied ExampleExecuteInterest1.xlsx workbook works
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/ExampleExecuteInterest1.xlsx')
        assert 'errors' not in status
        data = {}
        data['Years'] = 10
        data['Interest'] = 0.03
        data['Price'] = 125
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list)
        (decision, table, rule) = newData[-1]['Executed Rule']
        assert decision == 'Determine Compound Interest'
        assert table == 'Compute Interest'
        assert rule == '1'
        assert len(newData) == 11
        assert 'Result' in newData[-1]
        assert 'Price' in newData[-1]['Result']
        assert newData[-1]['Result']['Price'] == 167.98954741801523
        data['Years'] = 110
        data['Interest'] = 0.003
        data['Price'] = 125
        (status, newData) = dmnRules.decide(data)
        assert 'errors' in status

    def test_NoglossNodecision(self):
        '''
        Check that the no Glossary and no Decision example works
        Slightly more complicated because the hitPolicy is 'R'
        '''
        dmnRules = pyDMNrules.DMN()
        status = dmnRules.load('../pyDMNrules/TherapyNoglossNodecision.xlsx')
        assert 'errors' not in status
        data = {}
        data['Patient Age'] = 56
        data['Patient Allergies'] = ['Penicillin', 'Streptomycin']
        data['Patient Creatinine Level'] = 2.0
        data['Patient Weight'] = 78
        data['Patient Active Medication'] = 'Coumadin'
        data['Encounter Diagnosis'] = 'Acute Sinusitis'
        (status, newData) = dmnRules.decide(data)
        assert 'errors' not in status
        assert isinstance(newData, list)        # A list of dictionaries
        assert len(newData) == 8
        assert isinstance(newData[-1], dict)
        assert 'Executed Rule' in newData[-1]
        assert isinstance(newData[-1]['Executed Rule'], list)       # A list - one entry per rule run
        (decision, table, rule) = newData[-1]['Executed Rule'][-1]
        assert decision == 'Decide DecisionMedication'
        assert table == 'DecisionMedication'
        assert rule == 'warnInt'
        assert 'Result' in newData[-1]
        assert 'Recommended Medication' in newData[-1]['Result']
        assert newData[-1]['Result']['Recommended Medication'] == 'Levofloxacin'
        assert 'Recommended Dose' in newData[-1]['Result']
        assert newData[-1]['Result']['Recommended Dose'] == '500mg every 24 hours for 14 days'
        assert 'Warning' in newData[-1]['Result']
        assert newData[-1]['Result']['Warning'] is not None
        (testStatus, results) = dmnRules.test()
        assert len(results) == 3
        for i in range(len(results)):
            assert 'Mismatches' not in results[i]
        assert len(testStatus) == 3
        for i in range(len(testStatus)):
            assert 'errors' not in testStatus[i]

