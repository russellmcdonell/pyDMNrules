# pyDMNrules
An implementation DMN (Decision Model Notation) in Python, using the [pySFeel](https://github.com/russellmcdonell/pySFeel) and openxlrd modules.

DMN rules are read from an Excel workbook.
Then data, matching the input variables in the DMN rules is passed to the decide() function.
The returned data contains the decion.

The passed Excel workbook must contain the two tabs 'Glossary' and 'Decision'.
Other tabs contain as many DMN rules tables as necessary.

The 'Glossary' tab must contain a table headed 'Glossary'.
This table must contain three columns with the headings 'Variable', 'Business Concept' and 'Attribute'.
The data in these three columns describe the inputs and outputs associated with the 'Decision'.
'Variable' can be any text and is used to pass data to and from pyDMNrules.
'Business Concept' and 'Attribute' must be valid S-FEEL names, but may **not** contain the dot, or period character.
Items in the Glossary can be outputs of one decision table and inputs in the next decision table, which makes it easy to support complex business models.

| Glossary |                  |           |
|----------|------------------|-----------|
| Variable | Business Concept | Sector    |
| Customer | Customer         | sector    |
| OderSize | Order            | orderSize |
| Delivery |                  | delivery  |
| Discount | Discount         | discount  |

The 'Decision' tab must contain a table headed 'Decision' with the headings 'Decisions' and 'Execute Decision Table'.
The row contain a description and the name of a DMN rules table. pyDMNrules will execute each decision table in this order.

| Decision           |                            |
|--------------------|----------------------------|
| **Decisions**      | **Execute Decision Table** |
| Determine Discount | Discount                   |

The decision table(s) must exist on other spreadsheets in they Excel workbook. Decision tables have colums of inputs and columns of outut, with a double vertical line border between them. The headings for the input and output columns must be Variables from the Glossary.

# USAGE:

    import pyDMNrules
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
    print('Decision',repr(newData))
    if 'errors' in status:
        print('With errors', status['errors'])

newData will contain all the items listed in the Glossary, with their final assigned value.

    $ python3 HPV.py
    ExampleHPV.xlsx loaded
    Testing {'Participant Age': 36, 'In Test of Cure': True, 'Hysterectomy Flag': False, 'Cancer Flag': False, 'HPV-V': 'V0', 'Current Participant Risk Category': 'low'}
    Decision {'Immune Deficient Flag': None, 'Hysterectomy Flag': False, 'Cancer Flag': False, 'In Test of Cure': True, 'Participant Age': 36.0, 'Current Participant Risk Category': 'low', 'HPV-V': 'V0', 'Cyto-S': None, 'Cyto-E': None, 'Cyto-O': None, 'Collection Method': None, 'Test Risk Code': 'L', 'New Participant Risk Category': 'low', 'Participant Care Pathway': 'toBeDetermined', 'Next Rule': 'CervicalRisk2', 'Execute Rules': [('FirstTestOfCervicalRisk', '20')]}

# Testing

pyDMNrules reserves the spreadsheet name 'Test' which can be used for storing test data. The function test() read the test data and passed through the decide() function and assembles the results which are returned to the caller.

The 'Test' spreadsheet must contain unit test data tables. Each unit test data table must be labeled with a Business Concept from the Glossary. The heading of the table must be Variables from the Glossary associated with the Business Concept. The data, in each column, must be valid input data for the associated Variable. Each row is considered a unit test data set.

pyDMNrules also searches for a special table which must be called 'DMNrulesTests' which must exist. The 'DMNrulesTests' table has input columns followed by output columns, with a double vertical line border between them. The headings for the input columns must be Business Concepts from the Glossary. The data in input columns must be one based indexes into the matching unit test data table. The headings for the output columns must be Variables from the Glossary. The data in output columns will be valid data for the matching Variable and are the expected values returned by the decide() function.

The test() function will process each row in the 'DMNrulesTests' table, getting data from the unit test data tables and building a data{} dictionary. The test() function then passes this data{} dictionary to the decide() function. The returned newData{} is then compared to the output data for the matching test in the 'DMNrulesTests' table. The data{}, newData{} and a list of mismatched is then appended to the list of results, which is eventually passed back to the caller of the test() function.

The returned list is a list of dictionaries, one for each test in the 'DMNrulesTests' table. The key to this structure are
- 'Test ID' - the one based index into the 'DMNrulesTests' table which identifies which test which was run
- 'Annotations'(optional) - the list of annotation for this test - not present if no annotations were present in 'DMNrulesTests'
- 'data' - the dictionary of assembed data passed to the decide() function
- 'newData' - the dictionary returned by the decide() function
- 'dataAnnotations'(optional) - the list of annotations from the unit test data tables, for the unit test sets used in this test - not present if no annotations were persent in any of the unit test data tables for the selected unit test sets.
- 'Mismatches'(optional) - the list of mismatch reports, one for each output 'DMNrulesTests' table output value that did not match the value returned from the decide() function - not present if all the data retuned from the decide() function matched the values in the 'DMNrulesTest' table.

# Note
If an output should be the value of an input, then you can use the Variable from the Glossary. However, if the output is a manuipulation of an input, then you will need to use the internal name for that variable, being the Business Concept and the Attibute concateneted with the period character. That is 'Patient Age' is valid, but 'Patient Age + 5' is not. Instead you will need to use the syntax 'Patient.age + 5'

