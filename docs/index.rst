.. pyDMNrules documentation master file, created by
   sphinx-quickstart on Fri Jan 24 05:25:44 2020.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

========================
pyDMNrules Documentation
========================


.. toctree::
   :maxdepth: 2
   :caption: Contents:


The pyDMNrules functions
========================

.. py:module:: pyDMNrules

.. py:class:: DMN

   .. automethod:: load

   .. automethod:: use

   .. automethod:: getGlossary

   .. automethod:: getDecision

   .. automethod:: getSheets

   .. automethod:: decide

   .. automethod:: decidePandas

   .. automethod:: test

   .. automethod:: loadTest

   .. automethod:: useTest


pyDMNrules
==========

pyDMNrules build a rules engine from an Excel workbook which may contain a 'Glossary' tab an a 'Decision' tab.
Other tabs will contain the DMN rules tables use to build the rules engine.

Glossary - optional
===================
The 'Glossary' tab, if it exists, must contain a table headed 'Glossary'. A 'Glossary' table lets the user add additional documentation/clarification
to the input and output data. Data items can be grouped in to Business Concepts to better document where inputs come
from and where outputs will be going to. The 'Glossary' table must have exactly three columns with the headings 'Variable',
'Business Concept' and 'Attribute'. The data in these three columns describe the inputs and outputs associated with the 'Decision'.  

'Variable' can be any text and is used to pass data to and from pyDMNrules.  
NOTE: the name 'Execute' is a reserved name and cannot be used as a Variable name. The name 'Execute' is reserved as a heading for output columns, output rows or the output variable of a Cross Tabs decision table. The output value(s) in an 'Execute' output column, 'Execute' output row or 'Execute' Cross Tabs decision table must be the name of a Decision Table which will be invoked when the matching rule is triggered. If the output value in an 'Execute' output column, an 'Execute' output or an output cell in an 'Execute' Cross Tabs decision table is missing (and empty cell), then no Decision Table will be invoked whe the matching rule is triggered.

'Business Concept' and 'Attribute' must be valid FEEL names as the 'Variable' will be mapped to a Glossary item created by joining these two FEEL names
with a period character, as in 'Order.orderSize'.  
Vertically merged cells in the 'Business Concept' column are used to group multiple 'Variable'/'Attribute' pairs together into a single 'Businss Concept'.  
The 'Glossary' can have multiple columns of Annotations to describe things like data types, valid values etc.

Items in the Glossary can be outputs of one decision table and inputs in the next decision table, which makes it easy to support complex business models.

Input decision cells and Output cells can refer to input and output values either by their 'Variable' name, or by their Glossary item (FEEL) name, of 'BusinessConcept.Attribute'

========= ================ =========
**Glossary**
------------------------------------
Variable  Business Concept Attribute
========= ================ =========
Customer  Customer         sector
OrderSize Order            orderSize
Delivery                   delivery
Discount  Discount         discount
========= ================ =========

If a 'Glossary' tab is not present - if no Glossary is supplied, then pyDMNrules will create one using the input and output headings.
There will be a single 'Business Concept' called 'Data'.
Attributes will be created by replacing any space characters (' ') with an underscore character ('_') and then removing all **invalid** FEEL name characters.
This can cause conflicts. For instance ':' is a valid character in a Variable, but not in a FEEL name.
If there are two Varibles, being 'C' and 'C:' then both will be transformed to an Attribute of 'C' and a Glossary item of 'Data.C', which will be ambiguous.
This ambiguity can be avoided by supplying a 'Glossary' and explicitely mapping 'C:' to a valid FEEL name, such as 'Ccolon'.
The function getGlossary() will return the Glossary (either supplied or constructed).

Decision -  optional
====================
In order to make a decision, pyDMNrules needs to know where to start, what to do, when and why.
The 'Decision' tab, if it exists, will contain a table headed 'Decision' which will provide the required details.
The heading for the 'Desion' table will be none or more Variable names, a heading of 'Decisions', a heading of 'Execute Decision Table' and
none or more headings for annotations. The 'Decisions' column is documentation and contain a brief description of the DMN rules table
which is named in 'Execute Decision Table' column which is in effect just a list of DMN rules tables to be run in the specified order.
If there are any 'Variable' columns then these columns must contain valid DMN test as if they were DMN rules tables input columns.
pyDMNrules now knows where to start (with the first DMN rules table), what to do (each DMN rules table in turn), when (when all the 'Variable' tests are 'true')
and why (because all the 'Variable' tests are 'true')

=========== ================== ====================== =================
Decision
-----------------------------------------------------------------------
Variable(s) Decisions          Execute Decision Table Annotation(s)
=========== ================== ====================== =================
input test  Determine Discount Discount               Compute discount
=========== ================== ====================== =================

The 'Variable' cells are evaluated on the fly, which means that a preceeding DMN rules table can set or clear a 'Variable'
which can effect which DMN rules tables are included, and which DMN rules tables are excluded, in the final decision.

The Decision tab, and associated 'Decision' table are optional because in many instances, what to do is obvious.
If the workbook does not contain a 'Decision' tab, then pyDMNrules will construct a 'Decision' table.

If there is only one DMN rules table pyDMNrules will create a 'Decision' table with just one row.
There can be several DMN rules tables that are independant and hence the order doesn't matter.
If there's more than one DMN rules table, pyDMNrules looks for dependancies and works out an acceptable order.
To do this, pyDMNrules looks at the inputs of each DMN rules table and check to see if that input is also an outputs from another DMN rules table.
If DMN rules table 'A' has an input 'X' and DMN rules table 'B' has an output of 'X', then pyDMNrules will ensure that DMN rules table 'B' is
is placed in the constructed 'Decision' table before DMN rules table 'A'.
The default order for the remaining table will be the order in which pyDMNrules found them in the workbook.
The function getDecision() will return the 'Decision' table (either supplied or constructed).

Decision - recommended
----------------------
The 'Decision' table tells pyDMNrules exacty what to do.
Only the DMN rules tables named in the 'Decision' table will be run (plus any DMN rules tables 'Execute'd
as a result of running the named DMN rules tables).
Without a 'Decision' table, pyDMNrules will run everything, even old, obsolete or test DMN rules tables
that you left in the workbook by mistake.
The exact DMN rules tables which were run, and the order in which they were run is returned be the 'decide()' and 'decidePandas()' functions.

It may not matter the order in which the DMN rules tables are run.
None the less, it may make you decision logic easier for others to understand if they can imagine them being executed in a specific order.

Decision - required
-------------------
In certain circumstances a 'Decision' table may be required, because pyDMNrules is unable to determine the correct decision order.

+ DMN rules table 'A' may have 'X' as an input and 'Y' as an output and DMN rules table 'B' may have 'Y' as an input and 'X' as an output.
  This may be valid; you may calculate 'Y' from an 'X' approximation and then calculate an exact value of 'X' now that you have a valid value for 'Y'.
  To pyDMNrules, this looks like a deadly embrace which it cannot resolve.
+ DMN rules table 'K' may 'Execute' DMN rules table 'L', which in turn will 'Excute' DMN rules table 'K' - an iterative process.
  pyDMNrules cannot determine the correct starting point.
+ Input values can be used to programatically determine which DMN rules table to 'Execute'.
  If 'PersonType' and 'AddressType' are inputs, then a rule in a DMN rules table could 'Execute' - (PersonType + AddressType).
  There would need to be a set of DMN rules tables; for each valid combination of 'PersonType' and 'AddressType', each with the matching DMN rules table name.
  pyDMNrules cannot determine runtime logic; it will create a 'Decision' table containing all the DMN rules tables and run them all.
+ The decision may require a DMN rules table to be run more than once.
  A 'Decision' table created by pyDMNrules will have one row for each DMN rule table.
  The decision logic can legitimately require a rule to be run more than one.
  You cannot create iteration loops in pyDMNrules, but if you wanted to calculate the first five terms in a polynomial
  then you would repeat the one DMN rules table five times in the 'Decision' table.

All of these examples **require** a user created 'Decision' table in order to ensure the desired decision logic.
There may be other scenario requiring a user created 'Decision' table.
Hence the recommendation to provide one for anything other than the most trivial decisions.

DMN rules tables
================
The DMN rules table(s) listed under **Execute Decision Table** in the 'Decision' table (either supplied or constructed),
must exist on other spreadsheets in the Excel workbook and can be 'Decisions as Rows', 'Decisions as Columns' or 'Decision as Crosstab' rules tables.

Decision as Rows rules tables
-----------------------------
Decisions as Rows DMN rules tables have columns of inputs and columns of outputs,
with a double line vertical border delineating where input columns stop and output columns begin.
- The headings for the input and output columns must be Variables from the Glossary or, for an output column, is can be the word 'Execute'.
- The input test cells under the input heading must contain expressions about the input Variable,
which will evaluate to True or False, depending upon the value of the input Variable.
- The values in the output results cells, under the output headings, are the values that will be assigned to the output Variables, if all the input cells evaluate to True on the same row. If the output heading is the word 'Execute', then the output value must be the name of a Decision Table, which will be invoked immediately. Which means any output values already assigned can be overwritten by the invoked Decision Table, and any outputs created by the invoked Decision Table can be overwritten by subsequent outputs in the current Decision Table.

[ExampleRows.xlsx](https://github.com/russellmcdonell/pyDMNrules/blob/master/ExampleRows.xlsx) is an example fo a Decision as Rows DMN table.

Decision as Columns rules tables
--------------------------------
Decisions as Columns DMN rules tables have rows of inputs and rows of outputs,
with a double line horizontal border delineating where input rows stop and output rows begin.

+ The headings on the left hand side of the decision table, for the input and output rows, must be Variables from the Glossary.
+ The input test cells across the row must contain expressions about the input Variable,
  which will evalutate to True or False, depending upon the value of the input Variable.
+ The values in the output cells, across the row, are the values that will be assigned to the output Variable, if all the input cells evaluate to True in the same column.  If the output heading is the word 'Execute', then the output value must be the name of a Decision Table, which will be invoked immediately. Which means any output values already assigned can be overwritten by the invoked Decision Table, and any outputs created by the invoked Decision Table can be overwritten by subsequent outputs in the current Decision Table.

[ExampleColumns.xlsx](https://github.com/russellmcdonell/pyDMNrules/blob/master/ExampleColumns.xlsx) is an example fo a Decision as Columns DMN table.
### Decision as Crosstab rules tables
Decision as Crosstab DMN rules tables have one output heading at the top left of the table.

+ The horizontal and vertical headings, which must be Variables from the Glossary, are the input headings.
+ Under the horizontal headings and besides the vertical headings are the input test cells.
+ The output results cells form the body of the table and are assigned to the output Variable when the input cells in the same row and same column, all evaluate to True. If the output Variable is the word 'Execute', then the output value must be the name of a Decision Table, which will be invoked as part of completing the decision.

[ExampleCrosstab.xlsx](https://github.com/russellmcdonell/pyDMNrules/blob/master/ExampleCrosstab.xlsx) is an example fo a Decision as Crosstab DMN table.

The function getSheets() returns all the DMN rules tables.
Each DMN rules table is returned as an XHTML rendering, with all Variables converted to their FEEL name equivalents (BusinessConcept.Attribute).

Input cells, input 'Variable' and Input Tests
=============================================
An input cell in a DMN rules table is only part of the Input test. The other part is the input 'Variable'.
pyDMNrules has rules about how those two get combined in order to create an input test.

Simple Values
-------------
If the input cell contains a simple value then pyDMNrules will create an equality test; input "'Variable'" "equals" "input cell".
If the input cell contains a relational operator (<,>,!=,<=,>=) followed by a simple value, then pyDMNrules will use that relational operator.

List, 'not' and 'in'
--------------------
If an input cell contains only expressions separated by commas,
then pyDMNrules will interpret this as a list of unary tests and will create an "in function" test; "'Variable'" "in(" "input cell" ")".
This becomes a sequence of tests which returns True if any one of the tests are True. Each expression is treated as a 'Simple Value' - see above.

If you preface this simple list with 'not' then pyDMNrules will create a "not in function" test; "'Variable'" "not(in(" "input cell" "))"

However, sometime **the 'Variable' is the list**, and the test is whether or not the expression in the input cell is in the 'Variable' list.
To specify this form of test, the input cell should be a 'Simple Value' without an operator followed by ' in'. pyDMNrules will then create a "reverse in " test;
"input cell" "in " "'Variable'". Again, you can suffix the 'Simple Value' with " not in" and pyDMNrules will create "reverse not in " test;
"input cell" "not(in " "'Variable'" " )".    
NOTE: This is using the "in" operator, not the "in()" function.

If the input cell is a FEEL list or range, then pyDMNrules will interpret this as a test of "Variable" " in " list or range.
Again, you can suffix the list or range with " not in" and pyDMNrules will create "reverse not in " test;
"input cell" "not(in " "list or range" " )".

Input Tests and the Build-in Functions
++++++++++++++++++++++++++++++++++++++
Some FEEL functions [not(), odd(), even(), all() and any()] only ever take a single parameter and return 'true' or 'false'.
return 'true' or 'false' and hence you don't need to specify the parameter - is is the 'Variable'.
pyDMNrules will interpret "odd()" as "odd('Variable')".
[In output cells, all parameters for all functions, need to be fully specified.
To fully specify a 'Variable' from the Glossary use the internal name for the 'Variable',
being the 'Business Concept' and the 'Attibute' concateneted with the period character]

Some FEEL functions [starts with(), ends with() and list contains()] only ever take a two parameter and return 'true' or 'false'.
If you only specify one parameter then pyDMNrules will assume that 'Variable' is the first parameter and the parameter supplied
is the second parameter.

The matches() FEEL function can have either two or three parameters; a string and match pattern plus an optional regular expression flags parameter.
If you only supply one parameters then pyDMNrules will assume that the supplied parameter is the match pattern and will return mathes('Variable', parameter).
If you supply two parameters then pyDMNrules will assume that you have provided the match pattern and the optional flags parameter and will return
matches('Variable', parameter1, parameter2).  
Note: in an input cell matches("i", "i") will only return 'true if the input variable is 'i' or 'I'.

If that is not the test you want, and for all other functions, you will need to specify the test,
fully specify the 'Variable' and enclose any ambigous strings [those with spaces or commas] in double quotes ("").
[lower case(Python.output) = "hello world"]

Complex Input Tests
+++++++++++++++++++
pyDMNrules can get it wrong, because of the need to combine the input cell and the 'Variable' in various different ways,
especially if you have a very complex test. A simple, and often clearer approach is to use another DMN rules table
to compute the complex expression as an output; a DMNrules table with no inputs, only output(s) and only one rule.
Place this single row DMN rules table in the Decision before this complex test.
This single row DMN rules table will update an item in the Glossary.
Your complex input test DMN rules test becomes a simple test of that updated Glossary item.

Output cells and the Glossary
+++++++++++++++++++++++++++++
Output cells can contain constants (numbers, codes, string) or they can reflect the current value of one or more things in Glossary.
If an output value should be the value of an input, then you can use the 'Variable' from the Glossary.
However, if the output is a manuipulation of a Glossary value, then you will need to use the internal name for that variable,
being the 'Business Concept' and the 'Attibute' concateneted with the period character.
That is, 'Patient Age' is valid and will return the current value of 'Patient.age' from the Glossary, but 'Patient Age + 5' is not.
Instead you will need to use the syntax 'Patient.age + 5'

The 'Execute' Output Variable
+++++++++++++++++++++++++++++
The output variable named 'Execute' is reserved - it has a special meaning in pyDMNrules. A DMNrules table can have one or more output columns/rows
with the output variable name of 'Execute'. A crosstab DMNrules table can have 'Execute' as the name of it's output variable. Any cell associated with
an 'Execute' output variable must be the name of a DMNrules table. When a rule in a DMNrules table is triggered, and output values are being assigned,
no value will be assigned to any 'Execute' output variable.
Instead, when it it time to assign a value, the named DMNrules table will be invoked/run/executed. In effect, all the outputs from the name DMNrules table become
the output(s) for this output cell. Thus, you can have rules, within rules, within rules etc. A DMNrules table can even invoke/run/execute itself, although there
is a recursion limit of 100. However, you can have two identical tables, with different names, that call each other. Thus increasing the recursion limit to 200.
The recursion limit is not meant as an inhibitor, but rather as a safety factor to prevent runaway rules.

Strings and Numbers
+++++++++++++++++++
DMN uses a loosely typed language (S-FEEL). Python is a loosely typed language making Python a good fit for DMN.
pyDMNrules will attempt to interpret an Excel cell containing a string as a number, if it can,
but will revert back to treating it as a string if it cannot. Generally speaking this means that you can create your Excel worksheets
without having to be too careful. However some coding systems use numbers as codes, such as 'M-Male/F-Female/9-Unknown'.
When using code sets that have numbers as codes, or codes that start with a number, alway pass them a string to pyDMNrules
and always enter them with enclosing double quotes in your Excel worksheet 'M-Male/F-Female/"9"-Unknown'.    

Similarly, code that begin with 'P' could be misinterpreted as 'durations' (P42D is a duration of '42 days' in FEEL).
Again, these codes should always enclosed in double quotes in your Excel worksheet.

Usage
-----

::

    import pyDMNrules
    dmnRules = pyDMNrules.DMN()
    status = dmnRules.load('Example1.xlsx')
    if 'errors' in status:
        print('Example1.xlsx has errors', status['errors'])
        sys.exit(0)
    else:
        print('Example1.xlsx loaded')

    data = {}
    data['Applicant Age'] = 63
    data['Medical History'] = 'bad'
    (status, newData) = dmnRules.decide(data)
    if 'errors' in status:
        print('Failed')

Examples
--------
Examples (\*.py, \*.xlsx) can be found at [github](https://github.com/russellmcdonell/pyDMNrules)


Usage - decidePandas()
----------------------

::

   import pyDMNrules
   import pandas as pd
   import sys
   dmnRules = pyDMNrules.DMN()
   status = dmnRules.load('AN-SNAP rules (DMN).xlsx')
   if 'errors' in status:
      print('AN-SNAP rules (DMN).xlsx has errors', status['errors'])
      sys.exit(0)
   else:
        print('AN-SNAP rules (DMN).xlsx loaded')
   dataTypes = {'Patient_Age':int, 'Episode_Length_of_stay':int, 'Phase_Length_of_stay':int,
      'Phase_FIM_Motor_Score':int, 'Phase_FIM_Cognition_Score':int, 'Phase_RUG_ADL_Score':int,
      'Phase_w01':float, 'Phase_NWAU21':float,
      'Indigenous_Status':str, 'Care_Type':str, 'Funding_Source':str, 'Phase_Impairment_Code':str,
      'AROC_code':str, 'Phase_AN_SNAP_V4_0_code':str,
      'Same_day_addmitted_care':bool, 'Delerium_or_Dimentia':bool, 'RadioTherapy_Flag':bool, 'Dialysis_Flag':bool}
   dates = ['BirthDate', 'Admission_Date', 'Discharge_Date', 'Phase_Start_Date', 'Phase_End_Date']
   dfInput = pd.read_csv('subAcuteExtract.csv', dtype=dataTypes, parse_dates=dates)
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
   if dfStatus.where(dfStatus != 'no errors').count() > 0:
      print('has errors', dfStatus.loc['status' != 'no errors'])
      sys.exit(0)
   for index in dfResults.index:
      print(index, dfResults.loc[index, 'AN_SNAP_V4_code'])


The test() function
-------------------
pyDMNrules reserves the spreadsheet name 'Test' which can be used for storing test data.
The function test() reads the test data, passes it through the decide() function and assembles the results which are returned to the caller.

Unit Test data table(s)
+++++++++++++++++++++++
The 'Test' spreadsheet must contain Unit Test data tables.

+ Each Unit Test data table must be named with a 'Business Concept' from the Glossary. 
+ The heading of the table must be Variables from the Glossary associated with the Business Concept.
+ The data, in each column, must be valid input data for the associated Variable.
+ Each row is considered a unit test data set.
+ Unit Test data tables can have 'Annotations'


DMNrulesTest - the test that will be run
++++++++++++++++++++++++++++++++++++++++
pyDMNrules also searches the 'Test' worksheet for a special table named 'DMNrulesTests' which must exist.

+ The 'DMNrulesTests' table has input columns followed by output columns
  with a double line vertical border delineating where input columns stop and output columns begin.
+ The headings for the input columns must be Business Concepts from the Glossary.
+ The data in input columns must be one based indexes into the matching unit test data table.
+ The headings for the output columns must be Variables from the Glossary.
+ The data in output columns will be valid data for the matching Variable and are the expected values returned by the decide() function.
+ The 'DMNrulesTests' table can have 'Annotations'

The test() function will process each row in the 'DMNrulesTests' table,
getting data from the Unit Test data tables and building a data{} dictionary.
The test() function then passes this data{} dictionary to the decide() function.
The returned newData{} is then compared to the output data for the matching test in the 'DMNrulesTests' table.
The data{}, newData{} and a list of mismatches, if any, is then appended to the list of results,
which is eventually passed back to the caller of the test() function.

The returned list of results is a list of dictionaries, one for each test in the 'DMNrulesTests' table. The keys to this dictionary are

+ 'Test ID' - the one based index into the 'DMNrulesTests' table which identifies which test which was run
+ 'TestAnnotations'(optional) - the list of annotation for this test - not present if no annotations were present in 'DMNrulesTests'
+ 'data' - the dictionary of assembed data passed to the decide() function
+ 'newData' - the decision dictionary returned by the decide() function [see above] 'DataAnnotations'(optional) - the list of annotations from the unit test data tables,
  for the unit test sets used in this test - not present if no annotations were persent in any of the unit test data tables
  for the selected unit test sets.
+ 'Mismatches'(optional) - the list of mismatch reports,
  one for each 'DMNrulesTests' table output value that did not match the value
  returned from the decide() function - not present if all the data returned from the decide() function
  matched the values in the 'DMNrulesTest' table.


Usage - test()
--------------

::

    import pyDMNrules
    dmnRules = pyDMNrules.DMN()
    status = dmnRules.load('ExampleHPV.xlsx')
    if 'errors' in status:
        print('ExampleHPV.xlsx has errors', status['errors'])
        sys.exit(0)
    else:
        print('ExampleHPV.xlsx loaded')
    (testStatus, results) = dmnRules.test()
    for test in range(len(results)):
        if 'Mismatches' not in results[test]:
            print('Test ID', results[test]['Test ID'], 'passed')
        else:
            print('Test ID', results[test]['Test ID'], 'failed')
            for failure in range(len(results[test]['Mismatches'])):
                print(results[test]['Mismatches'][failure])
        if 'errors' in testStatus[test]:
            print('Failed')


Examples
--------
Examples (\*.py, \*.xlsx) can be found at [github](https://github.com/russellmcdonell/pyDMNrules)


Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
