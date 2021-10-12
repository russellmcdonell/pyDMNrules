.. pyDMNrules documentation master file, created by
   sphinx-quickstart on Fri Jan 24 05:25:44 2020.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

pyDMNrules Documentation
========================


.. toctree::
   :maxdepth: 2
   :caption: Contents:

the pyDMNrules functions
------------------------

.. py:module:: pyDMNrules

.. py:class:: DMN

   .. automethod:: load

.. py:class:: DMN

   .. automethod:: use

.. py:class:: DMN

   .. automethod:: decide

.. py:class:: DMN

   .. automethod:: loadTest

.. py:class:: DMN

   .. automethod:: useTest

.. py:class:: DMN

   .. automethod:: test



Input cells, input 'Variable' and Input Tests
---------------------------------------------
An input cell in a DMN rules table is only part of the Input test. The other part is the input 'Variable'.
pyDMNrules has rules about how those two get combined in order to create an input test.

Simple Values
+++++++++++++
If the input cell contains a simple value then pyDMNrules will create an equality test; input "'Variable'" "equals" "input cell".
If the input cell contains a relational operator (<,>,!=,<=,>=) followed by a simple value, then pyDMNrules will use that relational operator.

List, 'not' and 'in'
++++++++++++++++++++
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

Strings and Numbers
+++++++++++++++++++
DMN uses a loosely typed language (S-FEEL). Python is a loosely typed language making Python a good fit for DMN.
pyDMNrules will attempt to interpret an Excel cell containing a string as a number, if it can,
but will revert back to treating it as a string if it cannot. Generally speaking this means that you can type up you Excel worksheets
without having to be too careful. However some coding systems use numbers as codes, such as 'M-Male/F-Female/9-Unknown'.
When using code sets that have numbers as codes, alway pass them a string to pyDMNrule
and always enter them with enclosing double quotes in our Excel worksheet 'M-Male/F-Female/"9"-Unknown'.


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


.. py:class:: DMN

   .. automethod:: test

The test() function
-------------------
pyDMNrules reserves the spreadsheet name 'Test' which can be used for storing test data.
The function test() reads the test data, passes it through the decide() function and assembles the results which are returned to the caller.

Unit Test data table(s)
+++++++++++++++++++++++
The 'Test' spreadsheet must contain Unit Test data tables.

- Each Unit Test data table must be named with a 'Business Concept' from the Glossary.
  
- The heading of the table must be Variables from the Glossary associated with the Business Concept.

- The data, in each column, must be valid input data for the associated Variable.

- Each row is considered a unit test data set.

- Unit Test data tables can have 'Annotations'

DMNrulesTest - the test that will be run
++++++++++++++++++++++++++++++++++++++++
pyDMNrules also searches the 'Test' worksheet for a special table named 'DMNrulesTests' which must exist.

- The 'DMNrulesTests' table has input columns followed by output columns
  with a double line vertical border delineating where input columns stop and output columns begin.

- The headings for the input columns must be Business Concepts from the Glossary.

- The data in input columns must be one based indexes into the matching unit test data table.

- The headings for the output columns must be Variables from the Glossary.

- The data in output columns will be valid data for the matching Variable and are the expected values returned by the decide() function.

- The 'DMNrulesTests' table can have 'Annotations'

The test() function will process each row in the 'DMNrulesTests' table,
getting data from the Unit Test data tables and building a data{} dictionary.
The test() function then passes this data{} dictionary to the decide() function.
The returned newData{} is then compared to the output data for the matching test in the 'DMNrulesTests' table.
The data{}, newData{} and a list of mismatches, if any, is then appended to the list of results,
which is eventually passed back to the caller of the test() function.

The returned list of results is a list of dictionaries, one for each test in the 'DMNrulesTests' table. The keys to this dictionary are

- 'Test ID' - the one based index into the 'DMNrulesTests' table which identifies which test which was run

- 'TestAnnotations'(optional) - the list of annotation for this test - not present if no annotations were present in 'DMNrulesTests'

- 'data' - the dictionary of assembed data passed to the decide() function

- 'newData' - the decision dictionary returned by the decide() function [see above]

- 'DataAnnotations'(optional) - the list of annotations from the unit test data tables,
  for the unit test sets used in this test - not present if no annotations were persent in any of the unit test data tables
  for the selected unit test sets.

- 'Mismatches'(optional) - the list of mismatch reports,
  one for each 'DMNrulesTests' table output value that did not match the value
  returned from the decide() function - not present if all the data returned from the decide() function
  matched the values in the 'DMNrulesTest' table.

Usage
-----

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
