### 1.4.3 - Minor bug fix and features release - released to PyPI
 - Added support for the contains() function - thanks Leonard Budney (https://github.com/budney)
 - Fixed bug in getTableGlossary() - output rows/columns weren't being output
 - Fixed bug in replaceItems() - incorrect replacement for . operators
 - Fixed bugs when a Decision Table was embedded in a Collection type Decision Table
 - Made test() return all the Results, Decisions and Excuted Rules - thanks Leonard Budney (https://github.com/budney)
### 1.4.2 - Added support for contains() function - to be released to PyPI
 - Added support for contains() function - thank you budney (https://github.com/budney)
 - Fixed typo in getTableGlossary()
 - Fixed .hour, .minute and .second - now returns float()
 ### 1.4.1 - Bug fix release - released to PyPI
 - Fixed bug triggered by initalizing output values in a multi-decision table decision
 - Added getTableGlossary() for DecisionCentral
### 1.4.0 - BREAKING Bug fix release - released to PyPI
 - pySFeel changed parsing so that ranges with closed intervals are returned with round brackets (reverse facing square brackets are allowed on input, but will be converted to their equivalent round bracket)
   **NOTE BREAKING CHANGE:** this may effect some existing of tests
 - pySFeel fixed a bug in the equality test - now returns 'null' (Python value None) if values are of different data types
   **NOTE BREAKING CHANGE:** this may effected some existing tests
 - pySFeel fixed a bug in the string() function - now returns timezones for datetimes/time that have a timezone
   **NOTE BREAKING CHANGE:** this may effected some existing implementations
 - pySFeel fixed a bug in .weekday - now returns isoweekday() (now 1 - 7, was 0 - 6)
   **NOTE BREAKING CHANGE:** this may effected some existing implementations
 - Added decideTables() function for basing a decision on a specified list of Decision Tables.
 - pySFeel changed the any()/all() functions which now return False/True for empty lists.
 - Fixed bug when valid values (XML DMN) is a list of strings and there's a space after a comma
 - Fixed bug when input test is just not() function [argument must be a boolean expression]
 - Added suport for variable1 = variable2 being converted to variable1 in variable2 where variable2 is a list
 - Fixed bug when output value is a Context
 - pySFeel added support for 'instance of'
 - pySFeel added support fo single endpoint and operator ranges
 - pySFeel added support for named parameters in built-in functions
 - pySFeel added support for context scoped variables - {a:1,b:a+1} - 'a' is a valid variable, but only inside the context and only after it has been defined.
 - pySFeel added support for DMN 1.4 functions today() and now()
 - Tested with DMN-TCK [Test Conformance Kit]. Pass rate: 1569/1771 (89%) [97% at Conformance Level 2 - the basics]
### 1.3.20 - Bug fix for sort() - released to PyPI
 - function() variable were being wrapped in double quotes
### 1.3.19 -  - Added limited support for the sort() function - released to PyPI
 - Added limited support for the sort() function - only the anonymous form [ sort(expr, function(name0, name1) expression)) ]. And 'expression' is limited to 'name0 < name1' or 'name0 > name1'. If the list to be sorted ('expr) is a list of Contexts, the name0 and name1 must take the form of name0.attr and name1.attr, and 'attr' must be the same attribute for both name0 and name1.
 - Fixed bug in getSheets (adding defaultValues row/column to collection hit policies)
### 1.3.18 - WARNING - changed hitPolicy R and O - relased to PyPI
 - Hit Policy 'RULES ORDER' (R)' and 'OUTPUT ORDER' (O) now return each output variable as a list,
   as per the specification. Previously they return a list of 'Results' which was an error.
   HOWEVER, this required a rewrite of the TherapyNoglossNodecision.xlsx example.
 - All of the changes in this release are related to (limited) XML support and testing with DMN TCK (https://dmn-tck.github.io/tck/)
 - cleaned up XML support - added support for decisionTables in businessKnowledge
 - Removed Excel specific error messages
 - Allowed Business Concepts and Attributes in the Glossary to contain spaces (which get replaced with '_')
 - Added support for default output values (XML only)
 - Allowed Priority hit policy decision tables to have **no** valid values
 - Fixed filters of List Filters - e.g. EmployeeTable[name=LastName].deptNum[1]
 - 'item' now optional in List filters
 - Limited support for 'some/every in ... satifies expression'. 'expression' must be 'name relop expr'.
 - Tested with DMN TCK conformance tests [27/28 passed conformance Level 2]
### 1.3.17 - Added Beta XML support
 - Implemented automatic disambiguation of clashing, automatically derived, Attribute names.
 - Fixed output for COLLECTION, RULES ORDER and OUTPUT ORDER Hit Policies
### 1.3.16 - Added colours to Input and Output heading in getSheets() - realsed to PyPI
 - Centre aligned input tests and output value in getSheets()
 - Fixed error return when no rules loaded in getGlossary(), getDecision(), getSheets()
 - Preserved annotations in the Glossary
 - Added getGlossaryNames()
### 1.3.15 - Fixed bug with 'Execute' output columns - released to PyPI
 - Added validation for 'possible' tables - right structure, wrong headings
### 1.3.14 - Removed stray import of tkinter - released to PyPI
 - remove stray import of tkinter
 ### 1.3.13 - Fixed another bug - getSheets() - released to PyPI
* Error in headings with validity and annotation
### 1.3.12 - Fixed bug in getSheetsn() - release to PyPI
* getSheets() was not returning the first sheet
### 1.3.11 - Fixed bug in getDecision() - release to PyPI
* getDecision() was not returning all the headers
### 1.3.10 - Update to match pySFeel 1.3.7
* Removed arithmetic operators from valid FEEL names
### 1.3.9 - Bug fix (decidePandas) - released to PyPI
* decidePandas() was not honoring the 'strict' option
### 1.3.8 - Update decidePandas - released to PyPI
* Rewrite code to work around 'depricated' warnings
### 1.3.7 - Bug fix - released to PyPI
* Error when bad S-FEEL was being reported
### 1.3.6 Added getDecision() and getSheets() functions
* Fixed bug -  when constructing Glossary (none supplied) - output variables weren't being added to the Glossary
### 1.3.5 Made Decision optional
* Added getGlossary() function to aid web developers
### 1.3.4 Made Glossary optional - Added rules within rules functionality
* Changed test for Rules as Rows/Rules as Columns/CrossTab Rules (removed dependency on Glossary)
* Added glossary validation for output and inputs for CrossTab Rules (must be in provided glossary or valid for assembled glossary)
* Reserved the output Variable name 'Execute'.
    - Output cells for the 'Execute' output Variable must contain the name of a decision table
    - The named decision table will invoked/run/executed when the associated rule is triggered
    - Execution of the named decision table will occur when, normally, a value would have been assigned to the 'Execute' output variable
    - Outputs from the named decision table may overwrite outputs already assigned as part of the current rule
    - Outputs for the current rule, assigned after the named decision table has been invoked/run/execute may overwrite outputs from the named decsion table
    - A decision table can invode/run/execute itself, but there is a recursion limit of 100 - to prevent runaway rules
### 1.3.1 Added DMN 1.3 support
* Added support for the new DMN 1.3 built-in functions
    - is() and the Range comparison functions
* Added support for passing ranges as input values
    - ranges are passed as tuples of (end0, low0, high1, end1)    
    where end0 must be one of '(', '[' or ']' and end1 must be one of ')', ']', '['    
    and low0 and high1 must be the same data type, and a valid data type for conversion to FEEL
* Added support for passing string litterals (@"xxx") as input values
* pyDMNrules also inherits from pySFeel and there have been a number of enhancements to pySFeel
    - Support for literal strings (@"xxx") for dates/times/date-times/durations
    - Support for @region/location timezones for times and date-times
    - Support for attributes for date/times/date-times/durations/ranges    
    @"2021-12-2".year returns 2021
    - Plus some bug fixes
### 1.2.8 - Bug fixed
* fixed error message when failing input validity
* fixed error when a Pandas dataframe has invalid column names
* added 'strict' option to decidePandas() for forcing the return of only valid Pandas column names
### 1.2.7 - Bug fix
* fixed bug in fixed value comparisons of booleans
### 1.2.6 - Bug fix and optimization
* fixed datetime.date/time/datetime and datetime.timedelta data type errors in decidePandas() function
* fixed bug when assigning Glossary values to invalid FEEL (e.g. "string1" "string2" - typos in Variable names can create invalid FEEL)
* added optimization for testing input values against fixed values and outputing fixed output values

### 1.2.5 - Updated dependencies and documentation
### 1.2.4 - Added Pandas functionality
 * decidePandas(dfInput) passes each row of DataFrame dfInput through the decide() function.
    - Returns (dfStatus, dfResults, dfDecision)
        - dfStatus is a Series of 'no errors' or the 'errors' from the decision
        - dfResults is a DataFrame with one row for each row of dfInput, being 'Results' of the final decision
        - dfDecision is a DataFrame describing the decision for each row of dfInput (RuleName, TableName, RuleID, DecisionAnnotations, RuleAnnotations)

### 1.2.3 - Another bug fix release
 * fixed/improved substitution of BusinessConcept.Attribute for Variable in input and output cells when Decision tables are parsed
 * fixed substitution of actual value for BusinessConcept.Attribute when inputs are tested and output created in decide()
 * fixed bug in parsing Excel booleans (TRUE/FALSE)
### 1.2.2 - Fixed bug when variable referenced in a test/output and the variable and it's Business Concept shared the same name
### 1.2.1 - Fixed bug in in() function
### 1.2.0 - First candidate for a production release
* added lots of pytest tests
* fixed bug preventing the re-use of a decision in the list of decision tables
* fixed bug when assigning rule id in rules as columns decision tables
* added check that decision tables named on the Decision worksheet exist in the workbook
* rewrote test2sfeel
    - Added support for FEEL functions that return a boolean and have the input variable as a parameter.  
    They can now be specified without the input variable as a parameter.  
    e.g. "starts with("abc") will test if the input variable starts with 'abc'.    
    Supported functions are
        - not()
        - odd()
        - even()
        - all()
        - any()
        - starts with(string)
        - ends with(string)
        - list contains(value)
        - matches('string')
        - matches('string', 'flags')      
    Also supported is not(suportedFunction()), but definitely not not(not())
* Fixed validity checking
* Added support for Excel boolean cells (TRUE, FALSE) so long as they are the only value in a cell
* Added loadTest() function for loading the Test worksheet from a different workbook
* Added the use() function which take an openpyxl workbook object. Useful if you want to read other data from the workbook before using the rules
* Added the useTest() function which takes an openpyxl workbook object that contains the 'Test' worksheet. Useful if you want to run the tests and record the results in a separate workbook.
* Fixed checking of inputs and outputs against valid values

### 0.1.10 - the first release


