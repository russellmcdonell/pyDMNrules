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


