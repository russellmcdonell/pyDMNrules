
### 0.1.10 - the first release

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


