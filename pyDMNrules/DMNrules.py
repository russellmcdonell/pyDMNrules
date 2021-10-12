# -----------------------------------------------------------------------------
# pyDMNrules.py
# -----------------------------------------------------------------------------

import sys
import re
import csv
import datetime
import pySFeel
import openpyxl
from openpyxl import load_workbook
from openpyxl import utils

class DMN():


    def __init__(self):
        self.lexer = pySFeel.SFeelLexer()
        self.parser = pySFeel.SFeelParser()
        # self.glossary is a dictionary of dictionaries (one per variable).
        # self.glossary[variable]['item'] is BusinessConcept.Attribute.
        # self.glossary[variable]['concept'] is Business Concept
        self.glossary = {}
        self.glossaryItems = {}         # a dictonary - self.glossaryItems[BusinessConcept.Attribute] = Variable
        self.glossaryConcepts = {}      # a dictionary of BusinessConcepts
        self.glossaryLoaded = False
        self.isLoaded = False
        self.testIsLoaded = False
        self.errors = []
        self.warnings = []


    def sfeel(self, text):
        (status, returnVal) = self.parser.sFeelParse(text)
        if 'errors' in status:
            self.errors += status['errors']
        return returnVal


    def data2sfeel(self, coordinate, sheet, data, isTest):
        # Check that a string of text (data) is valid S-FEEL
        # Start by replacing all 'Variable's with their BusinessConcept.Attribute equivalents (which are valid S-FEEL)
        # Being careful not to replace any BusinessConcept.Attributes that already exist in data
        text = data
        for variable in self.glossary:
            item = self.glossary[variable]['item']
            if variable == self.glossary[variable]['concept']:       # The variable and it's business concept share the same name
                at = 0
                match = re.search(variable, text[at:])
                while match is not None:
                    replaceIt = True
                    if match.end() == len(text[at:]):
                        pass
                    elif text[at + match.end():at + match.end() + 1] != '.':
                        pass
                    else:
                        for otherItem in self.glossaryItems:
                            if text[at:].startswith(otherItem):
                                replaceIt = False
                                break
                    if replaceIt:
                        text = text[:at] + text[at:].replace(variable, item, 1)
                    at += match.end()
                    match = re.search(variable, text[at:])
            else:
                text = text.replace(variable, item)

        # Use the pySFeel tokenizer to look for strings that look like 'names', but aren't in the glossary
        isError = False
        tokens = self.lexer.tokenize(text)
        yaccTokens = []
        for token in tokens:
            if token.type == 'ERROR':
                if isTest:
                    return None
                if not isError:
                    self.errors.append("S-FEEL syntax error in text '{!s}' at '{!s}' on sheet '{!s}':{!s}".format(data, coordinate, sheet, token.value))
                    isError = True
            else:
                yaccTokens.append(token)
        thisData = ''
        # Step through the tokens
        for token in yaccTokens:
            if thisData != '':          # Empty tokens are white space
                thisData += ' '
            if token.type != 'NAME':    # If it doesn't look like a name then leave it alone
                thisData += token.value
            elif token.value in self.glossaryItems:     # If it's a fully qualified name (BusinessConcept.Attribute) then leave it alone
                thisData += token.value
            else:                           # Otherwise, assume it's a string that is missing it's double quotes
                thisData += '"' + token.value + '"'
        return thisData


    def test2sfeel(self, variable, coordinate, sheet, test):
        '''
    Combine the contents of an Excel cell (test) which is a string that can be combined
    with the Glossary variable ([Business Concept.Attribute]) to create a FEEL logical expression
    which pySFeel will evaluate to True or False    
        '''
        thisTest = str(test).strip()
        # Check for bad S-FEEL
        if len(thisTest) == 0:
            self.errors.append("Bad S-FEEL '{!r}' at '{!s}' on sheet '{!s}'".format(test, coordinate, sheet))
            return 'null'

        # Check for a comma separated list of strings - however, is only a list of tests if it is not an FEEL list
        listOfTests = []
        try:
            for row in csv.reader([thisTest], dialect=csv.excel, doublequote=False, escapechar='\\'):
                listOfTests = list(row)
        except:
            pass

        # Check for valid FEEL string
        wasString = False
        if (thisTest[0] == '"') and (thisTest[-1] == '"'):
            thisTest = thisTest[1:-1]
            wasString = True

        # Check to see if the 'Variable' is specified in the test.
        # If so, 'test' will have to be a fully specified S-FEEL expression.
        if thisTest.find(variable) != -1:
            return self.data2sfeel(coordinate, sheet, thisTest, False)

        # Check for the known one parameter FEEL functions that return True or False
        # If they are specified without any parameter, then assume variable is the parameter
        if thisTest == 'not()':         # if variable is Boolean, then negate it
            return self.data2sfeel(coordinate, sheet, 'not(' + variable + ')', False)

        # Check for the not() function which will reverse the test - we don't allow not(not())
        testIsNot = False
        if not wasString and thisTest.startswith('not(') and thisTest.endswith(')'):
            testIsNot = True
            thisTest = thisTest[4:-1].strip()

        if not wasString and thisTest == 'odd()':         # if variable is a number, then check that it is odd
            if testIsNot:
                return self.data2sfeel(coordinate, sheet, 'not(odd(' + variable + '))', False)
            else:
                return self.data2sfeel(coordinate, sheet, 'odd(' + variable + ')', False)
        if not wasString and thisTest == 'even()':        # if variable is a number, then check that it is ever
            if testIsNot:
                return self.data2sfeel(coordinate, sheet, 'not(even(' + variable + '))')
            else:
                return self.data2sfeel(coordinate, sheet, 'even(' + variable + ')', False)
        if not wasString and thisTest == 'all()':         # if variable is a list of Booleans, then check that all of them are True
            if testIsNot:
                return self.data2sfeel(coordinate, sheet, variable + 'not(all(' + variable + '))', False)
            else:
                return self.data2sfeel(coordinate, sheet, variable + 'all(' + variable + ')', False)
        if not wasString and thisTest == 'any()':         # if variable is a list of Booleans, then check if any of them are True
            if testIsNot:
                return self.data2sfeel(coordinate, sheet, variable + 'not(any(' + variable + '))', False)
            else:
                return self.data2sfeel(coordinate, sheet, variable + 'any(' + variable + ')', False)

        # Check for the known two parameter FEEL functions that return True or False
        match = re.match(r'^starts with\((.*)\)$', thisTest)
        if not wasString and match is not None:           # if variable is a string, then check that it starts with this string
            withString = match.group(1)
            if withString[0] != '"':    # make sure the second arguement is a string
                withString = '"' + withString
            if withString[-1] != '"':
                withString += '"'
            if testIsNot:
                return self.data2sfeel(coordinate, sheet, 'not(starts with(' + variable + ', ' + withString + '))', False)
            else:
                return self.data2sfeel(coordinate, sheet, 'starts with(' + variable + ', ' + withString + ')', False)
        match = re.match(r'^ends with\((.*)\)$', thisTest)
        if not wasString and match is not None:           # if variable is a string, then check that it ends with this string
            withString = match.group(1)
            if withString[0] != '"':    # make sure the second arguement is a string
                withString = '"' + withString
            if withString[-1] != '"':
                withString += '"'
            if testIsNot:
                return self.data2sfeel(coordinate, sheet, 'not(ends with(' + variable + ', ' + withString + '))', False)
            else:
                return self.data2sfeel(coordinate, sheet, 'ends with(' + variable + ', ' + withString + ')', False)
        match = re.match(r'^list contains\((.*)\)$', thisTest)
        if not wasString and match is not None:           # if variable is a list, then check that it contains this element
            return self.data2sfeel(coordinate, sheet, 'list contains(' + variable + ', ' + match.group(1) + ')', False)
        # And then the slightly more complex 'matches'
        match = re.match(r'^matches\((.*)\)$', thisTest)
        if not wasString and match is not None:
            # There can be one or two arguments
            try:
                for row in csv.reader([match.group(1)], dialect=csv.excel, doublequote=False, escapechar='\\'):
                    parameters = list(row)
            except:
                parameters = []
            if (len(parameters) == 1) or (len(parameters) == 2):
                if parameters[0][0] != '"':
                    parameters[0] = '"' + parameters[0]
                if parameters[0][-1] != '"':
                    parameters[0] += '"'
                if len(parameters) == 1:
                    return self.data2sfeel(coordinate, sheet, 'matches(' + variable + ', ' + parameters[0] + ')', False)
                if parameters[1][0] != '"':
                    parameters[1] = '"' + parameters[0]
                if parameters[1][-1] != '"':
                    parameters[1] += '"'
                if testIsNot:
                    return self.data2sfeel(coordinate, sheet, 'not(matches(' + variable + ', ' + parameters[0] + ', ' + parameters[1] + '))', False)
                else:
                    return self.data2sfeel(coordinate, sheet, 'matches(' + variable + ', ' + parameters[0] + ', ' + parameters[1] + ')', False)

        # Not a simple function.
        # So now we are trying to create an expression of 'variable' 'operator' 'test'
        # Start by checking if the operator is already supplied at the start of the 'test'
        if not wasString and ((thisTest[:2] == '!=') or (thisTest[:1] in ['<', '>', '='])):
            # Check for a comma separated list of operator/value expressions - these should use the in() function and become a list of tests
            if len(listOfTests) == 1:     # Not a comma separated list
                if testIsNot:
                    return self.data2sfeel(coordinate, sheet, 'not(' + variable + ' ' + thisTest + ')', False)
                else:
                    return self.data2sfeel(coordinate, sheet, variable + ' ' + thisTest, False)
        
        # Check for not()/not and in()/in
        # We can have 'variable' '[not] in' 'test'
        # or 'test' '[not] in' 'variable'
        variableIsNot = False
        if testIsNot:
            if thisTest.startswith('not '):
                self.errors.append("Bad S-FEEL '{!r}' at '{!s}' on sheet '{!s}'".format(test, coordinate, sheet))
                return 'null'
            variableIsNot = True
        elif not wasString and thisTest.startswith('not '):
            variableIsNot = True
            thisTest = thisTest[4:].strip()
        variableIsIn1 = False
        variableIsIn2 = False
        match = re.match(r'^in\s*\((.*)\)$', thisTest)
        if not wasString and match is not None:
            thisTest = match.group(1)
            variableIsIn2 = True
        elif not wasString and thisTest.startswith('in '):
            variableIsIn1 = True
            thisTest = thisTest[3:].strip()
        testIsIn = False
        if not wasString and thisTest.endswith(' in'):
            testIsIn = True
            thisTest = thisTest[:-3].strip()
        testIsNegated = False
        if not wasString and thisTest.endswith(' not'):
            testIsNegated = True
            thisTest = thisTest[:-4].strip()

        # Check for bad S-FEEL
        if len(thisTest) == 0:
            self.errors.append("Bad S-FEEL '{!r}' at '{!s}' on sheet '{!s}'".format(test, coordinate, sheet))
            return 'null'

        # Check for a list or a range, with 'in' either specified or implied
        if not wasString and (thisTest[0]  in ['[', '(']) and (thisTest[-1]  in [']', ')']):
            # Start putting the nots and ins back in - if they are valid
            if testIsIn:        # This is 'list/range' in variable - which is not supported
                self.errors.append("Bad S-FEEL '{!r}' at '{!s}' on sheet '{!s}'".format(test, coordinate, sheet))
                return 'null'
            if testIsNegated:       # variable in 'list/range' not - is invalid
                self.errors.append("Bad S-FEEL '{!r}' at '{!s}' on sheet '{!s}'".format(test, coordinate, sheet))
                return 'null'
            if variableIsIn2:       # variable in('list/range') - is invalid
                self.errors.append("Bad S-FEEL '{!r}' at '{!s}' on sheet '{!s}'".format(test, coordinate, sheet))
                return 'null'
            if variableIsNot:               # Not specified - in either specified or implied
                return self.data2sfeel(coordinate, sheet, variable + ' not(in ' + thisTest + ')', False)
            else:                   # in either specified or implied
                return self.data2sfeel(coordinate, sheet, variable + ' in ' + thisTest, False)

        # Check for a comma separated list - these should use the in() function and become a list of tests
        if len(listOfTests) > 1:
            origTest = thisTest
            if wasString:
                origTest = '"' + origTest + '"'
            theseTests = []
            for thisTest in listOfTests:
                aTest = thisTest.strip()
                # Any wrapping double quotes will have been removed, but any escaped double quotes will have been replace with just a non-escaped double quote
                if origTest[0] == '"':              # A string which may have been stripped of whitespace
                    bTest = aTest.replace('"', '\\"')
                    # We need to do some suffling and aligning to find out which test were wrapped in double quotes
                    testAt = origTest.find(bTest)                           # Must exist
                    quoteAt = origTest.find('"', testAt + len(aTest))       # Find the trailing double quote
                    theseTests.append(origTest[0:quoteAt + 1])
                    commaAt = origTest.find(',', quoteAt + 1)               # Find the next comma
                    if commaAt != -1:
                        origTest = origTest[commaAt + 1:].lstrip()
                    else:
                        origTest = ''
                else:                               # Not a 'string', but could be special value
                    if (aTest == 'true') or (aTest == 'false') or (aTest == 'null'):
                        theseTests.append(aTest)                    # Append a boolean
                    elif len(aTest) == 0:
                        theseTests.append(aTest)                    # Append an empty string
                    elif (aTest[:2] == '!=') or (aTest[:1] in ['<', '>', '=']):
                        theseTests.append(aTest)                    # Append an operator driven test
                    else:
                        try:
                            floatTest = float(aTest)
                            theseTests.append(aTest)                # Append a number
                        except:
                            theseTests.append('"' + aTest + '"')    # Append a code
                    testAt = origTest.find(aTest)                          # Must exist
                    commaAt = origTest.find(',', testAt + 1)               # Find the next comma
                    if commaAt != -1:
                        origTest = origTest[commaAt + 1:].lstrip()
                    else:
                        origTest = ''
            # Assemble the list of tests for the in() function
            newTests = []
            for i in range(len(theseTests)):
                newTests.append(self.data2sfeel(coordinate, sheet, theseTests[i], False))
            theseTests = newTests
            # Start putting the nots and ins back in - if they are valid
            if testIsIn:        # This is list of tests in variable - which is invalid S-FEEL
                self.errors.append("Bad S-FEEL '{!s}' at '{!s}' on sheet '{!s}'".format(test, coordinate, sheet))
                return 'null'
            if testIsNegated:       # variable in list of tests not in variable - which is invalid S-FEEL
                self.errors.append("Bad S-FEEL '{!s}' at '{!s}' on sheet '{!s}'".format(test, coordinate, sheet))
                return 'null'
            if variableIsIn1:       # variable in list of tests - is invalid - must be the in() function
                self.errors.append("Bad S-FEEL '{!s}' at '{!s}' on sheet '{!s}'".format(test, coordinate, sheet))
                return 'null'
            if variableIsNot:               # Not specified - either unary or function
                return variable + ' not( in(' + ','.join(theseTests) + '))'
            else:                   # in either specified or implied
                return variable + ' in(' + ','.join(theseTests) + ')'

        # If 'in' specified at the start of the 'test', the we will treat 'test' as a single entry list
        # as in 'in "abcd"'
        if variableIsIn1 or variableIsIn2:
            # Check we don't have jumbled up in's and not's
            if testIsIn:        # This is 'variable in test in variable' - which is not supported
                self.errors.append("Bad S-FEEL '{!s}' at '{!s}' on sheet '{!s}'".format(test, coordinate, sheet))
                return 'null'
            if testIsNegated:       # 'variable in test not' - is invalid
                self.errors.append("Bad S-FEEL '{!s}' at '{!s}' on sheet '{!s}'".format(test, coordinate, sheet))
                return 'null'
            if variableIsNot:               # 'not' - either specified or implied
                if variableIsIn1:
                    return self.data2sfeel(coordinate, sheet, variable + ' not(in ' + thisTest + ')', False)
                elif variableIsIn2:
                    return self.data2sfeel(coordinate, sheet, variable + ' not(in(' + thisTest + '))', False)
            else:
                if variableIsIn1:
                    return self.data2sfeel(coordinate, sheet, variable + ' in ' + thisTest, False)
                elif variableIsIn2:
                    return self.data2sfeel(coordinate, sheet, variable + ' in(' + thisTest + ')', False)

        # If 'in' or 'not in' specified at the end of the 'test', the we will treat 'test' as a single value
        # If it was a string, then it was passed as "value", otherwise it was passed as just value
        if testIsIn:
            if testIsNegated:               # 'not' was specified
                if wasString:
                    return self.data2sfeel(coordinate, sheet, '"' + thisTest + '"' + ' not(in ' + variable +')', False)
                else:
                    return self.data2sfeel(coordinate, sheet, thisTest + ' not(in ' + variable +')', False)
            else:
                if wasString:
                    return self.data2sfeel(coordinate, sheet, '"' + thisTest + '"' + ' in ' + variable, False)
                else:
                    return self.data2sfeel(coordinate, sheet, thisTest + ' in ' + variable, False)

        # All that is left is 'variable' '=' 'test'
        if wasString:
            return self.data2sfeel(coordinate, sheet, variable + ' = "' + thisTest + '"', False)
        else:
            return self.data2sfeel(coordinate, sheet, variable + ' = ' + thisTest, False)


    def list2sfeel(self, value):
        # Convert a Python list to a FEEL List - could have embedded lists and/or dictionaries
        newValue = '['
        for i in range(len(value)):
            if newValue != '[':
                newValue += ','
            if isinstance(value[i], list):
                newValue += self.list2sfeel(value[i])
            elif isinstance(value[i], dict):
                newValue += self.dict2sfeel(value[i])
            else:
                newValue += self.value2sfeel(value[i])
        return newValue + ']'

    def dict2sfeel(self, value):
        # Convert a Python dictionary to a FEEL Context - could have embedded lists and/or dictioaries
        newValue = '{'
        for key in value:
            if newValue != '{':
                newValue += ','
            newValue += '"' + key + '":'
            if isinstance(value[key], list):
                newValue += self.list2sfeel(value[key])
            elif isinstance(value[key], dict):
                newValue += self.dict2sfeel(value[key])
            else:
                newValue += self.value2sfeel(value[key])
        return newValue + '}'

    def excel2sfeel(self, value, isInput, coordinate, sheet):
        # Convert an Excel cell value into a FEEL equivalent
        # For non-strings this is a simple data conversion
        # For strings there are some recognised strings that are either DMN input rules (-)
        # or FEEL constants (true, false, null)
        # of a FEEL range, or a FEEL List, or a FEEL Context
        # All other strings could be FEEL expressions,
        # or they could be string constant that need to be wrapped in double quotes to make them valid FEEL
        if value is None:
            return 'null'
        elif isinstance(value, bool):
            if value:
                return 'true'
            else:
                return 'false'
        elif isinstance(value, float):
            return str(value)
        elif isinstance(value, int):
            return str(value)
        elif isinstance(value, str):
            if len(value) > 0:
                if isInput and (value == '-'):      # DMN skip input test token
                    pass
                elif (value == 'true') or (value == 'false') or (value == 'null'):      # FEEL string constant
                    pass
                elif re.match(r'^\s*(\[|\().*(\)|\])\s*$', value) is not None:           # FEEL range or List
                    pass
                elif re.match(r'^\s*\{.*\}\s*$', value) is not None:                      # FEEL Context
                    pass
                else:
                    listItems = []
                    try:            # Test for a comma separated list - could be a mixture of numbers, string, expressions
                        for row in csv.reader([value], dialect=csv.excel, doublequote=False, escapechar='\\'):
                            listItems = list(row)
                    except:
                        pass
                    if len(listItems) < 2:          # For lists we have to assume that the list items are correctly 'quoted'
                        sfeelText = self.data2sfeel(coordinate, sheet, value, True)
                        if sfeelText is None:               # Not valid FEEL - make this a FEEL string
                            if (value[0] != '"') or (value[-1] != '"'):
                                value = '"' + value + '"'
            else:
                value = '""'
            return value
        elif isinstance(value, datetime.date):
            return value.isoformat()
        elif isinstance(value, datetime.datetime):
            return value.isoformat(sep='T')
        elif isinstance(value, datetime.time):
            return value.isoformat()
        elif isinstance(value, datetime.timedelta):
            duration = value.total_seconds()
            secs = duration % 60
            duration = int(duration / 60)
            mins = duration % 60
            duration = int(duration / 60)
            hours = duration % 24
            days = int(duration / 24)
            return 'P%dDT%dH%dM%dS' % (days, hours, mins, secs)
        else:
            self.errors.append("Invalid Data '{!r}' at '{!s}' on sheet '{!s}' - not a valid S-FEEL data type".format(value, coordinate, sheet))
            return None


    def value2sfeel(self, value):
        # Converty a Python value to a FEEL equivalent
        if value is None:
            return 'null'
        elif isinstance(value, bool):
            if value:
                return 'true'
            else:
                return 'false'
        elif isinstance(value, float):
            return str(value)
        elif isinstance(value, int):
            return str(value)
        elif isinstance(value, str):
            if value in self.glossary:
                return value
            elif len(value) == 0:
                return '""'
            elif (value[0] == '"') and (value[-1] == '"'):
                return '"' + value[1:-1].replace('"', r'\"') + '"'
            else:
                return '"' + value.replace('"', r'\"') + '"'
        elif isinstance(value, list):
            return self.list2sfeel(value)
        elif isinstance(value, dict):
            return self.dict2sfeel(value)
        elif isinstance(value, datetime.date):
            return value.isoformat()
        elif isinstance(value, datetime.datetime):
            return value.isoformat(sep='T')
        elif isinstance(value, datetime.time):
            return value.isoformat()
        elif isinstance(value, datetime.timedelta):
            duration = value.total_seconds()
            secs = duration % 60
            duration = int(duration / 60)
            mins = duration % 60
            duration = int(duration / 60)
            hours = duration % 24
            days = int(duration / 24)
            return 'P%dDT%dH%dM%dS' % (days, hours, mins, secs)
        else:
            self.errors.append("Invalid Data '{!r}' - not a valid S-FEEL data type".format(value))
            return None


    def tableSize(self, cell):
        '''
        Determine the size of a table
        '''
        rows = 1
        cols = 0
        # The headers must not be null
        while cell.offset(row=1, column=cols).value is not None:
            coordinate = cell.offset(row=1, column=cols).coordinate
            for merged in self.mergedCells:
                if coordinate in merged:
                    cols += merged.max_col - merged.min_col + 1
                    break
            else:
                cols += 1

        # A row of all None is the end of the table
        inTable = True
        while inTable:
            inTable = False
            for col in range(cols):
                if cell.offset(row=rows, column=col).value is not None:
                    coordinate = cell.offset(row=rows).coordinate
                    for merged in self.mergedCells:
                        if coordinate in merged:
                            rows += merged.max_row - merged.min_row + 1
                            break
                    else:
                        rows += 1
                    inTable = True
                    break
        return (rows, cols)



    def parseDecionTable(self, cell, sheet):
        '''
        Parse a Decision Table
        '''
        startRow = cell.row
        startCol = cell.column
        table = cell.value
        coordinate = cell.coordinate
        table = str(table).strip()
        (rows, cols) = self.tableSize(cell)     # Find the length and width of the decision table
        if (rows == 1) and (cols == 0):
            # Empty table
            self.errors.append("Decision table '{!s}' at '{!s}' on sheet '{!s}' is empty".format(table, coordinate, sheet))
            return (rows, cols, -1)
        # print("Parsing Decision Table '{!s}' at '{!s}' on sheet '{!s}'".format(table, coordinate, sheet))
        self.rules[table] = []
        # Check the next cell down to determine the decision table layout
        # In 'Rules as Rows' layout this will be the hit policy
        # In 'Rules as Columns' layout this will be an input [which will be in the Glossary]
        # In 'Rules as Crosstab' layout this will be an output [which will be in the Glossary]
        thisCell = cell.offset(row=1).value
        nextCell = cell.offset(row=2).value
        if nextCell is not None:
            nextCell = str(nextCell).strip()
        if thisCell is None:
            # Empty table
            self.errors.append("Decision table '{!s}' at '{!s}' on sheet '{!s}' is empty".format(table, coordinate, sheet))
            return (rows, cols, -1)
        thisCell = str(thisCell).strip()
        if thisCell not in self.glossary:
            # print('Rules as Rows')
            # Rules as rows
            # Parse the heading
            inputColumns = outputColumns = 0
            doingValidity = False
            doingAnnotation = False
            coordinate = cell.offset(row=1).coordinate
            self.decisionTables[table]['inputColumns'] = []
            self.decisionTables[table]['inputValidity'] = []
            self.decisionTables[table]['outputColumns'] = []
            self.decisionTables[table]['outputValidity'] = []
            # Process all the columns on row 1
            for thisCol in range(cols):
                thisCell = cell.offset(row=1, column=thisCol).value
                coordinate = cell.offset(row=1, column=thisCol).coordinate
                if thisCol == 0:   # This should be the hit policy
                    if thisCell is None:
                        hitPolicy = 'U'
                    else:
                        thisCell = str(thisCell).strip()
                        hitPolicy = thisCell
                    if hitPolicy[0] not in ['U', 'A', 'P', 'F', 'C', 'O', 'R']:
                        self.errors.append("Invalid hit policy '{!s}' for table '{!s}'".format(hitPolicy, table))
                        return (rows, cols, -1)
                    if len(hitPolicy) != 1:
                        if (len(hitPolicy) != 2) or (hitPolicy[1] not in ['+', '<', '>', '#']):
                            self.errors.append("Invalid hit policy '{!s}' for table '{!s}'".format(hitPolicy, table))
                            return (rows, cols, -1)
                    self.decisionTables[table]['hitPolicy'] = hitPolicy
                    # Check if there is a second heading row (for the validity)
                    border = cell.offset(row=1, column=thisCol).border
                    if border.bottom.style != 'double':
                        border = cell.offset(row=2, column=thisCol).border
                        if border.top.style != 'double':
                            doingValidity = True
                    # Check if this is an output only decision table (no input columns)
                    if border.right.style == 'double':
                        doingInputs = False
                    else:
                        doingInputs = True
                    continue            # proceed to column 2
                # Process an input, output or annotation heading
                if thisCell is None:
                    if doingInputs:
                        self.errors.append("Missing Input heading in table '{!s}' at '{!s}' on sheet '{!s}'".format(table, coordinate, sheet))
                    elif not doingAnnotation:
                        self.errors.append("Missing Output heading in table '{!s}' at '{!s}' on sheet '{!s}'".format(table, coordinate, sheet))
                    else:
                        self.errors.append("Missing Annotation heading in table '{!s}' at '{!s}' on sheet '{!s}'".format(table, coordinate, sheet))
                    return (rows, cols, -1)
                thisCell = str(thisCell).strip()
                if not doingAnnotation:
                    # Check that this headings is in the Glossary - all input and output headings must be in the Glossary
                    if thisCell not in self.glossary:
                        if doingInputs:
                            self.errors.append("Input heading '{!s}' in table '{!s}' at '{!s}' on sheet '{!s}' is not in the Glossary".format(thisCell, table, coordinate, sheet))
                        else:
                            self.errors.append("Output heading '{!s}' in table '{!s}' at '{!s}' on sheet '{!s}' is not in the Glossary".format(thisCell, table, coordinate, sheet))
                        return (rows, cols, -1)
                if doingInputs:
                    inputColumns += 1
                    thisInput = len(self.decisionTables[table]['inputColumns'])
                    self.decisionTables[table]['inputColumns'].append({})
                    self.decisionTables[table]['inputColumns'][thisInput]['name'] = thisCell
                    # Check if we have hit the end of input columns (double border ending this cell, or starting next cell)
                    border = cell.offset(row=1, column=thisCol).border
                    if border.right.style == 'double':
                        doingInputs = False
                    border = cell.offset(row=1, column=thisCol + 1).border
                    if border.left.style == 'double':
                        doingInputs = False
                elif not doingAnnotation:       # this is an output heading
                    outputColumns += 1
                    thisOutput = len(self.decisionTables[table]['outputColumns'])
                    self.decisionTables[table]['outputColumns'].append({})
                    self.decisionTables[table]['outputColumns'][thisOutput]['name'] = thisCell
                    # Check if we have hit the end of output columns (double border ending this cell, or starting next cell)
                    border = cell.offset(row=1, column=thisCol).border
                    if border.right.style == 'double':
                        doingAnnotation = True
                    border = cell.offset(row=1, column=thisCol + 1).border
                    if border.left.style == 'double':
                        doingAnnotation = True
                    if doingAnnotation:
                        self.decisionTables[table]['annotation'] = []
                else:
                    self.decisionTables[table]['annotation'].append(thisCell)

            # Check that we at least has one output column in the headings
            if outputColumns == 0:
                self.errors.append("No Output column in table '{!s}' - missing double bar vertical border".format(table))
                return (rows, cols, -1)
            rulesRow = 2
            if doingValidity:
                # Parse the validity row
                doingInputs = True
                ranksFound = False          # A completely empty output validity row is not valid for hit policies 'P' and 'O'
                for thisCol in range(1, cols):
                    thisCell = cell.offset(row=2, column=thisCol).value
                    if thisCell is None:
                        if thisCol <= inputColumns:
                            self.decisionTables[table]['inputValidity'].append(None)
                        elif thisCol <= inputColumns + outputColumns:
                            self.decisionTables[table]['outputValidity'].append([])
                        continue
                    coordinate = cell.offset(row=2, column=thisCol).coordinate
                    if thisCol <= inputColumns:
                        inputName = self.decisionTables[table]['inputColumns'][thisCol - 1]['name']
                        variable = self.glossary[inputName]['item']
                        validityTests = self.excel2sfeel(thisCell, True, coordinate, sheet)
                        test = self.test2sfeel(variable, coordinate, sheet, validityTests)
                        self.decisionTables[table]['inputValidity'].append(test)
                    elif thisCol <= inputColumns + outputColumns:
                        ranksFound = True       # We have at least one output validity cell
                        self.decisionTables[table]['outputValidity'].append([])
                        if not isinstance(thisCell, str):
                            self.decisionTables[table]['outputValidity'][-1].append(thisCell)
                        else:
                            # Allow for multi-word strings, wrapped in double quotes with embedded commas - require csv.excel compliance
                            try:
                                for row in csv.reader([thisCell], dialect=csv.excel, doublequote=False, escapechar='\\'):
                                    validityTests = list(row)
                            except:
                                validityTests = [thisCell]
                            for validTest in validityTests:         # Each validity value should be a valid S-FEEL constant
                                sfeelText = self.data2sfeel(coordinate, sheet, validTest, True)
                                if sfeelText is None:               # Not valid FEEL - make this a FEEL string
                                    validValue = self.sfeel('{}'.format('"' + validTest + '"'))
                                else:
                                    validValue = self.sfeel('{}'.format(validTest))
                                self.decisionTables[table]['outputValidity'][-1].append(validValue)
                                if isinstance(validValue, float):
                                    self.decisionTables[table]['outputValidity'][-1].append(str(validValue))
                    else:
                        break
                doingValidity = False
                rulesRow += 1
                if (not ranksFound) and (self.decisionTables[table]['hitPolicy'] in ['P', 'O']):
                    self.errors.append("Decision table '{!s}' has hit policy '{!s}' but there is no ordered list of output values".format(
                        table, self.decisionTables[table]['hitPolicy']))
                    return (rows, cols, -1)
            elif self.decisionTables[table]['hitPolicy'] in ['P', 'O']:
                self.errors.append("Decision table '{!s}' has hit policy '{!s}' but there is no ordered list of output values".format(
                    table, self.decisionTables[table]['hitPolicy']))
                return (rows, cols, -1)
            else:       # Set up empty validity lists
                for thisCol in range(1, cols):
                    if thisCol <= inputColumns:
                        self.decisionTables[table]['inputValidity'].append(None)
                    elif thisCol <= inputColumns + outputColumns:
                        self.decisionTables[table]['outputValidity'].append([])
            lastTest = lastResult = {}
            # Parse the rules
            for thisRow in range(rulesRow, rows):
                thisCol = 0
                thisRule = len(self.rules[table])
                self.rules[table].append({})
                self.rules[table][thisRule]['tests'] = []
                self.rules[table][thisRule]['outputs'] = []
                if doingAnnotation:
                    self.rules[table][thisRule]['annotation'] = []
                for thisCol in range(cols):
                    thisCell = cell.offset(row=thisRow, column=thisCol).value
                    coordinate = cell.offset(row=thisRow, column=thisCol).coordinate
                    if thisCol == 0:
                        thisCell = str(thisCell).strip()
                        self.rules[table][thisRule]['ruleId'] = thisCell
                        continue

                    if thisCol <= inputColumns:
                        if thisCell is not None:
                            for merged in self.mergedCells:
                                if coordinate in merged:
                                    mergeCount = merged.max_row - merged.min_row
                                    break
                            else:
                                mergeCount = 0
                            # This is an input cell
                            thisCell = self.excel2sfeel(thisCell, True, coordinate, sheet)
                            if thisCell == '-':
                                lastTest[thisCol] = {}
                                lastTest[thisCol]['name'] = '.'
                                lastTest[thisCol]['mergeCount'] = mergeCount
                                continue
                            name = self.decisionTables[table]['inputColumns'][thisCol - 1]['name']
                            variable = self.glossary[name]['item']
                            test = self.test2sfeel(variable, coordinate, sheet, thisCell)
                            lastTest[thisCol] = {}
                            lastTest[thisCol]['name'] = name
                            lastTest[thisCol]['variable'] = variable
                            lastTest[thisCol]['test'] = test
                            lastTest[thisCol]['mergeCount'] = mergeCount
                        elif (thisCol in lastTest) and (lastTest[thisCol]['mergeCount'] > 0):
                            lastTest[thisCol]['mergeCount'] -= 1
                            name = lastTest[thisCol]['name']
                            if name == '.':
                                continue
                            variable = lastTest[thisCol]['variable']
                            test = lastTest[thisCol]['test']
                        else:
                            continue
                        # print("Setting test '{!s}' at '{!s}' on sheet '{!s}' to '{!s}'".format(name, coordinate, sheet, test))
                        self.rules[table][thisRule]['tests'].append((name, test, thisCol - 1))
                    elif thisCol <= inputColumns + outputColumns:
                        if thisCell is not None:
                            for merged in self.mergedCells:
                                if coordinate in merged:
                                    mergeCount = merged.max_row - merged.min_row
                                    break
                            else:
                                mergeCount = 0
                            # This is an output cell
                            name = self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]['name']
                            variable = self.glossary[name]['item']
                            result = self.excel2sfeel(thisCell, False, coordinate, sheet)
                            lastResult[thisCol] = {}
                            lastResult[thisCol]['name'] = name
                            lastResult[thisCol]['variable'] = variable
                            lastResult[thisCol]['result'] = result
                            lastResult[thisCol]['mergeCount'] = mergeCount
                        elif (thisCol in lastResult) and (lastResult[thisCol]['mergeCount'] > 0):
                            lastResult[thisCol]['mergeCount'] -= 1
                            name = lastResult[thisCol]['name']
                            variable = lastResult[thisCol]['variable']
                            result = lastResult[thisCol]['result']
                        else:
                            self.errors.append("Missing output value at '{!s}' on sheet '{!s}'".format(coordinate, sheet))
                            continue
                        rank = None
                        if self.decisionTables[table]['outputValidity'][thisCol - inputColumns - 1] != []:
                            try:
                                thisResult = float(result)
                            except:
                                thisResult = result
                            if thisResult in self.decisionTables[table]['outputValidity'][thisCol - inputColumns - 1]:
                                rank = self.decisionTables[table]['outputValidity'][thisCol - inputColumns - 1].index(thisResult)
                            else:
                                rank = -1
                        # print("Setting result '{!s}' at '{!s}' on sheet '{!s}' to '{!s}' with rank '{!s}'".format(name, coordinate, sheet, result, rank))
                        self.rules[table][thisRule]['outputs'].append((name, result, thisCol - inputColumns - 1, rank))
                    else:
                        self.rules[table][thisRule]['annotation'].append(thisCell)

        # Check for Rules as Columns
        elif (nextCell is not None) and (nextCell in self.glossary):
            # print('Rules as Columns')
            # Rules as columns
            # Parse the footer
            doingValidity = False
            doingAnnotation = False
            thisRow = rows - 1
            thisCol = 0
            # Process all the columns on the last row
            for thisCol in range(cols):
                thisCell = cell.offset(row=thisRow, column=thisCol).value
                if thisCol == 0:   # Should be hit policy
                    if thisCell is None:
                        hitPolicy = 'U'
                    else:
                        thisCell = str(thisCell).strip()
                        hitPolicy = thisCell
                    if (not isinstance(hitPolicy, str)) or (hitPolicy[0] not in ['U', 'A', 'P', 'F', 'C', 'O', 'R']):
                        self.errors.append("Invalid hit policy '{!s}' for table '{!s}'".format(hitPolicy, table))
                        return (rows, cols, -1)
                    if len(hitPolicy) != 1:
                        if (len(hitPolicy) != 2) or (hitPolicy[1] not in ['+', '<', '>', '#']):
                            self.errors.append("Invalid hit policy '{!s}' for table '{!s}'".format(hitPolicy, table))
                            return (rows, cols, -1)
                    self.decisionTables[table]['hitPolicy'] = hitPolicy
                    # Check if the second last row is validity
                    border = cell.offset(row=thisRow).border
                    if border.right.style != 'double':
                        border = cell.offset(row=thisRow, column=1).border
                        if border.left.style != 'double':
                            doingValidity = True
                elif (thisCol == 1) and doingValidity:
                    continue        # There is no validity next to the Hit Policy
                else:           # Process a rule id
                    thisRule = len(self.rules[table])
                    self.rules[table].append({})
                    self.rules[table][thisRule]['tests'] = []
                    self.rules[table][thisRule]['outputs'] = []
                    thisCell = str(thisCell).strip()
                    self.rules[table][thisRule]['ruleId'] = thisCell

            # Parse the heading
            inputRows = outputRows = 0
            self.decisionTables[table]['inputRows'] = []
            self.decisionTables[table]['inputValidity'] = []
            self.decisionTables[table]['outputRows'] = []
            self.decisionTables[table]['outputValidity'] = []
            doingInputs = True
            doingAnnotation = False
            for thisRow in range(1, rows - 1):
                thisCell = cell.offset(row=thisRow).value
                coordinate = cell.offset(row=thisRow).coordinate
                if thisCell is None:
                    if doingInputs:
                        self.errors.append("Missing Input heading in table '{!s}' at '{!s}' on sheet '{!s}'".format(table, coordinate, sheet))
                    elif not doingAnnotation:
                        self.errors.append("Missing Output heading in table '{!s}' at '{!s}' on sheet '{!s}'".format(table, coordinate, sheet))
                    else:
                        self.errors.append("Missing Annotation heading in table '{!s}' at '{!s}' on sheet '{!s}'".format(table, coordinate, sheet))
                    return (rows, cols, -1)
                thisCell = str(thisCell).strip()
                # Check that all the headings are in the Glossary
                if not doingAnnotation:
                    if thisCell not in self.glossary:
                        if doingInputs:
                            self.errors.append("Input heading '{!s}' in table '{!s}' at '{!s}' on sheet '{!s}' is not in the Glossary".format(thisCell, table, coordinate, sheet))
                        else:
                            self.errors.append("Output heading '{!s}' in table '{!s}' at '{!s}' on sheet '{!s}' is not in the Glossary".format(thisCell, table, coordinate, sheet))
                        return (rows, cols, -1)
                if doingInputs:
                    inputRows += 1
                    inRow = len(self.decisionTables[table]['inputRows'])
                    self.decisionTables[table]['inputRows'].append({})
                    self.decisionTables[table]['inputRows'][inRow]['name'] = thisCell
                    # Check if we have hit the end of input columns (double border ending this cell, or starting next cell)
                    border = cell.offset(row=thisRow).border
                    if border.bottom.style == 'double':
                        doingInputs = False
                    border = cell.offset(row=thisRow + 1).border
                    if border.top.style == 'double':
                        doingInputs = False
                elif not doingAnnotation:       # This is an output heading
                    outputRows += 1
                    outRow = len(self.decisionTables[table]['outputRows'])
                    self.decisionTables[table]['outputRows'].append({})
                    self.decisionTables[table]['outputRows'][outRow]['name'] = thisCell
                    # Check if we have hit the end of input columns (double border ending this cell, or starting next cell)
                    border = cell.offset(row=thisRow).border
                    if border.bottom.style == 'double':
                        doingAnnotation = True
                    border = cell.offset(row=thisRow + 1).border
                    if border.top.style == 'double':
                        doingAnnotation = True
                    if doingAnnotation:
                        self.decisionTables[table]['annotation'] = []
                else:
                    self.decisionTables[table]['annotation'].append(thisCell)

            # Check that we have at least one output column in the headings
            if outputRows == 0:
                self.errors.append("No Output row in table '{!s}' - missing double bar horizontal border".format(table))
                return (rows, cols, -1)

            rulesCol = 1
            if doingValidity:
                # Parse the validity column
                doingInputs = True
                ranksFound = False          # A completely empty output validity row is not valid for hit policies 'P' and 'O'
                for thisRow in range(1, rows - 1):
                    thisCell = cell.offset(row=thisRow, column=1).value
                    if thisCell is None:
                        if thisRow < inputRows:
                            self.decisionTables[table]['inputValidity'].append(None)
                        elif thisRow <= inputRows + outputRows:
                            self.decisionTables[table]['outputValidity'].append([])
                        continue
                    thisCell = str(thisCell).strip()
                    coordinate = cell.offset(row=thisRow, column=1).coordinate
                    if thisRow <= inputRows:
                        inputName = self.decisionTables[table]['inputRows'][thisRow - 1]['name']
                        variable = self.glossary[inputName]['item']
                        validityTests = self.excel2sfeel(thisCell, True, coordinate, sheet)
                        test = self.test2sfeel(variable, coordinate, sheet, validityTests)
                        self.decisionTables[table]['inputValidity'].append(test)
                    elif thisRow <= inputRows + outputRows:
                        ranksFound = True
                        self.decisionTables[table]['outputValidity'].append([])
                        if not isinstance(thisCell, str):
                            self.decisionTables[table]['outputValidity'][-1].append(thisCell)
                        else:
                            # Allow for multi-word strings, wrapped in double quotes with embedded commas - require csv.excel compliance
                            try:
                                for row in csv.reader([thisCell], dialect=csv.excel, doublequote=False, escapechar='\\'):
                                    validityTests = list(row)
                            except:
                                validityTests = [thisCell]
                            for validTest in validityTests:         # Each validity value should be a valid S-FEEL constant
                                sfeelText = self.data2sfeel(coordinate, sheet, validTest, True)
                                if sfeelText is None:               # Not valid FEEL - make this a FEEL string
                                    validValue = self.sfeel('{}'.format('"' + validTest + '"'))
                                else:
                                    validValue = self.sfeel('{}'.format(validTest))
                                self.decisionTables[table]['outputValidity'][-1].append(validValue)
                                if isinstance(validValue, float):
                                    self.decisionTables[table]['outputValidity'][-1].append(str(validValue))
                    else:
                        break
                rulesCol += 1
                doingValidity = False
                if (not ranksFound) and (self.decisionTables[table]['hitPolicy'] in ['P', 'O']):
                    self.errors.append("Decision table '{!s}' has hit policy '{!s}' but there is no ordered list of output values".format(
                        table, self.decisionTables[table]['hitPolicy']))
                    return (rows, cols, -1)
            elif self.decisionTables[table]['hitPolicy'] in ['P', 'O']:
                self.errors.append("Decision table '{!s}' has hit policy '{!s}' but there is no ordered list of output values".format(
                    table, self.decisionTables[table]['hitPolicy']))
                return (rows, cols, -1)
            else:       # Set up empty validity lists
                for thisRow in range(1, rows - 1):
                    if thisRow <= inputRows:
                        self.decisionTables[table]['inputValidity'].append(None)
                    elif thisRow <= inputRows + outputRows:
                        self.decisionTables[table]['outputValidity'].append([])

            # Parse the rules
            for thisRow in range(1, rows - 1):
                lastTest = lastResult = {}
                for thisCol in range(rulesCol, cols):
                    thisRule = thisCol - rulesCol
                    if doingAnnotation and ('annotation' not in self.rules[table][thisRule]):
                        self.rules[table][thisRule]['annotation'] = []
                    thisCell = cell.offset(row=thisRow, column=thisCol).value
                    coordinate = cell.offset(row=thisRow, column=thisCol).coordinate
                    if thisRow <= inputRows:
                        if thisCell is not None:
                            for merged in self.mergedCells:
                                if coordinate in merged:
                                    mergeCount = merged.max_col - merged.min_col
                                    break
                            else:
                                mergeCount = 0
                            # This is an input cell
                            thisCell = self.excel2sfeel(thisCell, True, coordinate, sheet)
                            if thisCell == '-':
                                lastTest = {}
                                lastTest['name'] = '.'
                                lastTest['mergeCount'] = mergeCount
                                continue
                            name = self.decisionTables[table]['inputRows'][thisRow - 1]['name']
                            variable = self.glossary[name]['item']
                            test = self.test2sfeel(variable, coordinate, sheet, thisCell)
                            lastTest = {}
                            lastTest['name'] = name
                            lastTest['variable'] = variable
                            lastTest['test'] = test
                            lastTest['mergeCount'] = mergeCount
                        elif ('mergeCount' in lastTest) and (lastTest['mergeCount'] > 0):
                            lastTest['mergeCount'] -= 1
                            name = lastTest['name']
                            if name == '.':
                                continue
                            variable = lastTest['variable']
                            test = lastTest['test']
                        else:
                            continue
                        # print("Setting test '{!s}' at '{!s}' on sheet '{!s}' to '{!s}'".format(name, coordinate, sheet, test))
                        self.rules[table][thisRule]['tests'].append((name, test, thisRow - 1))
                    elif thisRow <= inputRows + outputRows:
                        if thisCell is not None:
                            for merged in self.mergedCells:
                                if coordinate in merged:
                                    mergeCount = merged.max_col - merged.min_col
                                    break
                            else:
                                mergeCount = 0
                            # This is an output column
                            name = self.decisionTables[table]['outputRows'][thisRow - inputRows - 1]['name']
                            variable = self.glossary[name]['item']
                            result = self.excel2sfeel(thisCell, False, coordinate, sheet)
                            lastResult = {}
                            lastResult['name'] = name
                            lastResult['variable'] = variable
                            lastResult['result'] = result
                            lastResult['mergeCount'] = mergeCount
                        elif ('mergeCount' in lastResult) and (lastResult['mergeCount'] > 0):
                            lastResult['mergeCount'] -= 1
                            name = lastResult['name']
                            variable = lastResult['variable']
                            result = lastResult['result']
                        else:
                            self.errors.append("Missing output value at '{!s}' on sheet '{!s}'".format(coordinate, sheet))
                            continue
                        rank = None
                        if self.decisionTables[table]['outputValidity'][thisRow - inputRows - 1] != []:
                            try:
                                thisCellResult = float(result)
                            except:
                                thisCellResult = result
                            if thisCellResult in self.decisionTables[table]['outputValidity'][thisRow - inputRows - 1]:
                                rank = self.decisionTables[table]['outputValidity'][thisRow - inputRows - 1].index(thisCellResult)
                            else:
                                rank = -1
                        # print("Setting result '{!s}' at '{!s}' on sheet '{!s}' to '{!s}' with rank '{!s}'".format(name, coordinate, sheet, result, rank))
                        self.rules[table][thisRule]['outputs'].append((name, result, thisRow - inputRows - 1, rank))
                    else:
                        self.rules[table][thisRule]['annotation'].append(thisCell)

        else:
            # Rules as crosstab
            # This is the output, and the only output
            # print('Rules as Crosstab')
            thisCell = cell.offset(row=1).value
            outputVariable = str(thisCell).strip()
            # This should be merged cell - need a row and a column of variables, plus another row and column of tests (as a minimum)
            coordinate = cell.offset(row=1).coordinate
            for merged in self.mergedCells:
                if coordinate in merged:
                    width = merged.max_col - merged.min_col + 1
                    height = merged.max_row - merged.min_row + 1
                    break
            else:
                self.errors.append("Decision table '{!s}' - unknown DMN rules table type".format(table))
                return (rows, cols, -1)

            self.decisionTables[table]['hitPolicy'] = 'U'
            self.decisionTables[table]['inputColumns'] = []
            self.decisionTables[table]['inputValidity'] = []
            self.decisionTables[table]['inputValidity'].append(None)
            self.decisionTables[table]['inputRows'] = []
            self.decisionTables[table]['outputValidity'] = []
            self.decisionTables[table]['outputValidity'].append([])
            self.decisionTables[table]['output'] = {}
            self.decisionTables[table]['output']['name'] = outputVariable

            # Parse the horizontal heading
            coordinate = cell.offset(row=1, column=width).coordinate
            for merged in self.mergedCells:
                if coordinate in merged:
                    horizontalCols = merged.max_col - merged.min_col + 1
                    break
            else:
                horizontalCols = 1

            heading = cell.offset(row=1, column=width).value
            if heading is None:
                self.errors.append("Crosstab Decision table '{!s}' is missing a horizontal heading".format(table))
                return (rows, cols, -1)
            heading = str(heading).strip()
            if heading.find(',') != -1:
                inputs = heading.split(',')
            else:
                inputs = [heading]
            if len(inputs) < height - 1:
                self.errors.append("Crosstab Decision table '{!s}' is missing one or more rows of horizontal values".format(table))
                return (rows, cols, -1)
            elif len(inputs) > height - 1:
                self.errors.append("Crosstab Decision table '{!s}' has too many rows of horizontal values".format(table))
                return (rows, cols, -1)

            for thisVariable in range(height - 1):
                lastTest = {}
                for thisCol in range(horizontalCols):
                    if thisVariable == 0:
                        self.decisionTables[table]['inputColumns'].append({})
                        self.decisionTables[table]['inputColumns'][thisCol]['tests'] = []

                    thisCell = cell.offset(row=2 + thisVariable, column=width + thisCol).value
                    coordinate = cell.offset(row=2 + thisVariable, column=width + thisCol).coordinate
                    if thisCell is not None:
                        for merged in self.mergedCells:
                            if coordinate in merged:
                                mergeCount = merged.max_col - merged.min_col
                                break
                        else:
                            mergeCount = 0
                        # This is an input cell
                        thisCell = self.excel2sfeel(thisCell, True, coordinate, sheet)
                        if thisCell == '-':
                            lastTest[thisVariable] = {}
                            lastTest[thisVariable]['name'] = '.'
                            lastTest[thisVariable]['mergeCount'] = mergeCount
                            continue
                        name = inputs[thisVariable].strip()
                        variable = self.glossary[name]['item']
                        test = self.test2sfeel(variable, coordinate, sheet, thisCell)
                        lastTest[thisVariable] = {}
                        lastTest[thisVariable]['name'] = name
                        lastTest[thisVariable]['variable'] = variable
                        lastTest[thisVariable]['test'] = test
                        lastTest[thisVariable]['mergeCount'] = mergeCount
                    elif (thisVariable in lastTest) and (lastTest[thisVariable]['mergeCount'] > 0):
                        lastTest[thisVariable]['mergeCount'] -= 1
                        name = lastTest[thisVariable]['name']
                        if name == '.':
                            continue
                        variable = lastTest[thisVariable]['variable']
                        test = lastTest[thisVariable]['test']
                    else:
                        self.errors.append("Missing horizontal input test at '{!s}' on sheet '{!s}'".format(coordinate, sheet))
                        continue
                    # print("Setting horizontal test '{!s}' at '{!s}' on sheet '{!s}' to '{!s}'".format(name, coordinate, sheet, test))
                    self.decisionTables[table]['inputColumns'][thisCol]['tests'].append((name, test, 0))

            # Parse the vertical heading
            coordinate = cell.offset(row=1 + height).coordinate
            for merged in self.mergedCells:
                if coordinate in merged:
                    verticalRows = merged.max_row - merged.min_row + 1
                    break
            else:
                verticalRows = 1

            heading = cell.offset(row=1 + height).value
            if heading is None:
                self.errors.append("Crosstab Decision table '{!s}' is missing a vertical heading".format(table))
                return (rows, cols, -1)
            heading = str(heading).strip()
            if heading.find(',') != -1:
                inputs = heading.split(',')
            else:
                inputs = [heading]
            if len(inputs) < width - 1:
                self.errors.append("Crosstab Decision table '{!s}' is missing one or more columns of verticals".format(table))
                return (rows, cols, -1)
            elif len(inputs) > width - 1:
                self.errors.append("Crosstab Decision table '{!s}' has too many columns of vertical values".format(table))
                return (rows, cols, -1)

            for thisVariable in range(width - 1):
                lastTest = {}
                for thisRow in range(verticalRows):
                    if thisVariable == 0:
                        self.decisionTables[table]['inputRows'].append({})
                        self.decisionTables[table]['inputRows'][thisRow]['tests'] = []

                    thisCell = cell.offset(row=1 + height + thisRow, column=1 + thisVariable).value
                    coordinate = cell.offset(row=1 + height + thisRow, column=1 + thisVariable).coordinate
                    if thisCell is not None:
                        for merged in self.mergedCells:
                            if coordinate in merged:
                                mergeCount = merged.max_row - merged.min_row
                                break
                        else:
                            mergeCount = 0
                        # This is an input cell
                        thisCell = self.excel2sfeel(thisCell, True, coordinate, sheet)
                        if thisCell == '-':
                            lastTest[thisVariable] = {}
                            lastTest[thisVariable]['name'] = '.'
                            lastTest[thisVariable]['mergeCount'] = mergeCount
                            continue
                        name = inputs[thisVariable].strip()
                        variable = self.glossary[name]['item']
                        test = self.test2sfeel(variable, coordinate, sheet, thisCell)
                        lastTest[thisVariable] = {}
                        lastTest[thisVariable]['name'] = name
                        lastTest[thisVariable]['variable'] = variable
                        lastTest[thisVariable]['test'] = test
                        lastTest[thisVariable]['mergeCount'] = mergeCount
                    elif (thisVariable in lastTest) and (lastTest[thisVariable]['mergeCount'] > 0):
                        lastTest[thisVariable]['mergeCount'] -= 1
                        name = lastTest[thisVariable]['name']
                        if name == '.':
                            continue
                        variable = lastTest[thisVariable]['variable']
                        test = lastTest[thisVariable]['test']
                        thisCell = lastTest[thisVariable]['thisCell']
                    else:
                        self.errors.append("Missing vertical input test at '{!s}' on sheet '{!s}'".format(coordinate, sheet))
                        continue
                    # print("Setting veritical test '{!s}' at '{!s}' on sheet '{!s}' to '{!s}'".format(name, coordinate, sheet, test))
                    self.decisionTables[table]['inputRows'][thisRow]['tests'].append((name, test, 0))

            # Now build the Rules
            thisRule = 0
            for row in range(verticalRows):
                for col in range(horizontalCols):
                    self.rules[table].append({})
                    self.rules[table][thisRule]['ruleId'] = str(row + 1) + ':' + str(col + 1)
                    self.rules[table][thisRule]['tests'] = []
                    self.rules[table][thisRule]['outputs'] = []
                    self.rules[table][thisRule]['tests'] += self.decisionTables[table]['inputColumns'][col]['tests']
                    self.rules[table][thisRule]['tests'] += self.decisionTables[table]['inputRows'][row]['tests']
                    thisCell = cell.offset(row=1 + height + row, column=width + col).value
                    coordinate = cell.offset(row=1 + height + row, column=width + col).coordinate
                    if thisCell is None:
                        self.errors.append("Missing output result at '{!s}' on sheet '{!s}'".format(coordinate, sheet))
                        return (rows, cols, -1)
                    # This is an output cell
                    thisCell = self.excel2sfeel(thisCell, False, coordinate, sheet)
                    name = self.decisionTables[table]['output']['name']
                    variable = self.glossary[name]['item']
                    result = self.excel2sfeel(thisCell, False, coordinate, sheet)
                    # print("Setting result at '{!s}' on sheet '{!s}' to '{!s}'".format(coordinate, sheet, result))
                    self.rules[table][thisRule]['outputs'].append((name, result, 0, 0))
                    thisRule += 1

        return (rows, cols, len(self.rules[table]))


    def load(self, rulesBook):
        """
        Load a rulesBook

        This routine load an Excel workbook which must contain a 'Glossary' sheet,
        a 'Decision' sheet and other sheets containing DMN rules tables

        Args:
            param1 (str): The name of the Excel workbook (including path if it is not in the current working directory

        Returns:
            dict: status

            'status' is a dictionary of different status information.
            Currently only status['error'] is implemented.
            If the key 'error' is present in the status dictionary,
            then load() encountered one or more errors and status['error'] is the list of those errors

        """

        self.errors = []
        try:
            wb = load_workbook(filename=rulesBook)
        except Exception as e:
            self.errors.append("No readable workbook named '{!s}'!".format(rulesBook))
            status = {}
            status['errors'] = self.errors
            return status        
        return self.use(wb)

    def use(self, workbook):
        """
        Use a rules workbookook

        This routine uses an already loaded Excel workbook which must contain a 'Glossary' sheet,
        a 'Decision' sheet and other sheets containing DMN rules tables

        Args:
            param1 (openpyxl.workbook): An openpyxl workbook (either loaded with openpyxl or created using openpyxl)

        Returns:
            dict: status

            'status' is a dictionary of different status information.
            Currently only status['error'] is implemented.
            If the key 'error' is present in the status dictionary,
            then use() encountered one or more errors and status['error'] is the list of those errors

        """

        self.errors = []
        if not isinstance(workbook, openpyxl.Workbook):
            self.errors.append("workbook is not a valid openpyxl workbook")
            status = {}
            status['errors'] = self.errors
            return status

        self.wb = workbook

        # Read in the mandatory Glossary
        try:
            ws = self.wb['Glossary']
        except (KeyError):
            self.errors.append('No rulesBook sheet named Glossary!')
            status = {}
            status['errors'] = self.errors
            return status
        self.mergedCells = ws.merged_cells.ranges
        inGlossary = False
        self.glossary = {}
        self.glossaryItems = {}
        self.glossaryConcepts = {}
        for row in ws.rows:
            for cell in row:
                if not inGlossary:
                    thisCell = cell.value
                    if isinstance(thisCell, str):
                        if thisCell.startswith('Glossary'):
                            (rows, cols) = self.tableSize(cell)
                            if cols < 3:
                                self.errors.append('Invalid Glossary - not 3 columns wide')
                                status = {}
                                status['errors'] = self.errors
                                return status
                            inGlossary = True
                            break
                if inGlossary:
                    break
            if inGlossary:
                break
        if not inGlossary:
            self.errors.append('Glossary not found')
            status = {}
            status['errors'] = self.errors
            return status

        if cell.offset(row=1).value != 'Variable':
            self.errors.append("Missing Glossary heading - no 'Variable' column")
            status = {}
            status['errors'] = self.errors
            return status
        elif cell.offset(row=1, column=1).value != 'Business Concept':
            self.errors.append("Bad Glossary heading - missing column 'Business Concept'")
            status = {}
            status['errors'] = self.errors
            return status
        elif cell.offset(row=1, column=2).value != 'Attribute':
            self.errors.append("Bad Glossary heading - missing column 'Attribute'")
            status = {}
            status['errors'] = self.errors
            return status
        thisConcept = None
        for thisRow in range(2, rows):
            variable = cell.offset(row=thisRow).value
            coordinate = cell.offset(row=thisRow).coordinate
            if variable in self.glossary:
                self.errors.append("Variable '{!s}' with multiple definitions in Glossary at '{!s}'".format(variable, coordinate))
                status = {}
                status['errors'] = self.errors
                return status
            concept = cell.offset(row=thisRow, column=1).value
            coordinate = cell.offset(row=thisRow, column=1).coordinate
            if thisConcept is None:
                if concept is None:
                    self.errors.append("Missing Business Concept in Glossary at '{!s}'".format(coordinate))
                    status = {}
                    status['errors'] = self.errors
                    return status
            if concept is not None:
                if '.' in concept:
                    self.errors.append("Bad Business Concept '{!s}' in Glossary at '{!s}'".format(concept, coordinate))
                    status = {}
                    status['errors'] = self.errors
                    return status
                if concept in self.glossaryConcepts:
                    self.errors.append("Multiple definitions of Business Concept '{!s}' in Glossary at '{!s}'".format(concept, coordinate))
                    status = {}
                    status['errors'] = self.errors
                    return status
                self.glossaryConcepts[concept] = []
                thisConcept = concept
            attribute = cell.offset(row=thisRow, column=2).value
            coordinate = cell.offset(row=thisRow, column=2).coordinate
            if (attribute is None) or ('.' in attribute):
                self.errors.append("Bad Business Attribute '{!s}' for Variable '{!s}' in Business in Concept '{!s}' in Glossary at '{!s}'".format(attribute, variable, thisConcept, coordinate))
                status = {}
                status['errors'] = self.errors
                return status
            item = thisConcept + '.' + attribute
            self.glossary[variable] = {}
            self.glossary[variable]['item'] = item
            self.glossary[variable]['concept'] = thisConcept
            self.glossaryItems[item] = variable
            self.glossaryConcepts[thisConcept].append(variable)
        self.glossaryLoaded = True

        # Validate the glossary
        self.initGlossary()
        if len(self.errors) > 0:
            status = {}
            status['errors'] = self.errors
            return status

        # Read in the mandatory Decision
        try:
            ws = self.wb['Decision']
        except (KeyError):
            self.errors.append('No rulesBook sheet named Decision!')
            status = {}
            status['errors'] = self.errors
            return status
        self.mergedCells = ws.merged_cells.ranges
        inDecision = False
        endDecision = False
        decisionColumn = None
        self.decisions = []
        self.decisionTables = {}
        self.rules = {}
        for row in ws.rows:
            for cell in row:
                if not inDecision:
                    thisCell = cell.value
                    if isinstance(thisCell, str):
                        if thisCell.startswith('Decision'):
                            (rows, cols) = self.tableSize(cell)
                            if cols < 2:
                                self.errors.append('Invalid Decision - less than 2 columns wide')
                                status = {}
                                status['errors'] = self.errors
                                return status
                            inDecision = True
                            break
                if inDecision:
                    break
            if inDecision:
                break

        if not inDecision:
            self.errors.append('Decision not found')
            status = {}
            status['errors'] = self.errors
            return status
        inputColumns = 0
        inputVariables = []
        doingInputs = True
        doingDecisions = False
        for thisCol in range(cols):
            thisCell = cell.offset(row=1, column=thisCol).value
            coordinate = cell.offset(row=1, column=thisCol).coordinate
            inputVariables.append(thisCell)
            if doingInputs:
                # Check that all the headings are in the Glossary
                if thisCell == 'Decisions':
                    doingInputs = False
                    doingDecisions = True
                    continue
                if thisCell not in self.glossary:
                    self.errors.append("Input heading '{!s}' in the Decision table at '{!s}' is not in the Glossary".format(thisCell, coordinate))
                    status = {}
                    status['errors'] = self.errors
                    return status
                inputColumns += 1
            elif doingDecisions:
                if thisCell == 'Execute Decision Tables':
                    doingDecisions = False
                else:
                    self.errors.append("Bad heading '{!s}' in the Decision table at '{!s}'".format(thisCell, coordinate))
                    status = {}
                    status['errors'] = self.errors
                    return status
        if doingInputs:
            self.errors.append("Missing heading 'Decisions' in Decision table")
            status = {}
            status['errors'] = self.errors
            return status
        if doingDecisions:
            self.errors.append("Missing heading 'Execute Decision Table' in Decision table")
            status = {}
            status['errors'] = self.errors
            return status
        lastTest = {}
        for thisRow in range(2, rows):
            inputTests = []
            annotations = []
            for thisCol in range(cols):
                thisCell = cell.offset(row=thisRow, column=thisCol).value
                thisDataType = cell.offset(row=thisRow, column=thisCol).data_type
                coordinate = cell.offset(row=thisRow, column=thisCol).coordinate
                if thisCol < inputColumns:
                    if thisCell is not None:
                        for merged in self.mergedCells:
                            if coordinate in merged:
                                mergeCount = merged.max_row - merged.min_row
                                break
                        else:
                            mergeCount = 0
                        # This is an input value
                        thisCell = self.excel2sfeel(thisCell, True, coordinate, 'Decision')
                        if thisCell == '-':
                            lastTest[thisCol] = {}
                            lastTest[thisCol]['name'] = '.'
                            lastTest[thisCol]['mergeCount'] = mergeCount
                            continue
                        name = inputVariables[thisCol]
                        variable = self.glossary[name]['item']
                        test = self.test2sfeel(variable, coordinate, 'Decision', thisCell)
                        lastTest[thisCol] = {}
                        lastTest[thisCol]['name'] = name
                        lastTest[thisCol]['variable'] = variable
                        lastTest[thisCol]['test'] = test
                        lastTest[thisCol]['mergeCount'] = mergeCount
                    elif (thisCol in lastTest) and (lastTest[thisCol]['mergeCount'] > 0):
                        lastTest[thisCol]['mergeCount'] -= 1
                        name = lastTest[thisCol]['name']
                        if name == '.':
                            continue
                        variable = lastTest[thisCol]['variable']
                        test = lastTest[thisCol]['test']
                    else:
                        continue
                    inputTests.append((name, test))
                elif thisCol == inputColumns:
                    decision = cell.offset(row=thisRow, column=thisCol).value
                elif thisCol == inputColumns + 1:
                    table = cell.offset(row=thisRow, column=thisCol).value
                    coordinate = cell.offset(row=1, column=thisCol).coordinate
                    if (table in self.decisionTables) and (self.decisionTables[table]['name'] != decision):
                        self.errors.append("Execution Decision Table '{!s}' redefined in the Decision table at '{!s}'".format(table, coordinate))
                        status = {}
                        status['errors'] = self.errors
                        return status
                else:
                    name = inputVariables[thisCol]
                    annotations.append((name, thisCell))

            if table not in self.decisionTables:
                self.decisionTables[table] = {}
                self.decisionTables[table]['name'] = decision
            self.decisions.append((table, inputTests, annotations))

        # Now search for the Decision Tables
        self.rules = {}
        for sheet in self.wb.sheetnames:
            if sheet in ['Glossary', 'Decision', 'Test']:
                continue
            ws = self.wb[sheet]
            self.mergedCells = ws.merged_cells.ranges
            parsedRanges = []
            for row in ws.rows:
                for cell in row:
                    for i in range(len(parsedRanges)):
                        parsed = parsedRanges[i]
                        if cell.coordinate in parsed:
                            continue
                    thisCell = cell.value
                    if isinstance(thisCell, str):
                        if thisCell in self.decisionTables:
                            (rows, cols, rules) = self.parseDecionTable(cell, sheet)
                            if rules == -1:
                                status = {}
                                if len(self.errors) > 0:
                                    status['errors'] = self.errors
                                return status
                            elif rules == 0:
                                self.errors.append("Decision table '{!s}' has no rules".format(thisCell))
                                status = {}
                                status['errors'] = self.errors
                                return status
                            # Symbolically merge all the cells in this table
                            thisRow = cell.row
                            thisCol = cell.column
                            ws.merge_cells(start_row=thisRow, start_column=thisCol,
                                           end_row=thisRow + rows - 1, end_column=thisCol + cols - 1)
                            # Find this merge range
                            for thisMerged in self.mergedCells:
                                if (thisMerged.min_col == cell.column) and (thisMerged.min_row == cell.row):
                                    if thisMerged.max_col != (cell.column + cols - 1):
                                        continue
                                    if thisMerged.max_row != (cell.row + rows - 1):
                                        continue
                                    # Mark it as parsed
                                    parsedRanges.append(thisMerged)
                            break
        # Now check that every decision table has been found
        for (table, inputTests, decisionAnnotations) in self.decisions:
            if table not in self.rules:
                self.errors.append("Decision table '{!s}' not found".format(table))
                status = {}
                status['errors'] = self.errors
                return status

        self.isLoaded = True
        if 'Test'  in self.wb:
            self.wbt = self.wb
            self.testIsLoaded = True

        status = {}
        if len(self.errors) > 0:
            status['errors'] = self.errors
        return status

    def initGlossary(self):
        if not self.glossaryLoaded:
            self.errors.append('No rulesBook has been loaded')
            sys.exit(0)
        for item in self.glossaryItems:
            retVal = self.sfeel('{} <- null'.format(item))


    def replaceItems(self, text):
        # Replace any references to glossary items with their current value
        # If there are any, then 'text' will be a string (wrapped in "")
        # which 'must be' valid FEEL when the values are replace and the wrapping "" is removed
        newText = text
        if len(text) == 0:
            return text
        if (text[0] == '"') and (text[-1] == '"'):
            newText = newText[1:-1]
        oldText = newText
        # Start by replacing all 'Variable's with their BusinessConcept.Attribute equivalents (which are valid S-FEEL)
        # Being careful not to replace any BusinessConcept.Attributes that already exist in data
        for variable in self.glossary:
            item = self.glossary[variable]['item']
            if variable == self.glossary[variable]['concept']:       # The variable and it's business concept share the same name
                at = 0
                match = re.search(variable, newText[at:])
                while match is not None:
                    replaceIt = True
                    if match.end() == len(newText[at:]):
                        pass
                    elif newText[at + match.end():at + match.end() + 1] != '.':
                        pass
                    else:
                        for otherItem in self.glossaryItems:
                            if newText[at:].startswith(otherItem):
                                replaceIt = False
                                break
                    if replaceIt:
                        newText = newText[:at] + newText[at:].replace(variable, item, 1)
                    at += match.end()
                    match = re.search(variable, newText[at:])
            else:
                newText = newText.replace(variable, item)

        for item in self.glossaryItems:
            itemPattern = r'\b' + item.replace('.', r'\.') + r'\b'
            itemValue = self.sfeel('{}'.format(item))
            value = self.value2sfeel(itemValue)
            newText = re.sub(itemPattern, value, newText)
        if newText == oldText:
            return text
        else:
            return newText


    def decide(self, data):
        """
        Make a decision

        This routine runs the passed data through the loaded DMN rules and returns the decision

        Args:
            param1 (dict): The dictionary of data about which a decision is being made.

                - Each key in 'data' must match a 'Variable' in the Glossary
                - the matching 'Variable' in the Glossary will be set to the matching value for this key from the 'data' dictionary.
                - Any entry in the Glossary which does not have a key in the 'data' dictionary will be set to the value 'None'.
                - The values associated with each 'Variable' in the Glossary
                  will be the input values, for the input columns, in the first DMN rules table in the 'Decision'.

        Returns:
            tuple: (status, newData)

            status is a dictionary of different status information.
                Currently only status['error'] is implemented.
                If the key 'error' is present in the status dictionary,
                then decide() encountered one or more errors and status['error'] is the list of those errors

            newData

            * for a Single Hit Policy DMN rules table newData will be a decision dictionary of the decision.

            * for a Multi Hit Policy DMN rules tables newData will be a list of decison dictionaries; one for each matched rule.

            The keys to each decision dictionary are
                - 'Result' - for a Single Hit Policy DMN rules table, this will be a  dictionary where all the keys will be 'Variables'
                  from the Glossary and the matching the value will be the value of that 'Variable' after the decision was made.
                  For a Multi Hit Policy DMN rules table this will be a list of decision dictionaries, one for each matched rule.

                - 'Excuted Rule' - for a Single Hit Policy DMN rules table this will be
                  a tuple of the Decision Table Description ('Decisions' from the 'Decision' table),
                  the DMN rules table name ('Execute Decision Table' from the 'Decision' table)
                  and the Rule number for the rule that matched in that DMN rules Table.
                  For a Multi Hit Policy DMN rules table this will be the a list of tuples,
                  being the Decision Table Description, DMN rules table name and matched Rule number for each matching rule.

                - 'DecisionAnnotations'(optional) - list of tuples (heading, value) of the annotations
                  from the 'Decision' table named in the 'Executed Rule' tuple.

                - 'RuleAnnotations'(optional) - for a Single Hit Policy DMN rules table, this well be
                  a list of tuples (heading, value) of the annotations for the matching rule,
                  if there were any annotations for the matching rule.
                  For a Multi Hit Policy DMN rules table this will be a list of the lists of any annotations for each matching rule,
                  where an empty list means that the associated matching rule had no annotations.

            If the Decision table contains multiple rows (multiple DMN rules tables run sequentially in order to make the decision)
            then the returned 'newData' will be a list of decision dictionaries, with each containing
            the keys 'Result', 'Excuted Rule', 'DecisionAnnotations'(optional) and 'RuleAnnotations'(optional),
            being one list entry for each DMN rules table used from the Decision table.
            The final enty in this list is the final decision.
            All other entries are the intermediate states involved in making the final decision.

        """

        self.errors = []
        if not self.isLoaded:
            self.errors.append('No rulesBook has been loaded')
            status = {}
            status['errors'] = self.errors
            self.errors = []
            return (status, {})
        self.errors = []
        self.warnings = []
        self.initGlossary()
        validData = True
        for variable in data:
            if variable not in self.glossary:
                self.errors.append('variable ({!s}) not in Glossary'.format(variable))
                status = {}
                status['errors'] = self.errors
                self.errors = []
                return (status, {})
            item = self.glossary[variable]['item']
            value = data[variable]
            # Convert the passed Python data to it's FEEL equivalent
            value = self.value2sfeel(value)
            if value is None:
                validData = False
            else:
                # Store the S-FEEL value for this item
                retVal = self.sfeel('{} <- {}'.format(item, value))
        if not validData:
            self.errors.append("Input variable '{!s}' has is invalid S-FEEL value '{!s}'".format(variable, data[variable]))
            status = {}
            status['errors'] = self.errors
            self.errors = []
            return (status, {})

        # Process each decision table in order
        allResults = []
        for (table, inputTests, decisionAnnotations) in self.decisions:
            if len(inputTests) > 0:
                doDecision = True
                for (variable, test) in inputTests:
                    item = self.glossary[variable]['item']
                    itemValue = self.sfeel('{}'.format(item))
                    retVal = self.sfeel('{}'.format(test))
                    if not retVal:
                        doDecision = False
                        break
                if not doDecision:
                    continue
                
            ranks = []
            foundRule = None
            rankedRules = []
            for thisRule in range(len(self.rules[table])):
                for i in range(len(self.rules[table][thisRule]['tests'])):
                    (variable, test, inputIndex) = self.rules[table][thisRule]['tests'][i]
                    item = self.glossary[variable]['item']
                    itemValue = self.sfeel('{}'.format(item))
                    value = self.value2sfeel(itemValue)
                    if self.decisionTables[table]['inputValidity'][inputIndex] is not None:
                        testValidity = self.decisionTables[table]['inputValidity'][inputIndex]
                        retVal = self.sfeel('{}'.format(testValidity))
                        if not retVal:
                            self.errors.append('Variable {!s} has S-FEEL input value {!s} which does not match input validity list {!s}'.format(item, value, testValidity))
                            status = {}
                            status['errors'] = self.errors
                            self.errors = []
                            return (status, {})
                    retVal = self.sfeel(str(test))
                    if not retVal:
                        break
                else:
                    # We found a hit
                    if self.decisionTables[table]['hitPolicy'] in ['U', 'A', 'F']:
                        foundRule = thisRule
                        break
                    elif self.decisionTables[table]['hitPolicy'][0] in ['R', 'C']:
                        rankedRules.append(thisRule)
                    elif self.decisionTables[table]['hitPolicy'] in ['P', 'O']:
                        # Rank the multiple outputs
                        if ranks == []:
                            theseRanks = []
                            for i in range(len(self.rules[table][thisRule]['outputs'])):
                                (variable, result, outputIndex, rank) = self.rules[table][thisRule]['outputs'][i]
                                if rank is None:
                                    item = self.glossary[variable]['item']
                                    thisResult = self.sfeel('{}'.format(item))
                                    result = self.value2sfeel(thisResult)
                                    if isinstance(result, str):
                                        if (result[0] == '"') and (result[-1] == '"'):
                                            result = result[1:-1]
                                    if result in self.decisionTables[table]['outputValidity'][outputIndex]:
                                        rank = self.decisionTables[table]['outputValidity'][outputIndex].index(result)
                                    else:
                                        try:
                                            if float(result) in self.decisionTables[table]['outputValidity'][outputIndex]:
                                                rank = self.decisionTables[table]['outputValidity'][outputIndex].index(float(result))
                                        except:
                                            rank = -1
                                theseRanks.append(rank)
                            theseRanks.append(thisRule)
                            ranks.append(theseRanks)
                        else:
                            before = None
                            beforeFound = False
                            for i in len(range(ranks)):
                                for i in range(len(self.rules[table][thisRule]['outputs'])):
                                    (variable, result, outputIndex, rank) = self.rules[table][thisRule]['outputs'][i]
                                    if rank is None:
                                        item = self.glossary[variable]['item']
                                        thisResult = self.sfeel('{}'.format(item))
                                        result = self.value2sfeel(thisResult)
                                        if isinstance(result, str):
                                            if (result[0] == '"') and (result[-1] == '"'):
                                                result = result[1:-1]
                                        if result in self.decisionTables[table]['outputValidity'][outputIndex]:
                                            rank = self.decisionTables[table]['outputValidity'][outputIndex].index(result)
                                        else:
                                            try:
                                                if float(result) in self.decisionTables[table]['outputValidity'][outputIndex]:
                                                    rank = self.decisionTables[table]['outputValidity'][outputIndex].index(float(result))
                                            except:
                                                rank = -1
                                    if rank < ranks[i]:
                                        before = i
                                        break
                                    elif rank > ranks[i]:
                                        beforeFound = True
                                        break
                                if beforeFound:
                                    break
                            theseRanks.append(thisRule)
                            if not beforeFound:
                                ranks.append(theseRanks)
                            else:
                                ranks.insert(i, theseRanks)
            newData = {}
            newData['Result'] = {}
            annotations = []
            for variable in self.glossary:
                item = self.glossary[variable]['item']
                thisResult = self.sfeel('{}'.format(item))
                if isinstance(thisResult, str):
                    if (thisResult[0] == '"') and (thisResult[-1] == '"'):
                        thisResult = thisResult[1:-1]
                newData['Result'][variable] = thisResult
            if self.decisionTables[table]['hitPolicy'] in ['U', 'A', 'F']:
                if foundRule is None:
                    self.errors.append("No rules matched the input data for decision table '{!s}'".format(table))
                    status = {}
                    status['errors'] = self.errors
                    self.errors = []
                    return (status, {})
                else:
                    for i in range(len(self.rules[table][foundRule]['outputs'])):
                        (variable, result, outputIndex, rank) = self.rules[table][foundRule]['outputs'][i]
                        # result is a string of valid FEEL tokens, but my not be an invalid expression
                        result = self.replaceItems(result)      # Replace BusinessConcept.Attribute references with actual values
                        sfeelText = self.data2sfeel(None, None, result, True)           # See if this is valid S-FEEL
                        if sfeelText is None:               # Not valid FEEL - make this a FEEL string
                            result = '"' + result + '"'
                        else:                               # Valid FEEL tokens - check if it is a value FEEL expression
                            (status, returnVal) = self.parser.sFeelParse(result)
                            if 'errors' in status:          # No - so make it a string
                                result = '"' + result + '"'
                        item = self.glossary[variable]['item']
                        retVal = self.sfeel('{} <- {}'.format(item, result))
                        thisResult = self.sfeel('{}'.format(item))
                        if isinstance(thisResult, str):
                            if (thisResult[0] == '"') and (thisResult[-1] == '"'):
                                thisResult = thisResult[1:-1]
                        if self.decisionTables[table]['outputValidity'][outputIndex] != []:
                            validList = self.decisionTables[table]['outputValidity'][outputIndex]
                            if thisResult not in validList:
                                self.errors.append("Variable '{!s}' has output value '{!s}' which does not match validity list '{!s}'".format(variable, repr(thisResult), repr(validList)))
                                status = {}
                                status['errors'] = self.errors
                                self.errors = []
                                return (status, {})
                        newData['Result'][variable] = thisResult
                    ruleId = (self.decisionTables[table]['name'], table, str(self.rules[table][foundRule]['ruleId']))
                    if 'annotation' in self.decisionTables[table]:
                        for annotation in range(len(self.decisionTables[table]['annotation'])):
                            name = self.decisionTables[table]['annotation'][annotation]
                            text = self.rules[table][foundRule]['annotation'][annotation]
                            annotations.append((name, text))
                newData['Executed Rule'] = ruleId
                if len(decisionAnnotations) > 0:
                    newData['DecisionAnnotations'] = decisionAnnotations
                if len(annotations) > 0:
                    newData['RuleAnnotations'] = annotations
            elif self.decisionTables[table]['hitPolicy'][0] in ['R', 'C']:
                if len(rankedRules) == 0:
                    self.errors.append("No rules matched the input data for decision table '{!s}'".format(table))
                    status = {}
                    status['errors'] = self.errors
                    self.errors = []
                    return (status, {})
                else:
                    foundRule = rankedRules[0]
                    first = True
                    for (variable, result, rank) in self.rules[table][foundRule]['outputs']:
                        # result is a string of valid FEEL tokens, but my not be an invalid expression
                        result = self.replaceItems(result)      # Replace BusinessConcept.Attribute references with actual values
                        sfeelText = self.data2sfeel(None, None, result, True)           # See if this is valid S-FEEL
                        if sfeelText is None:               # Not valid FEEL - make this a FEEL string
                            result = '"' + result + '"'
                        else:                               # Valid FEEL tokens - check if it is a value FEEL expression
                            (status, returnVal) = self.parser.sFeelParse(result)
                            if 'errors' in status:          # No - so make it a string
                                result = '"' + result + '"'
                        item = self.glossary[variable]['item']
                        if first:
                            if variable not in newData['Result']:
                                if len(self.decisionTables[table]['hitPolicy']) == 1:
                                    newData['Result'][variable] = []
                                elif self.decisionTables[table]['hitPolicy'][1] in ['+', '#']:
                                    newData['Result'][variable] = 0
                                else:
                                    newData['Result'][item] = None
                            first = False
                        retVal = self.sfeel('{} <- {}'.format(item, result))
                        thisOutput = self.sfeel('{}'.format(item))
                        if isinstance(thisOutput, str):
                            if (thisOutput[0] == '"') and (thisOutput[-1] == '"'):
                                thisOutput = thisOutput[1:-1]
                        if len(self.decisionTables[table]['hitPolicy']) == 1:
                            newData['Result'][variable].append(thisOutput)
                        elif self.decisionTables[table]['hitPolicy'][1] == '+':
                            newData['Result'][variable] += thisOutput
                        elif self.decisionTables[table]['hitPolicy'][1] == '<':
                            if newData['Result'][variable] is None:
                                newData['Result'][variable] = thisOutput
                            elif thisOutput < newData['Result'][variable]:
                                newData['Result'][variable] = thisOutput
                        elif self.decisionTables[table]['hitPolicy'][1] == '>':
                            if newData['Result'][variable] is None:
                                newData['Result'][variable] = thisOutput
                            elif thisOutput > newData['Result'][variable]:
                                newData['Result'][variable] = thisOutput
                        else:
                            newData['Result'][variable] += 1
                    ruleId = (self.decisionTables[table]['name'], table, str(self.rules[table][foundRule]['ruleId']))
                    if 'annotation' in self.decisionTables[table]:
                        for annotation in range(len(self.decisionTables[table]['annotation'])):
                            name = self.decisionTables[table]['annotation'][annotation]
                            text = self.rules[table][foundRule]['annotation'][annotation]
                            annotations.append((name, text))
                    newData['Executed Rule'] = ruleId
                    if len(decisionAnnotations) > 0:
                        newData['DecisionAnnotations'] = decisionAnnotations
                    if len(annotations) > 0:
                        newData['RuleAnnotations'] = annotations
            elif self.decisionTables[table]['hitPolicy'][0] == 'P':
                if len(ranks) == 0:
                    self.errors.append("No rules matched the input data for decision table '{!s}'".format(table))
                    status = {}
                    status['errors'] = self.errors
                    self.errors = []
                    return (status, {})
                else:
                    foundRule = ranks[0][-1]
                    for (variable, result, rank) in self.rules[table][foundRule]['outputs']:
                        # result is a string of valid FEEL tokens, but my not be an invalid expression
                        result = self.replaceItems(result)      # Replace BusinessConcept.Attribute references with actual values
                        sfeelText = self.data2sfeel(None, None, result, True)           # See if this is valid S-FEEL
                        if sfeelText is None:               # Not valid FEEL - make this a FEEL string
                            result = '"' + result + '"'
                        else:                               # Valid FEEL tokens - check if it is a value FEEL expression
                            (status, returnVal) = self.parser.sFeelParse(result)
                            if 'errors' in status:          # No - so make it a string
                                result = '"' + result + '"'
                        item = self.glossary[variable]['item']
                        retVal = self.sfeel('{} <- {}'.format(item, result))
                        thisResult = self.sfeel('{}'.format(item))
                        if isinstance(thisResult, str):
                            if (thisResult[0] == '"') and (thisResult[-1] == '"'):
                                thisResult = thisResult[1:-1]
                        newData['Result'][variable] = thisResult
                    ruleId = (self.decisionTables[table]['name'], table, str(self.rules[table][foundRule]['ruleId']))
                    if 'annotation' in self.decisionTables[table]:
                        for annotation in range(len(self.decisionTables[table]['annotation'])):
                            name = self.decisionTables[table]['annotation'][annotation]
                            text = self.rules[table][foundRule]['annotation'][annotation]
                            annotations.append((name, text))
                    newData['Executed Rule'] = ruleId
                    if len(decisionAnnotations) > 0:
                        newData['DecisionAnnotations'] = decisionAnnotations
                    if len(annotations) > 0:
                        newData['RuleAnnotations'] = annotations
            elif self.decisionTables[table]['hitPolicy'][0] == 'O':
                if len(ranks) == 0:
                    self.errors.append("No rules matched the input data for decision table '{!s}'".format(table))
                    status = {}
                    status['errors'] = self.errors
                    self.errors = []
                    return (status, {})
                else:
                    ruleIds = []
                    haveAnnotations = False
                    for i in range(len(ranks)):
                        annotations.append([])
                        foundRule = ranks[i][-1]
                        for (variable, result, rank) in self.rules[table][foundRule]['outputs']:
                            # result is a string of valid FEEL tokens, but my not be an invalid expression
                            result = self.replaceItems(result)      # Replace BusinessConcept.Attribute references with actual values
                            sfeelText = self.data2sfeel(None, None, result, True)           # See if this is valid S-FEEL
                            if sfeelText is None:               # Not valid FEEL - make this a FEEL string
                                result = '"' + result + '"'
                            else:                               # Valid FEEL tokens - check if it is a value FEEL expression
                                (status, returnVal) = self.parser.sFeelParse(result)
                                if 'errors' in status:          # No - so make it a string
                                    result = '"' + result + '"'
                            item = self.glossary[variable]['item']
                            if item not in newData['Result']:
                                newData['Result'][variable] = []
                            retVal = self.sfeel('{} <- {}'.format(item, result))
                            thisResult = self.sfeel('{}'.format(item))
                            if isinstance(thisResult, str):
                                if (thisResult[0] == '"') and (thisResult[-1] == '"'):
                                    thisResult = thisResult[1:-1]
                            newData['Result'][variable].append(thisResult)
                        ruleIds.append(self.decisionTables[table]['name'], table + ':' + str(self.rules[table][foundRule]['ruleId']))
                        if 'annotation' in self.decisionTables[table]:
                            for annotation in range(len(self.decisionTables[table]['annotation'])):
                                name = self.decisionTables[table]['annotation'][annotation]
                                text = self.rules[table][foundRule]['annotation'][annotation]
                                annotations[i].append((name, text))
                                haveAnnotations = True
                    newData['Executed Rule'] = ruleIds
                    if len(decisionAnnotations) > 0:
                        newData['DecisionAnnotations'] = decisionAnnotations
                    if haveAnnotations:
                        newData['RuleAnnotations'] = annotations

            allResults.append(newData)

        status = {}
        if len(self.errors) > 0:
            status['errors'] = self.errors
            self.errors = []
        if len(allResults) == 1:
            return (status, newData)
        else:
            return (status, allResults)


    def loadTest(self, testBook):
        """
        Load a workbook that contains a Test worksheet, which defines the unit test for pyDMNrules.test()

        This routine loads an Excel workbook which must contain a 'Test' sheet

        Args:
            param1 (str): The name of the Excel workbook (including path if it is not in the current working directory

        Returns:
            dict: status

            'status' is a dictionary of different status information.
            Currently only status['error'] is implemented.
            If the key 'error' is present in the status dictionary,
            then loadTest() encountered one or more errors and status['error'] is the list of those errors

        """

        self.errors = []
        try:
            wb = load_workbook(filename=testBook)
        except Exception as e:
            self.errors.append("No readable workbook named '{!s}'!".format(testBook))
            status = {}
            status['errors'] = self.errors
            return status
        return self.useTest(wb)

    def useTest(self, workbook):
        """
        Use a workbookook that contains a Test worksheet, which defines the unit test for pyDMNrules.test()

        This routine uses an already loaded Excel workbook which must contain a 'Test' sheet

        Args:
            param1 (openpyxl.workbook): An openpyxl workbook (either loaded with openpyxl or created using openpyxl)

        Returns:
            dict: status

            'status' is a dictionary of different status information.
            Currently only status['error'] is implemented.
            If the key 'error' is present in the status dictionary,
            then load() encountered one or more errors and status['error'] is the list of those errors

        """

        self.errors = []
        if not isinstance(workbook, openpyxl.Workbook):
            self.errors.append("workbook is not a valid openpyxl workbook")
            status = {}
            status['errors'] = self.errors
            return status

        if 'Test' not in workbook:
            self.errors.append("No 'Test' sheet the workbook")
            status = {}
            status['errors'] = self.errors
            return status

        self.wbt = workbook
        self.testIsLoaded = True


    def test(self):
        """
        Run the test data through the decision

        This routine reads Unit Test data and a set of test (DMNrulesTests) from the 'Test' worksheet
        and runs the specified test data through the decide() function.
        Any descrepancies between the returned data and the expected data (as configured in DMNrulesTests)
        will be returned as a list of mismatches.

        Args:
            None: The spreadsheet 'Test' must exist in the load Excel workbook.

        Returns:
            tuple: (testStatus, results)

            testStatus is a list of dictionaries - being the 'status' returned by decide() for each test.
                If the 'status' returned from decide() contains the key 'error', then the returned newData will
                not be checked for mismatches.

            results is a list of dictionaries, one for each test in the 'DMNrulesTests' table.
                The keys to this dictionary are
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
                      matched the values in the 'DMNrulesTests' table.
        """

        if not self.testIsLoaded:
            self.errors.append("No 'Test' sheet has been loaded")
            status = {}
            status['errors'] = self.errors
            self.errors = []
            return (status, {})

        # Read in the Test worksheet
        try:
            ws = self.wbt['Test']
        except (KeyError):
            self.errors.append('No rulesBook sheet named Test!')
            status = {}
            status['errors'] = self.errors
            return status

        # Now search for the unit test data
        self.mergedCells = ws.merged_cells.ranges
        tests = {}
        testData = {}
        parsedRanges = []
        testsCell = None
        for row in ws.rows:
            for cell in row:
                for i in range(len(parsedRanges)):
                    parsed = parsedRanges[i]
                    if cell.coordinate in parsed:
                        continue
                # Skip the DMNrulesTests table if we have found it already
                if testsCell is not None:
                    if (cell.row >= testsRow) and (cell.row < testsRow + testsRows) and (cell.column >= testsCol) and (cell.column < testsCol + testsCols):
                        continue
                thisCell = cell.value
                coordinate = cell.coordinate
                if isinstance(thisCell, str):
                    # See if this is a 
                    (rows, cols) = self.tableSize(cell)
                    if (rows == 1) and (cols == 0):
                        continue
                    # Check if this is a unit test data table
                    if thisCell in self.glossaryConcepts:
                        # Parse a table of unit test data - the name of the table is a Glossary concept
                        concept = thisCell
                        inputColumns = 0
                        testData[concept] = {}
                        testData[concept]['heading'] = []       # List of headings
                        testData[concept]['unitData'] = []      # List of rows of unit data
                        testData[concept]['annotations'] = []   # List of rows of annotations
                        doingInputs = True
                        # Parse the horizontal heading for the variables
                        for thisCol in range(cols):
                            thisCell = cell.offset(row=1, column=thisCol).value
                            coordinate = cell.offset(row=1, column=thisCol).coordinate
                            if thisCell is None:
                                if doingInputs:
                                    self.errors.append("Missing Input heading in table at '{!s}' on sheet 'Test'".format(coordinate))
                                else:
                                    self.errors.append("Missing Annotation heading in table at '{!s}' on sheet 'Test'".format(coordinate))
                                status = {}
                                status['errors'] = self.errors
                                self.errors = []
                                return (status, {})
                            thisCell = str(thisCell).strip()
                            if doingInputs:
                                # Check that all the headings are in the Glossary
                                if thisCell not in self.glossary:
                                    self.errors.append("Input heading '{!s}' in table at '{!s}' on sheet 'Test' is not in the Glossary".format(thisCell, coordinate))
                                    status = {}
                                    status['errors'] = self.errors
                                    self.errors = []
                                    return (status, {})
                                # And that they belong to this Business Concept
                                if thisCell not in self.glossaryConcepts[concept]:
                                    if doingInputs:
                                        self.errors.append("Input heading '{!s}' in table at '{!s}' on sheet 'Test' is in the Glossary, but not in a Business Concept".format(thisCell, coordinate))
                                    status = {}
                                    status['errors'] = self.errors
                                    self.errors = []
                                    return (status, {})
                            # Save this heading - for inputs this is the variable for this column
                            testData[concept]['heading'].append(thisCell)
                            if doingInputs:
                                inputColumns += 1
                                border = cell.offset(row=1, column=thisCol).border
                                if border.right.style == 'double':
                                    doingInputs = False
                                border = cell.offset(row=1, column=thisCol + 1).border
                                if border.left.style == 'double':
                                    doingInputs = False

                        # Store the unit test data
                        for thisRow in range(2, rows):
                            testData[concept]['unitData'].append([])        # another row of unit test data
                            testData[concept]['annotations'].append([])     # another row of annotations for this row
                            for thisCol in range(cols):
                                heading = testData[concept]['heading'][thisCol]
                                thisCell = cell.offset(row=thisRow, column=thisCol).value
                                coordinate = cell.offset(row=thisRow, column=thisCol).coordinate
                                if thisCol < inputColumns:
                                    if thisCell is not None:
                                        if thisCell == 'true':
                                            value = True
                                        elif thisCell == 'false':
                                            value = False
                                        elif thisCell == 'null':
                                            value = None
                                        elif isinstance(thisCell, str):
                                            if ((thisCell[0] == '[') and (thisCell[-1] == ']')) or ((thisCell[0] == '{') and (thisCell[-1] == '}')):
                                                value = self.data2sfeel(coordinate, 'Test', thisCell, False)
                                                try:
                                                    value = eval(value)
                                                except Exception as e:
                                                    value = None
                                            elif (thisCell[0] == '"') and (thisCell[-1] == '"'):
                                                value = thisCell[1:-1].strip()
                                            else:
                                                value = thisCell.strip()
                                        else:
                                            value = thisCell
                                        testData[concept]['unitData'][thisRow - 2].append((heading, value))
                                elif thisCell is not None:
                                    testData[concept]['annotations'][thisRow - 2].append((heading, thisCell))

                    elif thisCell == 'DMNrulesTests':
                        testsCell = cell
                        testsRow = cell.row
                        testsCol = cell.column
                        testsRows = rows
                        testsCols = cols
                        continue
                    
                    # Symbolically merge all the cells in this test data table
                    thisRow = cell.row
                    thisCol = cell.column
                    ws.merge_cells(start_row=thisRow, start_column=thisCol,
                                   end_row=thisRow + rows - 1, end_column=thisCol + cols - 1)
                    # Find this merge range
                    for thisMerged in self.mergedCells:
                        if (thisMerged.min_col == cell.column) and (thisMerged.min_row == cell.row):
                            if thisMerged.max_col != (cell.column + cols - 1):
                                continue
                            if thisMerged.max_row != (cell.row + rows - 1):
                                continue
                            # Mark it as parsed
                            parsedRanges.append(thisMerged)

        # Now parse the tests configuration
        if testsCell is None:
            self.errors.append("No table 'DMNrulesTests' in 'Test' worksheet in rules book")
            status = {}
            status['errors'] = self.errors
            self.errors = []
            return (status, {})

        # Parse a table of tests Configuration
        cell = testsCell
        table = 'DMNrulesTests'
        rows = testsRows
        cols = testsCols
        if (rows == 1) and (cols == 0):
            self.errors.append("Table 'DMNrulesTest' in 'Test' is empty")
            status = {}
            status['errors'] = self.errors
            self.errors = []
            return (status, {})
        inputColumns = outputColumns = 0
        tests['headings'] = []      # The horizontal heading (concepts, variables, annotation)
        tests['inputColumns'] = []  # Rows of concept indexes
        tests['outputColumns'] = [] # Rows of variable output data
        tests['annotations'] = []   # The annotations for this test
        doingInputs = True
        doingAnnotation = False
        # Collect up all the headings and check that they are all valid
        for thisCol in range(cols):
            thisCell = cell.offset(row=1, column=thisCol).value
            coordinate = cell.offset(row=1, column=thisCol).coordinate
            if thisCell is None:
                if doingInputs:
                    self.errors.append("Missing Input heading in table '{!s}' at '{!s}' on sheet 'Test'".format(table, coordinate))
                if not doingAnnotation:
                    self.errors.append("Missing Output heading in table '{!s}' at '{!s}' on sheet 'Test'".format(table, coordinate))
                else:
                    self.errors.append("Missing Annotation heading in table '{!s}' at '{!s}' on sheet 'Test'".format(table, coordinate))
                status = {}
                status['errors'] = self.errors
                self.errors = []
                return (status, {})
            thisCell = str(thisCell).strip()
            # Check that the input and output headings are in the Glossary
            if doingInputs:
                if thisCell not in self.glossaryConcepts:
                    self.errors.append("Input heading '{!s}' in table '{!s}' in sheet 'Test' at '{!s}' on sheet 'Test' is not a Business Concept in the Glossary".format(thisCell, table, coordinate))
                    status = {}
                    status['errors'] = self.errors
                    self.errors = []
                    return (status, {})
                # Check that we have a table of unit test data for this concept
                if thisCell not in testData:
                    self.errors.append("No configured unit test data for Business Concept [heading '{!s}'] in table '{!s}' at '{!s}' on sheet 'Test'".format(thisCell, table, coordinate))
                    status = {}
                    status['errors'] = self.errors
                    self.errors = []
                    return (status, {})
            elif not doingAnnotation:
                if thisCell not in self.glossary:
                    self.errors.append("Output heading '{!s}' in table '{!s}' at '{!s}' on sheet 'Test' is not in the Glossary".format(thisCell, table, coordinate))
                    status = {}
                    status['errors'] = self.errors
                    self.errors = []
                    return (status, {})
            tests['headings'].append(thisCell)      # Save the heading
            if doingInputs:
                inputColumns += 1
                border = cell.offset(row=1, column=thisCol).border
                if border.right.style == 'double':
                    doingInputs = False
                border = cell.offset(row=1, column=thisCol + 1).border
                if border.left.style == 'double':
                    doingInputs = False
            elif not doingAnnotation:
                outputColumns += 1
                border = cell.offset(row=1, column=thisCol).border
                if border.right.style == 'double':
                    doingAnnotation = True
                border = cell.offset(row=1, column=thisCol + 1).border
                if border.left.style == 'double':
                    doingAnnotation = True

        # Check that we did find at least one output column
        if doingInputs:
            self.errors.append("No column of excected output data in table '{!s}'".format(table))
            status = {}
            status['errors'] = self.errors
            self.errors = []
            return (status, {})

        # Store the configuration for each test
        for thisRow in range(2, rows):
            thisTest = thisRow - 2
            tests['inputColumns'].append([])        # Horizontal tuples of (concept, index)
            tests['outputColumns'].append([])       # Horizontal tuples of (variable, data)
            tests['annotations'].append([])         # Horizonal array of annotations
            for thisCol in range(cols):
                heading = tests['headings'][thisCol]
                thisCell = cell.offset(row=thisRow, column=thisCol).value
                coordinate = cell.offset(row=thisRow, column=thisCol).coordinate
                # Only annotations can be blank
                if (thisCell is None) and (thisCol < inputColumns + outputColumns):
                    if thisCol < inputColumns:
                        self.errors.append("Missing input index in table '{!s}' at '{!s}' on sheet 'Test'".format(table, coordinate))
                    else:
                        self.errors.append("Missing output data in table '{!s}' at '{!s}' on sheet 'Test'".format(table, coordinate))
                    status = {}
                    status['errors'] = self.errors
                    self.errors = []
                    return (status, {})
                if thisCol < inputColumns:
                    try:
                        thisIndex = int(thisCell)
                    except:
                        self.errors.append("Invalid input index '{!s}' in table '{!s}' at '{!s}' on sheet 'Test'".format(thisCell, table, coordinate))
                        status = {}
                        status['errors'] = self.errors
                        self.errors = []
                        return (status, {})
                    if (thisIndex < 1) or (thisIndex > len(testData[heading]['unitData'])):
                        self.errors.append("Invalid input index '{!s}' in table '{!s}' at '{!s}' on sheet 'Test'".format(thisCell, table, coordinate))
                        status = {}
                        status['errors'] = self.errors
                        self.errors = []
                        return (status, {})
                    tests['inputColumns'][thisTest].append((heading, thisIndex))
                elif thisCol < inputColumns + outputColumns:
                    if thisCell == 'true':
                        value = True
                    elif thisCell == 'false':
                        value = False
                    elif thisCell == 'null':
                        value = None
                    elif isinstance(thisCell, str):
                        if ((thisCell[0] == '[') and (thisCell[-1] == ']')) or ((thisCell[0] == '{') and (thisCell[-1] == '}')):
                            value = self.data2sfeel(coordinate, 'Test', thisCell, False)
                            try:
                                value = eval(value)
                            except Exception as e:
                                value = None
                        elif (thisCell[0] == '"') and (thisCell[-1] == '"'):
                            value = thisCell[1:-1].strip()
                        else:
                            value = thisCell.strip()
                    else:
                        value = thisCell
                    tests['outputColumns'][thisTest].append((heading, value))
                    # tests['outputColumns'][thisTest].append((heading, thisCell))
                elif thisCell is not None:
                    tests['annotations'][thisTest].append((heading, thisCell))

        # Now run the tests
        results = []
        testStatus = []
        for thisTest in range(len(tests['inputColumns'])):
            results.append({})
            results[thisTest]['Test ID'] = thisTest + 1
            if len(tests['annotations'][thisTest]) > 0:
                results[thisTest]['TestAnnotations'] = tests['annotations'][thisTest]
            data = {}
            dataAnnotations = []
            for inputCol in range(inputColumns):
                (concept, thisIndex) = tests['inputColumns'][thisTest][inputCol]
                for thisData in range(len(testData[concept]['unitData'][thisIndex - 1])):
                    (variable, value) = testData[concept]['unitData'][thisIndex - 1][thisData]
                    data[variable] = value
                if len(testData[concept]['annotations'][thisIndex - 1]) > 0:
                   dataAnnotations.append(testData[concept]['annotations'][thisIndex - 1])
            results[thisTest]['data'] = data
            (status, newData) = self.decide(data)
            if isinstance(newData, list):
                newData = newData[-1]
            results[thisTest]['newData'] = newData
            results[thisTest]['status'] = status
            if len(dataAnnotations) > 0:
                results[thisTest]['DataAnnotations'] = dataAnnotations
            testStatus.append(status)
            if 'errors' in status:
                continue
            mismatches = []
            for outputCol in range(outputColumns):
                (heading, expected) = tests['outputColumns'][thisTest][outputCol]
                if heading not in newData['Result']:
                    mismatches.append("Variable '{!s}' not returned in newData['Result']{}".format(heading, '{}'))
                elif newData['Result'][heading] != expected:
                    mismatches.append("Mismatch: Variable '{!s}' returned '{!s}' but '{!s}' was expected".format(heading, newData['Result'][heading], expected))
            if len(mismatches) > 0:
                results[thisTest]['Mismatches'] = mismatches
        return(testStatus, results)


if __name__ == '__main__':

    dmnRules = DMN()
    status = dmnRules.load('Example1.xlsx')
    if 'errors' in status:
        print('Example1.xlsx has errors', status['errors'])
        sys.exit(0)
    else:
        print('Example1.xlsx loaded')

    data = {}
    data['Applicant Age'] = 63
    data['Medical History'] = 'bad'
    print('Testing',repr(data))
    (status, newData) = dmnRules.decide(data)
    print('Decision',repr(newData))
    if 'errors' in status:
        print('With errors', status['errors'])

    data['Applicant Age'] = 33
    data['Medical History'] = None
    print('Testing',repr(data))
    (status, newData) = dmnRules.decide(data)
    print('Decision',repr(newData))
    if 'errors' in status:
        print('With errors', status['errors'])

    data['Applicant Age'] = 13
    data['Medical History'] = 'good'
    print('Testing',repr(data))
    (status, newData) = dmnRules.decide(data)
    print('Decision',repr(newData))
    if 'errors' in status:
        print('With errors', status['errors'])

    data['Applicant Age'] = 13
    data['Medical History'] = 'bad'
    print('Testing',repr(data))
    (status, newData) = dmnRules.decide(data)
    print('Decision',repr(newData))
    if 'errors' in status:
        print('With errors', status['errors'])

    status = dmnRules.load('ExampleRows.xlsx')
    if 'errors' in status:
        print('ExampleRows.xlsx has errors', status['errors'])
        sys.exit(0)
    else:
        print('ExampleRows.xlsx loaded')

    data = {}
    data['Customer'] = 'Private'
    data['OrderSize'] = 9
    data['Delivery'] = 'slow'
    print('Testing',repr(data))
    (status, newData) = dmnRules.decide(data)
    print('Decision',repr(newData))
    if 'errors' in status:
        print('With errors', status['errors'])

    status = dmnRules.load('ExampleColumns.xlsx')
    if 'errors' in status:
        print('ExampleColumns.xlsx has errors', status['errors'])
        sys.exit(0)
    else:
        print('ExampleColumns.xlsx loaded')

    data = {}
    data['Customer'] = 'Private'
    data['OrderSize'] = 9
    data['Delivery'] = 'slow'
    print('Testing',repr(data))
    (status, newData) = dmnRules.decide(data)
    print('Decision',repr(newData))
    if 'errors' in status:
        print('With errors', status['errors'])

    status = dmnRules.load('ExampleCrosstab.xlsx')
    if 'errors' in status:
        print('ExampleCrosstab.xlsx has errors', status['errors'])
        sys.exit(0)
    else:
        print('ExampleCrosstab.xlsx loaded')

    data = {}
    data['Customer'] = 'Private'
    data['OrderSize'] = 9
    data['Delivery'] = 'slow'
    print('Testing',repr(data))
    (status, newData) = dmnRules.decide(data)
    print('Decision',repr(newData))
    if 'errors' in status:
        print('With errors', status['errors'])
