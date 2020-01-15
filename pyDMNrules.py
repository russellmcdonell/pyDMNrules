# -----------------------------------------------------------------------------
# pyDMNrules.py
# -----------------------------------------------------------------------------

import sys
import re
import datetime
import pySFeel
from openpyxl import load_workbook
from openpyxl import utils

class DMN():


    def __init__(self):
        self.lexer = pySFeel.SFeelLexer()
        self.parser = pySFeel.SFeelParser()
        self.glossaryLoaded = False
        self.isLoaded = False
        self.errors = []
        self.warnings = []


    def sfeel(self, text):
        (status, retVal) = self.parser.sFeelParse(text)
        if 'errors' in status:
            self.errors += status['errors']
        return retVal


    def data2sfeel(self, coordinate, data):
        if data in self.glossary:
            return self.glossary[data]['item']

        isError = False
        tokens = self.lexer.tokenize(data)
        yaccTokens = []
        for token in tokens:
            if token.type == 'ERROR':
                if not isError:
                    self.errors.append('S-FEEL syntax error at {!s}:{!s}'.format(coordinate, token.value))
            else:
                yaccTokens.append(token)
        thisData = ''
        for token in yaccTokens:
            if thisData != '':
                thisData += ' '
            if token.type != 'NAME':
                thisData += token.value
            elif token.value in self.glossaryItems:
                thisData += token.value
            else:
                thisData += '"' + token.value + '"'
        return thisData


    def test2sfeel(self, variable, coordinate, test):
        thisTest = test.strip()
        isNot = False
        if thisTest.startswith('not '):
            isNot = True
            thisTest = thisTest[4:].strip()
        isIn = False
        if thisTest[:2] == 'in':
            tmp = thisTest[2:].strip()
            if (tmp[1:] == '(') and (tmp[:-1] == ')'):
                # In function
                isIn = True
                thisTest = tmp[1:-1]
        if (thisTest[0] not in ['[', '(']) or (thisTest[-1] not in [']', ')']):
            # Not a list or range
            commaAt = thisTest.find(',')
            if commaAt == -1:
                # Not an list - should be an S-FEEL simple expression
                # Could be a string constant, but missing surrounding double quotes
                relOp = ''
                if thisTest[:2] in ['<=', '>=', '!=']:
                    relOp = thisTest[:2]
                    thisTest = thisTest[2:]
                elif thisTest[:1] in ['<', '>', '=']:
                    relOp = thisTest[:1]
                    thisTest = thisTest[1:]
                thisTest = self.data2sfeel(coordinate, thisTest)
                if isNot:
                    if isIn:
                        if relOp != '':
                            return variable + ' not in(' + relOp + ' ' + thisTest + ')'
                        else:
                            return variable + ' not in(' + thisTest + ')'
                    else:
                        if relOp != '':
                            return variable + ' not ' +relOp + ' ' + thisTest
                        else:
                            return variable + ' not ' + thisTest
                else:
                    if isIn:
                        if relOp != '':
                            return variable + ' in(' + relOp + ' ' + thisTest + ')'
                        else:
                            return variable + ' in(' + thisTest + ')'
                    else:
                        if relOp != '':
                            return variable + ' ' + relOp + ' ' + thisTest
                        else:
                            return variable + ' = ' + thisTest
            else:   # an unbracketed list
                # Either for an in() function, or an implied in() function
                theseTests = thisTest.split(',')
                for i in range(len(theseTests)):
                    aTest = theseTests[i].strip()
                    # Should be an S-FEEL simple expression
                    # Could be a string constant, but missing surrounding double quotes
                    theseTests[i] = self.data2sfeel(coordinate, aTest)
                thisTest = ','.join(theseTests)
                if isNot:
                    return variable + ' not in(' + thisTest + ')'
                else:
                    return variable + ' in(' + thisTest + ')'
        else:   # a list or range
            openBracket = thisTest[0]
            closeBracket = thisTest[-1]
            thisTest = thisTest[1:-1]
            elipseAt = thisTest.find('..')
            if elipseAt == -1:
                # Not a range, see if it's a list
                if (openBracket != '[') or (closeBracket != ']'):
                    # Neither - better be a valid bracketed S-FEEL expression
                    thisTest = self.data2sfeel(coordinate, thisTest)
                else:
                    theseTests = thisTest.split(',')
                    for i in range(len(theseTests)):
                        aTest = theseTests[i].strip()
                        # Should be an S-FEEL simple expression
                        # Could be a string constant, but missing surrounding double quotes
                        theseTests[i] = self.data2sfeel(coordinate, aTest)
                    thisTest = ','.join(theseTests)
                thisTest = openBracket + thisTest + closeBracket
                if isNot:
                    if isIn:
                        return variable + ' not in(' + thisTest + ')'
                    else:
                        return variable + ' not ' + thisTest
                else:
                    if isIn:
                        return variable + ' in(' + thisTest + ')'
                    else:
                        return variable + ' = ' + thisTest
            else:   # A range
                theseTests = thisTest.split('..')
                if len(theseTests) != 2:
                    self.errors.append('S-FEEL syntax error at {!s}:invalid range syntax'.format(coordinate))
                    return variable + ' = "' + thisTest + '"'
                for i in range(len(theseTests)):
                    aTest = theseTests[i].strip()
                    # what's left should be an S-FEEL simple expression
                    # Could be a string constant, but missing surrounding double quotes
                    theseTests[i] = self.data2sfeel(coordinate, aTest)
                thisTest = openBracket + '..'.join(theseTests) + closeBracket
                if isNot:
                    if isIn:
                        return variable + ' not in(' + thisTest + ')'
                    else:
                        return variable + ' not in ' + thisTest
                else:
                    if isIn:
                        return variable + ' in(' + thisTest + ')'
                    else:
                        return variable + ' in ' + thisTest


    def value2sfeel(self, value):
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
            else:
                return '"' + value + '"'
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
            hour = duration % 24
            days = int(duration / 24)
            return 'P%dDT%dH%dM%dS' % (days, hours, mins, secs)
        else:
            self.errors.append("Invalid Data '{!r}' - not a valid S-FEEL data type".format(value))
            return None


    def result2sfeel(self, variable, coordinate, result):
        thisResult = result.strip()
        thisResult = self.data2sfeel(coordinate, thisResult)
        (status, retVal) = self.parser.sFeelParse(thisResult)
        if 'errors' in status:
            self.errors.append("Invalid Output value '{!r}' at '{!s}'".format(result, coordinate))
            self.errors += status['errors']
        return thisResult


    def tableSize(self, cell):
        '''
        Determine the size of a table
        '''
        rows = 1
        cols = 0
        while cell.offset(row=rows).value is not None:
            coordinate = cell.offset(row=rows).coordinate
            for merged in self.mergedCells:
                if coordinate in merged:
                    rows += merged.max_row - merged.min_row + 1
                    break
            else:
                rows += 1
        while cell.offset(row=1, column=cols).value is not None:
            coordinate = cell.offset(row=1, column=cols).coordinate
            for merged in self.mergedCells:
                if coordinate in merged:
                    cols += merged.max_col - merged.min_col + 1
                    break
            else:
                cols += 1
        return (rows, cols)



    def parseDecionTable(self, cell):
        '''
        Parse a Decision Table
        '''
        (rows, cols) = self.tableSize(cell)
        startRow = cell.row
        startCol = cell.column
        table = cell.value
        coordinate = cell.coordinate
        table = str(table).strip()
        # print("Parsing Decision Table '{!s}' at '{!s}'".format(table, coordinate))
        self.rules[table] = []
        # Check the next cell down to determine the decision table layout
        thisCell = cell.offset(row=1).value
        nextCell = cell.offset(row=2).value
        if nextCell is not None:
            nextCell = str(nextCell).strip()
        if thisCell is None:
            # Empty table
            self.errors.append("Decision table '{!s}' at {!s}' is empty".format(table,coordinate))
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
            self.decisionTables[table]['outputColumns'] = []
            for thisCol in range(cols):
                thisCell = cell.offset(row=1, column=thisCol).value
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
                    border = cell.offset(row=1, column=thisCol).border
                    if border.bottom.style != 'double':
                        border = cell.offset(row=2, column=thisCol).border
                        if border.top.style != 'double':
                            doingValidity = True
                    doingInputs = True
                    continue
                if thisCell is None:
                    if doingInputs:
                        self.errors.append("Missing Input heading in table '{!s}' at '{!s}'".format(table, coordinate))
                    elif not doingAnnotation:
                        self.errors.append("Missing Output heading in table '{!s}' at '{!s}'".format(table, coordinate))
                    else:
                        self.errors.append("Missing Annotation heading in table '{!s}' at '{!s}'".format(table, coordinate))
                    return (rows, cols, -1)
                thisCell = str(thisCell).strip()
                coordinate = cell.offset(row=1, column=thisCol).coordinate
                if not doingAnnotation:
                    # Check that all the headings are in the Glossary
                    if thisCell not in self.glossary:
                        if doingInputs:
                            self.errors.append("Input heading '{!s}' in table '{!s}' at '{!s}' is not in the Glossary".format(thisCell, table, coordinate))
                        elif not doingAnnotation:
                            self.errors.append("Output heading '{!s}' in table '{!s}' at '{!s}' is not in the Glossary".format(thisCell, table, coordinate))
                        else:
                            self.errors.append("Annotation heading '{!s}' in table '{!s}' at '{!s}' is not in the Glossary".format(thisCell, table, coordinate))
                        return (rows, cols, -1)
                if doingInputs:
                    inputColumns += 1
                    thisInput = len(self.decisionTables[table]['inputColumns'])
                    self.decisionTables[table]['inputColumns'].append({})
                    self.decisionTables[table]['inputColumns'][thisInput]['name'] = thisCell
                    border = cell.offset(row=1, column=thisCol).border
                    if border.right.style == 'double':
                        doingInputs = False
                    border = cell.offset(row=1, column=thisCol + 1).border
                    if border.left.style == 'double':
                        doingInputs = False
                elif not doingAnnotation:
                    outputColumns += 1
                    thisOutput = len(self.decisionTables[table]['outputColumns'])
                    self.decisionTables[table]['outputColumns'].append({})
                    self.decisionTables[table]['outputColumns'][thisOutput]['name'] = thisCell
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

            if doingInputs:
                self.errors.append("No Output column in table '{!s}' - missing double bar vertical border".format(table))
                return (rows, cols, -1)
            rulesRow = 2
            if doingValidity:
                # Parse the validity row
                doingInputs = True
                ranksFound = False
                for thisCol in range(1, cols):
                    thisCell = cell.offset(row=2, column=thisCol).value
                    if thisCell is None:
                        thisCol += 1
                        continue
                    ranksFound = True
                    coordinate = cell.offset(row=2, column=thisCol).coordinate
                    if thisCol <= inputColumns:
                        name = self.decisionTables[table]['inputColumns'][thisCol - 1]['name']
                        variable = self.glossary[name]['item']
                        self.decisionTables[table]['inputColumns'][thisCol - 1]['validity'] = []
                        if not isinstance(thisCell, str):
                            self.decisionTables[table]['inputColumns'][thisCol - 1]['validity'].append(thisCell)
                        else:
                            validityTests = thisCell.split(',')
                            for validTest in validityTests:
                                thisTest = validTest.strip()
                                try:
                                    self.decisionTables[table]['inputColumns'][thisCol - 1]['validity'].append(float(thisTest))
                                except:
                                    self.decisionTables[table]['inputColumns'][thisCol - 1]['validity'].append(thisTest)
                    elif thisCol <= inputColumns + outputColumns:
                        name = self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]['name']
                        variable = self.glossary[name]['item']
                        self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]['validity'] = []
                        if not isinstance(thisCell, str):
                            self.decisionTables[table]['outputColumns'][thisCol - inputcolumns - 1]['validity'].append(thisCell)
                        else:
                            validityTests = thisCell.split(',')
                            for validTest in validityTests:
                                thisTest = validTest.strip()
                                try:
                                    self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]['validity'].append(float(thisTest))
                                except:
                                    self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]['validity'].append(thisTest)
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
                    if thisCol == 0:
                        thisCell = str(thisCell).strip()
                        self.rules[table][thisRule]['ruleId'] = thisCell
                        continue
                    coordinate = cell.offset(row=thisRow, column=thisCol).coordinate
                    if thisCol <= inputColumns:
                        if thisCell is not None:
                            for merged in self.mergedCells:
                                if coordinate in merged:
                                    mergeCount = merged.max_row - merged.min_row
                                    break
                            else:
                                mergeCount = 0
                            thisCell = str(thisCell).strip()
                            if thisCell == '-':
                                lastTest[thisCol] = {}
                                lastTest[thisCol]['name'] = '.'
                                lastTest[thisCol]['mergeCount'] = mergeCount
                                continue
                            name = self.decisionTables[table]['inputColumns'][thisCol - 1]['name']
                            variable = self.glossary[name]['item']
                            test = self.test2sfeel(variable, coordinate, thisCell)
                            lastTest[thisCol] = {}
                            lastTest[thisCol]['name'] = name
                            lastTest[thisCol]['variable'] = variable
                            lastTest[thisCol]['test'] = test
                            lastTest[thisCol]['thisCell'] = thisCell
                            lastTest[thisCol]['mergeCount'] = mergeCount
                        elif (thisCol in lastTest) and (lastTest[thisCol]['mergeCount'] > 0):
                            lastTest[thisCol]['mergeCount'] -= 1
                            name = lastTest[thisCol]['name']
                            if name == '.':
                                continue
                            variable = lastTest[thisCol]['variable']
                            test = lastTest[thisCol]['test']
                            thisCell = lastTest[thisCol]['thisCell']
                        else:
                            continue
                        if 'validity' in self.decisionTables[table]['inputColumns'][thisCol - 1]:
                            try:
                                thisValue = float(thisCell)
                            except:
                                thisValue = thisCell
                            if thisValue not in self.decisionTables[table]['inputColumns'][thisCol - 1]['validity']:
                                self.errors.append("Input test '{!s}' at '{!s}' is not in the input valid list '{!r}'".format(
                                    thisCell, coordinate, self.decisionTables[table]['inputColumns'][thisCol - 1]['validity']))
                        # print("Setting test '{!s}' at '{!s}' to '{!s}'".format(name, coordinate, test))
                        self.rules[table][thisRule]['tests'].append((name, test))
                    elif thisCol <= inputColumns + outputColumns:
                        if thisCell is not None:
                            for merged in self.mergedCells:
                                if coordinate in merged:
                                    mergeCount = merged.max_row - merged.min_row
                                    break
                            else:
                                mergeCount = 0
                            thisCell = str(thisCell).strip()
                            name = self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]['name']
                            variable = self.glossary[name]['item']
                            result = self.result2sfeel(variable, coordinate, thisCell)
                            lastResult[thisCol] = {}
                            lastResult[thisCol]['name'] = name
                            lastResult[thisCol]['variable'] = variable
                            lastResult[thisCol]['result'] = result
                            lastResult[thisCol]['thisCell'] = thisCell
                            lastResult[thisCol]['mergeCount'] = mergeCount
                        elif (thisCol in lastResult) and (lastResult[thisCol]['mergeCount'] > 0):
                            lastResult[thisCol]['mergeCount'] -= 1
                            name = lastResult[thisCol]['name']
                            variable = lastResult[thisCol]['variable']
                            result = lastResult[thisCol]['result']
                            thisCell = lastResult[thisCol]['thisCell']
                        else:
                            self.errors.append('Missing output value at {!s}'.format(coordinate))
                            continue
                        rank = None
                        if 'validity' in self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]:
                            try:
                                thisResult = float(thisCell)
                            except:
                                thisResult = thisCell
                            if thisResult not in self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]['validity']:
                                self.errors.append("Output value '{!s}' at '{!s}' is not in the output valid list '{!r}'".format(
                                    thisCell, coordinate, self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]['validity']))
                            else:
                                rank = self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]['validity'].index(thisResult)
                        # print("Setting result '{!s}' at '{!s}' to '{!s}' with rank '{!s}'".format(name, coordinate, result, rank))
                        self.rules[table][thisRule]['outputs'].append((name, result, rank))
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
                    border = cell.offset(row=thisRow).border
                    if border.right.style != 'double':
                        border = cell.offset(row=thisRow, column=1).border
                        if border.left.style != 'double':
                            doingValidity = True
                        thisCol += 1
                else:
                    thisCell = cell.offset(row=thisRow, column=thisCol).value
                    thisRule = len(self.rules[table])
                    self.rules[table].append({})
                    self.rules[table][thisRule]['tests'] = []
                    self.rules[table][thisRule]['outputs'] = []
                    if thisCell is not None:
                        thisCell = str(thisCell).strip()
                        self.rules[table][thisRule]['ruleId'] = thisCell

            # Parse the heading
            inputRows = outputRows = 0
            self.decisionTables[table]['inputRows'] = []
            self.decisionTables[table]['outputRows'] = []
            doingInputs = True
            for thisRow in range(1, rows - 1):
                thisCell = cell.offset(row=thisRow).value
                coordinate = cell.offset(row=thisRow).coordinate
                if thisCell is None:
                    if doingOutputs:
                        self.errors.append("Missing Output heading in table '{!s}' at '{!s}'".format(table, coordinate))
                    elif not doingAnnotation:
                        self.errors.append("Missing Input heading in table '{!s}' at '{!s}'".format(table, coordinate))
                    else:
                        self.errors.append("Missing Annotation heading in table '{!s}' at '{!s}'".format(table, coordinate))
                    return (rows, cols, -1)
                thisCell = str(thisCell).strip()
                # Check that all the headings are in the Glossary
                if not doingAnnotation:
                    if thisCell not in self.glossary:
                        if doingInputs:
                            self.errors.append("Input heading '{!s}' in table '{!s}' at '{!s}' is not in the Glossary".format(thisCell, table, coordinate))
                        else:
                            self.errors.append("Output heading '{!s}' in table '{!s}' at '{!s}' is not in the Glossary".format(thisCell, table, coordinate))
                        return (rows, cols, -1)
                if doingInputs:
                    inputRows += 1
                    inRow = len(self.decisionTables[table]['inputRows'])
                    self.decisionTables[table]['inputRows'].append({})
                    self.decisionTables[table]['inputRows'][inRow]['name'] = thisCell
                    border = cell.offset(row=thisRow).border
                    if border.bottom.style == 'double':
                        doingInputs = False
                    border = cell.offset(row=thisRow + 1).border
                    if border.top.style == 'double':
                        doingInputs = False
                elif not doingAnnotation:
                    outputRows += 1
                    outRow = len(self.decisionTables[table]['outputRows'])
                    self.decisionTables[table]['outputRows'].append({})
                    self.decisionTables[table]['outputRows'][outRow]['name'] = thisCell
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

            if doingInputs:
                self.errors.append("No Output row in table '{!s}' - missing double bar horizontal border".format(table))
                return (rows, cols, -1)

            rulesCol = 1
            if doingValidity:
                # Parse the validity column
                outputRow = inputRow = 0
                doingInputs = True
                ranksFound = False
                for thisRow in range(1, rows - 1):
                    thisCell = cell.offset(row=thisRow, column=1).value
                    if thisCell is None:
                        thisRow += 1
                        continue
                    thisCell = str(thisCell).strip()
                    ranksFound = True
                    coordinate = cell.offset(row=thisRow, column=1).coordinate
                    if inputRow < inputRows:
                        name = self.decisionTables[table]['inputRows'][inputRow]['name']
                        variable = self.glossary[name]['item']
                        self.decisionTables[table]['inputRows'][inputRow]['validity'] = []
                        if not isinstance(thisCell, str):
                            self.decisionTables[table]['inputRows'][inputRow]['validity'].append(thisCell)
                        else:
                            validityTests = thisCell.split(',')
                            for validTest in validityTests:
                                thisTest = validTest.strip()
                                try:
                                    self.decisionTables[table]['inputRows'][inputRow]['validity'].append(float(thisTest))
                                except:
                                    self.decisionTables[table]['inputRows'][inputRow]['validity'].append(thisTest)
                        inputRow += 1
                    elif outputRow < outputRows:
                        name = self.decisionTables[table]['outputRows'][outputRow]['name']
                        variable = self.glossary[name]['item']
                        self.decisionTables[table]['outputRows'][outputRow]['validity'] = []
                        if not isinstance(thisCell, str):
                            self.decisionTables[table]['outputRows'][outputRow]['validity'].append(thisCell)
                        else:
                            validityTests = thisCell.split(',')
                            for validTest in validityTests:
                                thisTest = validTest.strip()
                                try:
                                    self.decisionTables[table]['outputRows'][outputRow]['validity'].append(float(thisTest))
                                except:
                                    self.decisionTables[table]['outputRows'][outputRow]['validity'].append(thisTest)
                        outputRow += 1
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

            # Parse the rules
            outputRow = inputRow = 0
            for thisRow in range(1, rows - 1):
                lastTest = lastResult = {}
                for thisCol in range(rulesCol, cols):
                    thisRule = thisCol - rulesCol
                    if doingAnnotation and ('annotation' not in self.rules[table][thisRule]):
                        self.rules[table][thisRule]['annotation'] = []
                    thisCell = cell.offset(row=thisRow, column=thisCol).value
                    coordinate = cell.offset(row=thisRow, column=thisCol).coordinate
                    if inputRow < inputRows:
                        if thisCell is not None:
                            for merged in self.mergedCells:
                                if coordinate in merged:
                                    mergeCount = merged.max_col - merged.min_col
                                    break
                            else:
                                mergeCount = 0
                            thisCell = str(thisCell).strip()
                            if thisCell == '-':
                                lastTest = {}
                                lastTest['name'] = '.'
                                lastTest['mergeCount'] = mergeCount
                                continue
                            name = self.decisionTables[table]['inputRows'][inputRow]['name']
                            variable = self.glossary[name]['item']
                            test = self.test2sfeel(variable, coordinate, thisCell)
                            lastTest = {}
                            lastTest['name'] = name
                            lastTest['variable'] = variable
                            lastTest['test'] = test
                            lastTest['thisCell'] = thisCell
                            lastTest['mergeCount'] = mergeCount
                        elif ('mergeCount' in lastTest) and (lastTest['mergeCount'] > 0):
                            lastTest['mergeCount'] -= 1
                            name = lastTest['name']
                            if name == '.':
                                continue
                            variable = lastTest['variable']
                            test = lastTest['test']
                            thisCell = lastTest['thisCell']
                        else:
                            continue
                        if 'validity' in self.decisionTables[table]['inputRows'][inputRow]:
                            try:
                                thisValue = float(thisCell)
                            except:
                                thisValue = thisCell
                            if thisValue not in self.decisionTables[table]['inputRows'][inputRow]['validity']:
                                self.errors.append("Input test '{!s}' at '{!s}' is not in the input valid list '{!s}'".format(
                                    thisCell, coordinate, self.decisionTables[table]['inputRows'][inputRow]['validity']))
                        # print("Setting test '{!s}' at '{!s}' to '{!s}'".format(name, coordinate, test))
                        self.rules[table][thisRule]['tests'].append((name, test))
                    elif outputRow < outputRows:
                        if thisCell is not None:
                            for merged in self.mergedCells:
                                if coordinate in merged:
                                    mergeCount = merged.max_col - merged.min_col
                                    break
                            else:
                                mergeCount = 0
                            thisCell = str(thisCell).strip()
                            name = self.decisionTables[table]['outputRows'][outputRow]['name']
                            variable = self.glossary[name]['item']
                            result = self.result2sfeel(variable, coordinate, thisCell)
                            lastResult = {}
                            lastResult['name'] = name
                            lastResult['variable'] = variable
                            lastResult['result'] = result
                            lastResult['thisCell'] = thisCell
                            lastResult['mergeCount'] = mergeCount
                        elif ('mergeCount' in lastResult) and (lastResult['mergeCount'] > 0):
                            lastResult['mergeCount'] -= 1
                            name = lastResult['name']
                            variable = lastResult['variable']
                            result = lastResult['result']
                            thisCell = lastResult['thisCell']
                        else:
                            self.errors.append('Missing output value at {!s}'.format(coordinate))
                            outputRows += 1
                            continue
                        rank = None
                        if 'validity' in self.decisionTables[table]['outputRows'][outputRow]:
                            try:
                                thisCellResult = float(thisCell)
                            except:
                                thisCellResult = thisCell
                            if thisCellResult not in self.decisionTables[table]['outputRows'][outputRow]['validity']:
                                self.errors.append("Output value '{!s}' at '{!s}' is not in the output valid list '{!s}'".format(
                                    thisCell, coordinate, self.decisionTables[table]['outputRows'][outputRow]['validity']))
                            else:
                                rank = self.decisionTables[table]['outputRows'][outputRow]['validity'].index(thisCellResult)
                        # print("Setting result '{!s}' at '{!s}' to '{!s}' with rank '{!s}'".format(name, coordinate, result, rank))
                        self.rules[table][thisRule]['outputs'].append((name, result, rank))
                    else:
                        self.rules[table][thisRule]['annotation'].append(thisCell)
                if inputRow < inputRows:
                    inputRow += 1
                else:
                    outputRow += 1
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
                self.errors.append("Decision table '{!s}' - unknown type".format(table))
                return (rows, cols, -1)

            self.decisionTables[table]['hitPolicy'] = 'U'
            self.decisionTables[table]['inputColumns'] = []
            self.decisionTables[table]['inputRows'] = []
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
                        thisCell = str(thisCell).strip()
                        if thisCell == '-':
                            lastTest[thisVariable] = {}
                            lastTest[thisVariable]['name'] = '.'
                            lastTest[thisVariable]['mergeCount'] = mergeCount
                            continue
                        name = inputs[thisVariable].strip()
                        variable = self.glossary[name]['item']
                        test = self.test2sfeel(variable, coordinate, thisCell)
                        lastTest[thisVariable] = {}
                        lastTest[thisVariable]['name'] = name
                        lastTest[thisVariable]['variable'] = variable
                        lastTest[thisVariable]['test'] = test
                        lastTest[thisVariable]['thisCell'] = thisCell
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
                        self.errors.append('Missing horizontal input test at {!s}'.format(coordinate))
                        continue
                    # print("Setting horizontal test '{!s}' at '{!s}' to '{!s}'".format(name, coordinate, test))
                    self.decisionTables[table]['inputColumns'][thisCol]['tests'].append((name, test))

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
                        thisCell = str(thisCell).strip()
                        if thisCell == '-':
                            lastTest[thisVariable] = {}
                            lastTest[thisVariable]['name'] = '.'
                            lastTest[thisVariable]['mergeCount'] = mergeCount
                            continue
                        name = inputs[thisVariable].strip()
                        variable = self.glossary[name]['item']
                        test = self.test2sfeel(variable, coordinate, thisCell)
                        lastTest[thisVariable] = {}
                        lastTest[thisVariable]['name'] = name
                        lastTest[thisVariable]['variable'] = variable
                        lastTest[thisVariable]['test'] = test
                        lastTest[thisVariable]['thisCell'] = thisCell
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
                        self.errors.append('Missing vertical input test at {!s}'.format(coordinate))
                        continue
                    # print("Setting veritical test '{!s}' at '{!s}' to '{!s}'".format(name, coordinate, test))
                    self.decisionTables[table]['inputRows'][thisRow]['tests'].append((name, test))

            # Now build the Rules
            thisRule = 0
            for row in range(verticalRows):
                for col in range(horizontalCols):
                    self.rules[table].append({})
                    self.rules[table][thisRule]['ruleId'] = thisRule + 1
                    self.rules[table][thisRule]['tests'] = []
                    self.rules[table][thisRule]['outputs'] = []
                    self.rules[table][thisRule]['tests'] += self.decisionTables[table]['inputColumns'][col]['tests']
                    self.rules[table][thisRule]['tests'] += self.decisionTables[table]['inputRows'][row]['tests']
                    thisCell = cell.offset(row=1 + height + row, column=width + col).value
                    coordinate = cell.offset(row=1 + height + row, column=width + col).coordinate
                    if thisCell is None:
                        self.errors.append('Missing output result at {!s}'.format(coordinate))
                        return (rows, cols, -1)
                    thisCell = str(thisCell).strip()
                    name = self.decisionTables[table]['output']['name']
                    variable = self.glossary[name]['item']
                    result = self.result2sfeel(variable, coordinate, thisCell)
                    # print("Setting result at '{!s}' to '{!s}'".format(coordinate, result))
                    self.rules[table][thisRule]['outputs'].append((name, result, 0))
                    thisRule += 1

        return (rows, cols, len(self.rules[table]))


    def load(self, rulesBook):
        '''
        Load a rulesBook
        '''
        self.errors = []
        try:
            self.wb = load_workbook(filename=rulesBook)
        except (utils.exceptions.InvalidFileException, IOError):
            self.errors.append('No workbook named %s!', rulesBook)
            status = {}
            status['errors'] = self.errors
            return status

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
        endGlossary = False
        glossaryColumn = None
        self.glossary = {}
        self.glossaryItems = {}
        for row in ws.rows:
            for cell in row:
                if not inGlossary:
                    thisCell = cell.value
                    if isinstance(thisCell, str):
                        if thisCell.startswith('Glossary'):
                            coordinate = cell.coordinate
                            for merged in self.mergedCells:
                                if coordinate in merged:
                                    width = merged.max_col - merged.min_col
                                    if width != 2:
                                        self.errors.append('Invalid Glossary - not 3 columns wide')
                                        status = {}
                                        status['errors'] = self.errors
                                        return status
                                    break
                            else:
                                self.errors.append('Invalid Glossary - no merged heading')
                                status = {}
                                status['errors'] = self.errors
                                return status
                            inGlossary = True
                            glossaryColumn = cell.column
                            header = True
                            break
                if not inGlossary:
                    continue
                if cell.column < glossaryColumn:
                    continue
                if header:
                    if cell.value != 'Variable':
                        self.errors.append('Missing Glossary heading')
                        status = {}
                        status['errors'] = self.errors
                        return status
                    if cell.offset(column=1).value != 'Business Concept':
                        self.errors.append('Bad Glossary heading')
                        status = {}
                        status['errors'] = self.errors
                        return status
                    if cell.offset(column=2).value != 'Attribute':
                        self.errors.append('Bad Glossary heading')
                        status = {}
                        status['errors'] = self.errors
                        return status
                    thisConcept = None
                    header = False
                    break
                if (cell.column == glossaryColumn) and (cell.value is None):
                    endGlossary = True
                    break
                variable = cell.value
                if variable in self.glossary:
                    self.errors.append('Variable ({!s}) with multiple definitions in Glossary'.format(variable))
                    status = {}
                    status['errors'] = self.errors
                    return status
                concept = cell.offset(column=1).value
                attribute = cell.offset(column=2).value
                if thisConcept is None:
                    if concept is None:
                        self.errors.append('Missing Business Concept in Glossary')
                        status = {}
                        status['errors'] = self.errors
                        return status
                if concept is not None:
                    if '.' in concept:
                        self.errors.append('Bad Business Concept in Glossary:{!s}'.format(concept))
                        status = {}
                        status['errors'] = self.errors
                        return status
                    thisConcept = concept
                if (attribute is None) or ('.' in attribute):
                    self.errors.append('Bad Business Attribute ({!s}) for Variable ({!s}) in Business in Concept ({!s}) in Glossary:{!s}'.format(attribute, variable, thisConcept))
                    status = {}
                    status['errors'] = self.errors
                    return status
                item = thisConcept + '.' + attribute
                self.glossary[variable] = {}
                self.glossary[variable]['item'] = item
                self.glossary[variable]['concept'] = thisConcept
                self.glossaryItems[item] = variable
                break
            if endGlossary:
                break
        if not inGlossary:
            self.errors.append('Glossary not found')
            status = {}
            status['errors'] = self.errors
            return status
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
                            coordinate = cell.coordinate
                            for merged in self.mergedCells:
                                if coordinate in merged:
                                    width = merged.max_col - merged.min_col
                                    if width != 1:
                                        self.errors.append('Invalid Decision - not 2 columns wide')
                                        status = {}
                                        status['errors'] = self.errors
                                        return status
                                    break
                            else:
                                self.errors.append('Invalid Decision - no merged heading')
                                status = {}
                                status['errors'] = self.errors
                                return status
                            inDecision = True
                            decisionColumn = cell.column
                            header = True
                            break
                if not inDecision:
                    continue
                if cell.column < decisionColumn:
                    continue
                if header:
                    if cell.value != 'Decisions':
                        self.errors.append('Missing Decision heading')
                        status = {}
                        status['errors'] = self.errors
                        return status
                    if cell.offset(column=1).value != 'Execute Decision Tables':
                        self.errors.append('Bad Decision heading')
                        status = {}
                        status['errors'] = self.errors
                        return status
                    header = False
                    break
                if (cell.column == decisionColumn) and (cell.value is None):
                    endDecision = True
                    break
                decision = cell.value
                table = cell.offset(column=1).value
                if table in self.decisionTables:
                    self.errors.append('Execution Decision Table ({!s}) repeated in Decision'.format(table))
                    status = {}
                    status['errors'] = self.errors
                    return status
                self.decisionTables[table] = {}
                self.decisions.append((decision, table))
                break
            if endDecision:
                break
        if not inDecision:
            self.errors.append('Decision not found')
            status = {}
            status['errors'] = self.errors
            return status

        # Now search for the Decision Tables
        self.rules = {}
        for sheet in self.wb.sheetnames:
            if sheet in ['Glossary', 'Decision', 'Test']:
                continue
            ws = self.wb[sheet]
            self.mergedCells = ws.merged_cells.ranges
            parsedRanges = []
            for row in ws.rows:
                thisRow = row[0].row
                for cell in row:
                    for i in range(len(parsedRanges)):
                        parsed = parsedRanges[i]
                        if cell.coordinate in parsed:
                            continue
                    thisCell = cell.value
                    if isinstance(thisCell, str):
                        if thisCell in self.decisionTables:
                            (rows, cols, rules) = self.parseDecionTable(cell)
                            if rules == -1:
                                status = {}
                                if len(self.errors) > 0:
                                    status['errors'] = self.errors
                                return status
                            elif rules == 0:
                                print("Decision table '{!s}' has no rules".format(thisCell))
                                continue
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
        self.isLoaded = True
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


    def decide(self, data):
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
            value = self.value2sfeel(value)
            if value is None:
                validData = False
            else:
                # print('Setting variable ({!s}) [item ({!s})] to value ({!s})'.format(variable, item, value))
                retVal = self.sfeel('{} <- {}'.format(item, value))
        if not validData:
            status = {}
            status['errors'] = self.errors
            self.errors = []
            return (status, {})

        # Process each decision table in order
        newData = {}
        ruleIds = []
        annotations = []
        for variable in self.glossary:
            item = self.glossary[variable]['item']
            newData[variable] = self.sfeel('{}'.format(item))
        for table in self.decisionTables:
            ranks = []
            foundRule = None
            rankedRules = []
            for thisRule in range(len(self.rules[table])):
                for (variable, test) in self.rules[table][thisRule]['tests']:
                    item = self.glossary[variable]['item']
                    itemValue = self.sfeel('{}'.format(item))
                    retVal = self.sfeel('{}'.format(test))
                    # print("variable '{!s}' [data '{!s}'] with test '{!s}' returned '{!s}'".format(variable, itemValue, test, retVal))
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
                            for (variable, result, rank) in self.rules[table][thisRule]['outputs']:
                                theseRanks.append(rank)
                            theseRanks.append(thisRule)
                            ranks.append(theseRanks)
                        else:
                            before = None
                            beforeFound = False
                            for i in len(range(ranks)):
                                for (variable, result, rank) in self.rules[table][thisRule]['outputs']:
                                    if rank < ranks[i]:
                                        before = i
                                        break
                                    elif rank > ranks[i]:
                                        beforeFound = True
                                        break
                                if beforeFound:
                                    break
                            theseRanks = []
                            for (variable, result, rank) in self.rules[table][thisRule]['outputs']:
                                theseRanks.append(rank)
                            theseRanks.append(thisRule)
                            if not beforeFound:
                                ranks.append(theseRanks)
                            else:
                                ranks.insert(i, theseRanks)
            if self.decisionTables[table]['hitPolicy'] in ['U', 'A', 'F']:
                if foundRule is None:
                    self.errors.append("No rules matched the input data for decision table '{!s}'".format(table))
                    status = {}
                    status['errors'] = self.errors
                    self.errors = []
                    return (status, {})
                else:
                    for (variable, result, rank) in self.rules[table][foundRule]['outputs']:
                        item = self.glossary[variable]['item']
                        retVal = self.sfeel('{} <- {}'.format(item, result))
                        newData[variable] = self.sfeel('{}'.format(item))
                    ruleIds.append((table, str(self.rules[table][foundRule]['ruleId'])))
                    if 'annotation' in self.decisionTables[table]:
                        for annotation in range(len(self.decisionTables[table]['annotation'])):
                            name = self.decisionTables[table]['annotation'][annotation]
                            text = self.rules[table][foundRule]['annotation'][annotation]
                            annotations.append((name, text))
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
                        item = self.glossary[variable]['item']
                        if first:
                            if variable not in newData:
                                if len(self.decisionTables[table]['hitPolicy']) == 1:
                                    newData[variable] = []
                                elif self.decisionTables[table]['hitPolicy'][1] in ['+', '#']:
                                    newData[variable] = 0
                                else:
                                    newData[item] = None
                            first = False
                        retVal = self.sfeel('{} <- {}'.format(item, result))
                        thisOutput = self.sfeel('{}'.format(item))
                        if len(self.decisionTables[table]['hitPolicy']) == 1:
                            newData[variable].append(thisOutput)
                        elif self.decisionTables[table]['hitPolicy'][1] == '+':
                            newData[variable] += thisOutput
                        elif self.decisionTables[table]['hitPolicy'][1] == '<':
                            if newData[variable] is None:
                                newData[variable] = thisOutput
                            elif thisOutput < newData[variable]:
                                newData[variable] = thisOutput
                        elif self.decisionTables[table]['hitPolicy'][1] == '>':
                            if newData[variable] is None:
                                newData[variable] = thisOutput
                            elif thisOutput > newData[variable]:
                                newData[variable] = thisOutput
                        else:
                            newData[variable] += 1
                    ruleIds.append((table, str(self.rules[table][foundRule]['ruleId'])))
                    if 'annotation' in self.decisionTables[table]:
                        for annotation in range(len(self.decisionTables[table]['annotation'])):
                            name = self.decisionTables[table]['annotation'][annotation]
                            text = self.rules[table][foundRule]['annotation'][annotation]
                            annotations.append((name, text))
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
                        item = self.glossary[variable]['item']
                        retVal = self.sfeel('{} <- {}'.format(item, result))
                        newData[variable] = self.sfeel('{}'.format(item))
                    ruleIds.append((table, str(self.rules[table][foundRule]['ruleId'])))
                    if 'annotation' in self.decisionTables[table]:
                        for annotation in range(len(self.decisionTables[table]['annotation'])):
                            name = self.decisionTables[table]['annotation'][annotation]
                            text = self.rules[table][foundRule]['annotation'][annotation]
                            annotations.append((name, text))
            elif self.decisionTables[table]['hitPolicy'][0] == 'O':
                if len(ranks) == 0:
                    self.errors.append("No rules matched the input data for decision table '{!s}'".format(table))
                    status = {}
                    status['errors'] = self.errors
                    self.errors = []
                    return (status, {})
                else:
                    for i in range(len(ranks)):
                        foundRule = ranks[i][-1]
                        for (variable, result, rank) in self.rules[table][foundRule]['outputs']:
                            item = self.glossary[variable]['item']
                            if item not in newData:
                                newData[variable] = []
                            retVal = self.sfeel('{} <- {}'.format(item, result))
                            newData[variable].append(self.sfeel('{}'.format(item)))
                        ruleIds.append(table + ':' + str(self.rules[table][foundRule]['ruleId']))
                        if 'annotation' in self.decisionTables[table]:
                            for annotation in range(len(self.decisionTables[table]['annotation'])):
                                name = self.decisionTables[table]['annotation'][annotation]
                                text = self.rules[table][foundRule]['annotation'][annotation]
                                annotations.append((name, text))

        newData['Execute Rules'] = ruleIds
        if len(annotations) > 0:
            newData['Annotations'] = annotations
        status = {}
        if len(self.errors) > 0:
            status['errors'] = self.errors
            self.errors = []
        return (status, newData)


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
