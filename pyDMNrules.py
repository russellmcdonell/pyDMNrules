# -----------------------------------------------------------------------------
# pyDMNrules.py
# -----------------------------------------------------------------------------

import sys
import re
import datetime
import pySFeel
from openpyxl import load_workbook
from openpyxl import utils

class pyDMNrules():


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
        elif isinstance(value, str):
            if value in self.glossary:
                return value
            else:
                return '"' + value + '"'
        elif isinstance(value, int):
            return str(value)
        elif isinstance(value, float):
            return str(value)
        elif isinstance(value, bool):
            if value:
                return 'true'
            else:
                return 'false'
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


    def parseDecionTable(self, ws, cell):
        '''
        Parse a Decision Table
        '''
        rows = cols = 0
        startRow = cell.row
        startCol = cell.column
        table = cell.value
        coordinate = cell.coordinate
        # print("Parsing Decision Table '{!s}' at '{!s}'".format(table, coordinate))
        self.rules[table] = []
        # Check the next cell down to determine the decision table layout
        if cell.offset(row=1).value is None:
            # Empty table
            self.errors.append("Decision table '{!s}' at {!s}' is empty".format(table,coordinate))
            return (rows, cols, -1)
        if cell.offset(row=1).value not in self.glossary:
            # Rules as rows
            # Parse the heading
            rows += 1
            inputColumns = 0
            doingValidity = False
            self.decisionTables[table]['inputColumns'] = []
            self.decisionTables[table]['outputColumns'] = []
            while ws.cell(row=startRow + rows, column=startCol + cols).value != None:
                thisCell = ws.cell(row=startRow + rows, column=startCol + cols).value
                if cols == 0:   # This should be the hit policy
                    if thisCell is None:
                        hitPolicy = 'U'
                    else:
                        hitPolicy = thisCell
                    if (not isinstance(hitPolicy, str)) or (hitPolicy[0] not in ['U', 'A', 'P', 'F', 'C', 'O', 'R']):
                        self.errors.append("Invalid hit policy '{!s}' for table '{!s}'".format(hitPolicy, table))
                        return (rows, cols, -1)
                    if len(hitPolicy) != 1:
                        if (len(hitPolicy) != 2) or (hitPolicy[1] not in ['+', '<', '>', '#']):
                            self.errors.append("Invalid hit policy '{!s}' for table '{!s}'".format(hitPolicy, table))
                            return (rows, cols, -1)
                    self.decisionTables[table]['hitPolicy'] = hitPolicy
                    border = ws.cell(row=startRow + rows, column=startCol + cols).border
                    if border.bottom.style != 'double':
                        border = ws.cell(row=startRow + rows + 1, column=startCol + cols).border
                        if border.top.style != 'double':
                            doingValidity = True
                    doingInputs = True
                    cols += 1
                    continue
                # Check that all the headings are in the Glossary
                if thisCell not in self.glossary:
                    if doingInputs:
                        self.errors.append("Input heading '{!s}' in table '{!s}' is not in the Glossary".format(thisCell, table))
                    else:
                        self.errors.append("Output heading '{!s}' in table '{!s}' is not in the Glossary".format(thisCell, table))
                    return (rows, cols, -1)
                if doingInputs:
                    inputColumns += 1
                    thisCol = len(self.decisionTables[table]['inputColumns'])
                    self.decisionTables[table]['inputColumns'].append({})
                    self.decisionTables[table]['inputColumns'][thisCol]['name'] = thisCell
                    border = ws.cell(row=startRow + rows, column=startCol + cols).border
                    if border.right.style == 'double':
                        doingInputs = False
                    border = ws.cell(row=startRow + rows, column=startCol + cols + 1).border
                    if border.left.style == 'double':
                        doingInputs = False
                else:
                    thisCol = len(self.decisionTables[table]['outputColumns'])
                    self.decisionTables[table]['outputColumns'].append({})
                    self.decisionTables[table]['outputColumns'][thisCol]['name'] = thisCell
                cols += 1
                continue
            if doingInputs:
                self.errors.append("No Output column in table '{!s}' - missing double bar vertical border".format(table))
                return (rows, cols, -1)
            rows += 1
            if doingValidity:
                # Parse the validity row
                thisCol = 1
                doingInputs = True
                ranksFound = False
                while thisCol < cols:
                    thisCell = ws.cell(row=startRow + rows, column=startCol + thisCol).value
                    if thisCell is None:
                        thisCol += 1
                        continue
                    ranksFound = True
                    coordinate = ws.cell(row=startRow + rows, column=startCol + thisCol).coordinate
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
                    else:
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
                    thisCol += 1
                doingValidity = False
                rows += 1
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
            while ws.cell(row=startRow + rows, column=startCol).value != None:
                thisCol = 0
                thisRule = len(self.rules[table])
                self.rules[table].append({})
                self.rules[table][thisRule]['tests'] = []
                self.rules[table][thisRule]['outputs'] = []
                while thisCol < cols:
                    thisCell = ws.cell(row=startRow + rows, column=startCol + thisCol).value
                    if thisCol == 0:
                        self.rules[table][thisRule]['ruleId'] = thisCell
                        thisCol += 1
                        continue
                    coordinate = ws.cell(row=startRow + rows, column=startCol + thisCol).coordinate
                    if thisCol <= inputColumns:
                        if thisCell is not None:
                            thisCell = str(thisCell).strip()
                            if thisCell == '-':
                                if thisCol in lastTest:
                                    lastTest[thisCol]['name'] = '.'
                                thisCol += 1
                                continue
                            name = self.decisionTables[table]['inputColumns'][thisCol - 1]['name']
                            variable = self.glossary[name]['item']
                            test = self.test2sfeel(variable, coordinate, thisCell)
                            lastTest[thisCol] = {}
                            lastTest[thisCol]['name'] = name
                            lastTest[thisCol]['variable'] = variable
                            lastTest[thisCol]['test'] = test
                            lastTest[thisCol]['thisCell'] = thisCell
                        elif thisCol in lastTest:
                            name = lastTest[thisCol]['name']
                            if name == '.':
                                thisCol += 1
                                continue
                            variable = lastTest[thisCol]['variable']
                            test = lastTest[thisCol]['test']
                            thisCell = lastTest[thisCol]['thisCell']
                        else:
                            self.errors.append('Missing input test at {!s}'.format(coordinate))
                            thisCol += 1
                            continue
                        if 'validity' in self.decisionTables[table]['inputColumns'][thisCol - 1]:
                            thisValue = thisCell.strip()
                            try:
                                thisCellValue = float(thisValue)
                            except:
                                thisCellValue = thisValue
                            if thisCellValue not in self.decisionTables[table]['inputColumns'][thisCol - 1]['validity']:
                                self.errors.append("Input test '{!s}' at '{!s}' is not in the input valid list '{!s}'".format(
                                    thisCellValue, coordinate, self.decisionTables[table]['inputColumns'][thisCol - 1]['validity']))
                        # print("Setting test at '{!s}' to '{!s}'".format(coordinate, test))
                        self.rules[table][thisRule]['tests'].append((name, test))
                    else:
                        if thisCell is not None:
                            thisCell = str(thisCell).strip()
                            name = self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]['name']
                            variable = self.glossary[name]['item']
                            result = self.result2sfeel(variable, coordinate, thisCell)
                            lastResult[thisCol] = {}
                            lastResult[thisCol]['name'] = name
                            lastResult[thisCol]['variable'] = variable
                            lastResult[thisCol]['result'] = result
                        elif thisCol in lastResult:
                            name = lastResult[thisCol]['name']
                            variable = lastResult[thisCol]['variable']
                            result = lastResult[thisCol]['result']
                        else:
                            self.errors.append('Missing output value at {!s}'.format(coordinate))
                            thisCol += 1
                            continue
                        rank = None
                        if 'validity' in self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]:
                            thisResult = thisCell.strip()
                            try:
                                thisCellResult = float(thisResult)
                            except:
                                thisCellResult = thisResult
                            if thisCellResult not in self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]['validity']:
                                self.errors.append("Output value '{!s}' at '{!s}' is not in the output valid list '{!s}'".format(
                                    thisCellResult, coordinate, self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]['validity']))
                            else:
                                rank = self.decisionTables[table]['outputColumns'][thisCol - inputColumns - 1]['validity'].index(thisCellResult)
                        # print("Setting result at '{!s}' to '{!s}'".format(coordinate, result))
                        self.rules[table][thisRule]['outputs'].append((name, result, rank))
                    thisCol += 1
                    continue
                rows += 1
            return (rows, cols, len(self.rules[table]))
        elif cell.offset(row=2).value in self.glossary:
            # Rules as columns
            rows += 1
            # Search for the end of the table
            while ws.cell(row=startRow + rows, column=startCol + cols).value != None:
                rows += 1
            # Parse the footer
            doingValidity = False
            thisRow = startRow + rows - 1
            while ws.cell(row=thisRow, column=startCol + cols).value != None:
                thisCell = ws.cell(row=thisRow, column=startCol + cols).value
                if cols == 0:   # Should be hit policy
                    if thisCell is None:
                        hitPolicy = 'U'
                    else:
                        hitPolicy = thisCell
                    if (not isinstance(hitPolicy, str)) or (hitPolicy[0] not in ['U', 'A', 'P', 'F', 'C', 'O', 'R']):
                        self.errors.append("Invalid hit policy '{!s}' for table '{!s}'".format(hitPolicy, table))
                        return (rows, cols, -1)
                    if len(hitPolicy) != 1:
                        if (len(hitPolicy) != 2) or (hitPolicy[1] not in ['+', '<', '>', '#']):
                            self.errors.append("Invalid hit policy '{!s}' for table '{!s}'".format(hitPolicy, table))
                            return (rows, cols, -1)
                    self.decisionTables[table]['hitPolicy'] = hitPolicy
                    border = ws.cell(row=thisRow, column=startCol).border
                    if border.right.style != 'double':
                        border = ws.cell(row=thisRow, column=startCol + 1).border
                        if border.left.style != 'double':
                            doingValidity = True
                            cols += 1
                else:
                    thisCell = ws.cell(row=thisRow, column=startCol + cols).value
                    thisRule = len(self.rules[table])
                    self.rules[table].append({})
                    self.rules[table][thisRule]['tests'] = []
                    self.rules[table][thisRule]['outputs'] = []
                    self.rules[table][thisRule]['ruleId'] = thisCell
                cols += 1
            # Parse the heading
            inputRows = 0
            self.decisionTables[table]['inputRows'] = []
            self.decisionTables[table]['outputRows'] = []
            doingInputs = True
            thisRow = startRow + 1
            while thisRow < startRow + rows - 1:
                thisCell = ws.cell(row=thisRow, column=startCol).value
                # Check that all the headings are in the Glossary
                if thisCell not in self.glossary:
                    if doingOutputs:
                        self.errors.append("Output heading '{!s}' in table '{!s}' is not in the Glossary".format(thisCell, table))
                    else:
                        self.errors.append("Input heading '{!s}' in table '{!s}' is not in the Glossary".format(thisCell, table))
                    return (rows, cols, -1)
                if doingInputs:
                    inputRows += 1
                    inRow = len(self.decisionTables[table]['inputRows'])
                    self.decisionTables[table]['inputRows'].append({})
                    self.decisionTables[table]['inputRows'][inRow]['name'] = thisCell
                    border = ws.cell(row=thisRow, column=startCol).border
                    if border.bottom.style == 'double':
                        doingInputs = False
                    border = ws.cell(row=thisRow + 1, column=startCol).border
                    if border.top.style == 'double':
                        doingInputs = False
                else:
                    outRow = len(self.decisionTables[table]['outputRows'])
                    self.decisionTables[table]['outputRows'].append({})
                    self.decisionTables[table]['outputRows'][outRow]['name'] = thisCell
                thisRow += 1
                continue
            if doingInputs:
                self.errors.append("No Output row in table '{!s}' - missing double bar horizontal border".format(table))
                return (rows, cols, -1)
            rulesCol = 1
            if doingValidity:
                # Parse the validity row
                outputRow = inputRow = 0
                doingInputs = True
                ranksFound = False
                thisRow = startRow + 1
                while thisRow < startRow + rows - 1:
                    thisCell = ws.cell(row=thisRow, column=startCol + 1).value
                    if thisCell is None:
                        thisRow += 1
                        continue
                    ranksFound = True
                    coordinate = ws.cell(row=thisRow, column=startCol + 1).coordinate
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
                    else:
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
                    thisRow += 1
                doingValidity = False
                cols += 1
                rulesCol += 1
                if (not ranksFound) and (self.decisionTables[table]['hitPolicy'] in ['P', 'O']):
                    self.errors.append("Decision table '{!s}' has hit policy '{!s}' but there is no ordered list of output values".format(
                        table, self.decisionTables[table]['hitPolicy']))
                    return (rows, cols, -1)
            elif self.decisionTables[table]['hitPolicy'] in ['P', 'O']:
                self.errors.append("Decision table '{!s}' has hit policy '{!s}' but there is no ordered list of output values".format(
                    table, self.decisionTables[table]['hitPolicy']))
                return (rows, cols, -1)
            # Parse the rules
            thisRow = startRow + 1
            outputRow = inputRow = 0
            while thisRow < startRow + rows - 1:
                lastTest = lastResult = {}
                thisCol = rulesCol
                while thisCol < cols - 1:
                    thisRule = thisCol - rulesCol
                    thisCell = ws.cell(row=thisRow, column=startCol + thisCol).value
                    coordinate = ws.cell(row=thisRow, column=startCol + thisCol).coordinate
                    if inputRow < inputRows:
                        if thisCell is not None:
                            thisCell = str(thisCell).strip()
                            if thisCell == '-':
                                lastTest['name'] = '.'
                                thisCol += 1
                                continue
                            name = self.decisionTables[table]['inputRows'][inputRow]['name']
                            variable = self.glossary[name]['item']
                            test = self.test2sfeel(variable, coordinate, thisCell)
                            lastTest = {}
                            lastTest['name'] = name
                            lastTest['variable'] = variable
                            lastTest['test'] = test
                            lastTest['thisCell'] = thisCell
                        elif 'name' in lastTest:
                            name = lastTest['name']
                            if name == '.':
                                thisCol += 1
                                continue
                            variable = lastTest['variable']
                            test = lastTest['test']
                            thisCell = lastTest['thisCell']
                        else:
                            self.errors.append('Missing input test at {!s}'.format(coordinate))
                            thisCol += 1
                            continue
                        if 'validity' in self.decisionTables[table]['inputRows'][inputRow]:
                            thisValue = thisCell.strip()
                            try:
                                thisCellValue = float(thisValue)
                            except:
                                thisCellValue = thisValue
                            if thisCellValue not in self.decisionTables[table]['inputRows'][inputRow]['validity']:
                                self.errors.append("Input test '{!s}' at '{!s}' is not in the input valid list '{!s}'".format(
                                    thisCellValue, coordinate, self.decisionTables[table]['inputRows'][inputRow]['validity']))
                        # print("Setting test at '{!s}' to '{!s}'".format(coordinate, test))
                        self.rules[table][thisRule]['tests'].append((name, test))
                    else:
                        if thisCell is not None:
                            thisCell = str(thisCell).strip()
                            name = self.decisionTables[table]['outputRows'][outputRow]['name']
                            variable = self.glossary[name]['item']
                            result = self.result2sfeel(variable, coordinate, thisCell)
                            lastResult = {}
                            lastResult['name'] = name
                            lastResult['variable'] = variable
                            lastResult['result'] = result
                        elif 'name' in lastResult:
                            name = lastResult['name']
                            variable = lastResult['variable']
                            result = lastResult['result']
                        else:
                            self.errors.append('Missing output value at {!s}'.format(coordinate))
                            outputRows += 1
                            thisCol += 1
                            continue
                        rank = None
                        if 'validity' in self.decisionTables[table]['outputRows'][outputRow]:
                            thisResult = thisCell.strip()
                            try:
                                thisCellResult = float(thisResult)
                            except:
                                thisCellResult = thisResult
                            if thisCellResult not in self.decisionTables[table]['outputRows'][outputRow]['validity']:
                                self.errors.append("Output value '{!s}' at '{!s}' is not in the output valid list '{!s}'".format(
                                    thisCellResult, coordinate, self.decisionTables[table]['outputRows'][outputRow]['validity']))
                            else:
                                rank = self.decisionTables[table]['outputRows'][outputRow]['validity'].index(thisCellResult)
                        # print("Setting result at '{!s}' to '{!s}' with rank '{!s}'".format(coordinate, result, rank))
                        self.rules[table][thisRule]['outputs'].append((name, result, rank))
                    thisCol += 1
                    continue
                if inputRow < inputRows:
                    inputRow += 1
                else:
                    outputRow += 1
                thisRow += 1
            return (rows, cols, len(self.rules[table]))
        else:
            # Rules as crosstab
            # This is the output, and the only output
            rows += 1
            thisCell = cell.offset(row=1).value
            outputVariable = str(thisCell).strip()
            # This should be merged cell - need a row and a column of variables, plus another row and column of tests (as a minimum)
            mergedCells = ws.merged_cells.ranges
            for merged in mergedCells:
                if cell.offset(row=1).coordinate in merged:
                    width = merged.max_col - merged.min_col + 1
                    height = merged.max_row - merged.min_row + 1
                    break
            else:
                self.errors.append("Decision table '{!s}' - unknown type".format(table))
                return (rows, cols, -1)
            rows += height
            cols += width

            self.decisionTables[table]['hitPolicy'] = 'U'
            self.decisionTables[table]['inputColumns'] = []
            self.decisionTables[table]['inputRows'] = []
            self.decisionTables[table]['output'] = {}
            self.decisionTables[table]['output']['name'] = outputVariable

            # Parse the horizontal heading
            coordinate = cell.offset(row=1, column=width).coordinate
            for merged in mergedCells:
                if coordinate in merged:
                    horizontalCols = merged.max_col - merged.min_col + 1
                    break
            else:
                horizontalCols = 1

            heading = cell.offset(row=1, column=width).value
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

            thisVariable = 0
            while thisVariable < height - 1:
                thisCol = 0
                lastTest = {}
                while thisCol < horizontalCols:
                    if thisVariable == 0:
                        self.decisionTables[table]['inputColumns'].append({})
                        self.decisionTables[table]['inputColumns'][thisCol]['tests'] = []

                    thisCell = ws.cell(row=startRow + 2 + thisVariable, column=startCol + cols + thisCol).value
                    coordinate = ws.cell(row=startRow + 2 + thisVariable, column=startCol + cols + thisCol).coordinate
                    if thisCell is not None:
                        thisCell = str(thisCell).strip()
                        if thisCell == '-':
                            if thisVariable in lastTest:
                                lastTest[thisVariable]['name'] = '.'
                            thisCol += 1
                            continue
                        name = inputs[thisVariable].strip()
                        variable = self.glossary[name]['item']
                        test = self.test2sfeel(variable, coordinate, thisCell)
                        lastTest[thisVariable] = {}
                        lastTest[thisVariable]['name'] = name
                        lastTest[thisVariable]['variable'] = variable
                        lastTest[thisVariable]['test'] = test
                        lastTest[thisVariable]['thisCell'] = thisCell
                    elif thisVariable in lastTest:
                        name = lastTest[thisVariable]['name']
                        if name == '.':
                            thisCol += 1
                            continue
                        variable = lastTest[thisVariable]['variable']
                        test = lastTest[thisVariable]['test']
                        thisCell = lastTest[thisVariable]['thisCell']
                    else:
                        self.errors.append('Missing input test at {!s}'.format(coordinate))
                        thisCol += 1
                        continue
                    # print("Setting test at '{!s}' to '{!s}'".format(coordinate, test))
                    self.decisionTables[table]['inputColumns'][thisCol]['tests'].append((name, test))
                    thisCol += 1
                thisVariable += 1

            # Parse the vertical heading
            coordinate = cell.offset(row=1 + height).coordinate
            for merged in mergedCells:
                if coordinate in merged:
                    verticalRows = merged.max_row - merged.min_row + 1
                    break
            else:
                verticalRows = 1

            heading = cell.offset(row=1 + height).value
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

            thisVariable = 0
            while thisVariable < width - 1:
                thisRow = 0
                lastTest = {}
                while thisRow < verticalRows:
                    if thisVariable == 0:
                        self.decisionTables[table]['inputRows'].append({})
                        self.decisionTables[table]['inputRows'][thisRow]['tests'] = []

                    thisCell = ws.cell(row=startRow + 1 + height + thisRow, column=startCol + 1 + thisVariable).value
                    coordinate = ws.cell(row=startRow + 1 + height + thisRow, column=startCol + 1 + thisVariable).coordinate
                    if thisCell is not None:
                        thisCell = str(thisCell).strip()
                        if thisCell == '-':
                            if thisVariable in lastTest:
                                lastTest[thisVariable]['name'] = '.'
                            thisRow += 1
                            continue
                        name = inputs[thisVariable].strip()
                        variable = self.glossary[name]['item']
                        test = self.test2sfeel(variable, coordinate, thisCell)
                        lastTest[thisVariable] = {}
                        lastTest[thisVariable]['name'] = name
                        lastTest[thisVariable]['variable'] = variable
                        lastTest[thisVariable]['test'] = test
                        lastTest[thisVariable]['thisCell'] = thisCell
                    elif thisVariable in lastTest:
                        name = lastTest[thisVariable]['name']
                        if name == '.':
                            thisRow += 1
                            continue
                        variable = lastTest[thisVariable]['variable']
                        test = lastTest[thisVariable]['test']
                        thisCell = lastTest[thisVariable]['thisCell']
                    else:
                        self.errors.append('Missing input test at {!s}'.format(coordinate))
                        thisRow += 1
                        continue
                    # print("Setting test at '{!s}' to '{!s}'".format(coordinate, test))
                    self.decisionTables[table]['inputRows'][thisRow]['tests'].append((name, test))
                    thisRow += 1
                thisVariable += 1

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
                    thisCell = ws.cell(row=startRow + rows + row, column=startCol + cols + col).value
                    coordinate = ws.cell(row=startRow + rows + row, column=startCol + cols + col).coordinate
                    if thisCell is None:
                        self.errors.append('Missing output result at {!s}'.format(coordinate))
                        return (rows, cols, -1)
                    thisCell = str(thisCell).strip()
                    coordinate = ws.cell(row=startRow + rows + row, column=startCol + cols + col).coordinate
                    name = self.decisionTables[table]['output']['name']
                    variable = self.glossary[name]['item']
                    result = self.result2sfeel(variable, coordinate, thisCell)
                    # print("Setting result at '{!s}' to '{!s}'".format(coordinate, result))
                    self.rules[table][thisRule]['outputs'].append((name, result, 0))
                    thisRule += 1
            rows += verticalRows
            cols += horizontalCols

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
        mergedCells = ws.merged_cells.ranges
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
                            for merged in mergedCells:
                                if cell.coordinate in merged:
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
        mergedCells = ws.merged_cells.ranges
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
                            for merged in mergedCells:
                                if cell.coordinate in merged:
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
                            (rows, cols, rules) = self.parseDecionTable(ws, cell)
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
                            mergedCells = ws.merged_cells.ranges
                            for thisMerged in mergedCells:
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
                    ruleIds.append(table + ':' + str(self.rules[table][foundRule]['ruleId']))
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
                    ruleIds.append(table + ':' + str(self.rules[table][foundRule]['ruleId']))
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
                    ruleIds.append(table + ':' + str(self.rules[table][foundRule]['ruleId']))
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

        newData['Execute Rules'] = ruleIds
        status = {}
        if len(self.errors) > 0:
            status['errors'] = self.errors
            self.errors = []
        return (status, newData)


if __name__ == '__main__':

    dmnRules = pyDMNrules()
    status = dmnRules.load('Example1.xlsx')
    if 'errors' in status:
        print('With errors', status['errors'])

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
        print('With errors', status['errors'])

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
        print('With errors', status['errors'])

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
        print('With errors', status['errors'])
        sys.exit(0)

    data = {}
    data['Customer'] = 'Private'
    data['OrderSize'] = 9
    data['Delivery'] = 'slow'
    print('Testing',repr(data))
    (status, newData) = dmnRules.decide(data)
    print('Decision',repr(newData))
    if 'errors' in status:
        print('With errors', status['errors'])
