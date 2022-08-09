#!/usr/bin/env python

'''
A script to test pyDMNrules using the tck conformance suite of test


SYNOPSIS
$ python3 test_pyDMNrules_xml.py

REQUIRED

OPTIONS
-v loggingLevel|--verbose=loggingLevel
Set the level of logging that you want (defaut INFO).

-L logDir
The directory where the log file will be written.

-l logfile|--logfile=logfile
The name of a logging file where you want all messages captured.

'''

# Import all the modules that make life easy
import os
import sys
import glob
import logging
import argparse
import pyDMNrules
import pySFeel
import xml.etree.ElementTree as et
import re
import datetime
import dateutil

# The global variables
tckDir = '.'
lexer = pySFeel.SFeelLexer()
badFEELchars = u'[^?A-Z_a-z'
badFEELchars += u'\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u02FF'
badFEELchars += u'\u0370-\u037D\u037F-\u1FFF\u200C-\u200D\u2070-\u218F\u2C00-\u2FEF'
badFEELchars += u'\u3001-\uD7FF\uF900-\uFDCF\uFDF0\uFFFD'
badFEELchars += u'\U00010000-\U000EFFFF'
badFEELchars += u"0-9\u00B7\u0300-\u036F\u203F-\u2040\\.]"
decimalDigits = 8
ns = {}
rootPrefix = None
testNs = {}
testRootPrefix = None
brief = False
dictionary_outputs = ['0013-sort', '0069-feel-list', '0083-feel-unicode',     # Some tests expect dictionaries as outputs
                      '0057-feel-context', '0081-feel-getentries-function']

# This next section is plagurised from /usr/include/sysexits.h
EX_OK = 0        # successful termination
EX_WARN = 1        # non-fatal termination with warnings

EX_USAGE = 64        # command line usage error
EX_DATAERR = 65        # data format error
EX_NOINPUT = 66        # cannot open input
EX_NOUSER = 67        # addressee unknown
EX_NOHOST = 68        # host name unknown
EX_UNAVAILABLE = 69    # service unavailable
EX_SOFTWARE = 70    # internal software error
EX_OSERR = 71        # system error (e.g., can't fork)
EX_OSFILE = 72        # critical OS file missing
EX_CANTCREAT = 73    # can't create (user) output file
EX_IOERR = 74        # input/output error
EX_TEMPFAIL = 75    # temp failure; user is invited to retry
EX_PROTOCOL = 76    # remote error in protocol
EX_NOPERM = 77        # permission denied
EX_CONFIG = 78        # configuration error


def convertIn(thisValue, isTest):
    # convertIn converts data from a test file
    # The test file is XML, so the data will be a string
    # If this is a test input, and we are doing a FEEL test, then the returned value must be FEEL string
    # Otherwise is has to be a Python value
    if isinstance(thisValue, list):
        thisList = []
        for i in range(len(thisValue)):
            thisList.append(convertIn(thisValue[i], isTest))
        return thisList
    elif isinstance(thisValue, dict):
        thisContext = {}
        for key in thisValue:
            thisContext[key] = convertIn(thisValue[key], isTest)
        return thisContext
    if thisValue == '':
        return None
    if not isinstance(thisValue, str):
        print('convertIn: not a string:', thisValue)
        sys.stdout.flush()
        thisValue = str(thisValue)
    tokens = lexer.tokenize(thisValue)
    yaccTokens = []
    for token in tokens:
        yaccTokens.append(token)
    if len(yaccTokens) != 1:
        return thisValue
    elif yaccTokens[0].type == 'NUMBER':
        return float(thisValue)
    elif yaccTokens[0].type == 'BOOLEAN':
        if thisValue == 'true':
            return True
        elif thisValue == 'false':
            return False
        elif thisValue == 'null':
            return None
    elif yaccTokens[0].type == 'NAME':
        if thisValue == 'true':
            return True
        elif thisValue == 'True':
            return True
        elif thisValue == 'TRUE':
            return True
        elif thisValue == 'false':
            return False
        elif thisValue == 'False':
            return False
        elif thisValue == 'FALSE':
            return False
        elif thisValue == 'none':
            return None
        elif thisValue == 'None':
            return None
        elif thisValue == 'null':
            return None
        else:
            if isTest and ((thisValue[0] != '"') or (thisValue[-1] != '"')):
                return '"' + unicodeString(thisValue) + '"'
            else:
                return unicodeString(thisValue)
    elif yaccTokens[0].type == 'STRING':
        return unicodeString(thisValue[1:-1])
    elif yaccTokens[0].type == 'NUMBER':
        return float(thisValue)
    elif yaccTokens[0].type == 'DTDURATION':
        sign = 0
        if thisValue[0] == '-':
            sign = -1
            thisValue = thisValue[1:]     # skip -
        thisValue = thisValue[1:]         # skip P
        days = seconds = milliseconds = 0
        if thisValue.find('D') != -1:       # days is optional
            parts = thisValue.split('D')
            if len(parts[0]) > 0:
                days = int(parts[0])
            thisValue = parts[1]
        if len(thisValue) > 0:
            if thisValue[0] == 'T':
                thisValue = thisValue[1:]
        if len(thisValue) > 0:
            parts = thisValue.split('H')
            if len(parts) == 2:
                if len(parts[0]) > 0:
                    seconds = int(parts[0]) * 60 * 60
                thisValue = parts[1]
            parts = thisValue.split('M')
            if len(parts) == 2:
                if len(parts[0]) > 0:
                    seconds += int(parts[0]) * 60
                thisValue = parts[1]
            parts = thisValue.split('S')
            if len(parts) == 2:
                if len(parts[0]) > 0:
                    sPart = float(parts[0])
                    seconds += int(sPart)
                    milliseconds = int((sPart * 1000)) % 1000
        if sign == 0:
            return datetime.timedelta(days=days, seconds=seconds, milliseconds=milliseconds)
        else:
            return -datetime.timedelta(days=days, seconds=seconds, milliseconds=milliseconds)
    elif yaccTokens[0].type == 'YMDURATION':
        sign = 0
        if thisValue[0] == '-':
            sign = -1
            thisValue = thisValue[1:]     # skip -
        thisValue = thisValue[1:]         # skip P
        months = 0
        if len(thisValue) > 0:
            parts = thisValue.split('Y')
            if len(parts) == 2:
                if len(parts[0]) > 0:
                    months = int(parts[0]) * 12
                thisValue = parts[1]
        if len(thisValue) > 0:
            parts = thisValue.split('M')
            if len(parts) == 2:
                if len(parts[0]) > 0:
                    months += int(parts[0])
        if sign == 0:
            return int(months)
        else:
            return -int(months)
    elif yaccTokens[0].type == 'DATETIME':
        parts = thisValue.split('@')
        thisDateTime = dateutil.parser.parse(parts[0])
        if len(parts) > 1:
            thisZone = dateutil.tz.gettz(parts[1])
            if thisZone is not None:
                try:
                    thisDateTime = thisDateTime.replace(tzinfo=thisZone)
                except:
                    thisDateTime = thisDateTime
                thisDateTime = thisDateTime
        return thisDateTime
    elif yaccTokens[0].type == 'DATE':
        return dateutil.parser.parse(thisValue).date()
    elif yaccTokens[0].type == 'TIME':
        parts = thisValue.split('@')
        thisTime =  dateutil.parser.parse(parts[0]).timetz()     # A time with timezone
        if len(parts) > 1:
            thisZone = dateutil.tz.gettz(parts[1])
            if thisZone is not None:
                try:
                    thisTime = thisTime.replace(tzinfo=thisZone)
                except:
                    thisTime = thisTime
                thisTime = thisTime
        return thisTime
    else:
        return thisValue


def collectListTests(listElement):
    '''
    Collect at list of test values
    For FEEL the list of values need to be valid FEEL strings.
    For DMN the list of values needs to be Python values.
    '''

    global testNs, testRootPrefix

    items = listElement.findall(testRootPrefix + 'item', testNs)
    listData = []
    for item in items:
        value = item.find(testRootPrefix + 'value', testNs)
        newListElement = item.find(testRootPrefix + 'list', testNs)
        components = item.findall(testRootPrefix + 'component', testNs)
        if value is not None:
            if 'xsi' in testNs:
                if '{' + testNs['xsi'] + '}nil' in value.keys():
                    thisType = value.get('{' + testNs['xsi'] + '}nil')
                    if thisType == 'true':
                        listData.append(None)
                    else:
                        listData.append(convertIn(value.text, True))
                elif '{' + testNs['xsi'] + '}type' in value.keys():
                    thisType = value.get('{' + testNs['xsi'] + '}type')
                    if thisType == 'xsd:string':
                        if value.text is None:
                            listData.append('')
                        else:
                            listData.append(unicodeString(value.text))
                    elif (thisType == 'xsd:decimal') or (thisType == 'xsd:double'):
                        listData.append(float(value.text))
                    elif thisType == 'xsd:boolean':
                        if value.text == 'true':
                            listData.append(True)
                        else:
                            listData.append(False)
                    else:
                        listData.append(convertIn(value.text, True))
                else:
                    listData.append(convertIn(value.text, True))
            else:
                listData.append(convertIn(value.text, True))
        elif newListElement is not None:
            listData.append(collectListTests(newListElement))
            continue
        elif components != []:
            itemData = {}
            for component in components:
                itemData = itemData | collectContextTests(component)
            listData.append(itemData)
        else:
            print('Bad XML list item in tests file')
            logging.warning('Bad list item in XML tests file')
    return listData


def collectContextTests(component):
    '''
    Collect at dictionary of test values
    For FEEL the dictionary of values need to be valid FEEL contexts.
    For DMN the dictionary of values needs to be Python values.
    '''

    global testNs, testRootPrefix, dictionary_outputs

    context = {}
    variable = component.get('name')
    variable = unicodeString(variable)
    value = component.find(testRootPrefix + 'value', testNs)
    listElement = component.find(testRootPrefix + 'list', testNs)
    newComponent = component.find(testRootPrefix + 'component', testNs)
    if value is not None:
        if ('xsi' in testNs):
            if '{' + testNs['xsi'] + '}nil' in value.keys():
                thisType = value.get('{' + testNs['xsi'] + '}nil')
                if thisType == 'true':
                    context[variable] = 'null'
                else:
                    context[variable] = convertIn(value.text, True)
            elif '{' + testNs['xsi'] + '}type' in value.keys():
                thisType = value.get('{' + testNs['xsi'] + '}type')
                if thisType == 'xsd:string':
                    if value.text is None:
                        context[variable] = ''
                    else:
                        context[variable] = unicodeString(value.text)
                elif (thisType == 'xsd:decimal') or (thisType == 'xsd:double'):
                    context[variable] = float(value.text)
                elif thisType == 'xsd:boolean':
                    if value.text == 'true':
                        context[variable] = True
                    else:
                        context[variable] = False
                else:
                    context[variable] = convertIn(value.text, True)
            else:
                context[variable] = convertIn(value.text, True)
        else:
            context[variable] = convertIn(value.text, True)
    elif listElement is not None:
        context[variable] = collectListTests(listElement)
    elif newComponent is not None:
        if thisPattern in dictionary_outputs:
            context[variable] = {}
        newContexts = collectContextTests(newComponent)
        if thisPattern in dictionary_outputs:
            context[variable] |= newContexts
        else:
            for newVariable in newContexts:
                context[variable + '.' + newVariable] = newContexts[newVariable]
    else:
        print('Bad XML component in tests file')
        print(variable, component)
        print(component.keys)
        for child in component:
            print(child.tag, child.attrib)
        print(value)
        print(testNs)
        logging.warning('Bad component in XML tests file')
    return context


def collectTests():
    '''
    Collect the the tests and results from the tests XML file.
    For FEEL, tests become variables, to which values are assigned, using the FEEL <- assignment operator.
                           so the values need to be valid FEEL strings.
              results become Python values that match the values returned by pySFeel.
    For DMN, test become Python data that is passed to the pyDMNrules.decide() function.
              results become Python values that match the values returned by the pyDMNrules.decide() function.
    '''

    global testNs, testRootPrefix, pattern, thisPattern, dictionary_outputs

    tests = []
    testNames = []
    results = []
    decisionNames = []
    xmlFiles = glob.glob(pattern + '/' + thisPattern + '-test-01.xml')
    if len(xmlFiles) != 1:          # Check that we have a test XML file
        print('Missing XML test file for %s/%s' % (pattern, thisPattern))
        logging.warning('Missing XML tests file %s/%s.xml', pattern, thisPattern)
        return (tests, testNames, results, decisionNames)
    try:
        tree = et.parse(xmlFiles[0])
        root = tree.getroot()
    except:
        print('failed - Bad XML tests file ', pattern + '/' + thisPattern + '.xml (bad xml')
        logging.warning('failed - Bad XML tests file %s/%s.xml (bad xml)', pattern, thisPattern)
        return (tests, testNames, results, decisionNames)
    if root is None:
        print('failed - Bad XML tests file ', pattern + '/' + thisPattern + '.xml (missing root)')
        logging.warning('failed - Bad XML tests file %s/%s.xml (missing root)', pattern, thisPattern)
        return (tests, testNames, results, decisionNames)
    testRootPrefix = ''
    testRootNs = ''
    prefixEnd = (root.tag).find('}')
    testNs = {}         # Collect all the test namespaces
    for (event, elem) in et.iterparse(xmlFiles[0], events=['start-ns']):
        testNs[elem[0]] = elem[1]
        if prefixEnd != -1:
            if (elem[1] == root.tag[1:prefixEnd]):
                testRootNs = '{' + elem[1] + '}'
                if elem[0] != '':
                    testRootPrefix = elem[0] + ':'
        elif elem[0] == '':
            testRootNs = '{' + elem[1] + '}'
    if root.tag != testRootNs + 'testCases':
        print('failed - Bad XML tests file ', pattern + '/' + thisPattern + '.xml (no testCases)')
        logging.warning('failed - Bad XML tests file %s/%s.xml (no testCases)', pattern, thisPattern)
        return (tests, testNames, results, decisionNames)
    testNum = 1
    for test in root.findall(testRootPrefix + 'testCase', testNs):
        if 'id' in test.keys():
            id = test.get('id')
            if id.startswith('_'):
                testNames.append('{:03}'.format(testNum))
            else:
                testNames.append(test.get('id'))
        else:
            testNames.append(str(testNum))
        decisionNames.append([])
        testNum += 1
        data = {}       # The variable/value pairs
        for inputNode in test.findall(testRootPrefix + 'inputNode', testNs):
            variable = inputNode.get('name')
            variable = unicodeString(variable)
            value = inputNode.find(testRootPrefix + 'value', testNs)
            if value is not None:
                if 'xsi' in testNs:
                    if '{' + testNs['xsi'] + '}nil' in value.keys():
                        thisType = value.get('{' + testNs['xsi'] + '}nil')
                        if thisType == 'true':
                            data[variable] = None
                        else:
                            data[variable] = convertIn(value.text, True)
                    elif '{' + testNs['xsi'] + '}type' in value.keys():
                        thisType = value.get('{' + testNs['xsi'] + '}type')
                        if thisType == 'xsd:string':
                            if value.text is None:
                                data[variable] = ''
                            else:
                                data[variable] = unicodeString(value.text)
                        elif (thisType == 'xsd:decimal') or (thisType == 'xsd:double'):
                            data[variable] = float(value.text)
                        elif thisType == 'xsd:boolean':
                            if value.text == 'true':
                                data[variable] = True
                            else:
                                data[variable] = False
                        else:
                            data[variable] = convertIn(value.text, True)
                    else:
                        data[variable] = convertIn(value.text, True)
            listElement = inputNode.find(testRootPrefix + 'list', testNs)
            if listElement is not None:
                data[variable] = collectListTests(listElement)
                continue
            components = inputNode.findall(testRootPrefix + 'component', testNs)
            if components != []:
                if thisPattern in dictionary_outputs:
                    data[variable] = {}
                for component in components:
                    parts = collectContextTests(component)
                    if thisPattern in dictionary_outputs:
                        data[variable] |= parts
                    else:
                        for part in parts:
                            data[variable + '.' + part] = parts[part]
        tests.append(data)
        result = {}
        for resultNode in test.findall(testRootPrefix + 'resultNode', testNs):
            decisionName = resultNode.get('name')
            decisionNames[-1].append(decisionName)
            expected = resultNode.find(testRootPrefix + 'expected', testNs)
            value = expected.find(testRootPrefix + 'value', testNs)
            listElement = expected.find(testRootPrefix + 'list', testNs)
            components = expected.findall(testRootPrefix + 'component', testNs)
            if value is not None:
                variable = decisionName
                variable = unicodeString(variable)
                thisValue = value.text
                if 'xsi' in testNs:
                    if '{' + testNs['xsi'] + '}nil' in value.keys():
                        thisType = value.get('{' + testNs['xsi'] + '}nil')
                        if thisType == 'true':
                            result[variable] = None
                        else:
                            result[variable] = convertIn(thisValue, False)
                    elif '{' + testNs['xsi'] + '}type' in value.keys():
                        thisType = value.get('{' + testNs['xsi'] + '}type')
                        if thisType == 'xsd:string':
                            if thisValue is None:
                                result[variable] = ''
                            else:
                                result[variable] = unicodeString(str(thisValue))
                        elif (thisType == 'xsd:decimal') or (thisType == 'xsd:double'):
                            result[variable] = float(thisValue)
                        elif thisType == 'xsd:boolean':
                            if thisValue == 'true':
                                result[variable] = True
                            else:
                                result[variable] = False
                        else:
                            result[variable] = convertIn(thisValue, False)
                    else:
                        result[variable] = convertIn(thisValue, False)
                else:
                    result[variable] = convertIn(thisValue, False)
            elif listElement is not None:
                # pyDMNrules returns collections as a dictionary of lists - dmn-tck assumes a returned structure of a list of dictionaries
                if thisPattern in ['0109-ruleOrder-hitpolicy', '0110-outputOrder-hitpolicy', '0119-multi-collect-hitpolicy']:
                    result = collectListResults(listElement, decisionName)
                else:
                    result[decisionName] = collectListResults(listElement, decisionName)
            elif components != []:
                if thisPattern in dictionary_outputs:
                    variable = decisionName
                    result[variable] = {}
                for component in components:
                    if thisPattern in dictionary_outputs:
                        result[variable] |= { unicodeString(component.get('name')) : collectContextResults(component, variable) }
                    elif thisPattern in ['0096-feel-day-of-week-function', '0097-feel-month-of-year-function', '0098-feel-week-of-year-function',
                                         '1130-feel-interval']:
                        variable = unicodeString(component.get('name'))
                        result[variable] = collectContextResults(component, variable)
                    else:
                        variable = decisionName + '.' + unicodeString(component.get('name'))
                        result[variable] = collectContextResults(component, variable)
            else:
                print('failed - Bad XML tests file ', pattern + '/' + thisPattern + '.xml')
                logging.warning('failed - Bad XML tests file %s/%s.xml', pattern, thisPattern)
                return (tests, testNames, results, decisionNames)
        results.append(result)
    return (tests, testNames, results, decisionNames)


def collectListResults(listElement, variable):

    global testNs, testRootPrefix, thisPattern, dictionary_outputs

    items = listElement.findall(testRootPrefix + 'item', testNs)
    listData = []
    itemData = {}
    thisVariable = variable
    for item in items:
        value = item.find(testRootPrefix + 'value', testNs)
        newListElement = item.find(testRootPrefix + 'list', testNs)
        components = item.findall(testRootPrefix + 'component', testNs)
        if value is not None:
            if 'xsi' in testNs:
                if '{' + testNs['xsi'] + '}nil' in value.keys():
                    thisType = value.get('{' + testNs['xsi'] + '}nil')
                    if thisType == 'true':
                        listData.append(None)
                    else:
                        listData.append(convertIn(value.text, False))
                elif '{' + testNs['xsi'] + '}type' in value.keys():
                    thisType = value.get('{' + testNs['xsi'] + '}type')
                    if thisType == 'xsd:string':
                        if value.text is None:
                            listData.append('')
                        else:
                            listData.append(unicodeString(value.text))
                    elif (thisType == 'xsd:decimal') or (thisType == 'xsd:double'):
                        listData.append(float(value.text))
                    elif thisType == 'xsd:boolean':
                        if value.text == 'true':
                            listData.append(True)
                        else:
                            listData.append(False)
                    else:
                        listData.append(convertIn(value.text, False))
                else:
                    listData.append(convertIn(value.text, False))
            else:
                listData.append(convertIn(value.text, False))
        elif newListElement is not None:
            listData.append(collectListResults(newListElement, thisVariable))
        elif components != []:
            thisItemData = {}
            for component in components:
                if thisPattern in dictionary_outputs:
                    newVariable = component.get('name')
                    newVariable = unicodeString(newVariable)
                    thisItemData[newVariable]= collectContextResults(component, newVariable)
                else:
                    newVariable = variable + '.' + unicodeString(component.get('name'))
                    if newVariable not in itemData:
                        itemData[newVariable] = []
                    itemData[newVariable].append(collectContextResults(component, newVariable))
            if thisPattern in dictionary_outputs:
                listData.append(thisItemData)
        else:
            print('Bad XML list item in results file')
            logging.warning('Bad list item in XML results file')
    if itemData != {}:
       return itemData
    else:
        return listData


def collectContextResults(component, variable):

    global testNs, testRootPrefix, dictionary_outputs

    thisVariable = variable
    value = component.find(testRootPrefix + 'value', testNs)
    listElement = component.find(testRootPrefix + 'list', testNs)
    newComponents = component.findall(testRootPrefix + 'component', testNs)
    if value is not None:
        if ('xsi' in testNs):
            if '{' + testNs['xsi'] + '}nil' in value.keys():
                thisType = value.get('{' + testNs['xsi'] + '}nil')
                if thisType == 'true':
                    return 'null'
                elif '{' + testNs['xsi'] + '}type' in value.keys():
                    thisType = value.get('{' + testNs['xsi'] + '}type')
                    if thisType == 'xsd:string':
                        if value.text is None:
                            return ''
                        else:
                            return unicodeString(str(value.text))
                    elif (thisType == 'xsd:decimal') or (thisType == 'xsd:double'):
                        return float(value.text)
                    elif thisType == 'xsd:boolean':
                        if value == 'true':
                            return True
                        else:
                            return False
                    else:
                        return convertIn(value.text, False)
                else:
                    return convertIn(value.text, False)
            else:
                return convertIn(value.text, False)
        else:
            return unicodeString(value.text)
    elif listElement is not None:
        return collectListResults(listElement, variable)
    elif newComponents != []:
        itemData = {}
        for component in newComponents:
            if thisPattern in dictionary_outputs:
                newVariable = component.get('name')
                newVariable = unicodeString(newVariable)
            else:
                newVariable = variable + '.' + unicodeString(component.get('name'))
            itemData[newVariable] = collectContextResults(component, newVariable)
        return itemData
    else:
        print('Bad XML component in tests file')
        logging.warning('Bad component in XML tests file')
        return None


def logFailure(tests, results, data, newData, status):
    try:
        logging.debug('tests %s', str(tests))
    except:
        logging.debug('tests %s', str(tests).encode(errors='replace'))
    try:
        logging.debug('results %s', str(results))
    except:
        logging.debug('results %s', str(results).encode(errors='replace'))
    try:
        logging.debug('\tdata %s', str(data))
    except:
        logging.debug('\tdata %s', str(data).encode(errors='replace'))
    if 'errors' in status:
        if isinstance(status['errors'], list) and (len(status['errors']) > 0):
            for i in range(len(status['errors'])):
                if i == 0:
                    logging.debug("\tstatus['errors']\t[%s]", str(status['errors'][i]).encode(errors='replace'))
                else:
                    logging.debug("\t\t\t[%s]", str(status['errors'][i]).encode(errors='replace'))
        else:
            logging.debug("\tstatus['errors']\t[%s]", str(status['errors']).encode(errors='replace'))
    else:
        logging.debug("\tstatus[]\t%s", str(status).encode(errors='replace'))
    if isinstance(newData, list) and (len(newData) > 0):
        for i in range(len(newData)):
            if i == 0:
                logging.debug("\tnewData\t{'Result': %s", str(newData[i]).encode(errors='replace'))
            else:
                logging.debug("\t\t\t{'Result': %s", str(newData[i]).encode(errors='replace'))
    else:
        logging.debug('\tnewData\t%s', str(newData).encode(errors='replace'))
    logging.debug('')

def unicodeString(stringParam):
    if not isinstance(stringParam, str):
        return stringParam
    lastEnd = 0
    thisString = ''
    for match in re.finditer(r'(^|[^\\])\\(u([0-9A-Fa-f]{4})|U([0-9A-Fa-f]{6}))', stringParam):
        thisString += stringParam[lastEnd:match.end(1)]
        lastEnd = match.end(2)
        if match.group(3) is not None:
            thisGroup = match.group(3)
        else:
            thisGroup = match.group(4)
        thisHex = 0
        for i in range(len(thisGroup)):
            thisHex *= 16
            if thisGroup[i] in '0123456789':
                thisHex += ord(thisGroup[i]) - ord('0')
            elif thisGroup[i] in 'ABCDEF':
                thisHex += ord(thisGroup[i]) - ord('A') + 10
            elif thisGroup[i] in 'abcdef':
                thisHex += ord(thisGroup[i]) - ord('a') + 10
        thisString += chr(thisHex)
    thisString += stringParam[lastEnd:]
    return thisString


pattern = ''
thisPattern = ''

# The main code
if __name__ == '__main__':
    '''
The main code
Parse the command line arguments and set up general error logging.
Then process the file, named in the command line
    '''

    # Get the script name (without the '.py' extension)
    progName = os.path.basename(sys.argv[0])
    progName = progName[0:-3]        # Strip off the .py ending

    # Define the command line options
    parser = argparse.ArgumentParser(prog=progName)
    parser.add_argument ('-b', '--brief', action='store_true', help='Restrict output to brief report')
    parser.add_argument ('-v', '--verbose', dest='verbose', type=int, choices=range(0,5),
                         help='The level of logging\n\t0=CRITICAL,1=ERROR,2=WARNING,3=INFO,4=DEBUG')
    parser.add_argument ('-L', '--logDir', dest='logDir', default='.', help='The name of a logging directory')
    parser.add_argument ('-l', '--logFile', metavar='logFile', dest='logFile', help='The name of the logging file')

    # Parse the command line options
    args = parser.parse_args()
    brief = args.brief
    loggingLevel = args.verbose
    logDir = args.logDir
    logFile = args.logFile

    loggingLevel = args.verbose
    logDir = args.logDir
    logFile = args.logFile

    # Configure logging
    logging_levels = {0:logging.CRITICAL, 1:logging.ERROR, 2:logging.WARNING, 3:logging.INFO, 4:logging.DEBUG}
    logfmt = '%(filename)s [%(asctime)s]: %(message)s'
    if loggingLevel and (loggingLevel not in logging_levels) :
        sys.stderr.write('Error - invalid logging verbosity (%d)\n' % (loggingLevel))
        parser.print_usage(sys.stderr)
        sys.stderr.flush()
        sys.exit(EX_USAGE)
    if logFile :        # If sending to a file then check if the log directory exists
        # Check that the logDir exists
        if not os.path.isdir(logDir) :
            sys.stderr.write('Error - logDir (%s) does not exits\n' % (logDir))
            parser.print_usage(sys.stderr)
            sys.stderr.flush()
            sys.exit(EX_USAGE)
        if loggingLevel :
            logging.basicConfig(format=logfmt, datefmt='%d/%m/%y %H:%M:%S %p', level=logging_levels[loggingLevel],
                                filemode='w', filename=os.path.join(logDir, logFile))
        else :
            logging.basicConfig(format=logfmt, datefmt='%d/%m/%y %H:%M:%S %p',
                                filemode='w', filename=os.path.join(logDir, logFile))
        if not brief:
            print('Now logging to %s' % (os.path.join(logDir, logFile)))
        sys.stdout.flush()
    else :
        if loggingLevel :
            logging.basicConfig(format=logfmt, datefmt='%d/%m/%y %H:%M:%S %p', level=logging_levels[loggingLevel])
        else :
            logging.basicConfig(format=logfmt, datefmt='%d/%m/%y %H:%M:%S %p')
        if not brief:
            print('Now logging to sys.stderr')
        sys.stdout.flush()
    logging.info('Logging started')
    parser = pySFeel.SFeelParser()

    if brief:
        print('Level\tTest Group\tTest\tPass/Fail\tReason')

    # Process each conformance level
    for conformanceLevel in ['2', '3']:
    # for conformanceLevel in ['2']:
    # for conformanceLevel in ['3']:
        # Get all the patterns
        if not brief:
            print('Testing Conformance Level', conformanceLevel)
        patterns = glob.glob(tckDir + '/tck-master/TestCases/compliance-level-' + conformanceLevel + '/[0-9]*')
        # patterns = glob.glob(tckDir + '/tck-master/TestCases/compliance-level-' + conformanceLevel + '/1130*')
        logging.info('Testing Conformance Level %s', conformanceLevel)
        for pattern in patterns:
            badDMN = False
            if not brief:
                print('Testing Conformance Level', conformanceLevel)
                print('\tTesting', pattern)
                logging.info('Testing %s', pattern)
            thisPattern = os.path.basename(pattern)
            if not brief:
                print('\tTesting', thisPattern)
            logging.info('Testing %s', thisPattern)
            dmnFiles = glob.glob(pattern + '/' + thisPattern + '.dmn')
            if len(dmnFiles) != 1:
                if not brief:
                    print('Missing DMN file ', pattern + '/' + thisPattern + '.dmn')
                logging.warning('Missing DMN file %s/%s.dmn', pattern, thisPattern)
                badDMN = True
            xmlFile = open(dmnFiles[0], 'rt', newline='', encoding='utf-8')
            DMNtext = xmlFile.read()
            root = et.fromstring(DMNtext)
            ns = {}
            rootPrefix = ''
            rootNs = ''
            prefixEnd = (root.tag).find('}')
            for (event, elem) in et.iterparse(dmnFiles[0], events=['start-ns']):
                ns[elem[0]] = elem[1]
                if prefixEnd != -1:
                    if (elem[1] == root.tag[1:prefixEnd]):
                        rootNs = '{' + elem[1] + '}'
                        if elem[0] != '':
                            rootPrefix = elem[0] + ':'
                elif elem[0] == '':
                    rootNs = '{' + elem[1] + '}'
            if root.tag != rootNs + 'definitions':
                if not brief:
                    print('failed - Bad DMN file', dmnFiles[0], '- no definitions')
                logging.critical('failed - Bad DMN file %s - no definitions', dmnFiles[0])
                badDMN = True
            decisions = root.findall(rootPrefix + 'decision', ns)
            if decisions == []:
                if not brief:
                    print('failed - Bad DMN file', dmnFiles[0], '- no decisions')
                logging.critical('failed - Bad DMN file %s - no decisions', dmnFiles[0])
                badDMN = True
            dmnRules = pyDMNrules.DMN()
            status = dmnRules.useXML(DMNtext)
            if 'errors' in status:
                if not brief:
                    print('failed - Bad DMN file ', dmnFiles[0], '- could not be loaded by useXML()')
                logging.warning('failed - Bad DMN file %s - could not be loaded by useXML()', dmnFiles[0])
                logging.debug("\tstatus\t{'errors': [%s]}", status['errors'][0])
                logging.debug("\t\t{'errors': [%s]}", status['errors'][-1])
                badDMN = True
            (tests, testNames, results, decisionNames) = collectTests()
            if len(tests) > 0:
                for i in range(len(tests)):
                    failed = False
                    data = tests[i]
                    if not brief:
                        print('\t\tTest:', testNames[i], '(', decisionNames, ') testing - ', str(data))
                    logging.info('\t\tTest: %s ( %s ) testing - %s', testNames[i], decisionNames, str(data))
                    logging.info('\t\t\t expecting - %s', str(results[i]))
                    if badDMN:
                        print('{}\t{}\t{}\tfailed\tBad DMN file'.format(conformanceLevel, thisPattern, testNames[i]))
                        failed = True
                        continue
                    if thisPattern in ['0007-date-time', '0008-listGen',         # Some tests require previous decision outputs as inputs to subsequent decisions
                                       '0012-list-functions', '0016-some-every', '0020-vacation-days', '0035-test-structure-output']:
                        (status, newData) = dmnRules.decide(data)
                    else:
                        (status, newData) = dmnRules.decideTables(data, decisionNames[i])
                    if 'errors' in status:
                        if not brief:
                            try:
                                print('Test:', testNames[i], 'failed -\t', str(status['errors'][0]))
                            except:
                                print('Test:', testNames[i], 'failed -\t', str(status['errors'][0]).encode(errors='replace'))
                            try:
                                print('\t\t\t', str(status['errors'][-1]))
                            except:
                                print('\t\t\t', str(status['errors'][-1]).encode(errors='replace'))
                            try:
                                logging.info('Test: %s failed - %s', testNames[i], status['errors'][-1])
                            except:
                                logging.info('Test: %s failed - %s', testNames[i], str(status['errors'][-1]).encode(errors='replace'))
                            try:
                                logging.debug("\tstatus {'errors'[0]: [%s]}", status['errors'][0])
                            except:
                                logging.debug("\tstatus {'errors'[0]: [%s]}", str(status['errors'][0]).encode(errors='replace'))
                            try:
                                logging.debug("\tstatus {'errors'[-1]: [%s]}", status['errors'][-1])
                            except:
                                logging.debug("\tstatus {'errors'[-1]: [%s]}", str(status['errors'][-1]).encode(errors='replace'))
                            logFailure(tests, results, data, newData, status)
                        else:
                            print('{}\t{}\t{}\tfailed\t{}'.format(conformanceLevel, thisPattern, testNames[i], status['errors'][0]))
                            logFailure(tests, results, data, newData, status)
                            failed = True
                        continue
                    else:
                        logging.debug("\tstatus: %s", status)
                        if isinstance(newData, list) and (len(newData) > 0):
                            for j in range(len(newData)):
                                if j == 0:
                                    try:
                                        logging.debug("\tnewData\t%s", newData[j])
                                    except:
                                        logging.debug("\tnewData\t%s", str(newData[j]).encode(errors='replace'))
                                else:
                                    try:
                                        logging.debug("\t\t\t%s", newData[j])
                                    except:
                                        logging.debug("\t\t\t%s", str(newData[j]).encode(errors='replace'))
                        else:
                            try:
                                logging.debug('\tnewData\t%s', newData)
                            except:
                                logging.debug('\tnewData\t%s', str(newData).encode(errors='replace'))
                    if not isinstance(newData, list):
                        newData = [newData]
                    foundResults = []
                    for resultNum in range(len(newData)):
                        if isinstance(newData[resultNum]['Executed Rule'], list):
                            for j in range(len(newData[resultNum]['Executed Rule'])):
                                (thisDecision, thisTable, thisRule) = newData[resultNum]['Executed Rule'][j]
                                if thisTable in decisionNames[i]:
                                    foundResults.append(thisTable)
                        else:
                            (thisDecision, thisTable, thisRule) = newData[resultNum]['Executed Rule']
                            if thisTable in decisionNames[i]:
                                foundResults.append(thisTable)
                    missing = set(decisionNames[i]).difference(set(foundResults))
                    if len(missing) > 0:
                        print('{}\t{}\t{}\tfailed\tMissing results for Decision Tables {}'.format(conformanceLevel, thisPattern, testNames[i], missing))
                        failed = True
                        continue
                    for variable in results[i]:         # Check each variable
                        value = results[i][variable]
                        thisResult = len(newData) - 1
                        if variable not in newData[thisResult]['Result']:
                            if not brief:
                                print('Test:', testNames[i], 'failed - did not return a value for', variable)
                            else:
                                print('{}\t{}\t{}\tfailed\tdid not return a value for {}'.format(conformanceLevel, thisPattern, testNames[i], variable))
                                failed = True
                            logging.info('Test: %s failed - did not return a value for %s', testNames[i], variable)
                            logFailure(tests, results, data, newData, status)
                            break
                        if isinstance(value, list):     # We are expecting a list
                            if not isinstance(newData[thisResult]['Result'][variable], list):
                                if not brief:
                                    print('Test:', testNames[i], 'failed - returned value', newData[thisResult]['Result'][variable], 'for', variable, 'when', value, 'expected')
                                else:
                                    print('{}\t{}\t{}\tfailed\treturned value {} for variable {} when value {} expected'.format(conformanceLevel, thisPattern, testNames[i], newData[thisResult]['Result'][variable], variable, value))
                                    failed = True
                                logging.info('Test: %s failed - returned value %s for %s when %s expected', testNames[i], newData[thisResult]['Result'][variable],  variable, value)
                                logFailure(tests, results, data, newData, status)
                                break
                            for j in range(len(value)):     # Check each of the list items
                                thisValue = value[j]
                                if j >= len(newData[thisResult]['Result'][variable]):
                                    if not brief:
                                        print('Test:', testNames[i], 'failed - returned value', newData[thisResult]['Result'][variable], 'for', variable, 'when', value, 'expected')
                                    else:
                                        print('{}\t{}\t{}\tfailed\treturned value {} for variable {} when value {} expected'.format(conformanceLevel, thisPattern, testNames[i], newData[thisResult]['Result'][variable], variable, value))
                                        failed = True
                                    logging.info('Test: %s failed - returned value %s for %s when %s expected', testNames[i], newData[thisResult]['Result'][variable],  variable, value)
                                    logFailure(tests, results, data, newData, status)
                                    break
                                if isinstance(thisValue, float):
                                    thisValue = int(thisValue * (10 ** decimalDigits) + 0.5)/(10.0 ** decimalDigits)
                                    newData[thisResult]['Result'][variable][j] = int(float(newData[thisResult]['Result'][variable][j]) * (10 ** decimalDigits) + 0.5)/(10.0 ** decimalDigits)
                                if newData[thisResult]['Result'][variable][j] != thisValue:
                                    if not brief:
                                        print('Test:', testNames[i], 'failed - returned value', newData[thisResult]['Result'][variable], 'for', variable, 'when', value, 'expected')
                                    else:
                                        print('{}\t{}\t{}\tfailed\treturned value {} for variable {} when value {} expected'.format(conformanceLevel, thisPattern, testNames[i], newData[thisResult]['Result'][variable], variable, value))
                                        failed = True
                                    logging.info('Test: %s failed - returned value %s for %s when %s expected', testNames[i], newData[thisResult]['Result'][variable],  variable, value)
                                    logFailure(tests, results, data, newData, status)
                                    break
                            else:
                                continue
                            if not brief:
                                break
                        elif isinstance(value, dict):       # We are expecting a dictionary
                            if not isinstance(newData[thisResult]['Result'][variable], dict):
                                if not brief:
                                    print('Test:', testNames[i], 'failed - returned value', newData[thisResult]['Result'][variable], 'for', variable, 'when', value, 'expected')
                                else:
                                    print('{}\t{}\t{}\tfailed\treturned value {} for variable {} when value {} expected'.format(conformanceLevel, thisPattern, testNames[i], newData[thisResult]['Result'][variable], variable, value))
                                    failed = True
                                logging.info('Test: %s failed - returned value %s for %s when %s expected', testNames[i], newData[thisResult]['Result'][variable],  variable, value)
                                logFailure(tests, results, data, newData, status)
                                break
                            for thisVariable in value:      # Check each of the dictionary items
                                thisValue = value[thisVariable]
                                if thisVariable not in newData[thisResult]['Result'][variable]:
                                    if not brief:
                                        try:
                                            print('Test:', testNames[i], 'failed - did not return a value for', thisVariable)
                                        except:
                                            print('Test:', testNames[i], 'failed - did not return a value for', str(thisVariable).encode(errors='replace'))
                                    else:
                                        try:
                                            print('{}\t{}\t{}\tfailed\tdid not return a value for {}'.format(conformanceLevel, thisPattern, testNames[i], thisVariable))
                                        except:
                                            print('{}\t{}\t{}\tfailed\tdid not return a value for {}'.format(conformanceLevel, thisPattern, testNames[i], str(thisVariable).encode(errors='replace')))
                                        failed = True
                                    try:
                                        logging.info('Test: %s failed - did not return a value for %s', testNames[i], thisVariable)
                                    except:
                                        try:
                                            logging.info('Test: %s failed - did not return a value for %s', testNames[i], thisVariable)
                                        except:
                                            logging.info('Test: %s failed - did not return a value for %s', testNames[i], str(thisVariable).encode(errors='replace'))
                                    logFailure(tests, results, data, newData, status)
                                    break
                                if isinstance(thisValue, list):       # We are expecting a dictionary of lists
                                    if not isinstance(newData[thisResult]['Result'][thisVariable], list):
                                        if not brief:
                                            try:
                                                print('Test:', testNames[i], 'failed - returned value', newData[thisResult]['Result'][thisVariable], 'for', thisVariable, 'when', thisValue, 'expected')
                                            except:
                                                print('Test:', testNames[i], 'failed - returned value', str(newData[thisResult]['Result'][thisVariable]).encode(errors='replace'), 'for', str(thisVariable).encode(errors='replace'), 'when', str(thisValue).encode(errors='replace'), 'expected')
                                        else:
                                            try:
                                                print('{}\t{}\t{}\tfailed\treturned value {} for variable {} when value {} expected'.format(conformanceLevel, thisPattern, testNames[i], newData[thisResult]['Result'][thisVariable], thisVariable, thisValue))
                                            except:
                                                print('{}\t{}\t{}\tfailed\treturned value {} for variable {} when value {} expected'.format(conformanceLevel, thisPattern, testNames[i], str(newData[thisResult]['Result'][thisVariable]).encode(errors='replace'), str(thisVariable).encode(errors='replace'), str(thisValue).encode(errors='replace')))
                                            failed = True
                                        try:
                                            logging.info('Test: %s failed - returned value %s for %s when %s expected', testNames[i], newData[thisResult]['Result'][thisVariable],  thisVariable, thisValue)
                                        except:
                                            logging.info('Test: %s failed - returned value %s for %s when %s expected', testNames[i], str(newData[thisResult]['Result'][thisVariable]).encode(errors='replace'),  str(thisVariable).encode(errors='replace'), str(thisValue).encode(errors='replace'))
                                        logFailure(tests, results, data, newData, status)
                                        break
                                    for j in range(len(thisValue)):
                                        thisOneValue = thisValue[j]
                                        if j >= len(newData[thisResult]['Result'][thisVariable]):
                                            if not brief:
                                                print('Test:', testNames[i], 'failed - returned value', newData[thisResult]['Result'][variable], 'for', variable, 'when', value, 'expected')
                                            else:
                                                print('{}\t{}\t{}\tfailed\treturned value {} for variable {} when value {} expected'.format(conformanceLevel, thisPattern, testNames[i], newData[thisResult]['Result'][variable], variable, value))
                                                failed = True
                                            logging.info('Test: %s failed - returned value %s for %s when %s expected', testNames[i], newData[thisResult]['Result'][variable],  variable, value)
                                            logFailure(tests, results, data, newData, status)
                                            break
                                        if isinstance(thisOneValue, float):
                                            thisOneValue = int(thisOneValue * (10 ** decimalDigits) + 0.5)/(10.0 ** decimalDigits)
                                            newData[thisResult]['Result'][thisVariable][j] = int(float(newData[thisResult]['Result'][thisVariable][j]) * (10 ** decimalDigits) + 0.5)/(10.0 ** decimalDigits)
                                        if newData[thisResult]['Result'][thisVariable][j] != thisOneValue:
                                            if not brief:
                                                try:
                                                    print('Test:', testNames[i], 'failed - returned value', newData[thisResult]['Result'][thisVariable], 'for', thisVariable, 'when', thisValue, 'expected')
                                                except:
                                                    print('Test:', testNames[i], 'failed - returned value', str(newData[thisResult]['Result'][thisVariable]).encode(errors='replace'), 'for', str(thisVariable).encode(errors='replace'), 'when', str(thisValue).encode(errors='replace'), 'expected')
                                            else:
                                                try:
                                                    print('{}\t{}\t{}\tfailed\treturned value {} for variable {} when value {} expected'.format(conformanceLevel, thisPattern, testNames[i], newData[thisResult]['Result'][thisVariable]), thisVariable, thisValue)
                                                except:
                                                    print('{}\t{}\t{}\tfailed\treturned value {} for variable {} when value {} expected'.format(conformanceLevel, thisPattern, testNames[i], str(newData[thisResult]['Result'][thisVariable]).encode(errors='replace'), str(thisVariable).encode(errors='replace'), str(thisValue).encode(errors='replace')))
                                                failed = True
                                            try:
                                                logging.info('Test: %s failed - returned value %s for %s when %s expected', testNames[i], newData[thisResult]['Result'][thisVariable],  thisVariable, thisValue)
                                            except:
                                                logging.info('Test: %s failed - returned value %s for %s when %s expected', testNames[i], str(newData[thisResult]['Result'][thisVariable]).encode(errors='replace'),  str(thisVariable).encode(errors='replace'), str(thisValue).encode(errors='replace'))
                                            logFailure(tests, results, data, newData, status)
                                        break
                                else:           # Just a dictionary item
                                    if isinstance(thisValue, float):
                                        thisValue = int(thisValue * (10 ** decimalDigits) + 0.5)/(10.0 ** decimalDigits)
                                        if isinstance(newData[thisResult]['Result'][variable][thisVariable], float):
                                            newData[thisResult]['Result'][variable][thisVariable] = int(float(newData[thisResult]['Result'][variable][thisVariable]) * (10 ** decimalDigits) + 0.5)/(10.0 ** decimalDigits)
                                    if newData[thisResult]['Result'][variable][thisVariable] != thisValue:
                                        if not brief:
                                            try:
                                                print('Test:', testNames[i], 'failed - returned value', newData[thisResult]['Result'][variable][thisVariable], 'for', thisVariable, 'when', thisValue, 'expected')
                                            except:
                                                print('Test:', testNames[i], 'failed - returned value', str(newData[thisResult]['Result'][variable][thisVariable]).encode(errors='replace'), 'for', str(thisVariable).encode(errors='replace'), 'when', str(thisValue).encode(errors='replace'), 'expected')
                                        else:
                                            try:
                                                print('{}\t{}\t{}\tfailed\treturned value {} for variable {} when value {} expected'.format(conformanceLevel, thisPattern, testNames[i], newData[thisResult]['Result'][variable][thisVariable], thisVariable, thisValue))
                                            except:
                                                print('{}\t{}\t{}\tfailed\treturned value {} for variable {} when value {} expected'.format(conformanceLevel, thisPattern, testNames[i], str(newData[thisResult]['Result'][variable][thisVariable]).encode(errors='replace'), str(thisVariable).encode(errors='replace'), str(thisValue).encode(errors='replace')))
                                            failed = True
                                        try:
                                            logging.info('Test: %s failed - returned value %s for %s when %s expected', testNames[i], newData[thisResult]['Result'][variable][thisVariable],  thisVariable, thisValue)
                                        except:
                                            logging.info('Test: %s failed - returned value %s for %s when %s expected', testNames[i], str(newData[thisResult]['Result'][variable][thisVariable]).encode(errors='replace'),  str(thisVariable).encode(errors='replace'), str(thisValue).encode(errors='replace'))
                                        logFailure(tests, results, data, newData, status)
                                        break
                            else:
                                continue
                            if not brief:
                                break
                        else:       # Just a value, not a list or dictionary
                            if variable not in newData[thisResult]['Result']:
                                if not brief:
                                    try:
                                        print('Test:', testNames[i], 'failed - did not return a value for', variable)
                                    except:
                                        print('Test:', testNames[i], 'failed - did not return a value for', str(variable).encode(errors='replace'))
                                else:
                                    try:
                                        print('{}\t{}\t{}\tfailed\tdid not return a value for {}'.format(conformanceLevel, thisPattern, testNames[i], variable))
                                    except:
                                        print('{}\t{}\t{}\tfailed\tdid not return a value for {}'.format(conformanceLevel, thisPattern, testNames[i], str(variable).encode(errors='replace')))
                                    try:
                                        logging.debug("\t\tnewData:%s", newData[thisResult])
                                    except:
                                        logging.debug("\t\tnewData:%s", str(newData[thisResult]).encode(errors='replace'))
                                    failed = True
                                try:
                                    logging.info('Test: %s failed - did not return a value for %s', testNames[i], variable)
                                except:
                                    logging.info('Test: %s failed - did not return a value for %s', testNames[i], str(variable).encode(errors='replace'))
                                logFailure(tests, results, data, newData, status)
                                if not brief:
                                    break
                            if isinstance(value, float):
                                value = int(value * (10 ** decimalDigits) + 0.5)/(10.0 ** decimalDigits)
                                if isinstance(newData[thisResult]['Result'][variable], float):
                                    newData[thisResult]['Result'][variable] = int(float(newData[thisResult]['Result'][variable]) * (10 ** decimalDigits) + 0.5)/(10.0 ** decimalDigits)
                            elif isinstance(value, datetime.time) and isinstance(newData[thisResult]['Result'][variable], datetime.time):       # check for equivalent timezones (thisResult == +23)
                                resTime = newData[thisResult]['Result'][variable]
                                if (resTime.hour == value.hour) and (resTime.minute == value.minute) and (resTime.second == value.second) and (resTime.microsecond == value.microsecond):
                                    resTz = resTime.isoformat()[-6:]
                                    valueTz = value.isoformat()[-6:]
                                    resHr = int(resTz[1:3])
                                    resMin = int(resTz[4:6])
                                    valueHr = int(valueTz[1:3])
                                    valueMin = int(valueTz[4:6])
                                    if resTz[0] == '-':
                                        resTime = -(60 * resHr + resMin)
                                    else:
                                        resTime = 60 * resHr + resMin
                                    while resTime < 0:
                                        resTime += 24 * 60
                                    while resTime > 24 * 60:
                                        resTime -= 24 * 60
                                    if valueTz[0] == '-':
                                        valueTime = -(60 * valueHr + valueMin)
                                    else:
                                        valueTime = 60 * valueHr + valueMin
                                    while valueTime < 0:
                                        valueTime += 24 * 60
                                    while valueTime > 60 * 24:
                                        valueTime -= 24 * 60
                                    if resTime == valueTime:
                                        newData[thisResult]['Result'][variable] = value
                            if newData[thisResult]['Result'][variable] != value:
                                if not brief:
                                    try:
                                        print('Test:', testNames[i], 'failed - returned value', newData[thisResult]['Result'][variable], 'for', variable, 'when', value, 'expected')
                                    except:
                                        print('Test:', testNames[i], 'failed - returned value', newData[thisResult]['Result'][variable], 'for', str(variable).encode(errors='replace'), 'when', str(value).encode(errors='replace'), 'expected')
                                else:
                                    try:
                                        print('{}\t{}\t{}\tfailed\treturned value {} for variable {} when value {} expected'.format(conformanceLevel, thisPattern, testNames[i], newData[thisResult]['Result'][variable], variable, value))
                                    except:
                                        print('{}\t{}\t{}\tfailed\treturned value {} for variable {} when value {} expected'.format(conformanceLevel, thisPattern, testNames[i], str(newData[thisResult]['Result'][variable]).encode(errors='replace'), str(variable).encode(errors='replace'), value))
                                    failed = True
                                try:
                                    logging.info('Test: %s failed - returned value %s for %s when %s expected', testNames[i], newData[thisResult]['Result'][variable],  variable, value)
                                except:
                                    logging.info('Test: %s failed - returned value %s for %s when %s expected', testNames[i], str(newData[thisResult]['Result'][variable]).encode(errors='replace'),  str(variable).encode(errors='replace'), str(value).encode(errors='replace'))
                                logFailure(tests, results, data, newData, status)
                                if not brief:
                                    break
                    else:
                        if not failed:
                            if not brief:
                                print('\t\tTest:', testNames[i], 'passed')
                            else:
                                print('{}\t{}\t{}\tpassed'.format(conformanceLevel, thisPattern, testNames[i]))
                            logging.info('Test: %s passed', testNames[i])
                    if not brief:
                        break
            else:
                if not brief:
                    print('No tests in file ', pattern)
                else:
                    print('{}\t{}\t\tfailed\tNo tests in file'.format(conformanceLevel, thisPattern))
                logging.warning('No tests in file %s', pattern)
                continue

    sys.exit(EX_OK)
