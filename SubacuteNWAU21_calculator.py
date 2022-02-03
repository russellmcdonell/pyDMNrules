#!/usr/bin/env python

'''
A script to compute NWAU21 for a batch of Subacute episodes.


SYNOPSIS
$ python3 SubacuteNWAU21_calculator.py

REQUIRED
inFilename
The name of the Excel workbook of Subacute episode data (with headers) - the output of the grouper

outFilename
The name of the output Excel workbook


OPTIONS
'''

# Import all the modules that make life easy
import os
import sys
import argparse
import pandas as pd
import pyDMNrules


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
    parser.add_argument('inFilename', help='The name of the Excel workbook of grouped Subacute episode data (with headers)')
    parser.add_argument('outFilename', help='The name output Excel workbook')

    # Parse the command line options
    args = parser.parse_args()
    inFilename = args.inFilename
    outFilename = args.outFilename

    # Create the Subacute NWAU21 calculator
    NWAU21dmnRules = pyDMNrules.DMN()
    status = NWAU21dmnRules.load('Subacute NWAU21 calculator (DMN).xlsx')
    if status != {}:
        print(status)
        sys.exit(0)

    # Read in the grouped Subacute Episodic data
    # We need dtype because some of the things we code a 'codes' look like numbers
    dtype = {'AN-SNAP V4 code':str, 'Hospital Remoteness':str, 'Postcode':str, 'SA2':str,
             'Funding Source':str, 'Indigenous Status':str, 'State':str}
    dfInput = pd.read_excel(inFilename, dtype=dtype)

    # Assemble the Episode data - we need headings to map some of the column names
    headings = {'AN-SNAP V4 code':'AN-SNAP V4.0', 'Same-day admitted care':'Same Day Admission'}
    # map the Care Type to the NWAU21 Care Type codes
    dfInput.loc[dfInput['Care Type'] == 'Rehabilitation', 'Care Type'] = '6'
    dfInput.loc[dfInput['Care Type'] == 'GEM', 'Care Type'] = '8'
    dfInput.loc[dfInput['Care Type'] == 'Palliative Care', 'Care Type'] = '9'
    # Assemble the Patient data - Patient Age and Indigenous Status are in the extract.
    # Assemble the Funding data - State is in the extract.

    # Cost all the records
    (dfStatus, dfResults, dfDecision) = NWAU21dmnRules.decidePandas(dfInput, headings=headings)
    for index, value in dfStatus.items():
        if value != 'no errors':
            print('Error in row', index + 1, ':', value)
            print(dfResults.to_string())
            print(dfDecision.to_string())
            sys.exit(0)

    # Save the output
    dfOutput = pd.concat([dfInput, dfResults], axis=1)
    dfOutput.to_excel(outFilename, index=False)

