#!/usr/bin/env python

'''
A script to compute AN-SNAP version 4 codes for a batch of Subacute episodes.


SYNOPSIS
$ python3 AN-SNAP_V4_grouper.py

REQUIRED
inFilename
The name of the Excel workbook of subacute episode data (with headers)

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
    parser.add_argument('inFilename', help='The name of the Excel workbook of episode data (with headers)')
    parser.add_argument('outFilename', help='The name output Excel workbook')

    # Parse the command line options
    args = parser.parse_args()
    inFilename = args.inFilename
    outFilename = args.outFilename

    # Create the AN-SNAP version 4 grouper
    SNAPdmnRules = pyDMNrules.DMN()
    status = SNAPdmnRules.load('AN-SNAP V4 grouper (DMN).xlsx')
    if status != {}:
        print(status)
        sys.exit(0)

    # Read in the Episodic data - which must be already sorted in Patient UR, Phase Start Date ascending order
    dfInput = pd.read_excel(inFilename)

    # Assemble the Episode data - Admitted Flag, Care Type, Length of Stay, Same-day admitted care, Multidisciplinary and GEM clinic are in the extract
    dfInput['Long term care'] = False
    dfInput.loc[dfInput['Length of Stay'] >= 92, 'Long term care'] = True
    # Assemble the Patient data - Patient Age is in the extract. Patient Age Type is computed
    # Assemble the Rehab_and_nonAcute data - AROC code is in the extract
    # Assemble the Rehab_and_GEM data - FIM Motor score and FIM Cognition score are in the extract
    # Assemble the GEM_admitted data - Delerium or Dimentia is in the extract
    # Assemble the GEM_non_admitted data - Single Day of Care, Ongoing Pain and Clinic are in the extract
    # Assemble the PalliativeCare data - Phase Type is in the extract
    dfInput['First Phase'] = False
    thisPatient = thisEpisodeStartDate = None
    for index, row in dfInput.iterrows():
        if ((row['Patient UR'] != thisPatient) or (row['Episode Start Date'] != thisEpisodeStartDate)):
            thisPatient = row['Patient UR']
            thisEpisodeStartDate = row['Episode Start Date']
            if row['Phase Type'] == 'Unstable':
                dfInput.loc[index, 'First Phase'] = True
    # Assemble the Psychogeriatric data - HoNOS 65+ ADL, HoNOS 65 + Total, Focus of Care and Overactive Behaviour are in the extract
    # Assemble the PalliativeCare_and_NonAcute  data - RUG-ADL is in the extract
    # Assemble the NonAcute_and_Pyschogeriatric  data - Assessment Only is in the extract

    # Group all the records
    (dfStatus, dfResults, dfDecision) = SNAPdmnRules.decidePandas(dfInput)
    for index, value in dfStatus.items():
        if value != 'no errors':
            print('Error in row', index + 1, ':', value)
            sys.exit(0)

    # Save the output
    dfOutput = pd.concat([dfInput, dfResults], axis=1)
    dfOutput.to_excel(outFilename, index=False)
