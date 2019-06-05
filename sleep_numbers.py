""" Pulls information from excel spreadsheet and compiles those numbers into the stats excel spreadsheet. Written in
    Python 3.6 and depends on pandas, xlwings, datetime and math"""

import datetime # datetime.today
import math     # ceil
import sys

import pandas   # read_excel
import xlwings  # Book

""" Global variables for providers, techs and study type groupings """

PROVIDERS = ['kgw', 'qr', 'mb', 'jhf', 'ms']
TECHS = ['cfb', 'jls', 'kk', 'pgw', 'jjc', 'wmc', 'sl', 'kgw', 'kl']
PSG = ['Polysomnography', 'Polysomnography/wMSLT to follow', 'Polysomnography/wO2 end of study', 'Polysomnography/wRBD',
       'PostOp Polysomnography', 'Provent/wO2 Titration']

PSG_EEG = ['Polysomnography/wEEG']
PSG_ETCO2 = ['Polysomnography/wEEG-EtCO2', 'Polysomnography/wEtCO2']
SPLIT = ['Split-Night', 'Split-Night/wO2', 'Split-Night/wCPAP', 'Split-Night/wCPAP/wO2', 'Split-Night/wCPAP/VPAP',
         'Split-Night/wCPAP/VPAP/wO2', 'Split-Night/wCPAP/VPAP/ASV', 'Split-Night/wCPAP/VPAP/ASV/wO2',
         'Split-Night/wEEG', 'Split-Night/wEtCO2', 'Split-Night/wRBD', 'Split-Night/wVPAP/wO2', 'Split-Night/wVPAP/ASV',
         'Split-Night/wVPAP/ASV/wO2']

HST = ['ApneaLink +', 'ApneaLink Air', 'NOX-T3']
MSLT = ['MSLT', 'MSLT/wCPAP', 'MSLT/wCPAP/wO2']
MWT = ['MWT']
OAT = ['Matrix Titration', 'Oral Device Titration', 'Oral Device Titration/wO2',
       'Oral Device Titration/wO2w/PSG Follows', 'Polysomnography/wOral Device', 'PSG/wOral/wO2 end of study']

PAP = ['AdaptSV Titration', 'AdaptSV/wO2 Titration', 'BiPAP Trilogy', 'BiPAP Trilogy/wO2', 'C/V/ST/ASV Titration',
       'C/V/ST/ASV/wO2 Titration', 'C/VPAP/wRBD', 'C/VPAP/wO2/wRBD', 'CPAP Titration', 'CPAP Titration/wEEG',
       'CPAP Titration/wEEG/wO2', 'CPAP Titration/wEtCO2', 'CPAP Titration/wO2', 'CPAP Titration/wRBD',
       'CPAP Titration/wRBD/wO2', 'CPAP to Qualify for O2 - Medicare', 'CPAP to VPAP Titration',
       'CPAP to VPAP/wO2 Titration', 'CPAP/wOral Appliance', 'CPAP/wOral/wO2 Appliance', 'VPAP Titration/wEEG',
       'VPAP Titration/wEEG/wO2', 'VPAP ST Titration', 'VPAP ST/wO2 Titration', 'VPAP ST/ASV Titration',
       'VPAP ST/ASV/wO2 Titration', 'VPAP Titration', 'VPAP/wO2 Titration']

PAP_NAP = ['PAP-Nap']
FAILED_HST = ['Failed ApneaLink +', 'Failed ApneaLink Air', 'Failed NOX-T3']
NO_SHOW = ['Unable to tolerate CPAP', 'No Show', 'Rescheduled']
OTHER = ['WINX PSG', 'Provent Titration']
INSPIRE = ['Inspire']

""" Global variable for calculating the quarter """
NOW = datetime.date.today()
QUARTER = math.ceil(NOW.month / 3)

""" Globally opens excel to correct workbook """
#FILE = '\\\\Co.ihc.com\\swr\\DX\\Dept\\12600-28311\\Q&A for AASM\\2018-DRSDC-QA.xlsx'
FILE = "https://intermountainhealth.sharepoint.com/sites/DRMCSleepCenter/Shared Documents/IHC-SLEEP STUDIES-Y-T-D.xlsx"

try:
    WB = xlwings.Book(FILE)
except Exception as err:
    print(err)
    sys.exit()


def setup():

    """
    Opens the source excel workbook, reads all the data into a pandas dataframe, renames the columns and sorts
    the data into seperate dataframes by quarter
    @return 4 dataframes seprated by quarter:
    """

    srcfile = 'X:\\IHC-SLEEP STUDIES-Y-T-D.xlsx'

    # Make your DataFrame
    try:
        df = pandas.read_excel(srcfile, usecols='A, D, E, H:J, L, N, O, T', skiprows=9, header=None)
    except Exception as e:
        print(e)
        sys.exit()

    # Rename column names for ease of reference
    column_names = {
        0: 'Name',
        1: 'FIN',
        2: 'DOS',
        3: 'Delta Scored',
        4: 'Scoring Tech',
        5: 'Ordering Provider',
        6: 'Delta Interp',
        7: 'Delta FUV',
        8: 'Study Type',
        9: 'Failure',
    }
    df.rename(columns=column_names, inplace=True)

    # rename study types for easier sorting
    df['Study Type'].replace(HST, 'HST', inplace=True)
    df['Study Type'].replace(PSG, 'PSG', inplace=True)
    df['Study Type'].replace(PSG_EEG, 'PSG w/EEG', inplace=True)
    df['Study Type'].replace(PSG_ETCO2, 'PSG w/EtCO2', inplace=True)
    df['Study Type'].replace(SPLIT, 'Split', inplace=True)
    df['Study Type'].replace(MSLT, 'MSLT', inplace=True)
    df['Study Type'].replace(MWT, 'MWT', inplace=True)
    df['Study Type'].replace(OAT, 'OAT', inplace=True)
    df['Study Type'].replace(PAP, 'PAP', inplace=True)
    df['Study Type'].replace(PAP_NAP, 'PAP Nap', inplace=True)
    df['Study Type'].replace(FAILED_HST, 'Failed HST', inplace=True)
    df['Study Type'].replace(NO_SHOW, 'Failed in lab', inplace=True)
    df['Study Type'].replace(OTHER, 'Other', inplace=True)

    # setup masks for sorting by quarter
    # TODO: Setup so that I don't have to hard code the year
    q1mask = (df['DOS'] > '2018-1-1') & (df['DOS'] <= '2018-3-31')
    q1df = df.loc[q1mask]

    q2mask = (df['DOS'] > '2018-4-1') & (df['DOS'] <= '2018-6-30')
    q2df = df.loc[q2mask]

    q3mask = (df['DOS'] > '2018-7-1') & (df['DOS'] <= '2018-9-30')
    q3df = df.loc[q3mask]

    q4mask = (df['DOS'] > '2018-10-1') & (df['DOS'] <= '2018-12-31')
    q4df = df.loc[q4mask]

    return q1df, q2df, q3df, q4df


def calcHSTStats(df):

    """
    accepts dataframe and calculates the number of failed home studies and the average scoring time for
    successfull studies
    @param df:
    @return stats:
    """

    stats = []
    failed = df[df['Study Type'] == 'Failed HST']
    num_fails = failed.shape[0]
    df = df[df['Study Type'] == 'HST']
    scoring_ave = round(df['Delta Scored'].mean(), 2)
    interp_ave = round(df['Delta Interp'].mean(), 2)
    stats.append(scoring_ave)
    stats.append(interp_ave)
    stats.append(num_fails)
    return stats


def openSheet(sheet):

    """ Opens the excel sheet
    @param sheet:
    @return sheet:"""

    try:
        sheet = WB.sheets[sheet]
    except Exception:
        print("Sheet, {}, cannot be found".format(sheet))
        sys.exit()
    return sheet


def hstNumbersToExcel(stats, sheet):

    """ sends the stats to the correct location in excel
    @param stats, sheet:
    @return none:"""

    sheet = openSheet(sheet)
    sheet.range('V10').value = stats[0]
    sheet.range('V12').value = stats[1]
    sheet.range('V14').value = stats[2]


def calcStats(df):

    """ calculates stats for all of the studies done that quarter and saves them to a list
    @param df:
    @return stats:"""

    # Scoring stats
    scoring_ave = round(df['Delta Scored'].mean(), 2)
    scoring_stdev = round(df['Delta Scored'].std(), 2)
    scoring_min = df['Delta Scored'].min()
    scoring_max = df['Delta Scored'].max()
    scoring_stats = {
        'Ave': scoring_ave,
        "Stdv": scoring_stdev,
        'Min': scoring_min,
        'Max': scoring_max
    }
    # Interp stats
    interp_ave = round(df['Delta Interp'].mean(), 2)
    interp_stdev = round(df['Delta Interp'].std(), 2)
    interp_min = df['Delta Interp'].min()
    interp_max = df['Delta Interp'].max()
    interp_stats = {
        'Ave': interp_ave,
        "Stdv": interp_stdev,
        'Min': interp_min,
        'Max': interp_max
    }
    # list of tuples (what, score stat, interp stat)
    stats = [(k, scoring_stats[k], v) for k, v in interp_stats.items()]
    return stats


def numbersToExcel(stats, sheet):

    """ takes overall stats and sends them to the correct sheet in excel
    @param stats, sheet:
    @return none:"""

    scoring_cells = ['S34', 'S35', 'S36', 'S37']
    interp_cells = ['S38', 'S39', 'S40', 'S41']
    sheet = WB.sheets[sheet]
    a, scoring_stats, interp_stats = zip(*stats)
    for i in range(len(scoring_stats)):
        sheet.range(scoring_cells[i]).value = scoring_stats[i]
        sheet.range(interp_cells[i]).value = interp_stats[i]


def hstFailuresToExcel(df, sheet):

    """ takes dataframe and creates a list of the reasons that studies failed then sends them to excel
    @param dataframe, sheet:
    @return none:"""

    sheet = openSheet(sheet)

    failure_cells = ['D17', 'D18', 'D19', 'D20', 'D21', 'D22', 'D23', 'D24', 'D25', 'D26', 'D27', 'D28', 'D29', 'D30',
                     'D31', 'D32', 'D33', 'D34', 'D35']
    failures = df['Failure'].dropna().tolist()
    for i in range(len(failures)):
        sheet.range(failure_cells[i]).value = failures[i]


def printNumbers(q1df, q2df, q3df, q4df):

    """ setup stats for all studies and home studies by quarter and sends them to excel
    @param dataframes filtered by quarter, q1df, q2df, q3df, q4df:
    @return none:"""

    # stats are list of tuples (label, scoring stat, interp stat)
    if QUARTER > 1:
        q1_stats = calcStats(q1df)
        numbersToExcel(q1_stats, 'Q1')
        hstFailuresToExcel(q1df, 'HST - Q1')
        q1_hst_stats = calcHSTStats(q1df)
        hstNumbersToExcel(q1_hst_stats, 'HST - Q1')
        if QUARTER > 2:
            q2_stats = calcStats(q2df)
            numbersToExcel(q2_stats, 'Q2')
            hstFailuresToExcel(q2df, 'HST - Q2')
            q2_hst_stats = calcHSTStats(q2df)
            hstNumbersToExcel(q2_hst_stats, 'HST - Q2')
            if QUARTER > 3:
                q3_stats = calcStats(q3df)
                numbersToExcel(q3_stats, 'Q3')
                hstFailuresToExcel(q3df, 'HST - Q3')
                q3_hst_stats = calcHSTStats(q3df)
                hstNumbersToExcel(q3_hst_stats, 'HST - Q3')
                if QUARTER > 4:
                    q4_stats = calcStats(q4df)
                    numbersToExcel(q4_stats, 'Q4')
                    hstFailuresToExcel(q4df, 'HST - Q4')
                    q4_hst_stats = calcHSTStats(q4df)
                    hstNumbersToExcel(q4_hst_stats, 'HST - Q4')


def sendToExcel(studies, sheet, provider):

    """ sends the study counts to the correct cell in excel
    @param studies, sheet, provider:
    @return none:"""

    # excel sheet passed in from previous functions
    sheet = openSheet(sheet)

    # mapping for spreadsheet locations
    kgw_cells = {'PSG': 'K12', 'EEG': 'K13', 'ETCO2': 'K14', 'SPLIT': 'K15', 'HST': 'K16', 'MSLT': 'K17', 'MWT': 'K18',
                 'OAT': 'K19', 'Inspire': 'K20', 'PAP': 'K21', 'NAP': 'K22', 'FHST': 'K23', 'NS': 'K24'}
    mb_cells = {'PSG': 'P12', 'EEG': 'P13', 'ETCO2': 'P14', 'SPLIT': 'P15', 'HST': 'P16', 'MSLT': 'P17', 'MWT': 'P18',
                 'OAT': 'P19', 'Inspire': 'P20', 'PAP': 'P21', 'NAP': 'P22', 'FHST': 'P23', 'NS': 'P24'}
    qr_cells = {'PSG': 'U12', 'EEG': 'U13', 'ETCO2': 'U14', 'SPLIT': 'U15', 'HST': 'U16', 'MSLT': 'U17', 'MWT': 'U18',
                 'OAT': 'U19', 'Inspire': 'U20', 'PAP': 'U21', 'NAP': 'U22', 'FHST': 'U23', 'NS': 'U24'}
    jf_cells = {'PSG': 'AA12', 'EEG': 'AA13', 'ETCO2': 'AA14', 'SPLIT': 'AA15', 'HST': 'AA16', 'MSLT': 'AA17',
                'MWT': 'AA18', 'OAT': 'AA19', 'Inspire': 'AA20', 'PAP': 'AA21', 'NAP': 'AA22', 'FHST': 'AA23', 'NS': 'AA24'}
    dr_cells = {'PSG': 'AG12', 'EEG': 'AG13', 'ETCO2': 'AG14', 'SPLIT': 'AG15', 'HST': 'AG16', 'MSLT': 'AG17',
                'MWT': 'AG18', 'OAT': 'AG19', 'Inspire': 'AG20', 'PAP': 'AG21', 'NAP': 'AG22', 'FHST': 'AG23', 'NS': 'AG24'}

    # assign correct cell mapping by provider
    if provider == 'kgw':
        cells = kgw_cells
    elif provider == 'mb':
        cells = mb_cells
    elif provider == 'qr':
        cells = qr_cells
    elif provider == 'jf':
        cells = jf_cells
    elif provider == 'dr':
        cells = dr_cells
    else:
        print('Provider {} is unknown'.format(provider))
        return
        
    if studies[0] == 'PSG':
        sheet.range(cells['PSG']).value = studies[1]
    elif 'EEG' in studies[0]:
        sheet.range(cells['EEG']).value = studies[1]
    elif 'EtCO2' in studies[0]:
        sheet.range(cells['ETCO2']).value = studies[1]
    elif 'Split' in studies[0]:
        sheet.range(cells['SPLIT']).value = studies[1]
    elif studies[0] == 'HST':
        sheet.range(cells['HST']).value = studies[1]
    elif 'MSLT' in studies[0]:
        sheet.range(cells['MSLT']).value = studies[1]
    elif 'MWT' in studies[0]:
        sheet.range(cells['MWT']).value = studies[1]
    elif 'OAT' in studies[0]:
        sheet.range(cells['OAT']).value = studies[1]
    elif studies[0] == 'PAP':
        sheet.range(cells['PAP']).value = studies[1]
    elif studies[0] == 'PAP Nap':
        sheet.range(cells['NAP']).value = studies[1]
    elif studies[0] == 'Failed HST':
        sheet.range(cells['FHST']).value = studies[1]
    elif studies[0] == 'Failed in lab':
        sheet.range(cells['NS']).value = studies[1]
    elif studies[0] == 'Inspire':
        sheet.range(cells['Inspire']).value = studies[1]
    elif 'Other' in studies[0]:
        print(studies[0], studies[1])
    else:
        print('Study type, {} is not indexed'.format(studies[0]))
        return


def setupProviderNumbersForExport(kgw, jhf, qr, mb, dr, sheet):

    """ prepares the provider study counts to be sent to excel
    @param kgw, jhf, qr, mb, dr, sheet:
    @return none:"""

    for kgw_studies in kgw:
        sendToExcel(kgw_studies, sheet, 'kgw')
    for jf_studies in jhf:
        sendToExcel(jf_studies, sheet, 'jf')
    for qr_studies in qr:
        sendToExcel(qr_studies, sheet, 'qr')
    for mb_studies in mb:
        sendToExcel(mb_studies, sheet, 'mb')
    for dr_studies in dr:
        sendToExcel(dr_studies, sheet, 'dr')


def unpackProviderNumbers(df, sheet):

    """ converts the dataframe into a dictionary and then unpacks specific data into seperate lists
    @param df, sheet:
    @return none:"""

    kgw, jhf, qr, mb, dr = ([] for i in range(5))  # 5 empty lists
    studies = dict(df.size())
    for (a, k), v in studies.items():
        if a == 'kgw':
            kgw.append((k, v))
        elif a == 'jhf':
            jhf.append((k, v))
        elif a == 'qr':
            qr.append((k, v))
        elif a == 'mb':
            mb.append((k, v))
        elif a == 'ms':
            dr.append((k, v))
        else:
            print('Unknown provider')
    setupProviderNumbersForExport(kgw, jhf, qr, mb, dr, sheet)


def printNumbersByProvider(q1df, q2df, q3df, q4df):

    """ sends numbers by quarter for sorting and export to excel
    @param q1df, q2df, q3df, q4df:
    @return none:"""

    # Sort by provider and study type then print
    if QUARTER > 1:
        q1 = q1df.groupby(['Ordering Provider', 'Study Type'])
        unpackProviderNumbers(q1, 'Q1')
        if QUARTER > 2:
            q2 = q2df.groupby(['Ordering Provider', 'Study Type'])
            unpackProviderNumbers(q2, 'Q2')
            if QUARTER > 3:
                q3 = q3df.groupby(['Ordering Provider', 'Study Type'])
                unpackProviderNumbers(q3, 'Q3')
                if QUARTER > 4:
                    q4 = q4df.groupby(['Ordering Provider', 'Study Type'])
                    unpackProviderNumbers(q4, 'Q4')


if __name__ == "__main__":
    q1df, q2df, q3df, q4df = setup()
    printNumbers(q1df, q2df, q3df, q4df)
    printNumbersByProvider(q1df, q2df, q3df, q4df)
