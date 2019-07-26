"""
This is a test of validation_metrics.py against Excel spreadsheet calculations
provided by B. Pengra and S. Stehman. The intent is to verify that the return
values from the functions within validate.py match specific cell values from
the spreadsheet file. This supplies realistic values but may not test all edge
cases.

Authors: D. Wellington
"""

import xlrd
import numpy as np
import copy


def np_read_excel_column(sheet, cell_start, col_size):

    column_slice = sheet.col_slice(cell_start[1], cell_start[0], cell_start[0]+col_size)
    return np.array([x.value if x.ctype == xlrd.XL_CELL_NUMBER else 0 for x in column_slice])


def np_read_excel_row(sheet, cell_start, row_size):

    row_slice = sheet.row_slice(cell_start[0], cell_start[1], cell_start[1]+row_size)
    return np.array([x.value if x.ctype == xlrd.XL_CELL_NUMBER else 0 for x in row_slice])


def np_read_excel_matrix(sheet, cell_start, row_size, col_size):

    matrix = np.empty((0, row_size))
    for N in range(col_size):
        row = np_read_excel_row(sheet, (cell_start[0]+N, cell_start[1]), row_size)
        matrix = np.append(matrix, [row], axis=0)
    return matrix


def get_data_from_excel(sheet, start, shape):

    if shape == (1, 1):
        return sheet.cell_value(*start)
    elif shape[0] == 1:
        return np_read_excel_row(sheet, start, shape[1])
    elif shape[1] == 1:
        return np_read_excel_column(sheet, start, shape[0])
    else:
        return np_read_excel_matrix(sheet, start, shape[1], shape[0])


def main():

    import run_validation_metrics

    run_validation_metrics.main('plots/RefandMap-update.csv',
                                'plots/First50K_plots.xls',
                                '/lcmap_data/bulk/ancillary/NLCD/Original/nlcd_2001_landcover_2011_edition_2014_10_10/nlcd_2001_landcover_2011_edition_2014_10_10/nlcd_2001_landcover_2011_edition_2014_10_10.img',
                                'plots/MapCounts_prototype_full.csv',
                                outdir='test_sheets')

    DATA = {
        'N_CLASSES': 8,
        'N_YEARS': 32,
        'N_INVALID': 93
    }

    BOOKS_TO_CHECK = [
        {'file': 'test_sheets/1_Annual_LandCover_Agreement.xlsx',
         'ref_file': 'spreadsheets/bruce/1_Annual_PSE_Area&AgreementTables.xlsx',
         'sheets_to_skip': ['All Years'],
         'ref_sheets_to_skip': ['Template'],
         'cell_locations': [
             {'name': "Error matrix",
              'shape': (DATA['N_CLASSES'], DATA['N_CLASSES']),
              'ref_shape': (DATA['N_CLASSES'], DATA['N_CLASSES']),
              'start': (2, 2),
              'ref_start': (2, 2),
              },
             {'name': "Reference totals",
              'shape': (1, DATA['N_CLASSES']),
              'ref_shape': (1, DATA['N_CLASSES']),
              'start': (10, 2),
              'ref_start': (10, 2),
              },
             {'name': "Map totals",
              'shape': (DATA['N_CLASSES'], 1),
              'ref_shape': (DATA['N_CLASSES'], 1),
              'start': (2, 10),
              'ref_start': (2, 10),
              },
             {'name': "Total",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (10, 10),
              'ref_start': (10, 10),
              },
             {'name': "Producer's accuracy",
              'shape': (1, DATA['N_CLASSES']),
              'ref_shape': (1, DATA['N_CLASSES']),
              'start': (13, 2),
              'ref_start': (12, 2),
              },
             {'name': "Producer's standard error",
              'shape': (1, DATA['N_CLASSES']),
              'ref_shape': (1, DATA['N_CLASSES']),
              'start': (14, 2),
              'ref_start': (13, 2),
              },
             {'name': "User's accuracy",
              'shape': (DATA['N_CLASSES'], 1),
              'ref_shape': (DATA['N_CLASSES'], 1),
              'start': (2, 12),
              'ref_start': (2, 12),
              },
             {'name': "User's standard error",
              'shape': (DATA['N_CLASSES'], 1),
              'ref_shape': (DATA['N_CLASSES'], 1),
              'start': (2, 13),
              'ref_start': (2, 13),
              },
             {'name': "Poststratified producer's accuracy",
              'shape': (1, DATA['N_CLASSES']),
              'ref_shape': (1, DATA['N_CLASSES']),
              'start': (17, 2),
              'ref_start': (14, 2),
              },
             {'name': "Poststratified producer's standard error",
              'shape': (1, DATA['N_CLASSES']),
              'ref_shape': (1, DATA['N_CLASSES']),
              'start': (18, 2),
              'ref_start': (15, 2),
              },
             {'name': "Accuracy overall",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (13, 13),
              'ref_start': (13, 11),
              },
             {'name': "Poststratified accuracy overall",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (14, 13),
              'ref_start': (12, 11),
              },
             {'name': "Poststratified standard error overall",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (15, 13),
              'ref_start': (14, 13),
              },
             {'name': "Dice coefficients",
              'shape': (DATA['N_CLASSES'], 1),
              'ref_shape': (DATA['N_CLASSES'], 1),
              'start': (2, 15),
              'ref_start': (2, 15),
              },
             {'name': "Area estimate",
              'shape': (DATA['N_CLASSES'], 1),
              'ref_shape': (DATA['N_CLASSES'], 1),
              'start': (2, 17),
              'ref_start': (2, 26),
              },
             {'name': "Area estimate standard error",
              'shape': (DATA['N_CLASSES'], 1),
              'ref_shape': (DATA['N_CLASSES'], 1),
              'start': (2, 18),
              'ref_start': (2, 27),
              },
         ]},
        {'file': 'test_sheets/1_Annual_LandCover_Agreement.xlsx',
         'ref_file': 'spreadsheets/bruce/2_ErrorMatrix_AllYearsAggregated.xlsx',
         'sheets_to_skip': [str(x) for x in range(1985, 1985 + DATA['N_YEARS'])],
         'ref_sheets_to_skip': [],
         'cell_locations': [
             {'name': "Error matrix",
              'shape': (DATA['N_CLASSES'], DATA['N_CLASSES']),
              'ref_shape': (DATA['N_CLASSES'], DATA['N_CLASSES']),
              'start': (2, 2),
              'ref_start': (2, 2),
              },
             {'name': "Reference totals",
              'shape': (1, DATA['N_CLASSES']),
              'ref_shape': (1, DATA['N_CLASSES']),
              'start': (10, 2),
              'ref_start': (10, 2),
              },
             {'name': "Map totals",
              'shape': (DATA['N_CLASSES'], 1),
              'ref_shape': (DATA['N_CLASSES'], 1),
              'start': (2, 10),
              'ref_start': (2, 10),
              },
             {'name': "Total",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (10, 10),
              'ref_start': (10, 10),
              },
             {'name': "Producer's accuracy",
              'shape': (1, DATA['N_CLASSES']),
              'ref_shape': (1, DATA['N_CLASSES']),
              'start': (13, 2),
              'ref_start': (12, 2),
              },
             {'name': "User's accuracy",
              'shape': (DATA['N_CLASSES'], 1),
              'ref_shape': (DATA['N_CLASSES'], 1),
              'start': (2, 12),
              'ref_start': (2, 12),
              },
             {'name': "Accuracy overall",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (13, 13),
              'ref_start': (12, 11),
              },
             {'name': "Dice coefficients",
              'shape': (DATA['N_CLASSES'], 1),
              'ref_shape': (DATA['N_CLASSES'], 1),
              'start': (2, 15),
              'ref_start': (2, 15),
              },
         ]},
        {'file': 'test_sheets/2_Annual_LandCover_Agreement_Grouped.xlsx',
         'ref_file': 'spreadsheets/bruce/1b_AnnualProducersUsersOverallAgreementTables&Graphs.xlsx',
         'sheets_to_skip': ['Area'],
         'ref_sheets_to_skip': [],
         'cell_locations': [
             {'name': "User's",
              'shape': (DATA['N_CLASSES'], DATA['N_YEARS']),
              'ref_shape': (DATA['N_CLASSES'], DATA['N_YEARS']),
              'start': (1, 1),
              'ref_start': (1, 1),
              },
             {'name': "Producer's",
              'shape': (DATA['N_CLASSES'], DATA['N_YEARS']),
              'ref_shape': (DATA['N_CLASSES'], DATA['N_YEARS']),
              'start': (10, 1),
              'ref_start': (10, 1),
              },
             {'name': "Overall",
              'shape': (1, DATA['N_YEARS']),
              'ref_shape': (1, DATA['N_YEARS']),
              'start': (19, 1),
              'ref_start': (19, 1),
              },
             {'name': "Producer's - PSE",
              'shape': (DATA['N_CLASSES'], DATA['N_YEARS']),
              'ref_shape': (DATA['N_CLASSES'], DATA['N_YEARS']),
              'start': (21, 1),
              'ref_start': (22, 1),
              },
             {'name': "Overall - PSE",
              'shape': (4, DATA['N_YEARS']),
              'ref_shape': (4, DATA['N_YEARS']),
              'start': (30, 1),
              'ref_start': (31, 1),
              },
             {'name': "Overall Agreement per Class",
              'shape': (DATA['N_CLASSES'], DATA['N_YEARS']),
              'ref_shape': (DATA['N_CLASSES'], DATA['N_YEARS']),
              'start': (35, 1),
              'ref_start': (38, 1),
              },
         ]},
        {'file': 'test_sheets/2_Annual_LandCover_Agreement_Grouped.xlsx',
         'ref_file': 'spreadsheets/bruce/4_CompileAreaEstimatesPSE.xlsx',
         'sheets_to_skip': ['Agreement'],
         'ref_sheets_to_skip': ['Cover&SE_UsingPSE', 'Developed', 'Cropland', 'GrassShrub', 'TreeCover', 'Water',
                                'Wetland', 'Barren'],
         'cell_locations': [
             {'name': "Land Cover Proportions - PSE",
              'shape': (DATA['N_CLASSES'], DATA['N_YEARS']),
              'ref_shape': (DATA['N_CLASSES'], DATA['N_YEARS']),
              'start': (1, 1),
              'ref_start': (2, 1),
              },
             {'name': "Map Proportions",
              'shape': (DATA['N_CLASSES'], DATA['N_YEARS']),
              'ref_shape': (DATA['N_CLASSES'], DATA['N_YEARS']),
              'start': (10, 1),
              'ref_start': (15, 1),
              },
             {'name': "Map at Reference Proportions",
              'shape': (1, DATA['N_YEARS']),
              'ref_shape': (1, DATA['N_YEARS']),
              'start': (19, 1),
              'ref_start': (27, 1),
              },
             # {'name': "Reference Proportions",
             # 'shape': (DATA['N_CLASSES'], DATA['N_YEARS']),
             # 'ref_shape': (DATA['N_CLASSES'], DATA['N_YEARS']),
             # 'start': (28, 1),
             # 'ref_start': (39, 1),
             # },
         ]},
        {'file': 'test_sheets/3a_Annual_Change_Agreement_NominalYear.xlsx',
         'ref_file': 'spreadsheets/bruce/5_GenericLandCoverChange-NoChange_NominalYearOnly.xlsx',
         'sheets_to_skip': [str(x) for x in range(1986, 1986 + DATA['N_YEARS'] - 1)],
         'ref_sheets_to_skip': ['GenericLCChangeNoChange_Annual&'],
         'cell_locations': [
             {'name': "Error matrix",
              'shape': (2, 2),
              'ref_shape': (2, 2),
              'start': (2, 2),
              'ref_start': (6, 2),
              },
             {'name': "Reference totals",
              'shape': (1, 2),
              'ref_shape': (1, 2),
              'start': (4, 2),
              'ref_start': (9, 2),
              },
             {'name': "Map totals",
              'shape': (2, 1),
              'ref_shape': (2, 1),
              'start': (2, 4),
              'ref_start': (6, 5),
              },
             {'name': "Total",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (4, 4),
              'ref_start': (9, 5),
              },
             {'name': "Producer's accuracy",
              'shape': (1, 2),
              'ref_shape': (1, 2),
              'start': (7, 2),
              'ref_start': (11, 2),
              },
             {'name': "User's accuracy",
              'shape': (2, 1),
              'ref_shape': (2, 1),
              'start': (2, 6),
              'ref_start': (6, 6),
              },
             {'name': "Accuracy overall",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (7, 7),
              'ref_start': (11, 5),
              },
         ]},
        {'file': 'test_sheets/3a_Annual_Change_Agreement_NominalYear.xlsx',
         'ref_file': 'spreadsheets/bruce/6_GenericLandCoverChange_CreditForTemporalOffsetMatches.xlsx',
         'sheets_to_skip': [str(x) for x in range(1986, 1986 + DATA['N_YEARS'] - 1)],
         'ref_sheets_to_skip': [],
         'cell_locations': [
             {'name': "Error matrix",
              'shape': (2, 2),
              'ref_shape': (2, 2),
              'start': (2, 2),
              'ref_start': (5, 2),
              },
             {'name': "Reference totals",
              'shape': (1, 2),
              'ref_shape': (1, 2),
              'start': (4, 2),
              'ref_start': (8, 2),
              },
             {'name': "Map totals",
              'shape': (2, 1),
              'ref_shape': (2, 1),
              'start': (2, 4),
              'ref_start': (5, 5),
              },
             {'name': "Total",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (4, 4),
              'ref_start': (8, 5),
              },
             {'name': "Producer's accuracy",
              'shape': (1, 2),
              'ref_shape': (1, 2),
              'start': (7, 2),
              'ref_start': (10, 2),
              },
             {'name': "User's accuracy",
              'shape': (2, 1),
              'ref_shape': (2, 1),
              'start': (2, 6),
              'ref_start': (5, 6),
              },
             {'name': "Accuracy overall",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (7, 7),
              'ref_start': (10, 5),
              },
         ]},
        {'file': 'test_sheets/3b_Annual_Change_Agreement_OneYearOffsets.xlsx',
         'ref_file': 'spreadsheets/bruce/6_GenericLandCoverChange_CreditForTemporalOffsetMatches.xlsx',
         'sheets_to_skip': [str(x) for x in range(1986, 1986 + DATA['N_YEARS'] - 1)],
         'ref_sheets_to_skip': [],
         'cell_locations': [
             {'name': "Error matrix",
              'shape': (2, 2),
              'ref_shape': (2, 2),
              'start': (2, 2),
              'ref_start': (15, 2),
              },
             {'name': "Reference totals",
              'shape': (1, 2),
              'ref_shape': (1, 2),
              'start': (4, 2),
              'ref_start': (18, 2),
              },
             {'name': "Map totals",
              'shape': (2, 1),
              'ref_shape': (2, 1),
              'start': (2, 4),
              'ref_start': (15, 5),
              },
             {'name': "Total",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (4, 4),
              'ref_start': (18, 5),
              },
             {'name': "Producer's accuracy",
              'shape': (1, 2),
              'ref_shape': (1, 2),
              'start': (7, 2),
              'ref_start': (20, 2),
              },
             {'name': "User's accuracy",
              'shape': (2, 1),
              'ref_shape': (2, 1),
              'start': (2, 6),
              'ref_start': (15, 6),
              },
             {'name': "Accuracy overall",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (7, 7),
              'ref_start': (20, 5),
              },
         ]},
        {'file': 'test_sheets/3c_Annual_Change_Agreement_TwoYearOffsets.xlsx',
         'ref_file': 'spreadsheets/bruce/6_GenericLandCoverChange_CreditForTemporalOffsetMatches.xlsx',
         'sheets_to_skip': [str(x) for x in range(1986, 1986 + DATA['N_YEARS'] - 1)],
         'ref_sheets_to_skip': [],
         'cell_locations': [
             {'name': "Error matrix",
              'shape': (2, 2),
              'ref_shape': (2, 2),
              'start': (2, 2),
              'ref_start': (25, 2),
              },
             {'name': "Reference totals",
              'shape': (1, 2),
              'ref_shape': (1, 2),
              'start': (4, 2),
              'ref_start': (28, 2),
              },
             {'name': "Map totals",
              'shape': (2, 1),
              'ref_shape': (2, 1),
              'start': (2, 4),
              'ref_start': (25, 5),
              },
             {'name': "Total",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (4, 4),
              'ref_start': (28, 5),
              },
             {'name': "Producer's accuracy",
              'shape': (1, 2),
              'ref_shape': (1, 2),
              'start': (7, 2),
              'ref_start': (30, 2),
              },
             {'name': "User's accuracy",
              'shape': (2, 1),
              'ref_shape': (2, 1),
              'start': (2, 6),
              'ref_start': (25, 6),
              },
             {'name': "Accuracy overall",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (7, 7),
              'ref_start': (30, 5),
              },
         ]},
        {'file': 'test_sheets/5_Aggregated_Change_Agreement_Conversions.xlsx',
         'ref_file': 'spreadsheets/bruce/3_LC_Change_ErrorMatrix.xlsx',
         'sheets_to_skip': ['1-Year Offsets', '2-Year Offsets'],
         'ref_sheets_to_skip': [],
         'cell_locations': [
             {'name': "Error matrix",
              'shape': (DATA['N_CLASSES']**2, DATA['N_CLASSES']**2),
              'ref_shape': (48, 48),
              'start': (2, 2),
              'ref_start': (3, 3),
              },
             {'name': "Reference totals",
              'shape': (1, DATA['N_CLASSES']**2),
              'ref_shape': (1, 48),
              'start': (DATA['N_CLASSES']**2 + 2, 2),
              'ref_start': (48 + 3, 3),
              },
             {'name': "Map totals",
              'shape': (DATA['N_CLASSES']**2, 1),
              'ref_shape': (48, 1),
              'start': (2, DATA['N_CLASSES']**2 + 2),
              'ref_start': (3, 48 + 3),
              },
             {'name': "Total",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (DATA['N_CLASSES']**2 + 2, DATA['N_CLASSES']**2 + 2),
              'ref_start': (48 + 3, 48 + 3),
              },
             {'name': "Producer's accuracy",
              'shape': (1, DATA['N_CLASSES']**2),
              'ref_shape': (1, 48),
              'start': (DATA['N_CLASSES']**2 + 5, 2),
              'ref_start': (48 + 5, 3),
              },
             {'name': "User's accuracy",
              'shape': (DATA['N_CLASSES']**2, 1),
              'ref_shape': (48, 1),
              'start': (2, DATA['N_CLASSES']**2 + 4),
              'ref_start': (3, 48 + 5),
              },
             {'name': "Accuracy overall",
              'shape': (1, 1),
              'ref_shape': (1, 1),
              'start': (DATA['N_CLASSES']**2 + 5, DATA['N_CLASSES']**2 + 5),
              'ref_start': (48 + 6, 48 + 5),
              },
         ]},
    ]

    # Handle the comparison between the individual change years, which the reference has on one sheet
    change_dict = {
     'file': 'test_sheets/3a_Annual_Change_Agreement_NominalYear.xlsx',
     'ref_file': 'spreadsheets/bruce/5_GenericLandCoverChange-NoChange_NominalYearOnly.xlsx',
     'sheets_to_skip': ['All Years'],
     'ref_sheets_to_skip': ['Ch-NoCh_NominalYrs'],
     'cell_locations': [
         {'name': "Error matrix",
          'shape': (2, 2),
          'ref_shape': (2, 2),
          'start': (2, 2),
          'ref_start': (6, 2),
          },
         {'name': "Reference totals",
          'shape': (1, 2),
          'ref_shape': (1, 2),
          'start': (4, 2),
          'ref_start': (8, 2),
          },
         {'name': "Map totals",
          'shape': (2, 1),
          'ref_shape': (2, 1),
          'start': (2, 4),
          'ref_start': (6, 4),
          },
         {'name': "Total",
          'shape': (1, 1),
          'ref_shape': (1, 1),
          'start': (4, 4),
          'ref_start': (8, 4),
          },
         {'name': "Producer's accuracy",
          'shape': (1, 2),
          'ref_shape': (1, 2),
          'start': (7, 2),
          'ref_start': (10, 2),
          },
         {'name': "User's accuracy",
          'shape': (2, 1),
          'ref_shape': (2, 1),
          'start': (2, 6),
          'ref_start': (6, 6),
          },
         {'name': "Accuracy overall",
          'shape': (1, 1),
          'ref_shape': (1, 1),
          'start': (7, 7),
          'ref_start': (10, 4),
          },
     ]}

    for i, x in enumerate(range(1986, 1986 + DATA['N_YEARS'] - 1)):
        change_dict['sheets_to_skip'] = ['All Years'] + [str(y) for y in range(1986, 1986 + DATA['N_YEARS'] - 1)
                                                         if y != x]
        cell_locations = change_dict['cell_locations']
        if i != 0:
            for cell_dict in cell_locations:
                (x, y) = cell_dict['ref_start']
                cell_dict['ref_start'] = (x + 9, y) if i != 1 else (x + 10, y)

        BOOKS_TO_CHECK.append(copy.deepcopy(change_dict))

    # BEGIN CHECKS

    for report in BOOKS_TO_CHECK:
        print('')
        print('Checking ' + report['file'] + ' against ' + report['ref_file'])

        # Open the Excel file
        book = xlrd.open_workbook(report['file'])
        test_book = xlrd.open_workbook(report['ref_file'])

        sheets = [x for x in book.sheets() if x.name not in report['sheets_to_skip']]
        test_sheets = [x for x in test_book.sheets() if x.name not in report['ref_sheets_to_skip']]

        # Iterate over the sheets to check cell values against function return values
        for (sheet, test_sheet) in zip(sheets, test_sheets):

            for cell_data in report['cell_locations']:

                data_sheet = get_data_from_excel(sheet, cell_data['start'], cell_data['shape'])
                data_test = get_data_from_excel(test_sheet, cell_data['ref_start'], cell_data['ref_shape'])


                if report['file'] == 'test_sheets/5_Aggregated_Change_Agreement_Conversions.xlsx':

                    # For the LC Change Matrix, we need to excise blank rows and columns to match the reference sheets
                    if cell_data['shape'] != (1, 1):
                        col_headers = np.flip(get_data_from_excel(sheet, (1, 2), (1, DATA['N_CLASSES']**2)).astype(np.int))
                        col_headers_test = np.flip(get_data_from_excel(test_sheet, (1, 3), (1, 48)).astype(np.int))
                        for (i, col) in zip(range(len(col_headers) - 1, -1, -1), col_headers):
                            if col not in col_headers_test:
                                if (cell_data['shape'][0] != 1) and (cell_data['shape'][1] != 1):
                                    data_sheet = np.delete(data_sheet, i, axis=0)
                                    data_sheet = np.delete(data_sheet, i, axis=1)
                                else:
                                    data_sheet = np.delete(data_sheet, i)

                if (report['file'] == 'test_sheets/5_Aggregated_Change_Agreement_Conversions.xlsx') or \
                        (report['file'] == 'test_sheets/3a_Annual_Change_Agreement_NominalYear.xlsx') or \
                        (report['file'] == 'test_sheets/3b_Annual_Change_Agreement_OneYearOffsets.xlsx') or \
                        (report['file'] == 'test_sheets/3c_Annual_Change_Agreement_TwoYearOffsets.xlsx'):

                    if (report['file'] == 'test_sheets/5_Aggregated_Change_Agreement_Conversions.xlsx') or \
                            sheet.name == 'All Years':

                        # We also need to compensate for the fact that the reference sheets include invalid values
                        if cell_data['name'] == 'Total':
                            data_sheet += DATA['N_INVALID']

                        if cell_data['name'] == 'Accuracy overall':
                            total_dict = [x for x in report['cell_locations'] if x['name'] == 'Total'][0]
                            total = get_data_from_excel(sheet, total_dict['start'], total_dict['shape'])
                            data_sheet = (data_sheet * total + DATA['N_INVALID']) / (total + DATA['N_INVALID'])

                        if (report['file'] == 'test_sheets/3a_Annual_Change_Agreement_NominalYear.xlsx') or \
                            (report['file'] == 'test_sheets/3b_Annual_Change_Agreement_OneYearOffsets.xlsx') or \
                            (report['file'] == 'test_sheets/3c_Annual_Change_Agreement_TwoYearOffsets.xlsx'):

                            if cell_data['name'] == 'Error matrix':
                                data_sheet[0, 0] += DATA['N_INVALID']

                            if cell_data['name'] == 'Reference totals':
                                data_sheet[0] += DATA['N_INVALID']

                            if cell_data['name'] == 'Map totals':
                                data_sheet[0] += DATA['N_INVALID']

                            if cell_data['name'] == "User's accuracy":
                                total_dict = [x for x in report['cell_locations'] if x['name'] == 'Total'][0]
                                total = get_data_from_excel(sheet, total_dict['start'], total_dict['shape'])
                                data_sheet[0] = (data_sheet[0] * total + DATA['N_INVALID']) / \
                                                (total + DATA['N_INVALID'])

                            if cell_data['name'] == "Producer's accuracy":
                                total_dict = [x for x in report['cell_locations'] if x['name'] == 'Total'][0]
                                total = get_data_from_excel(sheet, total_dict['start'], total_dict['shape'])
                                data_sheet[0] = (data_sheet[0] * total + DATA['N_INVALID']) / \
                                                (total + DATA['N_INVALID'])

                    elif report['file'] == 'test_sheets/3a_Annual_Change_Agreement_NominalYear.xlsx':

                        if cell_data['name'] == 'Error matrix':
                            data_sheet[0, 0] = data_test[0, 0]  # Can't compare the No-change/No-change cells

                        if cell_data['name'] == "Reference totals":
                            data_sheet[0] = data_test[0]

                        if cell_data['name'] == "Map totals":
                            data_sheet[0] = data_test[0]

                        if cell_data['name'] == "User's accuracy":
                            data_sheet[0] = data_test[0]

                        if cell_data['name'] == "Producer's accuracy":
                            data_sheet[0] = data_test[0]

                        if cell_data['name'] == 'Accuracy overall':
                            continue

                        if cell_data['name'] == 'Total':
                            continue

                np.testing.assert_array_almost_equal(data_sheet, data_test,
                                                     err_msg=cell_data['name'] + " fails on sheet " + sheet.name)

            print(sheet.name+" ... ", end="")
        print('success!')


if __name__ == "__main__":
    main()
