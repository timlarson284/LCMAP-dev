"""
Various config settings and definitions
"""

import xlsxwriter

# Key for LCMAP land cover class codes
LCMAP_LAND_COVER_CLASSES = [
    'Developed',
    'Agriculture',
    'Grass/Shrub',
    'Forest',
    'Water',
    'Wetland',
    'Snow/Ice',
    'Barren',
]

CHANGE_NO_CHANGE = [
    'No Change',
    'Change',
]

CONVERSIONS = [
        101, 202, 303, 404, 505, 606, 707, 808,
        102, 103, 104, 105, 106, 107, 108,
        201, 203, 204, 205, 206, 207, 208,
        301, 302, 304, 305, 306, 307, 308,
        401, 402, 403, 405, 406, 407, 408,
        501, 502, 503, 504, 506, 507, 508,
        601, 602, 603, 604, 605, 607, 608,
        701, 702, 703, 704, 705, 706, 708,
        801, 802, 803, 804, 805, 806, 807,
    ]

# Cell locations and formatting dictionary
BOOK1_CELL_LOCATIONS = {
    'error_matrix': (2, 2, 'matrix', 'normal'),
    'users_accuracy': (2, 12, 'vertical', 'percent'),
    'users_standard_error': (2, 13, 'vertical', 'percent'),
    'producers_accuracy': (13, 2, 'horizontal', 'percent'),
    'producers_standard_error': (14, 2, 'horizontal', 'percent'),
    'poststratified_producers_accuracy': (17, 2, 'horizontal', 'percent'),
    'poststratified_producers_standard_error': (18, 2, 'horizontal', 'percent'),
    'overall_accuracy': (13, 13, 'cell', 'percent'),
    'poststratified_producers_accuracy_overall': (14, 13, 'cell', 'percent'),
    'poststratified_producers_standard_error_overall': (15, 13, 'cell', 'percent'),
    'poststratified_dice_coefficients': (2, 15, 'vertical', 'percent'),
    'area_proportion': (2, 17, 'vertical', 'percent'),
    'area_proportion_standard_error': (2, 18, 'vertical', 'percent'),
    'reference_totals': (10, 2, 'horizontal', 'normal'),
    'map_totals': (2, 10, 'vertical', 'normal'),
    'grand_total': (10, 10, 'cell', 'normal'),
}

# Cell locations and formatting dictionary
BOOK2_CELL_LOCATIONS = {
    'users_accuracy': (1, 1, 'matrix', 'int_percent_sm'),
    'producers_accuracy': (10, 1, 'matrix', 'int_percent_sm'),
    'overall_accuracy': (19, 1, 'horizontal', 'int_percent_sm_all'),
    'poststratified_producers_accuracy': (21, 1, 'matrix', 'int_percent_sm'),
    'poststratified_producers_accuracy_overall': (30, 1, 'horizontal', 'int_percent_sm_all'),
    'poststratified_producers_standard_error_overall': (31, 1, 'horizontal', 'int_percent_sm_err'),
    'poststratified_producers_accuracy_overall_upper': (32, 1, 'horizontal', 'int_percent_sm_all'),
    'poststratified_producers_accuracy_overall_lower': (33, 1, 'horizontal', 'int_percent_sm_all'),
    'poststratified_dice_coefficients': (35, 1, 'matrix', 'int_percent_sm'),
    'area_proportion': (1, 1, 'matrix', 'int_percent_sm_all'),
    'wh': (10, 1, 'matrix', 'int_percent_sm_all'),
    'class_proportions': (19, 1, 'matrix', 'int_percent_sm_all'),
    'reference_proportions': (28, 1, 'matrix', 'int_percent_sm_all'),
}

BOOK3_CELL_LOCATIONS = {
    'error_matrix': (2, 2, 'matrix', 'normal'),
    'users_accuracy': (2, 6, 'vertical', 'percent'),
    'users_standard_error': (2, 7, 'vertical', 'percent'),
    'producers_accuracy': (7, 2, 'horizontal', 'percent'),
    'producers_standard_error': (8, 2, 'horizontal', 'percent'),
    'overall_accuracy': (7, 7, 'cell', 'percent'),
    'reference_totals': (4, 2, 'horizontal', 'normal'),
    'map_totals': (2, 4, 'vertical', 'normal'),
    'grand_total': (4, 4, 'cell', 'normal'),
}

BOOK4_CELL_LOCATIONS = {
    'users_accuracy_0': (1, 1, 'matrix', 'int_percent_sm'),
    'producers_accuracy_0': (3, 1, 'matrix', 'int_percent_sm'),
    'overall_accuracy_0': (5, 1, 'horizontal', 'int_percent_sm_all'),
    'users_accuracy_1': (7, 1, 'matrix', 'int_percent_sm'),
    'producers_accuracy_1': (9, 1, 'matrix', 'int_percent_sm'),
    'overall_accuracy_1': (11, 1, 'horizontal', 'int_percent_sm_all'),
    'users_accuracy_2': (13, 1, 'matrix', 'int_percent_sm'),
    'producers_accuracy_2': (15, 1, 'matrix', 'int_percent_sm'),
    'overall_accuracy_2': (17, 1, 'horizontal', 'int_percent_sm_all'),
}

BOOK5_CELL_LOCATIONS = {
    'error_matrix': (2, 2, 'matrix', 'normal'),
    'users_accuracy': (2, 68, 'vertical', 'percent'),
    'users_standard_error': (2, 69, 'vertical', 'percent'),
    'producers_accuracy': (69, 2, 'horizontal', 'percent'),
    'producers_standard_error': (70, 2, 'horizontal', 'percent'),
    'overall_accuracy': (69, 69, 'cell', 'percent'),
    'reference_totals': (66, 2, 'horizontal', 'normal'),
    'map_totals': (2, 66, 'vertical', 'normal'),
    'grand_total': (66, 66, 'cell', 'normal'),
}


def get_formats(workbook):
    """
    Dictionary for shared Format objects.

    :param workbook: XlsxWriter Workbook object
    :return: Returns a dictionary of XlsxWriter Format objects
    """
    return {
        'normal': workbook.add_format(),
        'percent': workbook.add_format({'num_format': '0.00%'}),
        'int_percent_sm': workbook.add_format({'num_format': '0%', 'font_size': 9, 'align': 'center'}),
        'int_percent_sm_err': workbook.add_format({'num_format': '0.000%', 'font_size': 9, 'align': 'center'}),
        'int_percent_sm_all': workbook.add_format({'num_format': '0.0%', 'font_size': 9, 'align': 'center'}),
        'bold': workbook.add_format({'bold': True}),
        'bold_sm': workbook.add_format({'bold': True, 'font_size': 9}),
        'title': workbook.add_format({'bold': True, 'align': 'center'}),
        'highlight_title': workbook.add_format({'bold': True, 'align': 'center', 'pattern': 1,
                                                'bg_color': '#CCCCCC', 'text_wrap': True}),
        'rotated_title': workbook.add_format({'bold': True, 'align': 'vcenter', 'rotation': 90}),
        'half_rotated_title_highlight': workbook.add_format({'bold': True, 'align': 'vcenter', 'rotation': 45,
                                                             'pattern': 1, 'bg_color': '#CCCCCC'}),
    }


def write_workbook(file, list_data_by_sheet, list_sheet_names, formatter, cell_locations, **kwargs):
    """
    This function writes an output Excel file with the specified values and formatting information.

    :param file: Name of the xlsx file to be created with the output.
    :param list_data_by_sheet: A list of dictionaries with data for each worksheet.
    :param list_sheet_names: A list of sheet names.
    :param formatter: A reference to the function that writes formatting into each worksheet.
    :param cell_locations: A dictionary containing the locations and formatting matching the keywords in
    list_data_by_sheet dictionaries.
    :param kwargs: Other keyword arguments to pass to the formatter function.
    :return: Nothing, but writes an xlsx file.
    """

    workbook = xlsxwriter.Workbook(file, {'nan_inf_to_errors': True})

    for sheet_name, data in zip(list_sheet_names, list_data_by_sheet):
        worksheet = workbook.add_worksheet(sheet_name)
        formatter(workbook, worksheet, data, cell_locations, **kwargs)
        write_worksheet(workbook, worksheet, data, cell_locations)

    workbook.close()

    return


def write_worksheet(workbook, worksheet, valuedict, formatdict):
    """
    Writes worksheet data into defined cell locations.

    :param workbook: XlsxWriter Workbook object
    :param worksheet: XlsxWriter Worksheet object
    :param valuedict: A dictionary containing the data values.
    :param formatdict: A dictionary containing the location and formatting to write the data; keys need to match
    those in valuedict or no data will be written for that format entry.
    :return: Nothing, but modifies the worksheet.
    """

    formats = get_formats(workbook)

    for key, value in formatdict.items():
        if key not in valuedict:
            continue
        row, column, orientation, format_name = value
        if orientation == 'vertical':
            for i, x in enumerate(valuedict[key]):
                worksheet.write(row + i, column, x, formats[format_name])
        elif orientation == 'horizontal':
            for i, x in enumerate(valuedict[key]):
                worksheet.write(row, column + i, x, formats[format_name])
        elif orientation == 'matrix':
            for i, y in enumerate(valuedict[key]):
                for j, x in enumerate(y):
                    worksheet.write(row + i, column + j, x, formats[format_name])
        else:
            worksheet.write(row, column, valuedict[key], formats[format_name])

    return


def cell_range_border(workbook, worksheet, cell_range):
    """
    Helper function that puts borders around a range of cells. Probably a bad idea to use this with overlapping
    cell ranges.

    :param workbook: XlsxWriter Workbook object
    :param worksheet: XlsxWriter Worksheet object
    :param cell_range: An Excel cell range as a string, e.g., 'A3:A10'
    :return: Nothing, but modifies the worksheet.
    """

    row_x, col_x, row_y, col_y = cell_range

    # Because of how conditional formatting is applied, we have to handle the different edges/corners differently.
    left_top_corner = (row_x, col_x, row_x, col_x)
    right_top_corner = (row_x, col_y, row_x, col_y)
    left_bottom_corner = (row_y, col_x, row_y, col_x)
    right_bottom_corner = (row_y, col_y, row_y, col_y)

    left_noncorners = (row_x + 1, col_x, row_y - 1, col_x)
    right_noncorners = (row_x + 1, col_y, row_y - 1, col_y)
    top_noncorners = (row_x, col_x + 1, row_x, col_y - 1)
    bottom_noncorners = (row_y, col_x + 1, row_y, col_y - 1)

    left_border = workbook.add_format({'left': 2})
    top_border = workbook.add_format({'top': 2})
    right_border = workbook.add_format({'right': 2})
    bottom_border = workbook.add_format({'bottom': 2})

    upperleft_border = workbook.add_format({'left': 2, 'top': 2})
    lowerleft_border = workbook.add_format({'left': 2, 'bottom': 2})
    upperright_border = workbook.add_format({'right': 2, 'top': 2})
    lowerright_border = workbook.add_format({'right': 2, 'bottom': 2})

    row_border = workbook.add_format({'top': 2, 'bottom': 2})
    rowleft_border = workbook.add_format({'left': 2, 'top': 2, 'bottom': 2})
    rowright_border = workbook.add_format({'right': 2, 'top': 2, 'bottom': 2})

    column_border = workbook.add_format({'left': 2, 'right': 2})
    columntop_border = workbook.add_format({'top': 2, 'left': 2, 'right': 2})
    columnbottom_border = workbook.add_format({'bottom': 2, 'left': 2, 'right': 2})
    cell_border = workbook.add_format({'left': 2, 'right': 2, 'top': 2, 'bottom': 2})

    # Single cell formatting
    if (row_x, col_x) == (row_y, col_y):
        worksheet.conditional_format(*cell_range, {'type': 'no_errors', 'format': cell_border})
    # Single row formatting
    elif row_x == row_y:
        worksheet.conditional_format(
            *top_noncorners, {'type': 'no_errors', 'format': row_border})
        worksheet.conditional_format(
            *left_top_corner, {'type': 'no_errors', 'format': rowleft_border})
        worksheet.conditional_format(
            *right_top_corner, {'type': 'no_errors', 'format': rowright_border})
    # Single column formatting
    elif col_x == col_y:
        worksheet.conditional_format(
            *left_noncorners, {'type': 'no_errors', 'format': column_border})
        worksheet.conditional_format(
            *left_top_corner, {'type': 'no_errors', 'format': columntop_border})
        worksheet.conditional_format(
            *left_bottom_corner, {'type': 'no_errors', 'format': columnbottom_border})
    # 2D formatting
    else:
        worksheet.conditional_format(
            *left_noncorners, {'type': 'no_errors', 'format': left_border})
        worksheet.conditional_format(
            *top_noncorners, {'type': 'no_errors', 'format': top_border})
        worksheet.conditional_format(
            *right_noncorners, {'type': 'no_errors', 'format': right_border})
        worksheet.conditional_format(
            *bottom_noncorners, {'type': 'no_errors', 'format': bottom_border})
        worksheet.conditional_format(*left_top_corner, {'type': 'no_errors', 'format': upperleft_border})
        worksheet.conditional_format(*left_bottom_corner, {'type': 'no_errors', 'format': lowerleft_border})
        worksheet.conditional_format(*right_top_corner, {'type': 'no_errors', 'format': upperright_border})
        worksheet.conditional_format(*right_bottom_corner, {'type': 'no_errors', 'format': lowerright_border})

    return


def cell_diagonal_highlight(workbook, worksheet, cell_range):
    """
    Helper function that highlights cells along a diagonal of a range of cells.

    :param workbook: XlsxWriter Workbook object
    :param worksheet: XlsxWriter Worksheet object
    :param cell_range: An Excel cell range as a string, e.g., 'A3:A10'
    :return: Nothing, but modifies the worksheet.
    """

    highlight = workbook.add_format({'pattern': 1, 'bg_color': '#CCCCCC'})

    row_x, col_x, row_y, col_y = cell_range

    for i in range(row_y - row_x + 1):
        worksheet.conditional_format(row_x + i, col_x + i, row_x + i, col_x + i,
                                     {'type': 'no_errors', 'format': highlight})

    return


def book1_formatter(workbook, worksheet, valuedict, formatdict, categories=None):
    """
    Formatting script for first workbook in the report series.

    :param workbook: XlsxWriter Workbook object
    :param worksheet: XlsxWriter Worksheet object
    :param formatdict: A dictionary with tuples for starting row, column for data, the orientation that the data
    is to be written in, and a key for the format codes to write.
    :return: Nothing, but modifies the worksheet.
    """

    # TODO: Clean this up

    formats = get_formats(workbook)

    for i, value in enumerate(categories):
        worksheet.write(formatdict['error_matrix'][0] + i, formatdict['error_matrix'][0] - 1, value)
        worksheet.write(formatdict['error_matrix'][0] - 1, formatdict['error_matrix'][0] + i, value)

    n_categories = len(categories)

    row_headers_by_key = {
        'producers_accuracy': ("Accuracy", "bold"),
        'producers_standard_error': ("Std Err", "bold"),
        'poststratified_producers_accuracy': ("Accuracy", "bold"),
        'poststratified_producers_standard_error': ("Std Err", "bold"),
        'overall_accuracy': ("Accuracy", "bold"),
        'poststratified_producers_accuracy_overall': ("PSE Accuracy", "bold"),
        'poststratified_producers_standard_error_overall': ("PSE Std Err", "bold"),
        'reference_totals': ("Totals", "bold"),
    }

    column_headers_by_key = {
        'users_accuracy': ("Accuracy", "bold"),
        'users_standard_error': ("Std Err", "bold"),
        'poststratified_dice_coefficients': ("Dice Coefficients", "bold"),
        'area_proportion': ("Proportion", "bold"),
        'area_proportion_standard_error': ("Std Err", "bold"),
        'map_totals': ("Totals", "bold"),
    }

    merge_column_headers_by_key = {
        'producers_accuracy': ("Producer's", "title", n_categories, 1, 0),
        'poststratified_producers_accuracy': ("PSE Producer's", "title", n_categories, 1, 0),
        'users_accuracy': ("User's", "title", 2, 2, 0),
        'area_proportion': ("Area", "title", 2, 2, 0),
        'error_matrix': ("Reference", "title", n_categories, 2, 0),
        'overall_accuracy': ("Overall", "title", 2, 1, 1)
    }

    merge_row_headers_by_key = {
        'error_matrix': ("Map", "rotated_title", n_categories, 2),
    }

    cell_borders = [
        # Border around error matrix
        (formatdict['error_matrix'][0],
         formatdict['error_matrix'][1],
         formatdict['error_matrix'][0] + n_categories - 1,
         formatdict['error_matrix'][1] + n_categories - 1),
        # Border around map labels
        (formatdict['error_matrix'][0],
         formatdict['error_matrix'][1] - 1,
         formatdict['error_matrix'][0] + n_categories - 1,
         formatdict['error_matrix'][1] - 1),
        # Border around reference labels
        (formatdict['error_matrix'][0] - 1,
         formatdict['error_matrix'][1],
         formatdict['error_matrix'][0] - 1,
         formatdict['error_matrix'][1] + n_categories - 1),
        # Border in upper left corner of error matrix
        (formatdict['error_matrix'][0] - 1,
         formatdict['error_matrix'][1] - 1,
         formatdict['error_matrix'][0] - 1,
         formatdict['error_matrix'][1] - 1),
        # Border around reference totals header
        (formatdict['reference_totals'][0],
         formatdict['reference_totals'][1] - 1,
         formatdict['reference_totals'][0],
         formatdict['reference_totals'][1] - 1),
        # Border around reference totals
        (formatdict['reference_totals'][0],
         formatdict['reference_totals'][1],
         formatdict['reference_totals'][0],
         formatdict['reference_totals'][1] + n_categories - 1),
        # Border around map totals header
        (formatdict['map_totals'][0] - 1,
         formatdict['map_totals'][1],
         formatdict['map_totals'][0] - 1,
         formatdict['map_totals'][1]),
        # Border around map totals
        (formatdict['map_totals'][0],
         formatdict['map_totals'][1],
         formatdict['map_totals'][0] + n_categories - 1,
         formatdict['map_totals'][1]),
        # Border around grand total
        (formatdict['grand_total'][0],
         formatdict['grand_total'][1],
         formatdict['grand_total'][0],
         formatdict['grand_total'][1]),
    ]

    cell_diagonal_highlights = [
        (formatdict['error_matrix'][0],
         formatdict['error_matrix'][1],
         formatdict['error_matrix'][0] + n_categories - 1,
         formatdict['error_matrix'][1] + n_categories - 1),
    ]

    for header_key, values in row_headers_by_key.items():
        if header_key in formatdict:
            row, column, _, _ = formatdict[header_key]
            worksheet.write(row, column - 1, values[0], formats[values[1]])

    for header_key, values in column_headers_by_key.items():
        if header_key in formatdict:
            row, column, _, _ = formatdict[header_key]
            worksheet.write(row - 1, column, values[0], formats[values[1]])

    for header_key, values in merge_column_headers_by_key.items():
        if header_key in formatdict:
            row, column, _, _ = formatdict[header_key]
            worksheet.merge_range(row - values[3], column - values[4], row - values[3],
                                  column - values[4] + values[2] - 1, values[0], formats[values[1]])

    for row_key, values in merge_row_headers_by_key.items():
        if row_key in formatdict:
            row, column, _, _ = formatdict[row_key]
            worksheet.merge_range(row, column - values[3], row + values[2] - 1, column - values[3],
                                  values[0], formats[values[1]])

    for border in cell_borders:
        cell_range_border(workbook, worksheet, border)

    for diagonal_highlights in cell_diagonal_highlights:
        cell_diagonal_highlight(workbook, worksheet, diagonal_highlights)

    return


def book2_formatter(workbook, worksheet, valuedict, formatdict, years=None, categories=None):
    """
    Formatting script for second workbook in the report series.

    :param workbook: XlsxWriter Workbook object
    :param worksheet: XlsxWriter Worksheet object
    :param formatdict: A dictionary with tuples for starting row, column for data, the orientation that the data
    is to be written in, and a key for the format codes to write.
    :param years: A list of years to use as column headings.
    :return: Nothing, but modifies the workbook.
    """

    # TODO: Clean this up a bit
    formats = get_formats(workbook)

    keys_for_by_class = ['users_accuracy',
                         'producers_accuracy',
                         'poststratified_producers_accuracy',
                         'poststratified_dice_coefficients',
                         'area_proportion',
                         'wh',
                         'class_proportions',
                         'reference_proportions',
                         'users_accuracy_0',
                         'users_accuracy_1',
                         'users_accuracy_2',
                         'producers_accuracy_0',
                         'producers_accuracy_1',
                         'producers_accuracy_2',
                         ]

    keys_for_by_year = ['users_accuracy',
                        'producers_accuracy',
                        'poststratified_producers_accuracy',
                        'poststratified_dice_coefficients',
                        'area_proportion',
                        'wh',
                        'class_proportions',
                        'reference_proportions',
                        'overall_accuracy',
                        'poststratified_producers_accuracy_overall',
                        'users_accuracy_0',
                        'users_accuracy_1',
                        'users_accuracy_2',
                        ]

    row_headers_by_key = {'users_accuracy': ("User's Agreement", "User's"),
                          'producers_accuracy': ("Producer's Agreement", "Producer's"),
                          'overall_accuracy': ("Overall", "Overall Agreement"),
                          'poststratified_producers_accuracy': ("Producer's - PSE", "Producer's - (PSE)"),
                          'poststratified_producers_accuracy_overall': ("Overall - PSE", "Overall Agreement (PSE)"),
                          'poststratified_producers_standard_error_overall': ("", "SE (PSE)"),
                          'poststratified_producers_accuracy_overall_upper': ("", "Upper limit w/SE"),
                          'poststratified_producers_accuracy_overall_lower': ("", "Lower limit w/SE"),
                          'poststratified_dice_coefficients': ("Overall Agreement per Class (PSE - Dice Coefficient)",
                                                               "Overall"),
                          'area_proportion': ("Land Cover Proportions - PSE", "Area"),
                          'wh': ("Map Proportions", "Area"),
                          'class_proportions': ("Map at Reference Proportions", "Area"),
                          'reference_proportions': ("Reference Proportions", "Area"),
                          'users_accuracy_0': ("Nominal Year", "User's"),
                          'producers_accuracy_0': ("", "Producer's"),
                          'overall_accuracy_0': ("", "Overall Agreement"),
                          'users_accuracy_1': ("1-Year Offset", "User's"),
                          'producers_accuracy_1': ("", "Producer's"),
                          'overall_accuracy_1': ("", "Overall Agreement"),
                          'users_accuracy_2': ("2-Year Offset", "User's"),
                          'producers_accuracy_2': ("", "Producer's"),
                          'overall_accuracy_2': ("", "Overall Agreement"),
                          }

    worksheet.set_column(0, 0, 27)
    worksheet.set_column(1, len(years), 6)

    for key, value in formatdict.items():
        if key in valuedict.keys():
            if key in keys_for_by_year:
                worksheet.write(value[0] - 1, value[1] - 1, row_headers_by_key[key][0], formats['highlight_title'])
                worksheet.set_row(value[0] - 1, 30)
                for i, year in enumerate(years):
                    worksheet.write(value[0]-1, value[1]+i, int(year), formats['half_rotated_title_highlight'])
            if key in keys_for_by_class:
                for i, cat_value in enumerate(categories):
                    worksheet.write(value[0]+i, value[1]-1, cat_value+', '+row_headers_by_key[key][1],
                                    formats['bold_sm'])
            if key not in keys_for_by_class:
                worksheet.write(value[0], value[1]-1, row_headers_by_key[key][1], formats['bold_sm'])

    return
