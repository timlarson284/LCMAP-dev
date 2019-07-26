"""
Validation Metrics Script

This script is intended to output all validation reports from the work presented
by B. Pengra for the LCMAP reference data set validation efforts.

Author(s): D. Wellington
"""

import numpy as np

import validation_metrics
import validation_format


def annual_statistics(years, kwargs, all_years=True, each_year=True):
    """
    Loop through all the years in the data frame and provide statistics for each one, and optionally all.

    :param years:
    :param kwargs:
    :param all_years:
    :param each_year:
    :return:
    """

    statistics = []
    year_sets = []

    if all_years:
        year_sets += [years]

    if each_year:
        year_sets += [[x] for x in years]

    for year_list in year_sets:
        statistics.append(validation_metrics.statistics(years=year_list, **kwargs))

    return statistics


def grouped_by_key(all_data):
    """
    This function produces land cover agreement statistics regrouped so they can be parsed by metric instead of by
    year.

    :return: Returns a list with the dictionary of annual validation statistics reformatted from dictionaries
    for individual years, to a single dictionary with data for all years.
    """

    grouped = {}

    for data in all_data:
        for key, value in data.items():
            grouped.setdefault(key, np.array([]))
            if grouped[key].size == 0:
                grouped[key] = value
            else:
                if value.ndim == 0:
                    grouped[key] = np.append(grouped[key], value)
                elif value.ndim == 1:
                    grouped[key] = np.column_stack((grouped[key], value))
                elif value.ndim == 2:
                    grouped[key] = np.dstack((grouped[key], value))

    return grouped


def main(map_ref_file, plot_file, mask_file, histo_file, outdir='.'):
    """
    Generates the output reports (Excel spreadsheets) that provide validation metric calculations
    in a format modeled after the work by B. Pengra.
    """

    import validation_io

    def _combine_dicts(args):
        full_dict = {}
        for i, arg_dict in enumerate(args):
            for key in arg_dict.keys():
                arg_dict[key + '_' + str(i)] = arg_dict.pop(key)
            full_dict.update(arg_dict)
        return full_dict

    # Get the annualized crosswalked reference data with associated map data, and histogram data, from a file.
    ref_and_map = validation_io.load_ref_and_map(map_ref_file,
                                                 plot_file=plot_file,
                                                 mask_file=mask_file)
    histogram = validation_io.load_histogram_file(histo_file)

    years = np.unique(ref_and_map.image_year)
    str_years = [str(x) for x in years]

    accuracy_group = [
        'users_accuracy',
        'producers_accuracy',
        'overall_accuracy',
        'poststratified_producers_accuracy',
        'poststratified_producers_accuracy_overall',
        'poststratified_producers_standard_error_overall',
        'poststratified_producers_accuracy_overall_upper',
        'poststratified_producers_accuracy_overall_lower',
        'poststratified_dice_coefficients',
    ]

    area_group = [
        'area_proportion',
        'wh',
        'class_proportions',
        'reference_proportions',
    ]

    landcover_dict = {
        'data': ref_and_map,
        'axes': ['LC_Primary', 'Reference'],
        'categories': range(1, 9),
        'histogram': histogram,
        'histogram_columns': [str(x) for x in range(1, 9)],
        'data_year_header': 'image_year',
        'histogram_year_header': 'year',
    }

    change_dict = {
        'axes': ['MapChg', 'RefChg'],
        'categories': ['NoChg', 'Chg'],
        'data_year_header': 'image_year',
    }

    conversion_dict = {
        'axes': ['MapChgFromTo', 'RefChgFromTo'],
        'categories': [
            101, 202, 303, 404, 505, 606, 707, 808,
            102, 103, 104, 105, 106, 107, 108,
            201, 203, 204, 205, 206, 207, 208,
            301, 302, 304, 305, 306, 307, 308,
            401, 402, 403, 405, 406, 407, 408,
            501, 502, 503, 504, 506, 507, 508,
            601, 602, 603, 604, 605, 607, 608,
            701, 702, 703, 704, 705, 706, 708,
            801, 802, 803, 804, 805, 806, 807,
        ],
        'data_year_header': 'image_year',
    }

    annual_landcover_agreement = annual_statistics(years, landcover_dict)

    annual_change_agreement_list = []
    annual_conversion_agreement_list = []
    for i in range(3):
        print('Calculating change for allow_offset = '+str(i))
        data = {'data': validation_metrics.change_nochange(ref_and_map, allow_offset=i)}
        change_dict.update(data)
        conversion_dict.update(data)
        annual_change_agreement_list.append(annual_statistics(years[1:], change_dict))
        annual_conversion_agreement_list.append(annual_statistics(years[1:], conversion_dict, each_year=False))

    # 1_Annual_PSE_Area&AgreementTables.xlsx and 2_ErrorMatrix_AllYearsAggregated.xlsx
    validation_format.write_workbook(outdir+'/1_Annual_LandCover_Agreement.xlsx',
                                     annual_landcover_agreement,
                                     ['All Years'] + str_years,
                                     validation_format.book1_formatter,
                                     validation_format.BOOK1_CELL_LOCATIONS,
                                     categories=validation_format.LCMAP_LAND_COVER_CLASSES)

    # 1b_AnnualProducersUsersOverallAgreementTables&Graphs.xlsx and 4_CompileAreaEstimatesPSE.xlsx
    validation_format.write_workbook(outdir+'/2_Annual_LandCover_Agreement_Grouped.xlsx',
                                     [{key: grouped_by_key(annual_landcover_agreement[1:])[key] for
                                       key in accuracy_group},
                                      {key: grouped_by_key(annual_landcover_agreement[1:])[key] for
                                       key in area_group}],
                                     ['Agreement', 'Area'],
                                     validation_format.book2_formatter,
                                     validation_format.BOOK2_CELL_LOCATIONS,
                                     categories=validation_format.LCMAP_LAND_COVER_CLASSES, years=str_years)

    # 5_GenericLandCoverChange-NoChange_NominalYearOnly.xlsx
    validation_format.write_workbook(outdir+'/3a_Annual_Change_Agreement_NominalYear.xlsx',
                                     annual_change_agreement_list[0],
                                     ['All Years'] + str_years[1:],
                                     validation_format.book1_formatter,
                                     validation_format.BOOK3_CELL_LOCATIONS,
                                     categories=validation_format.CHANGE_NO_CHANGE)

    # 6_GenericLandCoverChange_CreditForTemporalOffsetMatches.xlsx
    validation_format.write_workbook(outdir+'/3b_Annual_Change_Agreement_OneYearOffsets.xlsx',
                                     annual_change_agreement_list[1],
                                     ['All Years'] + str_years[1:],
                                     validation_format.book1_formatter,
                                     validation_format.BOOK3_CELL_LOCATIONS,
                                     categories=validation_format.CHANGE_NO_CHANGE)

    # 6_GenericLandCoverChange_CreditForTemporalOffsetMatches.xlsx
    validation_format.write_workbook(outdir+'/3c_Annual_Change_Agreement_TwoYearOffsets.xlsx',
                                     annual_change_agreement_list[2],
                                     ['All Years'] + str_years[1:],
                                     validation_format.book1_formatter,
                                     validation_format.BOOK3_CELL_LOCATIONS,
                                     categories=validation_format.CHANGE_NO_CHANGE)

    validation_format.write_workbook(outdir+'/4_Annual_Change_Agreement_Grouped.xlsx',
                                     [_combine_dicts([grouped_by_key(x[1:]) for x in annual_change_agreement_list])],
                                     ['Change - No Change'],
                                     validation_format.book2_formatter,
                                     validation_format.BOOK4_CELL_LOCATIONS,
                                     categories=validation_format.CHANGE_NO_CHANGE, years=str_years[1:])

    # Equivalent to 3_LC_Change_ErrorMatrix.xlsx
    validation_format.write_workbook(outdir+'/5_Aggregated_Change_Agreement_Conversions.xlsx',
                                     [y for x in annual_conversion_agreement_list for y in x],
                                     ['Nominal', '1-Year Offsets', '2-Year Offsets'],
                                     validation_format.book1_formatter,
                                     validation_format.BOOK5_CELL_LOCATIONS,
                                     categories=validation_format.CONVERSIONS)

    # TODO:

    # 5_GenericLandCoverChange-NoChange_NominalYearOnly.xlsx - 5-year increments?

    # 7_TreeCoverLoss_NominalYear_Only.xlsx
    # This is a specialty change/no-change report - what to do about these?

    # 8_TreeCoverLoss_CreditForTemporalOffsetMatches.xlsx
    # Similarly with above

    return


if __name__ == "__main__":

    import validation_io

    main(validation_io.DEFAULT_MAP_REF_FILE,
         validation_io.DEFAULT_PLOT_FILE,
         validation_io.DEFAULT_MASK_FILE,
         validation_io.DEFAULT_HISTOGRAM_FILE)
