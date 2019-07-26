"""
These functions handle getting and formatting the reference data set for
validation.
Authors: D. Wellington, T. Larson, w/code copied from Kelcy's assessments.ipynb notebook
"""

import numpy as np
import pandas as pd

from osgeo import gdal
from osgeo.gdalconst import *
import struct
import sys

DEFAULT_MAP_REF_FILE = 'plots/RefandMap-final.csv'
DEFAULT_PLOT_FILE = 'plots/First50K_plots.xls'
DEFAULT_MASK_FILE = '/lcmap_data/bulk/ancillary/NLCD/Original/nlcd_2001_landcover_2011_edition_2014_10_10/nlcd_2001_landcover_2011_edition_2014_10_10/nlcd_2001_landcover_2011_edition_2014_10_10.img'
# DEFAULT_MASK_FILE = '/lcmap_data/bulk/assessment/ecoregions/masks/west_megaregion_mask.tif'
DEFAULT_HISTOGRAM_FILE = 'plots/MapCounts_prototype_full_w_h25v10.csv'


def load_ref_and_map(file, plot_file=None, mask_file=None):
    """
    This function loads the crosswalked/annualized/map-matched reference and map data from a file. The contents are
    mostly copied from Kelcy's assessments.ipynb notebook.

    :param file: File containing the reference and map data of a format equivalent to that in RefandMap-update.csv
    :return: A pandas Dataframe parsed from the csv file, with some modifications.
    """

    # Load the RefandMap data set that contains all the Reference and Map information lined up, by plot

    # Specify the column types pandas to ease loading
    dtypes = {'project_code': np.object,
              'plotid': np.int,
              'image_year': np.int32,
              'image_julday': np.int32,
              'interpreter': np.int16,
              'dominant_landuse': np.object,
              'secondary_landuse': np.object,
              'dominant_landuse_notes': np.object,
              'secondary_landuse_notes': np.object,
              'dominant_landcover': np.object,
              'second_landcover': np.object,
              'change_process': np.object,
              'change_process_notes': np.object,
              'LCMAP': np.object,
              'LCMAP_Change': np.object,
              'LCMAP_code': np.int,
              'CHANGE_code': np.float,  # 16
              'LCMAP_CH_code': np.int,
              'LCMAP_change_proc': np.float,
              'LCMAP_harvest_type': np.float,  # 19
              'Chg_ChangeDay': np.int32,
              'Chg_ChangeMag': np.float,
              'Chg_LastChange': np.int,
              'Chg_Quality': np.int8,
              'Chg_SegLength': np.int,
              'LC_Change': np.int8,
              'LC_Primary': np.int8,
              'LC_PrimeConf': np.int,
              'LC_Secondary': np.int8,
              'LC_SecondConf': np.int}

    # Map all the unique Land Cover labels in the Reference data set to LCMAP Map values
    # np.unique on the Reference column "LCMAP" after massaging
    # massaging includes: lower, rstrip
    lc_map = {'developed': 1,
              'mining': 1,
              'agriculture': 2,
              'grass': 3,
              'shrubs': 3,
              'forest': 4,
              'water': 5,
              'wetland': 6,
              'snow': 7,
              'ice_and_snow': 7,
              'barren': 8}

    refmap_df = pd.read_csv(file, nrows=0)
    dtypes = {key: value for (key, value) in dtypes.items() if key in refmap_df.columns}
    refmap_df = pd.read_csv(file, low_memory=False, dtype=dtypes)

    # if other object columns are giving issue, they may need to be explicitly cast to string as well ...
    refmap_df.loc[:, 'LCMAP'] = refmap_df.LCMAP.apply(lambda x: x.lower().rstrip())
    refmap_df.loc[:, 'LCMAP_Change'] = refmap_df.LCMAP_Change.apply(lambda x: str(x))
    refmap_df.loc[:, 'LCMAP_Change'] = refmap_df.LCMAP_Change.apply(lambda x: x.lower().rstrip())
    refmap_df.loc[:, 'Reference'] = refmap_df.LCMAP.apply(lambda x: lc_map[x])

    # Fix the floats ...
    if 'CHANGE_code' in refmap_df.columns:
        refmap_df.loc[:, 'CHANGE_code'] = refmap_df.CHANGE_code.fillna(0).astype(dtype='int')
    if 'LCMAP_change_proc' in refmap_df.columns:
        refmap_df.loc[:, 'LCMAP_change_proc'] = refmap_df.LCMAP_change_proc.fillna(0).astype(dtype='int')
    if 'LCMAP_harvest_type' in refmap_df.columns:
        refmap_df.loc[:, 'LCMAP_harvest_type'] = refmap_df.LCMAP_harvest_type.fillna(0).astype(dtype='int')

    # These are essentially useless for anything ...
    refmap_df.drop(['project_code', 'image_julday', 'interpreter'], axis=1, inplace=True)
    
    return filter_plots(refmap_df, plot_file, mask_file)


def load_histogram_file(file):
    return pd.read_csv(file)


def filter_plots(ref_df, plot_file, mask_file):

    def pt2fmt(pt):
        fmttypes = {
            GDT_Byte: 'B',
            GDT_Int16: 'h',
            GDT_UInt16: 'H',
            GDT_Int32: 'i',
            GDT_UInt32: 'I',
            GDT_Float32: 'f',
            GDT_Float64: 'f'
        }
        return fmttypes.get(pt, 'x')

    def readmask(x, y, gdal_ds):
        transf = gdal_ds.GetGeoTransform()
        band = gdal_ds.GetRasterBand(1)
        transfinv = gdal.InvGeoTransform(transf)

        px, py = gdal.ApplyGeoTransform(transfinv, x, y)
        structval = band.ReadRaster(int(px), int(py), 1, 1, buf_type=band.DataType)
        fmt = pt2fmt(band.DataType)
        intval = struct.unpack(fmt, structval)

        return round(intval[0], 2)  # intval is a tuple, length=1 as we only asked for 1 pixel value

    if (plot_file is None) or (mask_file is None):
        print('Plot and/or mask file not provided, not filtering reference data...')
        return ref_df

    first50_k_df = pd.read_excel(plot_file,
                                 sheet_name='First50K_plots',
                                 usecols=['x', 'y', 'plotid'])
    
    mask_ds = gdal.Open(mask_file, GA_ReadOnly)
    if mask_ds is None:
        print('Failed open file')
        sys.exit(1)
    ulx, xres, xskew, uly, yskew, yres = mask_ds.GetGeoTransform()
    cols = mask_ds.RasterXSize
    rows = mask_ds.RasterYSize
    lrx = ulx + (cols * xres)
    lry = uly + (rows * yres)
    
    unique_plots = ref_df.drop_duplicates(subset='plotid')
    first50_k_select = first50_k_df[first50_k_df.plotid.isin(unique_plots.plotid)]

    # FILTER PLOT DEFINITIONS BY MASK FILE EXTENT (the extents of the file, not the mask data), THEN BY DATA
    plotxy_mask_file_extent = first50_k_select[(first50_k_select['x'].astype(int) > ulx) &
                                               (first50_k_select['x'].astype(int) < lrx) &
                                               (first50_k_select['y'].astype(int) < uly) &
                                               (first50_k_select['y'].astype(int) > lry)]
    
    plotxy_final = pd.DataFrame(columns=['x', 'y', 'plotid'])  # empty dataframe
    for row in plotxy_mask_file_extent.itertuples():
        if readmask(row[1], row[2], mask_ds) > 0:
            plotxy_final = plotxy_final.append({'x': row[1], 'y': row[2], 'plotid': row[3]}, ignore_index=True)
    ref_map_final = ref_df[ref_df.plotid.isin(plotxy_final.plotid)].reset_index()
    
    return ref_map_final
