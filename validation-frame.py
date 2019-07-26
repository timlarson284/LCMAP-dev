import os
import time
import logging
from functools import lru_cache

import pandas as pd
from osgeo import gdal
import numpy as np

log = logging.getLogger()
__format = '%(asctime)s %(module)-10s::%(funcName)-20s - [%(lineno)-3d]%(message)s'
logging.basicConfig(level=logging.DEBUG,
                    format=__format,
                    datefmt='%Y-%m-%d %H:%M:%S')

_cu_tileaff = (-2565585, 150000, 0, 3314805, 0, -150000)
_root_maps = r'/lcmap_data/bulk/klsmith/test-runs/ccd_peeksize_zhe/test-maps/new'
_products = ('Chg_ChangeDay', 'Chg_ChangeMag', 'Chg_LastChange', 'Chg_Quality', 'Chg_SegLength',
             'LC_Change', 'LC_Primary', 'LC_PrimeConf', 'LC_Secondary', 'LC_SecondConf')


def main():
    ref_f = 'plots/lcmap_set1_27_postUSFS_vertex_crosswalked_annualized_assign_manualcorr.xlsx'
    ref_df = pd.read_excel(ref_f, sheet_name='lcmap_set1_27_postUSFS_vertex_c')

    pts_f = 'plots/First50K_plots.xls'
    pts_df = pd.read_excel(pts_f, sheet_name='First50K_plots')

    pts_df['hv'] = np.vectorize(paddedhv)(pts_df.x, pts_df.y)
    combined_df = pd.DataFrame()
    col_names = ['image_year']
    col_names.extend(_products)

    tiles = np.unique(pts_df.hv)
    cur = []
    tot = 0
    for tile in tiles:
        plots = np.unique(pts_df[pts_df.hv == tile].plotid)
        for plot in plots:
            plot, data = extractplot(plot, pts_df)
            if data is not None:
                if tile not in cur:
                    c = len(pts_df[pts_df.hv == tile])
                    tot += c
                    print(tile, c)
                    cur.append(tile)
                map_df = pd.DataFrame(list(data), columns=col_names)
                map_df['plotid'] = plot

                combined_df = combined_df.append(map_df)

    print(f'Total points: {tot}')
    ref_df = ref_df.merge(combined_df, on=['plotid', 'image_year'], how='left')
    mask = ref_df.LC_Primary.notna()
    refmap_mdf = ref_df[mask]
    refmap_mdf.to_csv('RefandMap.csv', index=False)


@lru_cache()
def transform_geo(x, y, affine):
    """
    Perform the affine transformation from a x/y coordinate to row/col
    space.

    Args:
        x: projected geo-spatial x coord
        y: projected geo-spatial y coord
        affine: gdal GeoTransform tuple

    Returns:
        containing pixel row/col
    """
    col = (x - affine[0] - affine[3] * affine[2]) / affine[1]
    row = (y - affine[3] - affine[0] * affine[4]) / affine[5]

    return int(row), int(col)


@lru_cache()
def transform_rc(row, col, affine):
    """
    Perform the affine transformation from a row/col coordinate to projected x/y
    space.

    Args:
        row: pixel/array row number
        col: pixel/array column number
        affine: gdal GeoTransform tuple

    Returns:
        x/y coordinate
    """
    x = affine[0] + col * affine[1] + row * affine[2]
    y = affine[3] + col * affine[4] + row * affine[5]

    return x, y


@lru_cache()
def determine_hv(x, y, affine=_cu_tileaff):
    """
    Determine the ARD tile H/V that contains the given coordinate.

    Args:
        x: projected geo-spatial x coord
        y: projected geo-spatial y coord
        affine: gdal GeoTransform tuple

    Returns:
        ARD tile h/v
    """
    return transform_geo(x, y, affine)[::-1]


@lru_cache()
def uladjust(x, y):
    """
    The x/y points are for center pixel mass but we care about upper-left for extracting raster data,
    this makes the appropriate adjustment.
    """
    return x - 15, y + 15


@lru_cache(maxsize=350)
def readtif(path, band=1):
    try:
        ds = gdal.Open(path, gdal.GA_ReadOnly)
        aff = ds.GetGeoTransform()
        data = ds.GetRasterBand(band).ReadAsArray()
    except:
        log.debug('Problem reading {}'.format(path))
        raise
    return data, aff


def readxy(path, x, y, band=1):
    arr, aff = readtif(path, band)
    row_off, col_off = transform_geo(x, y, aff)
    return arr[row_off, col_off]


@lru_cache()
def paths(root, strfilter):
    return sorted([os.path.join(root, f) for f in os.listdir(root) if strfilter in f and f[-4:] == '.tif'])


def extractprod(x, y, root, prod):
    ps = paths(root, prod)
    return [readxy(f, x, y) for f in ps]


def extractpt(x, y, root_maps=_root_maps, products=_products):
    x, y = uladjust(x, y)
    h, v = determine_hv(x, y)
    root = os.path.join(root_maps, 'h{:02}v{:02}'.format(h, v))
    if not os.path.exists(root):
        return

    ret = [extractprod(x, y, root, p) for p in products]

    return zip(range(1985, 2018), *ret)


def extractplot(plotid, masterdf):
    mask = masterdf.plotid == plotid
    x = masterdf[mask].x.values[0]
    y = masterdf[mask].y.values[0]
    return plotid, extractpt(x, y)


def paddedhv(x, y):
    h, v = determine_hv(x, y)
    return 'h{:02}v{:02}'.format(h, v)


if __name__ == '__main__':
    t1 = time.time()
    main()
    print(time.time() - t1)
