import time
import os

from osgeo import gdal
import numpy as np
import pandas as pd

import validation_io


_nlcdpath = r'/lcmap_data/bulk/ancillary/NLCD/Original/nlcd_2001_landcover_2011_edition_2014_10_10/nlcd_2001_landcover_2011_edition_2014_10_10/nlcd_2001_landcover_2011_edition_2014_10_10.img'
_root_maps = r'/lcmap_data/bulk/klsmith/test-runs/ccd_peeksize_zhe/test-maps/new'
# _exclude = ['h25v10']
_exclude = []

# _region_mask = r'/lcmap_data/bulk/assessment/ecoregions/masks/west_megaregion_mask.tif'
_region_mask = _nlcdpath


def main():

    comb_df = pd.DataFrame()
    for tile in os.listdir(_root_maps):
        if tile not in _exclude:
            print(f'Working tile: {tile}')
            ps = paths(os.path.join(_root_maps, tile), 'LC_Primary')
            h, v = hv_fromdir(tile)
            maskNLCD = clip_filemask(h, v, _nlcdpath)
            maskREGION = clip_filemask(h, v, _region_mask)
            mask = (maskNLCD & maskREGION)

            for year in ps:
                print('Pulling number for {}'.format(os.path.split(year)[-1]))
                yr = year_frompath(year)
                hist = histogram(year, mask)
                data = {k: v for k, v in zip(hist[0], hist[1])}
                data['year'] = yr
                data['tile'] = tile
                temp_df = pd.DataFrame(data, index=[0])
                comb_df = comb_df.append(temp_df)

    print('Saving to CSV')
    comb_df = comb_df.loc[:, ['tile', 'year', 0, 1, 2, 3, 4, 5, 6, 7, 8]]
    comb_df.to_csv('MapCounts.csv', index=False)


def histogram(path, mask):
    ds = gdal.Open(path)
    arr = ds.GetRasterBand(1).ReadAsArray()
    if mask is not None:
        h = np.unique(arr[mask], return_counts=True)
    else:
        h = np.unique(arr, return_counts=True)
    return h


def year_frompath(path):
    parts = path.split('_')
    return int(parts[-1][:4])


def paths(root, strfilter):
    return sorted([os.path.join(root, f) for f in os.listdir(root) if strfilter in f and f[-4:] == '.tif'])


def hv_affine(h, v):
    xmin = -2565585 + h * 5000 * 30
    ymax = 3314805 - v * 5000 * 30
    return xmin, 30, 0, ymax, 0, -30


def hv_fromdir(dirname):
    parts = dirname.split('v')
    return int(parts[0][1:]), int(parts[1])


def geoto_rowcol(x, y, affine):
    # ul_x x_res rot_1 ul_y rot_2 y_res
    col = (x - affine[0]) / affine[1]
    row = (y - affine[3]) / affine[5]

    return int(row), int(col)


def clip_filemask(h, v, path):
    arr = np.zeros(shape=(5000, 5000), dtype=np.int)

    ulx, _, _, uly, _, _ = hv_affine(h, v)

    ds = gdal.Open(path, gdal.GA_ReadOnly)
    aff = ds.GetGeoTransform()
    col_cnt = ds.RasterXSize
    row_cnt = ds.RasterYSize
    row, col = geoto_rowcol(ulx, uly, aff)

    xsize, ysize = (5000, 5000)
    arr_rst, arr_cst = (0, 0)
    arr_rsp, arr_csp = (5000, 5000)
    r_col, r_row = (col, row)

    if row < 0:
        arr_rst = abs(row)
        r_row = 0
        ysize = 5000 - arr_rst

    if col < 0:
        arr_cst = abs(col)
        r_col = 0
        xsize = 5000 - arr_cst

    if col + xsize >= col_cnt:
        xsize = col_cnt - col
        arr_csp = xsize
    if row + ysize >= row_cnt:
        ysize = row_cnt - row
        arr_rsp = ysize

    arr[arr_rst:arr_rsp, arr_cst:arr_csp] = ds.GetRasterBand(1).ReadAsArray(r_col, r_row, xsize, ysize)

    return arr > 0


if __name__ == '__main__':
    t1 = time.time()
    main()
    print(time.time() - t1)

