import time
import datetime as dt
import base64
from functools import lru_cache
import multiprocessing as mp
from functools import partial

import requests
import numpy as np
import xarray as xr
from osgeo import ogr, gdal
import skimage.exposure as ex

conus_url = 'http://lcmap-test.cr.usgs.gov/ARD_CU_C01_V01'
#conus_url = 'http://lcmap-test.cr.usgs.gov/ARD_AK_C01_V01'
aux_url = 'http://lcmap-test.cr.usgs.gov/AUX_CU_V01'

ardgroups = {'sr_reds': ['LC08_SRB4', 'LE07_SRB3', 'LT05_SRB3', 'LT04_SRB3'],
             'sr_greens': ['LC08_SRB3', 'LE07_SRB2', 'LT05_SRB2', 'LT04_SRB2'],
             'sr_blues': ['LC08_SRB2', 'LE07_SRB1', 'LT05_SRB1', 'LT04_SRB1'],
             'sr_nirs': ['LC08_SRB5', 'LE07_SRB4', 'LT05_SRB4', 'LT04_SRB4'],
             'sr_swir1s': ['LC08_SRB6', 'LE07_SRB5', 'LT05_SRB5', 'LT04_SRB5'],
             'sr_swir2s': ['LC08_SRB7', 'LE07_SRB7', 'LT05_SRB7', 'LT04_SRB7'],
             'thermals': ['LC08_BTB10', 'LE07_BTB6', 'LT05_BTB6', 'LT04_BTB6'],
             'qas': ['LC08_PIXELQA', 'LE07_PIXELQA', 'LT05_PIXELQA', 'LT04_PIXELQA']}

qamap = {'fill': 0,
         'clear': 1,
         'water': 2,
         'shadow': 3,
         'snow': 4,
         'cloud': 5,
         'cloud_conf1': 6,
         'cloud_conf2': 7,
         'cirrus1': 8,
         'cirrus2': 9,
         'occulsion': 10}

def retry(retries):
    def retry_dec(func):
        def wrapper(*args, **kwargs):
            count = 1
            while True:
                try:
                    return func(*args, **kwargs)
                except:
                    count += 1
                    if count > retries:
#                         log.debug('Retry limit exceeded')
                        raise
                    time.sleep(5)
        return wrapper
    return retry_dec

@retry(20)
def getchips(x, y, acquired, ubid, resource=conus_url):
    """
    Make a request to the HTTP API for some chip data.
    """
    chip_url = f'{resource}/chips'
    resp = requests.get(chip_url, params={'x': x,
                                          'y': y,
                                          'acquired': acquired,
                                          'ubid': ubid})
    if not resp.ok:
        resp.raise_for_status()
    
    return resp.json()

@lru_cache()
def getregistry(resource=conus_url):
    """
    Retrieve the spec registry from the API.
    """
    reg_url = f'{resource}/registry'
    return requests.get(reg_url).json()

@lru_cache()
def getspec(ubid):
    """
    Retrieve the appropriate spec information for the corresponding ubid.
    """
    registry = getregistry()
    return next(filter(lambda x: x['ubid'] == ubid, registry), None)

def tonumpy(chip):
    """
    Convert the data response to a numpy array.
    """
    spec = getspec(chip['ubid'])
    data = base64.b64decode(chip['data'])

    chip['data'] = np.frombuffer(data, spec['data_type'].lower()).reshape(*spec['data_shape'])

    return chip

def requestchips(x, y, acquired, ubid, resource=conus_url):
    """
    Helper func to wrap the data conversion around the http response.
    """
    return [tonumpy(c) for c in getchips(x, y, acquired, ubid, resource)]

@lru_cache()
def getgrids(resource=conus_url):
    """
    Retrieve the tile and chip definitions for the grid (geospatial transformation information)
    from the API.
    """
    grid_url = f'{resource}/grid'
    return requests.get(grid_url).json()

@lru_cache()
def getsnap(x, y, resource=conus_url):
    """
    Resource to provide the containing chip and tile upper left coordinates.
    """
    snap_url = f'{resource}/grid/snap'
    return requests.get(snap_url, params={'x': x, 'y': y}).json()

@lru_cache()
def getgrid(grid, resource=conus_url):
    """
    Pull specific grid definition from the list of grids.
    """
    grids = getgrids(resource)
    return next(filter(lambda x: x['name'] == grid, grids), None)

def unscale_sr(sr_arr):
    """
    Unscale and clip SR data.
    """
    out = sr_arr * 0.0001
    out[out < 0] = 0
    out[out > 1] = 1
    
    return out

def align(inx, iny):
    """
    Aligns the coordinate to the chip grid.
    """
    x, y = getsnap(inx, iny)['chip']['proj-pt']
    return int(x), int(y)

def zoomout(x, y, factor=1):
    """
    Generate a list coordinates, centered on a given one.
    """
    ul = (x - 3000 * factor, y + 3000 * factor)
    lr = (x + 3000 * factor, y - 3000 * factor)
    
    
    return [(x, y) for x in range(ul[0], lr[0] + 3000, 3000)
            for y in range(ul[1], lr[1] - 3000, -3000)]

def mosaicdate(coord_ls, date, ubid):
    """
    Create a single date mosaic for single ubid.
    coord_ls is assumed to align to the chip grid.
    """
    ulx, uly = find_ul(coord_ls)
    affine = buildaffine(ulx, uly)
    acq = '/'.join([date, date])
    arr = np.zeros(shape=findrowscols(coord_ls))
    
    for coord in coord_ls:
        data = requestchips(coord[0], coord[1], acq, ubid)[0]
        r, c = transform_geo(data['x'], data['y'], affine)
        arr[r:r + 100, c:c + 100] = data['data']
        
    return arr

def mp_mosaicdate(coord_ls, date, ubid, cpu, large=False):
    """
    A rewrite of mosaicdate, except to use multiple cpu's or threads.
    """
    ulx, uly = find_ul(coord_ls)
    affine = buildaffine(ulx, uly)
    acq = '/'.join([date, date])
    if large:
        arr = np.memmap('temp.dat', shape=findrowscols(coord_ls), dtype=np.int16)
    else:
        arr = np.zeros(shape=findrowscols(coord_ls), dtype=np.int16)
    print(affine)
    print(arr.shape)
    
    func = partial(requestchips,
                   acquired=acq,
                   ubid=ubid)
    
    with mp.Pool(cpu) as pool:
        data = pool.starmap(func, coord_ls)
        
    for chip in data:
        if not chip:
            continue
        r, c = transform_geo(chip[0]['x'], chip[0]['y'], affine)
        arr[r:r + 100, c:c + 100] = chip[0]['data']
        
    return arr

def find_ul(coord_ls):
    """
    Helper function to identify the upper left coordinate from a list of coordinates.
    """
    xs, ys = zip(*coord_ls)
    return min(xs), max(ys)

def find_lr(coord_ls):
    """
    Helper function to identify the lower left coordinate from a list of coordinates.
    Note this is not the LR of the extent.
    """
    xs, ys = zip(*coord_ls)
    return max(xs), min(ys)

def findrowscols(coord_ls):
    """
    Find the total number of rows and cols contained in the coord_ls.
    """
    extent = np.array(find_ul(coord_ls)) - np.array(find_lr(coord_ls))
    col, row = np.abs(extent / 30) + 100
    return int(row), int(col)

# Functions to help with affine transformations
def transform_geo(x, y, affine):
    """
    Transform from a geospatial (x, y) to (row, col).
    """
    col = (x - affine[0]) / affine[1]
    row = (y - affine[3]) / affine[5]
    return int(row), int(col)

def transform_rowcol(row, col, affine):
    """
    Tranform from (row, col) to geospatial (x, y).
    """
    x = affine[0] + col * affine[1]
    y = affine[3] + row * affine[5]
    return x, y

def buildaffine(ulx, uly, size=30):
    """
    Returns a standard GDAL geotransform affine tuple for 30m pixels.
    """
    return (ulx, size, 0, uly, 0, -size)

def rasterize(shapepath, pixel_size=30):
    """
    Adapted from:
    http://pcjericks.github.io/py-gdalogr-cookbook/vector_layers.html#convert-vector-layer-to-array
    """
    # Open the data source and read in the extent
    source_ds = ogr.Open(shapepath)
    source_layer = source_ds.GetLayer()
    source_srs = source_layer.GetSpatialRef()
    x_min, x_max, y_min, y_max = source_layer.GetExtent()
    x_min, y_max = align(x_min, y_max)
    x_max, y_min = align(x_max, y_min)
    x_max += 3000 # Align will pull it in, so let's buffer it back out.
    y_min -= 3000

    # Create the destination data source
    x_res = int((x_max - x_min) / pixel_size)
    y_res = int((y_max - y_min) / pixel_size)
    target_ds = gdal.GetDriverByName('MEM').Create('', x_res, y_res, gdal.GDT_Byte)
    target_ds.SetGeoTransform((x_min, pixel_size, 0, y_max, 0, -pixel_size))
    band = target_ds.GetRasterBand(1)
    # band.SetNoDataValue(NoData_value)

    # Rasterize
    gdal.RasterizeLayer(target_ds, [1], source_layer, burn_values=[1])

    # Read as array
    array = band.ReadAsArray()
    
    source_ds = None
    target_ds = None
    
    return array, (x_min, y_max, x_max, y_min)

def buildrequestls(trutharr, ulx, uly):
    """
    Build a list of what chips to actually request based on an array of 0 or other.
    Assumes a 30m pixel size and the ulx/uly have already been aligned.
    """
    aff = buildaffine(ulx, uly, 30)
    rows, cols = trutharr.shape
    truth = trutharr.astype(np.bool)

    return [transform_rowcol(row, col, aff) 
            for row in range(0, rows, 100) 
            for col in range(0, cols, 100)
            if np.any(truth[row:row + 100, col:col + 100])]
                
def todatetime(chipdate):
    """
    Convert the a Chip's acquisition date time string to a python datetime object.
    """
    return dt.datetime.strptime(chipdate[:10], '%Y-%m-%d')

def requestgroup(x, y, acq, group):
    """
    Request all ubids in an associated grouping.
    """
    ret = []
    for u in group:
        data = requestchips(x, y, acq, u)
        if data:
            ret.extend([d for d in data])
        
    return ret

def chipsasxr(chips, name=None):
    """
    Takes a set of chips converts it to a nice xarray dataframe.
    Assumes all chips are for the same x/y.
    """
    grid = getgrid('chip')
    
    return xr.DataArray(np.stack([c['data'] for c in chips]),
                        dims=['acquired', 'y', 'x'],
                        coords={'acquired': [todatetime(c['acquired']) for c in chips],
                                'y': np.arange(chips[0]['y'], chips[0]['y'] - 3000, -30),
                                'x': np.arange(chips[0]['x'], chips[0]['x'] + 3000, 30),
                                'chip_x': chips[0]['x'],
                                'chip_y': chips[0]['y'],
                                'projection': grid['proj']},
                       name=name)

def nnextract(arr, row1, col1, row2, col2):
    """
    Extract values from an array along a given line, should emulate Nearest Neighbor (needs to be vetted) ...
    This is for C-ordered arrays
    """
    l = int(np.hypot(col2 - col1, row2 - row1))
    idx_col, idx_row = np.linspace(col1, col2, l), np.linspace(row1, row2, l)
    
    return arr[idx_row.astype(np.int), idx_col.astype(np.int)]

def xrndvi(reddf, nirdf, qadf, qamap=qamap):
    """
    Calculate a NDVI xarray dataframe.
    """
    nirm = nirdf.where((qadf == qamap['clear']) | (qadf == qamap['water']))
    redm = reddf.where((qadf == qamap['clear']) | (qadf == qamap['water']))
    return 100 * (nirm - redm) / (nirm + redm) + 100

def xrmaxndvi(spectdf, ndvidf):
    """
    Filter the spectral data frame by the max NDVI.
    """
    return spectdf.where(ndvidf==ndvidf.max('acquired'), drop=False).max('acquired')

def unpackqa(packed, qamap=qamap):
    """
    Unpack the pixel QA into a set heirarchy for an array.
    fill > cloud > shadow > snow > water > clear
    """
    # Assumed all fill unless told otherwise
    unpacked = np.full(packed.shape, qamap['fill'])

    unpacked[packed & 1 << qamap['clear'] > 0] = qamap['clear']
    unpacked[packed & 1 << qamap['water'] > 0] = qamap['water']
    unpacked[packed & 1 << qamap['snow'] > 0] = qamap['snow']
    unpacked[packed & 1 << qamap['shadow'] > 0] = qamap['shadow']
    unpacked[packed & 1 << qamap['cloud'] > 0] = qamap['cloud']
    
    return unpacked

def unpackqachip(qachip, qamap=qamap):
    """
    Unpack a pixel QA chip's data into a heirarchy of values.
    """
    qachip['data'] = unpackqa(qachip['data'], qamap)
    return qachip

def unpackqachips(qachips, qamap=qamap):
    """
    Unpack a sequence of pixel QA chips.
    """
    return [unpackqachip(q) for q in qachips]

def requestqa(x, y, acq, qas=ardgroups['qas'], qamap=qamap):
    """
    Helper function to request and unpack the pixel QA chips.
    """
    data = unpackqachips(requestgroup(x, y, acq, qas))
    
    return [d for d in data]

def maxattr(dfls, attr):
    obj = max(dfls, key=lambda x: int(getattr(x, attr)))
    return int(getattr(obj, attr))

def minattr(dfls, attr):
    obj = min(dfls, key=lambda x: int(getattr(x, attr)))
    return int(getattr(obj, attr))

def sortxrls(dfls):
    return sorted(dfls, key=lambda c: (int(getattr(c, 'chip_x')), 
                                       int(getattr(c, 'chip_y'))))

def xrmosaic(dfls, extent=None):
    if extent is None:
        ulx = minattr(dfls, 'chip_x')
        uly = maxattr(dfls, 'chip_y')
        lrx = maxattr(dfls, 'chip_x')
        lry = minattr(dfls, 'chip_y')
        rows = int((uly - lry) / 30) + 100
        cols = int((lrx - ulx) / 30) + 100
        
    else:
        ulx, uly, lrx, lry = extent
        rows = int((uly - lry) / 30)
        cols = int((lrx - ulx) / 30)
    
    
    out = np.zeros(shape=(rows, cols), dtype=np.int16)
    aff = buildaffine(extent[0], extent[1], 30)
    
    for df in dfls:
        if df.size == 0:
            continue
        r, c = transform_geo(int(df.chip_x), int(df.chip_y), aff)
        out[r:r + 100, c:c + 100] = df.values
        
    return out