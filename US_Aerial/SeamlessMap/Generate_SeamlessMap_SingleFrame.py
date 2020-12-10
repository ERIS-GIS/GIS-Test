#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      JLoucks
#
# Created:     09/06/2020
# Copyright:   (c) JLoucks 2020
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import os, csv
import exifread
import re
import csv
import arcpy
def get_type(path):
    if 'INDEX' in path or 'index' in path:
        return 'INDEX'
    else:
        return 'IMAGE'
def get_extension(path):
    extensions = ['.tif','.jpg','.sid','.png','.tiff','.jpeg','.tif','.jp2']
    foundext = []
    for ext in extensions:
        if os.path.exists(path.replace('.TAB',ext)):
            foundext.append(ext)
    if len(foundext) == 0:
        return 'NO IMAGE'
    elif len(foundext) == 1:
        return foundext[0]
    else:
        if '.tif' in foundext:
            return '.tif'
        elif '.jpg' in foundext:
            return '.jpg'
        else:
            return 'DUPLICATES - '+str(foundext)
def get_year(filename,netpath):
    for parse in re.split("\s|(?<!\d)[,.](?!\d)|\\\| |_|\.|\-",netpath):
        if len(parse) == 4 and unicode(parse,'utf-8').isnumeric() and '19' in parse:
            return parse
            break
        else:
            for parse in filename.split('_'):
                if len(parse) == 2 and unicode(parse,'utf-8').isnumeric():
                    if parse in ['00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19']:
                        return '20'+parse
                        break
                    else:
                        return '19'+parse
                        break

def get_source(filename,netpath):
    sources = ['NAIP',
    'USDA',
    'USGS',
    'DOT',
    'ASCS',
    'WALLACE',
    'TOBIN',
    'FAIRCHILD',
    'CAPCOG',
    'COSA',
    'NCTCOG',
    'AMS',
    'HGACOG',
    'TXDOT',
    'ACE',
    'LDOT',
    'ANMAN',
    'ODOT',
    'COF',
    'RGIS',
    'NHAP',
    'NAPP',
    'NASA',
    'VGIN',
    'FDOT',
    'USAF',
    'AMMAN',
    'CTDEP',
    'haha',
    'KUCERA',
    'CAC',
    'SLCO',
    'RISPP',
    'NDOT',
    'AMI',
    'MNDOT',
    'CAS',
    'USN',
    'TVA',
    'MCAO',
    'NCDOT',
    'CH2M',
    'NPS',
    'PAI',
    'BAS',
    'UU',
    'ACA',
    'CDOT',
    'IBWC',
    'NRCS',
    'WDNR',
    'NOAA'
    ]
    foundsource = []
    for i in sources:
        if i in re.split("\s|(?<!\d)[,.](?!\d)|\\\| |_|\.|\-",netpath):
            foundsource.append(i)
    if len(foundsource) == 1:
        return foundsource[0]
    elif len(foundsource) > 1:
        if foundsource[0] == foundsource[1]:
            return foundsource[0]
    elif len(re.split("\s|(?<!\d)[,.](?!\d)|\\\| |_|\.|\-",netpath)) > 2:
        if len(re.split("\s|(?<!\d)[,.](?!\d)|\\\| |_|\.|\-",netpath)[-2]) > 14:
            return 'USGS'
    else:
        return 'different source detected' + str(foundsource)

def get_id(filename):
    return 'id not generated'
def get_filemetadata(filepath,ext):
    datadpi = 'NA'
    databits = 'NA'
    datacomp = 'NA'
    datawidth = 'NA'
    datalength = 'NA'
    imagepath = filepath.replace('.TAB',ext)
    try:
        f = open(imagepath, 'rb')
        tags = exifread.process_file(f)
    except UnicodeEncodeError:
        tags = {}
    except TypeError:
        tags = {}
    for datakey in tags.keys():
        if datakey in ['Image XResolution','Image BitsPerSample','Image ImageWidth','Image Compression','Image ImageLength']:
            if datakey == 'Image XResolution':
                datadpi = str(tags['Image XResolution'])
            elif datakey == 'Image BitsPerSample':
                databits = str(tags['Image BitsPerSample'])
            elif datakey == 'Image Compression':
                datacomp = str(tags['Image Compression'])
            elif datakey == 'Image ImageWidth':
                datawidth = str(tags['Image ImageWidth'])
            elif datakey == 'Image ImageLength':
                datalength = str(tags['Image ImageLength'])
    return [datadpi,databits,datacomp,datawidth,datalength]
def get_filesize(imagepath,ext):
    imagepath = imagepath.replace('.TAB',ext)
    return str(os.stat(imagepath).st_size)
def get_imagepath(imagepath,ext):
    imagepath = imagepath.replace('.TAB',ext)
    return imagepath
def get_spatialres(imagepath,ext):
    try:
        imagepath = imagepath.replace('.TAB',ext)

        cellsizeX = arcpy.GetRasterProperties_management(imagepath,'CELLSIZEX')
        cellsizeY = arcpy.GetRasterProperties_management(imagepath,'CELLSIZEY')
        if cellsizeY > cellsizeX:
            return str(cellsizeY)
        else:
            return str(cellsizeX)
    except Exception:
        return '0'


master_xl = r"C:\Users\JLoucks\Desktop\historical\historical_11202020.csv"
data = []
v = open(master_xl)
r = csv.reader(v)
row0 = r.next()
row0.append('netpath')
row0.append('type')
row0.append('ext')
row0.append('georef')
row0.append('year')
row0.append('source')
row0.append('id')
row0.append('dpi')
row0.append('bits')
row0.append('compression')
row0.append('width')
row0.append('length')
row0.append('filesize')
row0.append('imagepath')
row0.append('spatial_resolution')
data.append(row0)
for item in r:
    netpath = item[0].replace('U:',r'\\cabcvan1nas003\historical')
    filename = item[1]
    item.append(netpath)
    item.append(get_type(netpath))
    item.append(get_extension(netpath))
    item.append('Y')
    item.append(get_year(filename,netpath))
    item.append(get_source(filename,netpath))
    item.append(get_id(filename))

    if get_extension(netpath) in ['.tif','.jpg','.sid','.png','.tiff','.jpeg','.tif','.jp2']:
        metadata = get_filemetadata(netpath, get_extension(netpath))
        item.append(metadata[0])
        item.append(metadata[1])
        item.append(metadata[2])
        item.append(metadata[3])
        item.append(metadata[4])
        item.append(get_filesize(netpath,get_extension(netpath)))
        item.append(get_imagepath(netpath,get_extension(netpath)))
        item.append(get_spatialres(netpath,get_extension(netpath)))
    else:
        item.append('NA')
        item.append('NA')
        item.append('NA')
        item.append('NA')
        item.append('NA')
        item.append('NA')
        item.append('NO IMAGE')
        item.append('No readable spatial res')
    #print item
    data.append(item)
with open(r"C:\Users\JLoucks\Desktop\historical\meta_historical_11202020.csv",'wb') as result_file:
    wr = csv.writer(result_file,dialect='excel')
    wr.writerows(data)
result_file.close()





