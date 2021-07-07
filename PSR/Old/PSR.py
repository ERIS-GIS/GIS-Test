#-------------------------------------------------------------------------------
# Name:        Physical Setting Report
#-------------------------------------------------------------------------------

import shutil, csv
import cx_Oracle,urllib, glob
import arcpy, os, numpy
from datetime import datetime
import getDirectionText
import gc, time, timeit
import traceback
from numpy import gradient
from numpy import arctan2, arctan, sqrt
import PSR_config
import json
import ssl

def returnUniqueSetString_musym(tableName):
    data = arcpy.da.TableToNumPyArray(tableName, ['mukey', 'musym'])
    uniques = numpy.unique(data[data['musym']!='NOTCOM']['mukey'])
    if len(uniques) == 0:
        return ''
    else:
        myString = '('
        for item in uniques:
            myString = myString + "'" + str(item) + "', "
        myString = myString[0:-2] + ")"
        return myString

def returnUniqueSetString(tableName, fieldName):
    data = arcpy.da.TableToNumPyArray(tableName, [fieldName])
    uniques = numpy.unique(data[fieldName])
    if len(uniques) == 0:
        return ''
    else:
        myString = '('
        for item in uniques:
            myString = myString + "'" + str(item) + "', "
        myString = myString[0:-2] + ")"
        return myString

#check if an array contain the same values
def checkIfUniqueValue(myArray):
    value = myArray[0]
    for i in range(0,len(myArray)):
        if(myArray[i] != value):
            return False
    return True

def returnMapUnitAttribute(dataarray, mukey, attributeName):   #water, urban land is not in dataarray, so will return '?'
    #data = dataarray[dataarray['mukey'] == mukey][attributeName[0:10]]   #0:10 is to account for truncating of the field names in .dbf
    data = dataarray[dataarray['mukey'] == mukey][attributeName]
    if (len(data) == 0):
        return "?"
    else:
        if(checkIfUniqueValue):
            if (attributeName == 'brockdepmin' or attributeName == 'wtdepannmin'):
                if data[0] == -99:
                    return 'null'
                else:
                    return str(data[0]) + 'cm'
            return str(data[0])  #will convert to str no matter what type
        else:
            return "****ERROR****"

def returnComponentAttribute_rvindicatorY(dataarray,mukey):
    resultarray = []
    dataarray1 = dataarray[dataarray['mukey'] == mukey]
    data = dataarray1[dataarray1['majcompflag'] =='Yes']      # 'majcompfla' needs to be used for .dbf table
    comps = data[['cokey','compname','comppct_r']]
    comps_sorted = numpy.sort(numpy.unique(comps), order = 'comppct_r')[::-1]     #[::-1] gives descending order
    for comp in comps_sorted:
        horizonarray = []
        keyname = comp[1] + '('+str(comp[2])+'%)'
        horizonarray.append([keyname])

        selection = data[data['cokey']==comp[0]][['mukey','cokey','compname','comppct_r','hzname','hzdept_r','hzdepb_r','texdesc']]
        selection_sorted = numpy.sort(selection, order = 'hzdept_r')
        for item in selection_sorted:
            horizonlabel = 'horizon ' + item['hzname'] + '(' + str(item['hzdept_r']) + 'cm to '+ str(item['hzdepb_r']) + 'cm)'
            horizonTexture = item['texdesc']
            horizonarray.append([horizonlabel,horizonTexture])
        resultarray.append(horizonarray)

    return resultarray

def returnComponentAttribute(dataarray,mukey):
    resultarray = []
    dataarray1 = dataarray[dataarray['mukey'] == mukey]
    data = dataarray1[dataarray1['majcompflag'] =='Yes']      # 'majcompfla' needs to be used for .dbf table
    comps = data[['cokey','compname','comppct_r', 'rv']]
    comps_sorted = numpy.sort(numpy.unique(comps), order = 'comppct_r')[::-1]     #[::-1] gives descending order
    comps_sorted_rvYes = comps_sorted[comps_sorted['rv'] == 'Yes']     # there are only two values in 'rv' field: Yes and No
    for comp in comps_sorted_rvYes:
        horizonarray = []
        keyname = comp[1] + '('+str(comp[2])+'%)'
        horizonarray.append([keyname])

        data_rvYes = data[data['rv']== 'Yes']
        selection = data_rvYes[data_rvYes['cokey']==comp[0]][['mukey','cokey','compname','comppct_r','hzname','hzdept_r','hzdepb_r','texdesc']]
        selection_sorted = numpy.sort(selection, order = 'hzdept_r')
        for item in selection_sorted:
            horizonlabel = 'horizon ' + item['hzname'] + '(' + str(item['hzdept_r']) + 'cm to '+ str(item['hzdepb_r']) + 'cm)'
            horizonTexture = item['texdesc']
            horizonarray.append([horizonlabel,horizonTexture])
        if len(selection_sorted)> 0:
            resultarray.append(horizonarray)
        else:
            horizonarray.append(['No representative horizons available.',''])
            resultarray.append(horizonarray)

    return resultarray

def addBuffertoMxd(bufferName,thedf):    # note: buffer is a shapefile, the name doesn't contain .shp

    bufferLayer = arcpy.mapping.Layer(bufferlyrfile)
    bufferLayer.replaceDataSource(scratch_folder,"SHAPEFILE_WORKSPACE",bufferName)
    arcpy.mapping.AddLayer(thedf,bufferLayer,"Top")
    thedf.extent = bufferLayer.getSelectedExtent(False)
    thedf.scale = thedf.scale * 1.1

def addOrdergeomtoMxd(ordergeomName, thedf):
    orderGeomLayer = arcpy.mapping.Layer(orderGeomlyrfile)
    orderGeomLayer.replaceDataSource(scratch_folder,"SHAPEFILE_WORKSPACE",ordergeomName)
    arcpy.mapping.AddLayer(thedf,orderGeomLayer,"Top")

def getElevation(dataset,fields):
    pntlist={}
    with arcpy.da.SearchCursor(dataset,fields) as uc:
        for row in uc:
            pntlist[row[2]]=(row[0],row[1])
    del uc

    params={}
    params['XYs']=pntlist
    params = urllib.urlencode(params)
    inhouse_esri_geocoder = r"https://gisserverprod.glaciermedia.ca/arcgis/rest/services/GPTools_temp/pntElevation2/GPServer/pntElevation2/execute?env%3AoutSR=&env%3AprocessSR=&returnZ=false&returnM=false&f=pjson"
    context = ssl._create_unverified_context()
    f = urllib.urlopen(inhouse_esri_geocoder,params,context=context)
    results =  json.loads(f.read())
    result = eval( results['results'][0]['value'])

    check_field = arcpy.ListFields(dataset,"Elevation")
    if len(check_field)==0:
        arcpy.AddField_management(dataset, "Elevation", "DOUBLE", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")
    with arcpy.da.UpdateCursor(dataset,["Elevation"]) as uc:
        for row in uc:
            row[0]=-999
            uc.updateRow(row)
    del uc

    with arcpy.da.UpdateCursor(dataset,['Elevation',fields[-1]]) as uc:
        for row in uc:
            if result[row[1]] !='':
                row[0]= result[row[1]]
                uc.updateRow(row)
    del row
    return dataset
try:

###############################################################################################################
    #parameters to change for deployment
    connectionString = PSR_config.connectionString#'ERIS_GIS/gis295@GMTEST.glaciermedia.inc'
    report_path = PSR_config.report_path#"\\cabcvan1obi002\ErisData\Reports\test\noninstant_reports"
    viewer_path = PSR_config.viewer_path#"\\CABCVAN1OBI002\ErisData\Reports\test\viewer"
    upload_link = PSR_config.upload_link#"http://CABCVAN1OBI002/ErisInt/BIPublisherPortal/Viewer.svc/"
    #production: upload_link = r"http://CABCVAN1OBI002/ErisInt/BIPublisherPortal_prod/Viewer.svc/"
    reportcheck_path = PSR_config.reportcheck_path#'\\cabcvan1obi002\ErisData\Reports\test\reportcheck'
###############################################################################################################

    OrderIDText = arcpy.GetParameterAsText(0)
    # scratch_gdb = arcpy.env.scratchGDB
    scratch_folder = arcpy.env.scratchFolder
    arcpy.env.workspace = scratch_folder
    arcpy.env.overwriteOutput = True
    arcpy.env.OverWriteOutput = True

# LOCAL #######################################################################################################
    OrderIDText = '1129956'
    scratch_gdb = arcpy.CreateFileGDB_management(scratch_folder,r"scratch_gdb.gdb")   # for tables to make Querytable
    scratch_gdb = os.path.join(scratch_folder,r"scratch_gdb.gdb")
    arcpy.AddMessage(scratch_folder)
###############################################################################################################
    try:
        start = timeit.default_timer()
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        cur.execute("select order_num, address1, city, provstate from orders where order_id =" + OrderIDText)
        t = cur.fetchone()

        OrderNumText = str(t[0])
        AddressText = str(t[1])+","+str(t[2])+","+str(t[3])
        ProvStateText = str(t[3])

        cur.execute("select geometry_type, geometry, radius_type  from eris_order_geometry where order_id =" + OrderIDText)
        t = cur.fetchone()

        cur.callproc('eris_psr.ClearOrder', (OrderIDText,))

        OrderType = str(t[0])
        OrderCoord = eval(str(t[1]))
        RadiusType = str(t[2])
    finally:
        cur.close()
        con.close()

    print "--- starting " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    ##bufferDist_topo = "2 MILES"
    ##bufferDist_flood = "1 MILES"
    ##bufferDist_wetland = "1 MILES"
    ##bufferDist_geol = "1 MILES"
    ##bufferDist_soil = "0.25 MILES"
    ##bufferDist_wwells = "0.5 MILES"
    ##bufferDist_ogw = "0.5 MILES"
    ##bufferDist_Radon = "1 MILES"

    searchRadius = {}
    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        cur.execute("select DS_OID, SEARCH_RADIUS, REPORT_SOURCE from order_radius_psr where order_id =" + str(OrderIDText))
        items = cur.fetchall()

        for t in items:
            dsoid = t[0]
            radius = t[1]
            reportsource = t[2]

            searchRadius[str(dsoid)] = float(radius)
    finally:
        cur.close()
        con.close()

    # 9334 SSURGO
    # 10683 FEMA FLOOD
    # 10684 US WETLAND
    # 10685 US GEOLOGY
    # 10688 RADON ZONE
    # 10689 INDOOR RADON
    # others:
    # 5937 PCWS
    # 8710 PWS
    # 10093 PWSV
    # 10209 WATER WELL
    # 9144 OGW  (one state source..)
    # 10061 OGW (another state source..)
    srGoogle = arcpy.SpatialReference(3857)   #web mercator
    srWGS84 = arcpy.SpatialReference(4326)   #WGS84
    
    bufferDist_topo = '1 MILES'
    bufferDist_flood = str(searchRadius['10683']) + ' MILES'
    bufferDist_wetland = str(searchRadius['10684']) + ' MILES'
    bufferDist_geol = str(searchRadius['10685']) + ' MILES'
    bufferDist_soil = str(searchRadius['9334']) + ' MILES'
    buffer_dist_sp = "0.5 MILES" #survey and pipeline
    #bufferDist_ogw = "0.5 MILES"
    bufferDist_radon = str(searchRadius['10689']) + ' MILES'    # use teh indoor radon one

    dsoid_wells = []
    dsoid_wells_maxradius = '10093'     # 10093 is a federal source, PWSV
    for key in searchRadius:
        if key not in ['9334', '10683', '10684', '10685', '10688','10689', '10695', '10696']:       #10695 is US topo, 10696 is HTMC, 10688 and 10689 are radons
            dsoid_wells.append(key)
            if (searchRadius[key] > searchRadius[dsoid_wells_maxradius]):
                dsoid_wells_maxradius = key
            #radii_wells.append(searchRadius[key])

    connectionPath = PSR_config.connectionPath

    orderGeomlyrfile_point = PSR_config.orderGeomlyrfile_point#"E:\GISData\PSR\python\mxd\SiteMaker.lyr"
    orderGeomlyrfile_polyline = PSR_config.orderGeomlyrfile_polyline#"E:\GISData\PSR\python\mxd\orderLine.lyr"
    orderGeomlyrfile_polygon = PSR_config.orderGeomlyrfile_polygon#"E:\GISData\PSR\python\mxd\orderPoly.lyr"
    bufferlyrfile = PSR_config.bufferlyrfile#"E:\GISData\PSR\python\mxd\buffer.lyr"
    topowhitelyrfile =PSR_config.topowhitelyrfile# r"E:\GISData\PSR\python\mxd\topo_white.lyr"
    gridlyrfile = PSR_config.gridlyrfile#"E:\GISData\PSR\python\mxd\Grid_hollow.lyr"
    relieflyrfile = PSR_config.relieflyrfile#"E:\GISData\PSR\python\mxd\relief.lyr"

    masterlyr_topo = PSR_config.masterlyr_topo#"E:\GISData\Topo_USA\masterfile\CellGrid_7_5_Minute.shp"
    # data_topo = PSR_config.data_topo#"E:\GISData\Topo_USA\masterfile\Cell_PolygonAll.shp"
    csvfile_topo = PSR_config.csvfile_topo#"E:\GISData\Topo_USA\masterfile\All_USTopo_T_7.5_gda_results.csv"
    tifdir_topo = PSR_config.tifdir_topo#"\\cabcvan1fpr009\DATA_GIS\USGS_currentTopo_Geotiff"
    data_shadedrelief = PSR_config.data_shadedrelief#"\\cabcvan1fpr009\DATA_GIS\US_DEM\CellGrid_1X1Degree_NW.shp"

    data_geol = PSR_config.data_geol#'E:\GISData\Data\PSR\PSR.gdb\GEOL_DD_MERGE'
    data_flood = PSR_config.data_flood#'E:\GISData\Data\PSR\PSR.gdb\S_Fld_haz_Ar_merged'
    data_floodpanel = PSR_config.data_floodpanel#'E:\GISData\Data\PSR\PSR.gdb\S_FIRM_PAN_MERGED'
    data_wetland = PSR_config.data_wetland#'E:\GISData\Data\PSR\PSR.gdb\Merged_wetland_Final'
    eris_wells = PSR_config.eris_wells#"E:\GISData\PSR\python\mxd\ErisWellSites.lyr"   #which contains water, oil/gas wells etc.

    path_shadedrelief = PSR_config.path_shadedrelief#"\\cabcvan1fpr009\DATA_GIS\US_DEM\hillshade13"
    datalyr_wetland = PSR_config.datalyr_wetland#"E:\GISData\PSR\python\mxd\wetland.lyr"
    # datalyr_wetlandNY = PSR_config.datalyr_wetlandNY
    datalyr_wetlandNYkml = PSR_config.datalyr_wetlandNYkml#u'E:\\GISData\\PSR\\python\\mxd\\wetlandNY_kml.lyr'
    datalyr_wetlandNYAPAkml = PSR_config.datalyr_wetlandNYAPAkml#r"E:\GISData\PSR\python\mxd\wetlandNYAPA_kml.lyr"
    datalyr_plumetacoma = PSR_config.datalyr_plumetacoma#r"E:\GISData\PSR\python\mxd\Plume.lyr"
    datalyr_flood = PSR_config.datalyr_flood#"E:\GISData\PSR\python\mxd\flood.lyr"
    datalyr_geology = PSR_config.datalyr_geology#"E:\GISData\PSR\python\mxd\geology.lyr"
    datalyr_contour = PSR_config.datalyr_contour#"E:\GISData\PSR\python\mxd\contours_largescale.lyr"

    imgdir_dem = PSR_config.imgdir_dem#"\\cabcvan1fpr009\DATA_GIS\US_DEM\DEM13"
    imgdir_demCA = PSR_config.imgdir_demCA#r"\\cabcvan1fpr009\US_DEM\DEM1"
    masterlyr_dem = PSR_config.masterlyr_dem#"\\cabcvan1fpr009\DATA_GIS\US_DEM\CellGrid_1X1Degree_NW_imagename_update.shp"
    masterlyr_demCA =PSR_config.masterlyr_demCA #r"\\cabcvan1fpr009\US_DEM\Canada_DEM_edited.shp"
    masterlyr_states = PSR_config.masterlyr_states#"E:\GISData\PSR\python\mxd\USStates.lyr"
    masterlyr_counties = PSR_config.masterlyr_counties#"E:\GISData\PSR\python\mxd\USCounties.lyr"
    masterlyr_cities = PSR_config.masterlyr_cities#"E:\GISData\PSR\python\mxd\USCities.lyr"
    masterlyr_NHTowns = PSR_config.masterlyr_NHTowns#"E:\GISData\PSR\python\mxd\NHTowns.lyr"
    masterlyr_zipcodes = PSR_config.masterlyr_zipcodes#"E:\GISData\PSR\python\mxd\USZipcodes.lyr"

    mxdfile_topo = PSR_config.mxdfile_topo#"E:\GISData\PSR\python\mxd\topo.mxd"
    mxdfile_topo_Tacoma = PSR_config.mxdfile_topo_Tacoma#"E:\GISData\PSR\python\mxd\topo.mxd"
    mxdMMfile_topo = PSR_config.mxdMMfile_topo#"E:\GISData\PSR\python\mxd\topoMM.mxd"
    mxdMMfile_topo_Tacoma = PSR_config.mxdMMfile_topo_Tacoma #r"E:\GISData\PSR\python\mxd\topoMM_tacoma.mxd"
    mxdfile_relief =  PSR_config.mxdfile_relief#"E:\GISData\PSR\python\mxd\shadedrelief.mxd"
    mxdMMfile_relief =  PSR_config.mxdMMfile_relief#"E:\GISData\PSR\python\mxd\shadedreliefMM.mxd"
    mxdfile_wetland = PSR_config.mxdfile_wetland#"E:\GISData\PSR\python\mxd\wetland.mxd"
    mxdfile_wetlandNY = PSR_config.mxdfile_wetlandNY#"E:\GISData\PSR\python\mxd\wetland.mxd"
    mxdMMfile_wetland = PSR_config.mxdMMfile_wetland#"E:\GISData\PSR\python\mxd\wetlandMM.mxd"
    mxdMMfile_wetlandNY = PSR_config.mxdMMfile_wetlandNY
    mxdfile_flood = PSR_config.mxdfile_flood#"E:\GISData\PSR\python\mxd\flood.mxd"
    mxdMMfile_flood = PSR_config.mxdMMfile_flood#"E:\GISData\PSR\python\mxd\floodMM.mxd"
    mxdfile_geol = PSR_config.mxdfile_geol#"E:\GISData\PSR\python\mxd\geology.mxd"
    mxdMMfile_geol = PSR_config.mxdMMfile_geol#"E:\GISData\PSR\python\mxd\geologyMM.mxd"
    mxdfile_soil = PSR_config.mxdfile_soil#"E:\GISData\PSR\python\mxd\soil.mxd"
    mxdMMfile_soil = PSR_config.mxdMMfile_soil#"E:\GISData\PSR\python\mxd\soilMM.mxd"
    mxdfile_wells = PSR_config.mxdfile_wells#"E:\GISData\PSR\python\mxd\wells.mxd"
    mxdMMfile_wells = PSR_config.mxdMMfile_wells#"E:\GISData\PSR\python\mxd\wellsMM.mxd"

    outputjpg_topo = os.path.join(scratch_folder, OrderNumText+'_US_TOPO.jpg')
    outputjpg_relief = os.path.join(scratch_folder, OrderNumText+'_US_RELIEF.jpg')
    outputjpg_wetland = os.path.join(scratch_folder, OrderNumText+'_US_WETL.jpg')
    outputjpg_wetlandNY = os.path.join(scratch_folder, OrderNumText+'_NY_WETL.jpg')
    outputjpg_flood = os.path.join(scratch_folder, OrderNumText+'_US_FLOOD.jpg')
    outputjpg_soil = os.path.join(scratch_folder, OrderNumText+'_US_SOIL.jpg')
    outputjpg_geol = os.path.join(scratch_folder, OrderNumText+'_US_GEOL.jpg')
    outputjpg_wells = os.path.join(scratch_folder, OrderNumText+'_US_WELLS.jpg')
    outputjpg_sp = os.path.join(scratch_folder, OrderNumText+'_US_SURVEY_PIPELINE.jpg')

    srGCS83 = PSR_config.srGCS83#arcpy.SpatialReference(os.path.join(connectionPath, r"projections\GCSNorthAmerican1983.prj"))



    erisid = 0

    point = arcpy.Point()
    array = arcpy.Array()
    sr = arcpy.SpatialReference()
    sr.factoryCode = 4269  # requires input geometry is in 4269
    sr.XYTolerance = .00000001
    sr.scaleFactor = 2000
    sr.create()
    featureList = []
    for feature in OrderCoord:
        # For each coordinate pair, set the x,y properties and add to the Array object.
        for coordPair in feature:
            point.X = coordPair[0]
            point.Y = coordPair[1]
            sr.setDomain (point.X, point.X, point.Y, point.Y)
            array.add(point)
        if OrderType.lower()== 'point':
            feat = arcpy.Multipoint(array, sr)
        elif OrderType.lower() =='polyline':
            feat  = arcpy.Polyline(array, sr)
        else :
            feat = arcpy.Polygon(array,sr)
        array.removeAll()

        # Append to the list of Polygon objects
        featureList.append(feat)

    orderGeometry= os.path.join(scratch_folder,"orderGeometry.shp")
    arcpy.CopyFeatures_management(featureList, orderGeometry)
    del featureList
    arcpy.DefineProjection_management(orderGeometry, srGCS83)

    arcpy.AddField_management(orderGeometry, "xCentroid", "DOUBLE", 18, 11)
    arcpy.AddField_management(orderGeometry, "yCentroid", "DOUBLE", 18, 11)

    xExpression = '!SHAPE.CENTROID.X!'
    yExpression = '!SHAPE.CENTROID.Y!'

    arcpy.CalculateField_management(orderGeometry, 'xCentroid', xExpression, "PYTHON_9.3")
    arcpy.CalculateField_management(orderGeometry, 'yCentroid', yExpression, "PYTHON_9.3")

    arcpy.AddField_management(orderGeometry, "UTM", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
    arcpy.CalculateUTMZone_cartography(orderGeometry, 'UTM')
    UT= arcpy.SearchCursor(orderGeometry)
    UTMvalue = ''
    Lat_Y = 0
    Lon_X = 0
    for row in UT:
        UTMvalue = str(row.getValue('UTM'))[41:43]
        Lat_Y = row.getValue('yCentroid')
        Lon_X = row.getValue('xCentroid')
    del UT
    if UTMvalue[0]=='0':
        UTMvalue=' '+UTMvalue[1:]
    out_coordinate_system = arcpy.SpatialReference('NAD 1983 UTM Zone %sN'%UTMvalue)

    orderGeometryPR = os.path.join(scratch_folder, "ordergeoNamePR.shp")
    arcpy.Project_management(orderGeometry, orderGeometryPR, out_coordinate_system)
    arcpy.AddField_management(orderGeometryPR, "xCenUTM", "DOUBLE", 18, 11)
    arcpy.AddField_management(orderGeometryPR, "yCenUTM", "DOUBLE", 18, 11)

    xExpression = '!SHAPE.CENTROID.X!'
    yExpression = '!SHAPE.CENTROID.Y!'

    arcpy.CalculateField_management(orderGeometryPR, 'xCenUTM', xExpression, "PYTHON_9.3")
    arcpy.CalculateField_management(orderGeometryPR, 'yCenUTM', yExpression, "PYTHON_9.3")

    del point
    del array

    ##in_rows = arcpy.SearchCursor(orderGeometryPR)
    ##for in_row in in_rows:
    ##    xCentroid = in_row.xCentroid
    ##    yCentroid = in_row.yCentroid
    ##del in_row
    ##del in_rows

    if OrderType.lower()== 'point':
        orderGeomlyrfile = orderGeomlyrfile_point
    elif OrderType.lower() =='polyline':
        orderGeomlyrfile = orderGeomlyrfile_polyline
    else:
        orderGeomlyrfile = orderGeomlyrfile_polygon

    spatialRef = out_coordinate_system

    # determine if needs to be multipage
    # according to Raf: will be multipage if line is over 1/4 mile, or polygon is over 1 sq miles
    # need to check the extent of the geometry
    geomExtent = arcpy.Describe(orderGeometryPR).extent
    multipage_topo = False
    multipage_relief = False
    multipage_wetland = False
    multipage_flood = False
    multipage_geology = False
    multipage_soil = False
    multipage_wells = False
    multipage_sp = False # survey and pipeline
    
    need_viewer = 'N'
    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        cur.execute("select psr_viewer from order_viewer where order_id =" + str(OrderIDText))
        t = cur.fetchone()
        if t != None:
            need_viewer = t[0]
        if need_viewer:
            viewerdir_kml = os.path.join(scratch_folder,OrderNumText+'_psrkml')
            if not os.path.exists(viewerdir_kml):
                os.mkdir(viewerdir_kml)

    finally:
        cur.close()
        con.close()

    gridsize = "2 MILES"
    if geomExtent.width > 1300 or geomExtent.height > 1300:
        multipage_wetland = True
        multipage_flood = True
        multipage_geology = True
        multipage_soil = True
        multipage_topo = True
        multipage_relief = True
        multipage_topo = True
        multipage_wells = True
        multipage_sp = True
    if geomExtent.width > 500 or geomExtent.height > 500:
        multipage_topo = True
        multipage_relief = True
        multipage_topo = True
        multipage_wells = True
    # multipage_sp = True
# Survey and pipeline
    print("--- starting survey & pipeline report " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    buffer_sp_fc = os.path.join(scratch_folder,"buffer_sp.shp")
    arcpy.Buffer_analysis(orderGeometryPR, buffer_sp_fc, buffer_dist_sp)

    point = arcpy.Point()
    array = arcpy.Array()
    feature_list = []

    width = arcpy.Describe(buffer_sp_fc).extent.width/2
    height = arcpy.Describe(buffer_sp_fc).extent.height/2

    if (width/height > 7/7):    #7/7 now since adjusted the frame to square
        # wider shape
        height = width/7*7
    else:
        # longer shape
        width = height/7*7
    xCentroid = (arcpy.Describe(buffer_sp_fc).extent.XMax + arcpy.Describe(buffer_sp_fc).extent.XMin)/2
    yCentroid = (arcpy.Describe(buffer_sp_fc).extent.YMax + arcpy.Describe(buffer_sp_fc).extent.YMin)/2
    
    if multipage_sp == True:
        width = width + 6400     #add 2 miles to each side, for multipage
        height = height + 6400   #add 2 miles to each side, for multipage

    point.X = xCentroid-width
    point.Y = yCentroid+height
    array.add(point)
    point.X = xCentroid+width
    point.Y = yCentroid+height
    array.add(point)
    point.X = xCentroid+width
    point.Y = yCentroid-height
    array.add(point)
    point.X = xCentroid-width
    point.Y = yCentroid-height
    array.add(point)
    point.X = xCentroid-width
    point.Y = yCentroid+height
    array.add(point)
    feat = arcpy.Polygon(array,spatialRef)
    array.removeAll()
    feature_list.append(feat)
    
    data_frame_sp_fc = os.path.join(scratch_folder, "data_frame_sp.shp")
    arcpy.CopyFeatures_management(feature_list, data_frame_sp_fc)
    data_frame_desc = arcpy.Describe(data_frame_sp_fc)

    mxd_sp = arcpy.mapping.MapDocument(PSR_config.mxd_survey_pipeline)
    # select survey and pipleline feature(s) layers in mxd file
    pipeline_lyr =  arcpy.mapping.ListLayers(mxd_sp)[0]
    survey_lyr = arcpy.mapping.ListLayers(mxd_sp)[1]
    
    arcpy.SelectLayerByLocation_management(pipeline_lyr, 'intersect',data_frame_sp_fc)
    count_pipeline = int((arcpy.GetCount_management(pipeline_lyr).getOutput(0)))
   
    arcpy.SelectLayerByLocation_management(survey_lyr, 'intersect',data_frame_sp_fc)
    count_survey = int((arcpy.GetCount_management(survey_lyr).getOutput(0)))

    if count_pipeline > 0 or count_survey > 0:
        survey_cursor = arcpy.SearchCursor(survey_lyr) # use it for inserting in DB 
        pipeline_cursor = arcpy.SearchCursor(pipeline_lyr) # use it for inserting in DB 
        
        df_sp = arcpy.mapping.ListDataFrames(mxd_sp,"*")[0]
        df_sp.spatialRef = out_coordinate_system
        df_sp.extent = data_frame_desc.extent
        addBuffertoMxd("buffer_sp",df_sp)
        addOrdergeomtoMxd("ordergeoNamePR", df_sp)
        for lyr in arcpy.mapping.ListLayers(mxd_sp):
            # clear selections
            if lyr.isFeatureLayer:
                arcpy.SelectLayerByAttribute_management(lyr, "CLEAR_SELECTION")
        
        if not multipage_sp:
            mxd_sp.saveACopy(os.path.join(scratch_folder,"mxd_sp.mxd"))
            mxd_sp_tmp = arcpy.mapping.MapDocument(os.path.join(scratch_folder,"mxd_sp.mxd"))
            arcpy.mapping.ExportToJPEG(mxd_sp_tmp, outputjpg_sp, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
            shutil.copy(outputjpg_sp, os.path.join(report_path, 'PSRmaps', OrderNumText))

        else:
            grid_name = 'grid_lyr_sp'
            grid_lyr_shp = os.path.join(scratch_gdb, grid_name)
            arcpy.GridIndexFeatures_cartography(grid_lyr_shp, buffer_sp_fc, "", "", "", gridsize, gridsize)
            
            mxd_mm_sp = arcpy.mapping.MapDocument(PSR_config.mxd_survey_pipeline_mm)
            df_mm_sp = arcpy.mapping.ListDataFrames(mxd_mm_sp,"*")[0]
            df_mm_sp.spatialReference = out_coordinate_system
            # part 1: the overview map
            # add grid layer
            grid_layer = arcpy.mapping.Layer(gridlyrfile)
            grid_layer.replaceDataSource(scratch_gdb,"FILEGDB_WORKSPACE","grid_lyr_sp")
            arcpy.mapping.AddLayer(df_sp,grid_layer,"Top")
            
            df_sp.extent = grid_layer.getExtent()
            df_sp.scale = df_sp.scale * 1.1
            
            mxd_sp.saveACopy(os.path.join(scratch_folder, "mxd_sp.mxd"))
            mxd_sp_tmp = arcpy.mapping.MapDocument(os.path.join(scratch_folder,"mxd_sp.mxd"))
            arcpy.mapping.ExportToJPEG(mxd_sp_tmp, outputjpg_sp, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            
            if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
                shutil.copy(outputjpg_sp, os.path.join(report_path, 'PSRmaps', OrderNumText))
                del mxd_sp
                del df_sp
            shutil.copy(outputjpg_sp, os.path.join(report_path, 'PSRmaps', OrderNumText))
            
            # part 2: the data driven pages
            page = 1
            page = int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))  + page

            addBuffertoMxd("buffer_sp",df_mm_sp)
            addOrdergeomtoMxd("ordergeoNamePR", df_mm_sp)

            grid_layer_mm = arcpy.mapping.ListLayers(mxd_mm_sp,"Grid" ,df_mm_sp)[0]
            grid_layer_mm.replaceDataSource(scratch_gdb, "FILEGDB_WORKSPACE","grid_lyr_sp")
            arcpy.CalculateAdjacentFields_cartography(grid_lyr_shp, 'PageNumber')
            mxd_mm_sp.saveACopy(os.path.join(scratch_folder, "mxd_mm_sp.mxd"))
            
            for i in range(1,int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))+1):
    	        arcpy.SelectLayerByAttribute_management(grid_layer_mm, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
                df_mm_sp.extent = grid_layer_mm.getSelectedExtent(True)
                df_mm_sp.scale = df_mm_sp.scale * 1.1
                arcpy.SelectLayerByAttribute_management(grid_layer_mm, "CLEAR_SELECTION")

                title_text = arcpy.mapping.ListLayoutElements(mxd_mm_sp, "TEXT_ELEMENT", "title")[0]
                title_text.text = "Survey & Pipeline - Page " + str(i)
                title_text.elementPositionX = 0.468
                arcpy.RefreshTOC()

                arcpy.mapping.ExportToJPEG(mxd_mm_sp, outputjpg_sp[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
                if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                    os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
                shutil.copy(outputjpg_sp[0:-4]+str(i)+".jpg", os.path.join(report_path, 'PSRmaps', OrderNumText))
            del mxd_mm_sp
            del df_mm_sp
    else:
        arcpy.AddMessage('There is no survey and pipeline data for this property')
    ### Save Survey and Pipeline data in the DB
    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()
        ### Insert in order_detail_psr and eris_flex_reporting_psr tables
        
        ### Survey
        for surv_row in survey_cursor:
            cur.callproc('eris_psr.InsertOrderDetail', (OrderIDText, erisid,'17251'))
            erisid += 1
        ### Pipeline
        for pipe_row in pipeline_cursor:
            cur.callproc('eris_psr.InsertOrderDetail', (OrderIDText, erisid,'17250'))
            query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '17250',2,'',1, 'Pipe Line ID',pipe_row.PLINE_ID))
            query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '17250',2,'',2, 'Status',pipe_row.STATUS_CD))
            query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '17250',2,'',3, 'T4 Permit NO',pipe_row.T4PERMIT))
            query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '17250',2,'',4, 'Commodity',pipe_row.COMMODITY1))
            query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '17250',2,'',5, 'Cmdty Desc',pipe_row.CMDTY_DESC))
            query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '17250',2,'',6, 'Operator',pipe_row.OPER_NM))
            query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '17250',2,'',7, 'System Name',pipe_row.SYS_NM))
            query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '17250',2,'',7, 'Diameter (inches)',pipe_row.DIAMETER))
            erisid += 1
         
        ### Insert in eris_maps_psr table
        if os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText, OrderNumText+'_US_SURVEY_PIPELINE.jpg')):
            query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'SURVEY_PIPELINE', OrderNumText+'_US_SURVEY_PIPELINE.jpg', 1))
            if multipage_sp == True:
                for i in range(1,page):
                    query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'SURVEY_PIPELINE', OrderNumText+'_US_SURVEY_PIPELINE'+str(i)+'.jpg', i+1))

        else:
            print( 'No survey and pipleline map is available')
    finally:
        cur.close()
    ### Generate KML if need viewer
    if need_viewer:
        df_sp.spatialRef = srWGS84
        
        #re-focus using Buffer layer for multipage
        if multipage_sp:
            buffer_layer = arcpy.mapping.ListLayers(mxd_sp, "Buffer", df_sp)[0]
            df_sp.extent = buffer_layer.getSelectedExtent(False)
            df_sp.scale = df_sp.scale * 1.1
        df_as_feature = arcpy.Polygon(arcpy.Array([df_sp.extent.lowerLeft, df_sp.extent.lowerRight, df_sp.extent.upperRight, df_sp.extent.upperLeft]),
					df_sp.spatialReference)
        sp_df_extent = os.path.join(scratch_gdb,"sp_df_extent_WGS84")
        arcpy.Project_management(df_as_feature, sp_df_extent, srWGS84)
        ### Select pipeline by dataframe extent to generate kml file
        arcpy.SelectLayerByLocation_management(pipeline_lyr, 'intersect',sp_df_extent)
        if int((arcpy.GetCount_management(pipeline_lyr).getOutput(0))) > 0:
            pipeline_visible_fields = ['PLINE_ID','STATUS_CD','STATUS_CD','T4PERMIT','COMMODITY1','CMDTY_DESC','OPER_NM','SYS_NM','DIAMETER']
            pipeline_field_info = ""
            pipeline_field_list = arcpy.ListFields(pipeline_lyr.dataSource)
            for field in pipeline_field_list:
                if field.name in pipeline_visible_fields:
                        pipeline_field_info = pipeline_field_info + field.name + " " + field.name + " VISIBLE;"
                else:
                    pipeline_field_info = pipeline_field_info + field.name + " " + field.name + " HIDDEN;"
            # save Pipeline layer 
            pipeline_lyr_file = os.path.join(scratch_folder,'pipeline.lyr')
            arcpy.SaveToLayerFile_management(pipeline_lyr_file, pipeline_lyr_file, "ABSOLUTE")
            
            pipeline_tmp_lyr = arcpy.MakeFeatureLayer_management(pipeline_lyr, 'pipeline_tmp_lyr', "", "", pipeline_field_info[:-1])
            arcpy.ApplySymbologyFromLayer_management(pipeline_tmp_lyr, pipeline_lyr_file)
            
            arcpy.LayerToKML_conversion(pipeline_tmp_lyr, os.path.join(viewerdir_kml,"pipeline.kmz"))
        else:
            arcpy.AddMessage('No Pipeline data for creating KML')
            
        ### Select Survey by dataframe extent to generate kml file
        arcpy.SelectLayerByLocation_management(survey_lyr, 'intersect',sp_df_extent)
        if int((arcpy.GetCount_management(survey_lyr).getOutput(0))) > 0:
            survey_visible_fields = ['ANUM2','L1SURNAM']
            survey_field_info = ""
            survey_field_list = arcpy.ListFields(survey_lyr.dataSource)
            for field in survey_field_list:
                if field.name in survey_visible_fields:
                        survey_field_info = survey_field_info + field.name + " " + field.name + " VISIBLE;"
                else:
                    survey_field_info = survey_field_info + field.name + " " + field.name + " HIDDEN;"
            # save Survey layer 
            survey_lyr_file = os.path.join(scratch_folder,'survey.lyr')
            arcpy.SaveToLayerFile_management(survey_lyr, survey_lyr_file, "ABSOLUTE")
            surveye_tmp_lyr = arcpy.MakeFeatureLayer_management(survey_lyr_file, 'surveye_tmp_lyr', "", "", survey_field_info[:-1])
            arcpy.ApplySymbologyFromLayer_management(surveye_tmp_lyr, survey_lyr_file)
            arcpy.LayerToKML_conversion(surveye_tmp_lyr, os.path.join(viewerdir_kml,"survey.kmz"))
        else:
            arcpy.AddMessage('No Pipeline data for creating KML')
    
    del mxd_sp
    del mxd_sp_tmp
    del df_sp

# current Topo map, no attributes ----------------------------------------------------------------------------------
    print "--- starting Topo section " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    bufferSHP_topo = os.path.join(scratch_folder,"buffer_topo.shp")
    arcpy.Buffer_analysis(orderGeometryPR, bufferSHP_topo, bufferDist_topo)

    point = arcpy.Point()
    array = arcpy.Array()
    featureList = []

    width = arcpy.Describe(bufferSHP_topo).extent.width/2
    height = arcpy.Describe(bufferSHP_topo).extent.height/2

    if (width/height > 7/7):    #7/7 now since adjusted the frame to square
        # wider shape
        height = width/7*7
    else:
        # longer shape
        width = height/7*7
    xCentroid = (arcpy.Describe(bufferSHP_topo).extent.XMax + arcpy.Describe(bufferSHP_topo).extent.XMin)/2
    yCentroid = (arcpy.Describe(bufferSHP_topo).extent.YMax + arcpy.Describe(bufferSHP_topo).extent.YMin)/2

    if multipage_topo == True:
        width = width + 6400     #add 2 miles to each side, for multipage
        height = height + 6400   #add 2 miles to each side, for multipage

    point.X = xCentroid-width
    point.Y = yCentroid+height
    array.add(point)
    point.X = xCentroid+width
    point.Y = yCentroid+height
    array.add(point)
    point.X = xCentroid+width
    point.Y = yCentroid-height
    array.add(point)
    point.X = xCentroid-width
    point.Y = yCentroid-height
    array.add(point)
    point.X = xCentroid-width
    point.Y = yCentroid+height
    array.add(point)
    feat = arcpy.Polygon(array,spatialRef)
    array.removeAll()
    featureList.append(feat)
    clipFrame_topo = os.path.join(scratch_folder, "clipFrame_topo.shp")
    arcpy.CopyFeatures_management(featureList, clipFrame_topo)

    masterLayer_topo = arcpy.mapping.Layer(masterlyr_topo)
    arcpy.SelectLayerByLocation_management(masterLayer_topo,'intersect',clipFrame_topo)

    if(int((arcpy.GetCount_management(masterLayer_topo).getOutput(0))) ==0):

        print "NO records selected"
        masterLayer_topo = None

    else:
        cellids_selected = []
        cellsizes = []
        # loop through the relevant records, locate the selected cell IDs
        rows = arcpy.SearchCursor(masterLayer_topo)    # loop through the selected records
        for row in rows:
            cellid = str(int(row.getValue("CELL_ID")))
            cellids_selected.append(cellid)
        del row
        del rows
        masterLayer_topo = None

        infomatrix = []
        with open(csvfile_topo, "rb") as f:
            reader = csv.reader(f)
            for row in reader:
                if row[9] in cellids_selected:
                    pdfname = row[15].strip()

                    #for current topos, read the year from the geopdf file name
                    templist = pdfname.split("_")
                    year2use = templist[len(templist)-3][0:4]

                    if year2use[0:2] != "20":
                        print "################### Error in the year of the map!!"

                    print row[9] + " " + row[5] + "  " + row[15] + "  " + year2use
                    infomatrix.append([row[9],row[5],row[15],year2use])

        mxd_topo = arcpy.mapping.MapDocument(mxdfile_topo) if ProvStateText !='WA' else arcpy.mapping.MapDocument(mxdfile_topo_Tacoma)#mxdfile_topo_Tacoma
        df_topo = arcpy.mapping.ListDataFrames(mxd_topo,"*")[0]
        df_topo.spatialReference = spatialRef

        if multipage_topo == True:
            mxdMM_topo = arcpy.mapping.MapDocument(mxdMMfile_topo) if ProvStateText !='WA' else arcpy.mapping.MapDocument(mxdMMfile_topo_Tacoma)
            dfMM_topo = arcpy.mapping.ListDataFrames(mxdMM_topo,"*")[0]
            dfMM_topo.spatialReference = spatialRef

        topofile = topowhitelyrfile
        quadrangles =""
        for item in infomatrix:
            pdfname = item[2]
            tifname = pdfname[0:-4]   # note without .tif part
            tifname_bk = tifname
            year = item[3]
            if os.path.exists(os.path.join(tifdir_topo,tifname+ "_t.tif")):
                if '.' in tifname:
                    tifname = tifname.replace('.','')

                # need to make a local copy of the tif file for fast data source replacement
                namecomps = tifname.split('_')
                namecomps.insert(-2,year)
                newtifname = '_'.join(namecomps)

                shutil.copyfile(os.path.join(tifdir_topo,tifname_bk+"_t.tif"),os.path.join(scratch_folder,newtifname+'.tif'))

                topoLayer = arcpy.mapping.Layer(topofile)
                topoLayer.replaceDataSource(scratch_folder, "RASTER_WORKSPACE", newtifname)
                topoLayer.name = newtifname
                arcpy.mapping.AddLayer(df_topo, topoLayer, "BOTTOM")
                if multipage_topo == True:
                    arcpy.mapping.AddLayer(dfMM_topo, topoLayer, "BOTTOM")

                comps = pdfname.split('_')
                quadname = " ".join(comps[1:len(comps)-3])+","+comps[0]

                if quadrangles =="":
                    quadrangles = quadname
                else:
                    quadrangles = quadrangles + "; " + quadname
                topoLayer = None

            else:
                print "tif file doesn't exist " + tifname
                if not os.path.exists(tifdir_topo):
                    print "tif dir doesn't exist " + tifdir_topo
                else:
                    print "tif dir does exist " + tifdir_topo
        if 'topoLayer' in locals():                 # possibly no topo returned. Seen one for EDR Alaska order. = True even for topoLayer = None
            del topoLayer
            addBuffertoMxd("buffer_topo",df_topo)
            addOrdergeomtoMxd("ordergeoNamePR", df_topo)

            yearTextE = arcpy.mapping.ListLayoutElements(mxd_topo, "TEXT_ELEMENT", "year")[0]
            yearTextE.text = "Current USGS Topo (" +year+ ")"
            # yearTextE.text = "Current USGS Topo"
            yearTextE.elementPositionX = 0.4959

            quadrangleTextE = arcpy.mapping.ListLayoutElements(mxd_topo, "TEXT_ELEMENT", "quadrangle")[0]
            quadrangleTextE.text = "Quadrangle(s): " + quadrangles

            sourceTextE = arcpy.mapping.ListLayoutElements(mxd_topo, "TEXT_ELEMENT", "source")[0]
            sourceTextE.text = "Source: USGS 7.5 Minute Topographic Map"

            arcpy.RefreshTOC()

            if multipage_topo == False:
                arcpy.mapping.ExportToJPEG(mxd_topo, outputjpg_topo, "PAGE_LAYOUT")#, resolution=200, jpeg_quality=90)
                if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                    os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
                shutil.copy(outputjpg_topo, os.path.join(report_path, 'PSRmaps', OrderNumText))

                mxd_topo.saveACopy(os.path.join(scratch_folder,"mxd_topo.mxd"))
                del mxd_topo
                del df_topo

            else:     #multipage
                gridlr = "gridlr_topo"   #gdb feature class doesn't work, could be a bug. So use .shp
                gridlrshp = os.path.join(scratch_gdb, gridlr)
                arcpy.GridIndexFeatures_cartography(gridlrshp, bufferSHP_topo, "", "", "", gridsize, gridsize)  #note the tool takes featureclass name only, not the full path

                # part 1: the overview map
                # add grid layer
                gridLayer = arcpy.mapping.Layer(gridlyrfile)
                gridLayer.replaceDataSource(scratch_gdb,"FILEGDB_WORKSPACE","gridlr_topo")
                arcpy.mapping.AddLayer(df_topo,gridLayer,"Top")

                df_topo.extent = gridLayer.getExtent()
                df_topo.scale = df_topo.scale * 1.1

                mxd_topo.saveACopy(os.path.join(scratch_folder, "mxd_topo.mxd"))
                arcpy.mapping.ExportToJPEG(mxd_topo, outputjpg_topo, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
                if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                    os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
                shutil.copy(outputjpg_topo, os.path.join(report_path, 'PSRmaps', OrderNumText))
                del mxd_topo
                del df_topo

                # part 2: the data driven pages
                page = 1
                page = int(arcpy.GetCount_management(gridlrshp).getOutput(0))  + page

                addBuffertoMxd("buffer_topo",dfMM_topo)
                addOrdergeomtoMxd("ordergeoNamePR", dfMM_topo)

                gridlayerMM = arcpy.mapping.ListLayers(mxdMM_topo,"Grid" ,dfMM_topo)[0]
                gridlayerMM.replaceDataSource(scratch_gdb, "FILEGDB_WORKSPACE","gridlr_topo")
                arcpy.CalculateAdjacentFields_cartography(gridlrshp, 'PageNumber')
                mxdMM_topo.saveACopy(os.path.join(scratch_folder, "mxdMM_topo.mxd"))

                for i in range(1,int(arcpy.GetCount_management(gridlrshp).getOutput(0))+1):
                    arcpy.SelectLayerByAttribute_management(gridlayerMM, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
                    dfMM_topo.extent = gridlayerMM.getSelectedExtent(True)
                    dfMM_topo.scale = dfMM_topo.scale * 1.1

                    # might want to select the quad name again
                    quadrangles_mm = ""
                    images = arcpy.mapping.ListLayers(mxdMM_topo, "*TM_geo", dfMM_topo)
                    for image in images:
                        if image.getExtent().overlaps(gridlayerMM.getSelectedExtent(True)) or image.getExtent().contains(gridlayerMM.getSelectedExtent(True)):
                            temp = image.name.split('_20')[0]    #e.g. VA_Port_Royal
                            comps = temp.split('_')
                            quadname = " ".join(comps[1:len(comps)])+","+comps[0]

                            if quadrangles_mm =="":
                                quadrangles_mm = quadname
                            else:
                                quadrangles_mm = quadrangles_mm + "; " + quadname

                    arcpy.SelectLayerByAttribute_management(gridlayerMM, "CLEAR_SELECTION")

                    yearTextE = arcpy.mapping.ListLayoutElements(mxdMM_topo, "TEXT_ELEMENT", "year")[0]
                    yearTextE.text = "Current USGS Topo - Page " + str(i)
                    yearTextE.elementPositionX = 0.4959

                    quadrangleTextE = arcpy.mapping.ListLayoutElements(mxdMM_topo, "TEXT_ELEMENT", "quadrangle")[0]
                    quadrangleTextE.text = "Quadrangle(s): " + quadrangles_mm

                    sourceTextE = arcpy.mapping.ListLayoutElements(mxdMM_topo, "TEXT_ELEMENT", "source")[0]
                    sourceTextE.text = "Source: USGS 7.5 Minute Topographic Map"

                    arcpy.RefreshTOC()

                    arcpy.mapping.ExportToJPEG(mxdMM_topo, outputjpg_topo[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
                    if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                        os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
                    shutil.copy(outputjpg_topo[0:-4]+str(i)+".jpg", os.path.join(report_path, 'PSRmaps', OrderNumText))

                del mxdMM_topo
                del dfMM_topo

# shaded relief map ----------------------------------------------------------------------------------------
    print "--- starting relief " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    mxd_relief = arcpy.mapping.MapDocument(mxdfile_relief)
    df_relief = arcpy.mapping.ListDataFrames(mxd_relief,"*")[0]
    df_relief.spatialReference = spatialRef

    point = arcpy.Point()
    array = arcpy.Array()
    featureList = []

    addBuffertoMxd("buffer_topo",df_relief)
    addOrdergeomtoMxd("ordergeoNamePR", df_relief)
    # locate and add relevant shadedrelief tiles
    width = arcpy.Describe(bufferSHP_topo).extent.width/2
    height = arcpy.Describe(bufferSHP_topo).extent.height/2

    if (width/height > 5/4.4):
        # wider shape
        height = width/5*4.4
    else:
        # longer shape
        width = height/4.4*5

    xCentroid = (arcpy.Describe(bufferSHP_topo).extent.XMax + arcpy.Describe(bufferSHP_topo).extent.XMin)/2
    yCentroid = (arcpy.Describe(bufferSHP_topo).extent.YMax + arcpy.Describe(bufferSHP_topo).extent.YMin)/2

    width = width + 6400     #add 2 miles to each side, for multipage
    height = height + 6400   #add 2 miles to each side, for multipage

    point.X = xCentroid-width
    point.Y = yCentroid+height
    array.add(point)
    point.X = xCentroid+width
    point.Y = yCentroid+height
    array.add(point)
    point.X = xCentroid+width
    point.Y = yCentroid-height
    array.add(point)
    point.X = xCentroid-width
    point.Y = yCentroid-height
    array.add(point)
    point.X = xCentroid-width
    point.Y = yCentroid+height
    array.add(point)
    feat = arcpy.Polygon(array,spatialRef)
    array.removeAll()
    featureList.append(feat)
    clipFrame_relief = os.path.join(scratch_folder, "clipFrame_relief.shp")
    arcpy.CopyFeatures_management(featureList, clipFrame_relief)

    masterLayer_relief = arcpy.mapping.Layer(masterlyr_dem)
    arcpy.SelectLayerByLocation_management(masterLayer_relief,'intersect',clipFrame_relief)
    print "after selectLayerByLocation "+time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())

    cellids_selected = []
    if(int((arcpy.GetCount_management(masterLayer_relief).getOutput(0))) ==0):

        print "NO records selected"
        masterLayer_relief = None

    else:
        cellid = ''
        # loop through the relevant records, locate the selected cell IDs
        rows = arcpy.SearchCursor(masterLayer_relief)    # loop through the selected records
        for row in rows:
            cellid = str(row.getValue("image_name")).strip()
            if cellid !='':
                cellids_selected.append(cellid)
        del row
        del rows
        masterLayer_relief = None
        print "Before adding data sources" + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
        for item in cellids_selected:
            item =item[:-4]
            reliefLayer = arcpy.mapping.Layer(relieflyrfile)
            shutil.copyfile(os.path.join(path_shadedrelief,item+'_hs.img'),os.path.join(scratch_folder,item+'_hs.img'))
            reliefLayer.replaceDataSource(scratch_folder,"RASTER_WORKSPACE",item+'_hs.img')
            reliefLayer.name = item
            arcpy.mapping.AddLayer(df_relief, reliefLayer, "BOTTOM")
            reliefLayer = None

    arcpy.RefreshActiveView()

    if multipage_relief == False:
        mxd_relief.saveACopy(os.path.join(scratch_folder,"mxd_relief.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_relief, outputjpg_relief, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
            os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
        shutil.copy(outputjpg_relief, os.path.join(report_path, 'PSRmaps', OrderNumText))

        del mxd_relief
        del df_relief
    else:     # multipage
        gridlr = "gridlr_relief"   #gdb feature class doesn't work, could be a bug. So use .shp
        gridlrshp = os.path.join(scratch_gdb, gridlr)
        arcpy.GridIndexFeatures_cartography(gridlrshp, bufferSHP_topo, "", "", "", gridsize, gridsize)  #note the tool takes featureclass name only, not the full path

        # part 1: the overview map
        # add grid layer
        gridLayer = arcpy.mapping.Layer(gridlyrfile)
        gridLayer.replaceDataSource(scratch_gdb,"FILEGDB_WORKSPACE","gridlr_relief")
        arcpy.mapping.AddLayer(df_relief,gridLayer,"Top")

        df_relief.extent = gridLayer.getExtent()
        df_relief.scale = df_relief.scale * 1.1

        mxd_relief.saveACopy(os.path.join(scratch_folder, "mxd_relief.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_relief, outputjpg_relief, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
            os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
        shutil.copy(outputjpg_relief, os.path.join(report_path, 'PSRmaps', OrderNumText))
        del mxd_relief
        del df_relief

        # part 2: the data driven pages
        page = 1

        page = int(arcpy.GetCount_management(gridlrshp).getOutput(0))  + page
        mxdMM_relief = arcpy.mapping.MapDocument(mxdMMfile_relief)

        dfMM_relief = arcpy.mapping.ListDataFrames(mxdMM_relief,"*")[0]
        dfMM_relief.spatialReference = spatialRef
        addBuffertoMxd("buffer_topo",dfMM_relief)
        addOrdergeomtoMxd("ordergeoNamePR", dfMM_relief)

        gridlayerMM = arcpy.mapping.ListLayers(mxdMM_relief,"Grid" ,dfMM_relief)[0]
        gridlayerMM.replaceDataSource(scratch_gdb, "FILEGDB_WORKSPACE","gridlr_relief")
        arcpy.CalculateAdjacentFields_cartography(gridlrshp, 'PageNumber')
        mxdMM_relief.saveACopy(os.path.join(scratch_folder, "mxdMM_relief.mxd"))

        for item in cellids_selected:
            item =item[:-4]
            reliefLayer = arcpy.mapping.Layer(relieflyrfile)
            shutil.copyfile(os.path.join(path_shadedrelief,item+'_hs.img'),os.path.join(scratch_folder,item+'_hs.img'))   #make a local copy, will make it run faster
            reliefLayer.replaceDataSource(scratch_folder,"RASTER_WORKSPACE",item+'_hs.img')
            reliefLayer.name = item
            arcpy.mapping.AddLayer(dfMM_relief, reliefLayer, "BOTTOM")

        for i in range(1,int(arcpy.GetCount_management(gridlrshp).getOutput(0))+1):
            arcpy.SelectLayerByAttribute_management(gridlayerMM, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
            dfMM_relief.extent = gridlayerMM.getSelectedExtent(True)
            dfMM_relief.scale = dfMM_relief.scale * 1.1
            arcpy.SelectLayerByAttribute_management(gridlayerMM, "CLEAR_SELECTION")

            arcpy.mapping.ExportToJPEG(mxdMM_relief, outputjpg_relief[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
            shutil.copy(outputjpg_relief[0:-4]+str(i)+".jpg", os.path.join(report_path, 'PSRmaps', OrderNumText))
        del mxdMM_relief
        del dfMM_relief

    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()
        if os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText, OrderNumText+'_US_TOPO.jpg')):
            query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'TOPO', OrderNumText+'_US_TOPO.jpg', 1))
            if multipage_topo == True:
                for i in range(1,page):
                    query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'TOPO', OrderNumText+'_US_TOPO'+str(i)+'.jpg', i+1))

        else:
            print "No Topo map is available"

        if os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText, OrderNumText+'_US_RELIEF.jpg')):
            query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'RELIEF', OrderNumText+'_US_RELIEF.jpg', 1))
            if multipage_relief == True:
                for i in range(1,page):
                    query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'RELIEF', OrderNumText+'_US_RELIEF'+str(i)+'.jpg', i+1))

        else:
            print "No Relief map is available"

    finally:
        cur.close()
        con.close()

# Wetland Map only, no attributes ---------------------------------------------------------------------------------
    print "--- starting Wetland section " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    bufferSHP_wetland = os.path.join(scratch_folder,"buffer_wetland.shp")
    arcpy.Buffer_analysis(orderGeometryPR, bufferSHP_wetland, bufferDist_wetland)

    mxd_wetland = arcpy.mapping.MapDocument(mxdfile_wetland)
    df_wetland = arcpy.mapping.ListDataFrames(mxd_wetland,"big")[0]
    df_wetland.spatialReference = spatialRef
    df_wetlandsmall = arcpy.mapping.ListDataFrames(mxd_wetland,"small")[0]
    df_wetlandsmall.spatialReference = spatialRef
    del df_wetlandsmall

    addBuffertoMxd("buffer_wetland",df_wetland)
    addOrdergeomtoMxd("ordergeoNamePR", df_wetland)

    # print the maps
    if multipage_wetland == False:
        mxd_wetland.saveACopy(os.path.join(scratch_folder, "mxd_wetland.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_wetland, outputjpg_wetland, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
            os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
        shutil.copy(outputjpg_wetland, os.path.join(report_path, 'PSRmaps', OrderNumText))
        del mxd_wetland
        del df_wetland

    else:    # multipage
        gridlr = "gridlr_wetland"   #gdb feature class doesn't work, could be a bug. So use .shp
        gridlrshp = os.path.join(scratch_gdb, gridlr)
        arcpy.GridIndexFeatures_cartography(gridlrshp, bufferSHP_wetland, "", "", "", gridsize, gridsize)  #note the tool takes featureclass name only, not the full path

        # part 1: the overview map
        # add grid layer
        gridLayer = arcpy.mapping.Layer(gridlyrfile)
        gridLayer.replaceDataSource(scratch_gdb,"FILEGDB_WORKSPACE","gridlr_wetland")
        arcpy.mapping.AddLayer(df_wetland,gridLayer,"Top")

        df_wetland.extent = gridLayer.getExtent()
        df_wetland.scale = df_wetland.scale * 1.1

        mxd_wetland.saveACopy(os.path.join(scratch_folder, "mxd_wetland.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_wetland, outputjpg_wetland, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
            os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
        shutil.copy(outputjpg_wetland, os.path.join(report_path, 'PSRmaps', OrderNumText))

        del mxd_wetland
        del df_wetland

        # part 2: the data driven pages
        page = 1

        page = int(arcpy.GetCount_management(gridlrshp).getOutput(0))  + page
        mxdMM_wetland = arcpy.mapping.MapDocument(mxdMMfile_wetland)

        dfMM_wetland = arcpy.mapping.ListDataFrames(mxdMM_wetland,"big")[0]
        dfMM_wetland.spatialReference = spatialRef
        addBuffertoMxd("buffer_wetland",dfMM_wetland)
        addOrdergeomtoMxd("ordergeoNamePR", dfMM_wetland)
        gridlayerMM = arcpy.mapping.ListLayers(mxdMM_wetland,"Grid" ,dfMM_wetland)[0]
        gridlayerMM.replaceDataSource(scratch_gdb, "FILEGDB_WORKSPACE","gridlr_wetland")
        arcpy.CalculateAdjacentFields_cartography(gridlrshp, 'PageNumber')
        mxdMM_wetland.saveACopy(os.path.join(scratch_folder, "mxdMM_wetland.mxd"))

        for i in range(1,int(arcpy.GetCount_management(gridlrshp).getOutput(0))+1):
            arcpy.SelectLayerByAttribute_management(gridlayerMM, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
            dfMM_wetland.extent = gridlayerMM.getSelectedExtent(True)
            dfMM_wetland.scale = dfMM_wetland.scale * 1.1
            arcpy.SelectLayerByAttribute_management(gridlayerMM, "CLEAR_SELECTION")

            titleTextE = arcpy.mapping.ListLayoutElements(mxdMM_wetland, "TEXT_ELEMENT", "title")[0]
            titleTextE.text = "Wetland Type - Page " + str(i)
            titleTextE.elementPositionX = 0.468
            arcpy.RefreshTOC()

            arcpy.mapping.ExportToJPEG(mxdMM_wetland, outputjpg_wetland[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
            shutil.copy(outputjpg_wetland[0:-4]+str(i)+".jpg", os.path.join(report_path, 'PSRmaps', OrderNumText))
        del mxdMM_wetland
        del dfMM_wetland

    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        ###cur.callproc('eris_psr.ClearOrder', (OrderIDText,))
        query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'WETLAND', OrderNumText+'_US_WETL.jpg', 1))
        if multipage_wetland == True:
            for i in range(1,page):
                query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'WETLAND', OrderNumText+'_US_WETL'+str(i)+'.jpg', i+1))

    finally:
        cur.close()
        con.close()

# NY Wetland Map only, no attributes ---------------------------------------------------------------------------------
    if ProvStateText =='NY':

        print "--- starting NY Wetland section " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
        bufferSHP_wetland = os.path.join(scratch_folder,"buffer_wetland.shp")
        # arcpy.Buffer_analysis(orderGeometryPR, bufferSHP_wetland, bufferDist_wetland)

        mxd_wetlandNY = arcpy.mapping.MapDocument(mxdfile_wetlandNY)
        df_wetlandNY = arcpy.mapping.ListDataFrames(mxd_wetlandNY,"big")[0]
        df_wetlandNY.spatialReference = spatialRef

        addBuffertoMxd("buffer_wetland",df_wetlandNY)
        addOrdergeomtoMxd("ordergeoNamePR", df_wetlandNY)

        # print the maps
        if multipage_wetland == False:
            mxd_wetlandNY.saveACopy(os.path.join(scratch_folder, "mxd_wetlandNY.mxd"))
            arcpy.mapping.ExportToJPEG(mxd_wetlandNY, outputjpg_wetlandNY, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            shutil.copy(outputjpg_wetlandNY, os.path.join(report_path, 'PSRmaps', OrderNumText))
            del mxd_wetlandNY
            del df_wetlandNY

        else:    # multipage
            gridlr = "gridlr_wetland"   #gdb feature class doesn't work, could be a bug. So use .shp
            gridlrshp = os.path.join(scratch_gdb, gridlr)
            #arcpy.GridIndexFeatures_cartography(gridlrshp, bufferSHP_wetland, "", "", "", gridsize, gridsize)  #note the tool takes featureclass name only, not the full path

            # part 1: the overview map
            #add grid layer
            gridLayer = arcpy.mapping.Layer(gridlyrfile)
            gridLayer.replaceDataSource(scratch_gdb,"FILEGDB_WORKSPACE","gridlr_wetland")
            arcpy.mapping.AddLayer(df_wetlandNY,gridLayer,"Top")

            df_wetlandNY.extent = gridLayer.getExtent()
            df_wetlandNY.scale = df_wetlandNY.scale * 1.1

            mxd_wetlandNY.saveACopy(os.path.join(scratch_folder, "mxd_wetlandNY.mxd"))
            arcpy.mapping.ExportToJPEG(mxd_wetlandNY, outputjpg_wetlandNY, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
            shutil.copy(outputjpg_wetlandNY, os.path.join(report_path, 'PSRmaps', OrderNumText))

            del mxd_wetlandNY
            del df_wetlandNY

            # part 2: the data driven pages
            page = 1

            page = int(arcpy.GetCount_management(gridlrshp).getOutput(0))  + page
            mxdMM_wetlandNY = arcpy.mapping.MapDocument(mxdMMfile_wetlandNY)

            dfMM_wetlandNY = arcpy.mapping.ListDataFrames(mxdMM_wetlandNY,"big")[0]
            dfMM_wetlandNY.spatialReference = spatialRef
            addBuffertoMxd("buffer_wetland",dfMM_wetlandNY)
            addOrdergeomtoMxd("ordergeoNamePR", dfMM_wetlandNY)
            gridlayerMM = arcpy.mapping.ListLayers(mxdMM_wetlandNY,"Grid" ,dfMM_wetlandNY)[0]
            gridlayerMM.replaceDataSource(scratch_gdb, "FILEGDB_WORKSPACE","gridlr_wetland")
            arcpy.CalculateAdjacentFields_cartography(gridlrshp, 'PageNumber')
            mxdMM_wetlandNY.saveACopy(os.path.join(scratch_folder, "mxdMM_wetlandNY.mxd"))

            for i in range(1,int(arcpy.GetCount_management(gridlrshp).getOutput(0))+1):
                arcpy.SelectLayerByAttribute_management(gridlayerMM, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
                dfMM_wetlandNY.extent = gridlayerMM.getSelectedExtent(True)
                dfMM_wetlandNY.scale = dfMM_wetlandNY.scale * 1.1
                arcpy.SelectLayerByAttribute_management(gridlayerMM, "CLEAR_SELECTION")

                titleTextE = arcpy.mapping.ListLayoutElements(mxdMM_wetlandNY, "TEXT_ELEMENT", "title")[0]
                titleTextE.text = "NY Wetland Type - Page " + str(i)
                titleTextE.elementPositionX = 0.468
                arcpy.RefreshTOC()

                arcpy.mapping.ExportToJPEG(mxdMM_wetlandNY, outputjpg_wetlandNY[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
                if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                    os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
                shutil.copy(outputjpg_wetlandNY[0:-4]+str(i)+".jpg", os.path.join(report_path, 'PSRmaps', OrderNumText))
            del mxdMM_wetlandNY
            del dfMM_wetlandNY

        try:
            con = cx_Oracle.connect(connectionString)
            cur = con.cursor()

            ###cur.callproc('eris_psr.ClearOrder', (OrderIDText,))
            query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'WETLAND', OrderNumText+'_NY_WETL.jpg', 1))
            if multipage_wetland == True:
                for i in range(1,page):
                    query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'WETLAND', OrderNumText+'_NY_WETL'+str(i)+'.jpg', i+1))

        finally:
            cur.close()
            con.close()

# Floodplain -----------------------------------------------------------------------------------------------------
    print "--- starting floodplain " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())

    bufferSHP_flood = os.path.join(scratch_folder,"buffer_flood.shp")
    arcpy.Buffer_analysis(orderGeometryPR, bufferSHP_flood, bufferDist_flood)

    # clip a layer for attribute retrieval, and zoom to the right area on geology mxd.
    flood_clip =os.path.join(scratch_gdb,'flood')   #better keep in file geodatabase due to content length in certain columns
    arcpy.Clip_analysis(data_flood, bufferSHP_flood, flood_clip)
    del data_flood

    floodpanel_clip =os.path.join(scratch_gdb,'floodpanel')   #better keep in file geodatabase due to content length in certain columns
    arcpy.Clip_analysis(data_floodpanel, bufferSHP_flood, floodpanel_clip)
    del data_floodpanel

    arcpy.Statistics_analysis(flood_clip, os.path.join(scratch_folder,"summary_flood.dbf"), [['FLD_ZONE','FIRST'], ['ZONE_SUBTY','FIRST']],'ERIS_CLASS')
    arcpy.Sort_management(os.path.join(scratch_folder,"summary_flood.dbf"), os.path.join(scratch_folder,"summary1_flood.dbf"), [["ERIS_CLASS", "ASCENDING"]])

    print "right before reading mxdfile_flood " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    mxd_flood = arcpy.mapping.MapDocument(mxdfile_flood)

    print "right before reading df_flood " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    df_flood = arcpy.mapping.ListDataFrames(mxd_flood,"Flood*")[0]
    df_flood.spatialReference = spatialRef

    print "right before reading df_floodsmall " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    df_floodsmall = arcpy.mapping.ListDataFrames(mxd_flood,"Study*")[0]
    df_floodsmall.spatialReference = spatialRef
    del df_floodsmall

    addBuffertoMxd("buffer_flood",df_flood)
    addOrdergeomtoMxd("ordergeoNamePR", df_flood)

    arcpy.RefreshActiveView();

    if multipage_flood == False:
        mxd_flood.saveACopy(os.path.join(scratch_folder, "mxd_flood.mxd"))       #<-- this line seems to take huge amount of memory, up to 1G. possibly due to df SR change
        arcpy.mapping.ExportToJPEG(mxd_flood, outputjpg_flood, "PAGE_LAYOUT", resolution=150, jpeg_quality=85)
        if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
            os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
        shutil.copy(outputjpg_flood, os.path.join(report_path, 'PSRmaps', OrderNumText))
        del mxd_flood
        del df_flood

    else:    # multipage
        gridlr = "gridlr_flood"   #gdb feature class doesn't work, could be a bug. So use .shp
        gridlrshp = os.path.join(scratch_gdb, gridlr)
        arcpy.GridIndexFeatures_cartography(gridlrshp, bufferSHP_flood, "", "", "", gridsize, gridsize)  #note the tool takes featureclass name only, not the full path

        # part 1: the overview map
        # add grid layer
        gridLayer = arcpy.mapping.Layer(gridlyrfile)
        gridLayer.replaceDataSource(scratch_gdb,"FILEGDB_WORKSPACE","gridlr_flood")
        arcpy.mapping.AddLayer(df_flood,gridLayer,"Top")

        df_flood.extent = gridLayer.getExtent()
        df_flood.scale = df_flood.scale * 1.1

        mxd_flood.saveACopy(os.path.join(scratch_folder, "mxd_flood.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_flood, outputjpg_flood, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
            os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
        shutil.copy(outputjpg_flood, os.path.join(report_path, 'PSRmaps', OrderNumText))

        del mxd_flood
        del df_flood

        # part 2: the data driven pages
        page = 1

        page = int(arcpy.GetCount_management(gridlrshp).getOutput(0))  + page
        mxdMM_flood = arcpy.mapping.MapDocument(mxdMMfile_flood)

        dfMM_flood = arcpy.mapping.ListDataFrames(mxdMM_flood,"Flood*")[0]
        dfMM_flood.spatialReference = spatialRef
        addBuffertoMxd("buffer_flood",dfMM_flood)
        addOrdergeomtoMxd("ordergeoNamePR", dfMM_flood)
        gridlayerMM = arcpy.mapping.ListLayers(mxdMM_flood,"Grid" ,dfMM_flood)[0]
        gridlayerMM.replaceDataSource(scratch_gdb, "FILEGDB_WORKSPACE","gridlr_flood")
        arcpy.CalculateAdjacentFields_cartography(gridlrshp, 'PageNumber')
        mxdMM_flood.saveACopy(os.path.join(scratch_folder, "mxdMM_flood.mxd"))

        for i in range(1,int(arcpy.GetCount_management(gridlrshp).getOutput(0))+1):
            arcpy.SelectLayerByAttribute_management(gridlayerMM, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
            dfMM_flood.extent = gridlayerMM.getSelectedExtent(True)
            dfMM_flood.scale = dfMM_flood.scale * 1.1
            arcpy.SelectLayerByAttribute_management(gridlayerMM, "CLEAR_SELECTION")

            titleTextE = arcpy.mapping.ListLayoutElements(mxdMM_flood, "TEXT_ELEMENT", "title")[0]
            titleTextE.text = "Flood Hazard Zones - Page " + str(i)
            titleTextE.elementPositionX = 0.5946
            arcpy.RefreshTOC()

            arcpy.mapping.ExportToJPEG(mxdMM_flood, outputjpg_flood[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
            shutil.copy(outputjpg_flood[0:-4]+str(i)+".jpg", os.path.join(report_path, 'PSRmaps', OrderNumText))
        del mxdMM_flood
        del dfMM_flood

    flood_IDs=[]
    availPanels = ''
    if (int(arcpy.GetCount_management(os.path.join(scratch_folder,"summary1_flood.dbf")).getOutput(0))== 0):
        # no floodplain records selected....
        print 'No floodplain records are selected....'
        if (int(arcpy.GetCount_management(floodpanel_clip).getOutput(0))== 0):
            # no panel available, means no data
            print 'no panels available in the area'

        else:
            # panel available, just not records in area
            in_rows = arcpy.SearchCursor(floodpanel_clip)
            for in_row in in_rows:
                print ": " + in_row.FIRM_PAN    #panel number
                print in_row.EFF_DATE      #effective date

                availPanels = availPanels + in_row.FIRM_PAN+'(effective:' + str(in_row.EFF_DATE)[0:10]+') '
            del in_row
            del in_rows

        try:
            con = cx_Oracle.connect(connectionString)
            cur = con.cursor()

            if len(availPanels) > 0:
                erisid = erisid+1
                print 'erisid for availPanels is ' + str(erisid)
                cur.callproc('eris_psr.InsertOrderDetail', (OrderIDText, erisid,'10683'))
                query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '10683', 2, 'N', 1, 'Available FIRM Panels in area: ', availPanels))
            query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'FLOOD', OrderNumText+'_US_FLOOD.jpg', 1))

        finally:
            cur.close()
            con.close()
    else:
        in_rows = arcpy.SearchCursor(floodpanel_clip)
        for in_row in in_rows:
            print ": " + in_row.FIRM_PAN    #panel number
            print in_row.EFF_DATE      #effective date

            availPanels = availPanels + in_row.FIRM_PAN+'(effective:' + str(in_row.EFF_DATE)[0:10]+') '
        # del in_row
        del in_rows

        try:
            con = cx_Oracle.connect(connectionString)
            cur = con.cursor()
            flood_IDs =[]
            ###cur.callproc('eris_psr.ClearOrder', (OrderIDText,))
            in_rows = arcpy.SearchCursor(os.path.join(scratch_folder,"summary1_flood.dbf"))
            erisid = erisid + 1
            cur.callproc('eris_psr.InsertOrderDetail', (OrderIDText, erisid,'10683'))
            query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '10683', 2, 'N', 1, 'Available FIRM Panels in area: ', availPanels))
            for in_row in in_rows:
                # note the column changed in summary dbf
                print ": " + in_row.ERIS_CLASS    #eris label
                print in_row.FIRST_FLD_      #zone type
                print in_row.FIRST_ZONE   #subtype

                erisid = erisid + 1
                flood_IDs.append([in_row.ERIS_CLASS,erisid])
                ###cur.callproc('eris_psr.ClearOrder', (OrderIDText,))
                cur.callproc('eris_psr.InsertOrderDetail', (OrderIDText, erisid,'10683'))
                query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '10683', 2, 'S1', 1, "Flood Zone " + in_row.ERIS_CLASS, ''))
                query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '10683', 2, 'N', 2, 'Zone: ', in_row.FIRST_FLD_))
                query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '10683', 2, 'N', 3, 'Zone subtype: ', in_row.FIRST_ZONE))

                #query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '10683', 2, 'N', 2, 'Zone tye: ', in_row.FIRST_FLD_))
                #query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '10683', 2, 'N', 3, 'Zone Subtype: ', in_row.FIRST_ZONE))

            del in_row
            del in_rows

            query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'FLOOD', OrderNumText+'_US_FLOOD.jpg', 1))
            if multipage_flood == True:
                for i in range(1,page):
                    query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'FLOOD', OrderNumText+'_US_FLOOD'+str(i)+'.jpg', i+1))

            #result = cur.callfunc('eris_psr.CreateReport', str, (OrderIDText,))

        finally:
            cur.close()
            con.close()

# GEOLOGY REPORT -------------------------------------------------------------------------------------------------
    print "--- starting geology " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    bufferSHP_geol = os.path.join(scratch_folder,"buffer_geol.shp")
    arcpy.Buffer_analysis(orderGeometryPR, bufferSHP_geol, bufferDist_geol)

    # clip a layer for attribute retrieval, and zoom to the right area on geology mxd.
    geol_clip =os.path.join(scratch_gdb,'geology')   #better keep in file geodatabase due to content length in certain columns
    arcpy.Clip_analysis(data_geol, bufferSHP_geol, geol_clip)

    arcpy.Statistics_analysis(geol_clip, os.path.join(scratch_folder,"summary_geol.dbf"), [['UNIT_NAME','FIRST'], ['UNIT_AGE','FIRST'], ['ROCKTYPE1','FIRST'], ['ROCKTYPE2','FIRST'], ['UNITDESC','FIRST'], ['ERIS_KEY_1','FIRST']],'ORIG_LABEL')
    arcpy.Sort_management(os.path.join(scratch_folder,"summary_geol.dbf"), os.path.join(scratch_folder,"summary1_geol.dbf"), [["ORIG_LABEL", "ASCENDING"]])
    # seqarray = arcpy.da.TableToNumPyArray(os.path.join(scratch_folder,'summary1_geol.dbf'), '*')

    mxd_geol = arcpy.mapping.MapDocument(mxdfile_geol)
    df_geol = arcpy.mapping.ListDataFrames(mxd_geol,"*")[0]
    df_geol.spatialReference = spatialRef

    addBuffertoMxd("buffer_geol",df_geol)
    addOrdergeomtoMxd("ordergeoNamePR", df_geol)

    # print the maps
    if multipage_geology == False:
        #df.scale = 5000
        mxd_geol.saveACopy(os.path.join(scratch_folder, "mxd_geol.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_geol, outputjpg_geol, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
            os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
        shutil.copy(outputjpg_geol, os.path.join(report_path, 'PSRmaps', OrderNumText))
        del mxd_geol
        del df_geol

    else:    # multipage
        gridlr = "gridlr_geol"   #gdb feature class doesn't work, could be a bug. So use .shp
        gridlrshp = os.path.join(scratch_gdb, gridlr)
        arcpy.GridIndexFeatures_cartography(gridlrshp, bufferSHP_geol, "", "", "", gridsize, gridsize)  #note the tool takes featureclass name only, not the full path

        # part 1: the overview map
        # add grid layer
        gridLayer = arcpy.mapping.Layer(gridlyrfile)
        gridLayer.replaceDataSource(scratch_gdb,"FILEGDB_WORKSPACE","gridlr_geol")
        arcpy.mapping.AddLayer(df_geol,gridLayer,"Top")

        df_geol.extent = gridLayer.getExtent()
        df_geol.scale = df_geol.scale * 1.1

        mxd_geol.saveACopy(os.path.join(scratch_folder, "mxd_geol.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_geol, outputjpg_geol, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
            os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
        shutil.copy(outputjpg_geol, os.path.join(report_path, 'PSRmaps', OrderNumText))

        del mxd_geol
        del df_geol

        # part 2: the data driven pages
        page = 1

        page = int(arcpy.GetCount_management(gridlrshp).getOutput(0))  + page
        mxdMM_geol = arcpy.mapping.MapDocument(mxdMMfile_geol)

        dfMM_geol = arcpy.mapping.ListDataFrames(mxdMM_geol,"*")[0]
        dfMM_geol.spatialReference = spatialRef
        addBuffertoMxd("buffer_geol",dfMM_geol)
        addOrdergeomtoMxd("ordergeoNamePR", dfMM_geol)

        gridlayerMM = arcpy.mapping.ListLayers(mxdMM_geol,"Grid" ,dfMM_geol)[0]
        gridlayerMM.replaceDataSource(scratch_gdb, "FILEGDB_WORKSPACE","gridlr_geol")
        arcpy.CalculateAdjacentFields_cartography(gridlrshp, 'PageNumber')
        mxdMM_geol.saveACopy(os.path.join(scratch_folder, "mxdMM_geol.mxd"))

        for i in range(1,int(arcpy.GetCount_management(gridlrshp).getOutput(0))+1):
            arcpy.SelectLayerByAttribute_management(gridlayerMM, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
            dfMM_geol.extent = gridlayerMM.getSelectedExtent(True)
            dfMM_geol.scale = dfMM_geol.scale * 1.1
            arcpy.SelectLayerByAttribute_management(gridlayerMM, "CLEAR_SELECTION")

            titleTextE = arcpy.mapping.ListLayoutElements(mxdMM_geol, "TEXT_ELEMENT", "title")[0]
            titleTextE.text = "Geologic Units - Page " + str(i)
            titleTextE.elementPositionX = 0.6303
            arcpy.RefreshTOC()

            arcpy.mapping.ExportToJPEG(mxdMM_geol, outputjpg_geol[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
            shutil.copy(outputjpg_geol[0:-4]+str(i)+".jpg", os.path.join(report_path, 'PSRmaps', OrderNumText))
        del mxdMM_geol
        del dfMM_geol

    if (int(arcpy.GetCount_management(os.path.join(scratch_folder,"summary1_geol.dbf")).getOutput(0))== 0):
        # no geology polygon selected...., need to send in map only
        print 'No geology polygon is selected....'
        try:
            con = cx_Oracle.connect(connectionString)
            cur = con.cursor()

            query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'GEOL', OrderNumText+'_US_GEOL.jpg', 1))          #note type 'SOIL' or 'GEOL' is used internally

        finally:
            cur.close()
            con.close()
    else:
        try:
            con = cx_Oracle.connect(connectionString)
            cur = con.cursor()
            geology_IDs = []

            in_rows = arcpy.SearchCursor(os.path.join(scratch_gdb,"geology"))
            for in_row in in_rows:
                # note the column changed in summary dbf
                print "Unit label is: " + in_row.ORIG_LABEL
                print in_row.UNIT_NAME      # unit name
                print in_row.UNIT_AGE       # unit age
                print in_row.ROCKTYPE1      # rocktype 1
                print in_row.ROCKTYPE2      # rocktype2
                print in_row.UNITDESC       # unit description
                print in_row.ERIS_KEY_1     # eris key created from upper(unit_link)
                erisid = erisid + 1
                geology_IDs.append([in_row.ERIS_KEY_1,erisid])
                cur.callproc('eris_psr.InsertOrderDetail', (OrderIDText, erisid,'10685'))
                query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '10685', 2, 'S1', 1, 'Geologic Unit ' + in_row.ORIG_LABEL, ''))
                query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '10685', 2, 'N', 2, 'Unit Name: ', in_row.UNIT_NAME))
                query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '10685', 2, 'N', 3, 'Unit Age: ', in_row.UNIT_AGE))
                query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '10685', 2, 'N', 4, 'Primary Rock Type: ', in_row.ROCKTYPE1))
                query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '10685', 2, 'N', 5, 'Secondary Rock Type: ', in_row.ROCKTYPE2))
                if in_row.UNITDESC == None:
                    nodescr = 'No description available.'
                    query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '10685', 2, 'N', 6, 'Unit Description: ', nodescr))
                else:
                    query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '10685', 2, 'N', 6, 'Unit Description: ', in_row.UNITDESC.encode('utf-8')))
            del in_row
            del in_rows

            query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'GEOL', OrderNumText+'_US_GEOL.jpg', 1))          #note type 'SOIL' or 'GEOL' is used internally
            if multipage_geology == True:
                for i in range(1,page):
                    query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'GEOL', OrderNumText+'_US_GEOL'+str(i)+'.jpg', i+1))
            # result = cur.callfunc('eris_psr.CreateReport', str, (OrderIDText,))
            # if result == '{"RunReportResult":"OK"}':
            #     print 'report generation success'
            # else:
            #     print 'report generation failure'

        finally:
            cur.close()
            con.close()

# SOIL REPORT ----------------------------------------------------------------------------------------------------------
    print "--- starting Soil section " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    siteState = ProvStateText
    if siteState == 'HI':
        datapath_soil =PSR_config.datapath_soil_HI#r'\\cabcvan1fpr009\DATA_GIS\SSURGO\CONUS_2015\gSSURGO_HI.gdb'
    elif siteState == 'AK':
        datapath_soil =PSR_config.datapath_soil_AK#r'\\cabcvan1fpr009\DATA_GIS\SSURGO\CONUS_2015\gSSURGO_AK.gdb'
    else:
        datapath_soil =PSR_config.datapath_soil_CONUS#r'\\cabcvan1fpr009\DATA_GIS\SSURGO\CONUS_2015\gSSURGO_CONUS_10m.gdb'

    table_muaggatt = os.path.join(datapath_soil,'muaggatt')
    table_component = os.path.join(datapath_soil,'component')
    table_chorizon = os.path.join(datapath_soil,'chorizon')
    table_chtexturegrp = os.path.join(datapath_soil,'chtexturegrp')
    masterfile = os.path.join(datapath_soil,'MUPOLYGON')
    arcpy.MakeFeatureLayer_management(masterfile,'masterLayer')

    fc_soils = os.path.join(scratch_folder,"soils.shp")
    fc_soils_PR = os.path.join(scratch_folder, "soilsPR.shp")
    # fc_soils_m = os.path.join(scratch_folder,"soilsPR_m.shp")
    stable_muaggatt = os.path.join(scratch_gdb,"muaggatt")
    stable_component = os.path.join(scratch_gdb,"component")
    stable_chorizon = os.path.join(scratch_gdb,"chorizon")
    stable_chtexturegrp = os.path.join(scratch_gdb,"chtexturegrp")

    bufferSHP_soil = os.path.join(scratch_folder,"buffer_soil.shp")
    arcpy.Buffer_analysis(orderGeometryPR, bufferSHP_soil, bufferDist_soil)
    arcpy.Clip_analysis(masterfile,bufferSHP_soil,fc_soils)
    arcpy.MakeFeatureLayer_management(fc_soils,'soillayer')

    hydrologic_dict = PSR_config.hydrologic_dict
##    {
##        "A":'Soils in this group have low runoff potential when thoroughly wet. Water is transmitted freely through the soil.',
##        "B":'Soils in this group have moderately low runoff potential when thoroughly wet. Water transmission through the soil is unimpeded.',
##        "C":'Soils in this group have moderately high runoff potential when thoroughly wet. Water transmission through the soil is somewhat restricted.',
##        "D":'Soils in this group have high runoff potential when thoroughly wet. Water movement through the soil is restricted or very restricted.',
##        "A/D":'These soils have low runoff potential when drained and high runoff potential when undrained.',
##        "B/D":'These soils have moderately low runoff potential when drained and high runoff potential when undrained.',
##        "C/D":'These soils have moderately high runoff potential when drained and high runoff potential when undrained.',
##        }

    hydric_dict = PSR_config.hydric_dict
##    {
##        '1':'All hydric',
##        '2':'Not hydric',
##        '3':'Partially hydric',
##        '4':'Unknown',
##        }

    if (int(arcpy.GetCount_management('soillayer').getOutput(0)) == 0):   # no soil polygons selected
        print 'no polygons selected'
        try:
            con = cx_Oracle.connect(connectionString)
            cur = con.cursor()

            erisid = erisid + 1
            query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '9334', 2, 'N', 1, 'No soil data available in the project area.', ''))

        finally:
            cur.close()
            con.close()

    else:
        arcpy.Project_management(fc_soils, fc_soils_PR, out_coordinate_system)

        # create map keys
        # arcpy.SpatialJoin_analysis(fc_soils_PR, orderGeometryPR, fc_soils_m, "JOIN_ONE_TO_MANY", "KEEP_ALL","#", "CLOSEST","5000 Kilometers", "Distance")   # this is the reported distance
        # arcpy.AddField_management(fc_soils_m, "label", "TEXT", "", "", "", "", "NON_NULLABLE", "REQUIRED", "")
        # arcpy.Statistics_analysis(fc_soils_PR, os.path.join(scratch_folder,"summary_soil.dbf"), [['mukey','FIRST']],g_ESRI_variable_14)
        arcpy.Statistics_analysis(fc_soils_PR, os.path.join(scratch_folder,"summary_soil.dbf"), [['mukey','FIRST'],["Shape_Area","SUM"]],'musym')
        arcpy.Sort_management(os.path.join(scratch_folder,"summary_soil.dbf"), os.path.join(scratch_folder,"summary1_soil.dbf"), [["musym", "ASCENDING"]])
        seqarray = arcpy.da.TableToNumPyArray(os.path.join(scratch_folder,'summary1_soil.dbf'), '*')    #note: it could contain 'NOTCOM' record

        # retrieve attributes
        unique_MuKeys = returnUniqueSetString_musym(fc_soils)
        if(len(unique_MuKeys)>0):    # special case: order only returns one "NOTCOM" category, filter out
            whereClause_selectTable = "muaggatt.mukey in " + unique_MuKeys
            arcpy.TableSelect_analysis(table_muaggatt, stable_muaggatt, whereClause_selectTable)

            whereClause_selectTable = "component.mukey in " + unique_MuKeys
            arcpy.TableSelect_analysis(table_component, stable_component, whereClause_selectTable)

            unique_CoKeys = returnUniqueSetString(stable_component, 'cokey')
            whereClause_selectTable = "chorizon.cokey in " + unique_CoKeys
            arcpy.TableSelect_analysis(table_chorizon, stable_chorizon, whereClause_selectTable)

            unique_CHKeys = returnUniqueSetString(stable_chorizon,'chkey')
            if len(unique_CHKeys) > 0:       # special case: e.g. there is only one Urban Land polygon
                whereClause_selectTable = "chorizon.chkey in " + unique_CHKeys
                arcpy.TableSelect_analysis(table_chtexturegrp, stable_chtexturegrp, whereClause_selectTable)

                tablelist = [stable_muaggatt, stable_component,stable_chorizon, stable_chtexturegrp]
                fieldlist  = PSR_config.fc_soils_fieldlist#[['muaggatt.mukey','mukey'], ['muaggatt.musym','musym'], ['muaggatt.muname','muname'],['muaggatt.drclassdcd','drclassdcd'],['muaggatt.hydgrpdcd','hydgrpdcd'],['muaggatt.hydclprs','hydclprs'], ['muaggatt.brockdepmin','brockdepmin'], ['muaggatt.wtdepannmin','wtdepannmin'], ['component.cokey','cokey'],['component.compname','compname'], ['component.comppct_r','comppct_r'], ['component.majcompflag','majcompflag'],['chorizon.chkey','chkey'],['chorizon.hzname','hzname'],['chorizon.hzdept_r','hzdept_r'],['chorizon.hzdepb_r','hzdepb_r'], ['chtexturegrp.chtgkey','chtgkey'], ['chtexturegrp.texdesc1','texdesc'], ['chtexturegrp.rvindicator','rv']]
                keylist = PSR_config.fc_soils_keylist#['muaggatt.mukey', 'component.cokey','chorizon.chkey','chtexturegrp.chtgkey']
                #whereClause_queryTable = "muaggatt.mukey = component.mukey and component.cokey = chorizon.cokey and chorizon.chkey = chtexturegrp.chkey and chtexturegrp.rvindicator = 'Yes'"
                whereClause_queryTable = PSR_config.fc_soils_whereClause_queryTable#"muaggatt.mukey = component.mukey and component.cokey = chorizon.cokey and chorizon.chkey = chtexturegrp.chkey"
                #Query tables may only be created using data from a geodatabase or an OLE DB connection
                queryTableResult = arcpy.MakeQueryTable_management(tablelist,'queryTable','USE_KEY_FIELDS', keylist, fieldlist, whereClause_queryTable)  #note: outTable is a table view and won't persist

                arcpy.TableToTable_conversion('queryTable',scratch_gdb, 'soilTable')  #note: 1. <null> values will be retained using .gdb, will be converted to 0 using .dbf; 2. domain values, if there are any, will be retained by using .gdb

                dataarray = arcpy.da.TableToNumPyArray(os.path.join(scratch_gdb,'soilTable'), '*', null_value = -99)

        reportdata = []
        for i in range (0, len(seqarray)):
            mapunitdata = {}
            mukey = seqarray['FIRST_MUKE'][i]   #note the column name in the .dbf output was cut off
            print '***** map unit ' + str(i)
            print 'musym is ' + str(seqarray['MUSYM'][i])
            print 'mukey is ' + str(mukey)
            mapunitdata['Seq'] = str(i+1)    # note i starts from 0, but we want labels to start from 1

            if (seqarray['MUSYM'][i].upper() == 'NOTCOM'):
                mapunitdata['Map Unit Name'] = 'No Digital Data Available'
                mapunitdata['Mukey'] = mukey
                mapunitdata['Musym'] = 'NOTCOM'
            else:
                if 'dataarray' not in locals():           #there is only one special polygon(urban land or water)
                    cursor = arcpy.SearchCursor(stable_muaggatt, "mukey = '" + str(mukey) + "'")
                    row = cursor.next()
                    mapunitdata['Map Unit Name'] = row.muname
                    print '  map unit name: ' + row.muname
                    mapunitdata['Mukey'] = mukey          #note
                    mapunitdata['Musym'] = row.musym
                    row = None
                    cursor = None

                elif ((returnMapUnitAttribute(dataarray, mukey, 'muname')).upper() == '?'):  #Water or Unrban Land
                    cursor = arcpy.SearchCursor(stable_muaggatt, "mukey = '" + str(mukey) + "'")
                    row = cursor.next()
                    mapunitdata['Map Unit Name'] = row.muname
                    print '  map unit name: ' + row.muname
                    mapunitdata['Mukey'] = mukey          #note
                    mapunitdata['Musym'] = row.musym
                    row = None
                    cursor = None
                else:
                    mapunitdata['Mukey'] = returnMapUnitAttribute(dataarray, mukey, 'mukey')
                    mapunitdata['Musym'] = returnMapUnitAttribute(dataarray, mukey, 'musym')
                    mapunitdata['Map Unit Name'] = returnMapUnitAttribute(dataarray, mukey, 'muname')
                    mapunitdata['Drainage Class - Dominant'] = returnMapUnitAttribute(dataarray, mukey, 'drclassdcd')
                    mapunitdata['Hydrologic Group - Dominant'] = returnMapUnitAttribute(dataarray, mukey, 'hydgrpdcd')
                    mapunitdata['Hydric Classification - Presence'] = returnMapUnitAttribute(dataarray, mukey, 'hydclprs')
                    mapunitdata['Bedrock Depth - Min'] = returnMapUnitAttribute(dataarray, mukey, 'brockdepmin')
                    mapunitdata['Watertable Depth - Annual Min'] = returnMapUnitAttribute(dataarray, mukey, 'wtdepannmin')

                    componentdata = returnComponentAttribute(dataarray,mukey)
                    mapunitdata['component'] = componentdata
            mapunitdata["Soil_Percent"]  ="%s"%round(seqarray['SUM_Shape_'][i]/sum(seqarray['SUM_Shape_'])*100,2)+r'%'
            reportdata.append(mapunitdata)

        for mapunit in reportdata:
            print 'mapunit name: ' + mapunit['Map Unit Name']
            if 'component' in mapunit.keys():
                print 'Major component info are printed below'
                for comp in mapunit['component']:
                    print '    component name is ' + comp[0][0]
                    for i in range(1,len(comp)):
                        print '      '+comp[i][0] +': '+ comp[i][1]

        # create the map
        point = arcpy.Point()
        array = arcpy.Array()
        featureList = []

        width = arcpy.Describe(bufferSHP_soil).extent.width/2
        height = arcpy.Describe(bufferSHP_soil).extent.height/2
        if (width > 662 or height > 662):
            if (width/height > 1):
               # buffer has a wider shape
               width = width * 1.1
               height = width

            else:
                # buffer has a vertically elonged shape
                height = height * 1.1
                width = height
        else:
            width = 662*1.1
            height = 662*1.1
        width = width + 6400     #add 2 miles to each side, for multipage soil
        height = height + 6400   #add 2 miles to each side, for multipage soil
        xCentroid = (arcpy.Describe(bufferSHP_soil).extent.XMax + arcpy.Describe(bufferSHP_soil).extent.XMin)/2
        yCentroid = (arcpy.Describe(bufferSHP_soil).extent.YMax + arcpy.Describe(bufferSHP_soil).extent.YMin)/2
        point.X = xCentroid-width
        point.Y = yCentroid+height
        array.add(point)
        point.X = xCentroid+width
        point.Y = yCentroid+height
        array.add(point)
        point.X = xCentroid+width
        point.Y = yCentroid-height
        array.add(point)
        point.X = xCentroid-width
        point.Y = yCentroid-height
        array.add(point)
        point.X = xCentroid-width
        point.Y = yCentroid+height
        array.add(point)
        feat = arcpy.Polygon(array,spatialRef)
        array.removeAll()
        featureList.append(feat)
        clipFrame = os.path.join(scratch_folder, "clipFrame.shp")
        arcpy.CopyFeatures_management(featureList, clipFrame)

        arcpy.MakeFeatureLayer_management(clipFrame,'ClipLayer')
        arcpy.Clip_analysis('masterLayer','ClipLayer',os.path.join(scratch_folder, "soil_disp.shp"))

        # add another column to soil_disp just for symbology purpose
        arcpy.AddField_management(os.path.join(scratch_folder, 'soil_disp.shp'), "FIDCP", "TEXT", "", "", "", "", "NON_NULLABLE", "REQUIRED", "")

        arcpy.CalculateField_management(os.path.join(scratch_folder, 'soil_disp.shp'), 'FIDCP', '!FID!', "PYTHON_9.3")

        mxd_soil = arcpy.mapping.MapDocument(mxdfile_soil)
        df_soil = arcpy.mapping.ListDataFrames(mxd_soil,"*")[0]
        df_soil.spatialReference = spatialRef

        lyr = arcpy.mapping.ListLayers(mxd_soil, "SSURGO*", df_soil)[0]
        lyr.replaceDataSource(scratch_folder,"SHAPEFILE_WORKSPACE", "soil_disp")
        lyr.symbology.addAllValues()
        soillyr = lyr

        addBuffertoMxd("buffer_soil", df_soil)
        addOrdergeomtoMxd("ordergeoNamePR", df_soil)

        if multipage_soil == False:
            mxd_soil.saveACopy(os.path.join(scratch_folder, "mxd_soil.mxd"))
            arcpy.mapping.ExportToJPEG(mxd_soil, outputjpg_soil, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
            shutil.copy(outputjpg_soil, os.path.join(report_path, 'PSRmaps', OrderNumText))
            del mxd_soil
            del df_soil

        else:   # multipage
            gridlr = "gridlr_soil"   #gdb feature class doesn't work, could be a bug. So use .shp
            gridlrshp = os.path.join(scratch_gdb, gridlr)
            arcpy.GridIndexFeatures_cartography(gridlrshp, bufferSHP_soil, "", "", "", gridsize, gridsize)  #note the tool takes featureclass name only, not the full path

            # part 1: the overview map
            # add grid layer
            gridLayer = arcpy.mapping.Layer(gridlyrfile)
            gridLayer.replaceDataSource(scratch_gdb,"FILEGDB_WORKSPACE","gridlr_soil")
            arcpy.mapping.AddLayer(df_soil,gridLayer,"Top")

            df_soil.extent = gridLayer.getExtent()
            df_soil.scale = df_soil.scale * 1.1

            mxd_soil.saveACopy(os.path.join(scratch_folder, "mxd_soil.mxd"))
            arcpy.mapping.ExportToJPEG(mxd_soil, outputjpg_soil, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
            shutil.copy(outputjpg_soil, os.path.join(report_path, 'PSRmaps', OrderNumText))
            del mxd_soil
            del df_soil

            # part 2: the data driven pages maps
            page = 1

            page = int(arcpy.GetCount_management(gridlrshp).getOutput(0))  + page
            mxdMM_soil = arcpy.mapping.MapDocument(mxdMMfile_soil)

            dfMM_soil = arcpy.mapping.ListDataFrames(mxdMM_soil,"*")[0]
            dfMM_soil.spatialReference = spatialRef
            addBuffertoMxd("buffer_soil",dfMM_soil)
            addOrdergeomtoMxd("ordergeoNamePR", dfMM_soil)
            lyr = arcpy.mapping.ListLayers(mxdMM_soil, "SSURGO*", dfMM_soil)[0]
            lyr.replaceDataSource(scratch_folder,"SHAPEFILE_WORKSPACE", "soil_disp")
            lyr.symbology.addAllValues()
            soillyr = lyr

            gridlayerMM = arcpy.mapping.ListLayers(mxdMM_soil,"Grid" ,dfMM_soil)[0]
            gridlayerMM.replaceDataSource(scratch_gdb, "FILEGDB_WORKSPACE","gridlr_soil")
            arcpy.CalculateAdjacentFields_cartography(gridlrshp, 'PageNumber')
            mxdMM_soil.saveACopy(os.path.join(scratch_folder, "mxdMM_soil.mxd"))

            for i in range(1,int(arcpy.GetCount_management(gridlrshp).getOutput(0))+1):
                arcpy.SelectLayerByAttribute_management(gridlayerMM, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
                dfMM_soil.extent = gridlayerMM.getSelectedExtent(True)
                dfMM_soil.scale = dfMM_soil.scale * 1.1
                arcpy.SelectLayerByAttribute_management(gridlayerMM, "CLEAR_SELECTION")

                titleTextE = arcpy.mapping.ListLayoutElements(mxdMM_soil, "TEXT_ELEMENT", "title")[0]
                titleTextE.text = "SSURGO Soils - Page " + str(i)
                titleTextE.elementPositionX = 0.6156
                arcpy.RefreshTOC()

                arcpy.mapping.ExportToJPEG(mxdMM_soil, outputjpg_soil[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
                if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                    os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
                shutil.copy(outputjpg_soil[0:-4]+str(i)+".jpg", os.path.join(report_path, 'PSRmaps', OrderNumText))
            del mxdMM_soil
            del dfMM_soil

        try:
            con = cx_Oracle.connect(connectionString)
            cur = con.cursor()
            soil_IDs = []
            ###cur.callproc('eris_psr.ClearOrder', (OrderIDText,))

            for mapunit in reportdata:
                erisid = erisid + 1
                mukey = str(mapunit['Mukey'])
                soil_IDs.append([mapunit['Musym'],erisid])
                cur.callproc('eris_psr.InsertOrderDetail', [OrderIDText, erisid,'9334',mukey])
                query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '9334', 2, 'S1', 1, 'Map Unit ' + mapunit['Musym'] + " (%s)"%mapunit["Soil_Percent"], ''))
                query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '9334', 2, 'N', 2, 'Map Unit Name:', mapunit['Map Unit Name']))
                if (len(mapunit) < 6):    #for Water, Urbanland and Gravel Pits
                    query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '9334', 2, 'N', 3, 'No more attributes available for this map unit',''))
                else:           # not do for Water or urban land
                    query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '9334', 2, 'N', 3, 'Bedrock Depth - Min:',  mapunit['Bedrock Depth - Min']))
                    query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '9334', 2, 'N', 4, 'Watertable Depth - Annual Min:', mapunit['Watertable Depth - Annual Min']))
                    if (mapunit['Drainage Class - Dominant'] == '-99'):
                        query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '9334', 2, 'N', 5, 'Drainage Class - Dominant:', 'null'))
                    else:
                        query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '9334', 2, 'N', 5, 'Drainage Class - Dominant:', mapunit['Drainage Class - Dominant']))
                    if (mapunit['Hydrologic Group - Dominant'] == '-99'):
                        query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '9334', 2, 'N', 6, 'Hydrologic Group - Dominant:', 'null'))
                    else:
                        query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '9334', 2, 'N', 6, 'Hydrologic Group - Dominant:', mapunit['Hydrologic Group - Dominant'] + ' - ' +hydrologic_dict[mapunit['Hydrologic Group - Dominant']]))
                    query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '9334', 2, 'N', 7, 'Major components are printed below', ''))

                    k = 7
                    if 'component' in mapunit.keys():
                        k = k + 1
                        for comp in mapunit['component']:
                            query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '9334', 2, 'S2', k, comp[0][0],''))
                            for i in range(1,len(comp)):
                                k = k+1
                                query = cur.callproc('eris_psr.InsertFlexRep', (OrderIDText, erisid, '9334', 2, 'S3', k, comp[i][0], comp[i][1]))

            #old: query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'SOIL'))
            query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'SOIL', OrderNumText+'_US_SOIL.jpg', 1))
            if multipage_soil == True:
                for i in range(1,page):
                    query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'SOIL', OrderNumText+'_US_SOIL'+str(i)+'.jpg', i+1))
            #result = cur.callfunc('eris_psr.CreateReport', str, (OrderIDText,))
            # example: InsertMap(411578, ?SOIL?, ?20131002005_US_SOIL.jpg?, 1)
            #result = cur.callfunc('eris_psr.CreateReport', str, (OrderIDText,))
            #if result == '{"RunReportResult":"OK"}':
            #    print 'report generation success'
            #else:
            #    print 'report generation failure'

        finally:
            cur.close()
            con.close()

# Water Wells and Oil and Gas Wells ------------------------------------------------------------------------------------------
    print "--- starting WaterWells section " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    in_rows = arcpy.SearchCursor(orderGeometryPR)
    orderCentreSHP = os.path.join(scratch_folder, "SiteMarkerPR.shp")
    point1 = arcpy.Point()
    array1 = arcpy.Array()
    featureList = []
    arcpy.CreateFeatureclass_management(scratch_folder, "SiteMarkerPR.shp", "POINT", "", "DISABLED", "DISABLED", spatialRef)

    cursor = arcpy.InsertCursor(orderCentreSHP)
    feat = cursor.newRow()
    for in_row in in_rows:
        # Set X and Y for start and end points
        point1.X = in_row.xCenUTM
        point1.Y = in_row.yCenUTM
        array1.add(point1)

        centerpoint = arcpy.Multipoint(array1)
        array1.removeAll()
        featureList.append(centerpoint)
        feat.shape = point1
        cursor.insertRow(feat)
    del feat
    del cursor
    del in_row
    del in_rows
    del point1
    del array1

    arcpy.AddField_management(orderCentreSHP, "Lon_X", "DOUBLE", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")
    arcpy.AddField_management(orderCentreSHP, "Lat_Y", "DOUBLE", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")

    # prepare for elevation calculation
    arcpy.CalculateField_management(orderCentreSHP, 'Lon_X', Lon_X, "PYTHON_9.3", "")
    arcpy.CalculateField_management(orderCentreSHP, 'Lat_Y', Lat_Y, "PYTHON_9.3", "")
    arcpy.ImportToolbox(PSR_config.tbx)
    orderCentreSHP = getElevation(orderCentreSHP,["Lon_X","Lat_Y","Id"])##orderCentreSHP = arcpy.inhouseElevation_ERIS(orderCentreSHP).getOutput(0)
    Call_Google = ''
    rows = arcpy.SearchCursor(orderCentreSHP)
    for row in rows:
        if row.Elevation == -999:
            Call_Google = 'YES'
            break
        else:
            print row.Elevation
    del row
    del rows
    if Call_Google == 'YES':
        orderCentreSHP = arcpy.googleElevation_ERIS(orderCentreSHP).getOutput(0)
##    orderCentreSHPPR = os.path.join(scratch_folder, "SiteMarkerPR.shp")
##    arcpy.Project_management(orderCentreSHP,orderCentreSHPPR,out_coordinate_system)
##    orderCentreSHP = orderCentreSHPPR
    arcpy.AddXY_management(orderCentreSHP)

    mergelist = []
    for dsoid in dsoid_wells:
        bufferSHP_wells = os.path.join(scratch_folder,"buffer_"+dsoid+".shp")
        arcpy.Buffer_analysis(orderGeometryPR, bufferSHP_wells, str(searchRadius[dsoid])+" MILES")
        wells_clip = os.path.join(scratch_folder,'wellsclip_'+dsoid+'.shp')

        arcpy.Clip_analysis(eris_wells, bufferSHP_wells, wells_clip)
        arcpy.Select_analysis(wells_clip, os.path.join(scratch_folder,'wellsselected_'+dsoid+'.shp'), "DS_OID ="+dsoid)
        mergelist.append(os.path.join(scratch_folder,'wellsselected_'+dsoid+'.shp'))
    wells_merge = os.path.join(scratch_folder, "wells_merge.shp")
    arcpy.Merge_management(mergelist, wells_merge)
    del eris_wells
    print "--- WaterWells section, after merge " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    # Calculate Distance with integration and spatial join- can be easily done with Distance tool along with direction if ArcInfo or Advanced license
    wells_mergePR= os.path.join(scratch_folder,"wells_mergePR.shp")
    arcpy.Project_management(wells_merge, wells_mergePR, out_coordinate_system)
    arcpy.Integrate_management(wells_mergePR, ".5 Meters")

    arcpy.AddField_management(orderGeometryPR, "Elevation", "DOUBLE", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")
    cursor = arcpy.SearchCursor(orderCentreSHP)
    row = cursor.next()
    elev_marker = row.getValue("Elevation")
    del cursor
    del row
    arcpy.CalculateField_management(orderGeometryPR, 'Elevation', eval(str(elev_marker)), "PYTHON_9.3", "")

    # add distance to selected wells
    wells_sj= os.path.join(scratch_folder,"wells_sj.shp")
    wells_sja= os.path.join(scratch_folder,"wells_sja.shp")
    arcpy.SpatialJoin_analysis(wells_mergePR, orderGeometryPR, wells_sj, "JOIN_ONE_TO_MANY", "KEEP_ALL","#", "CLOSEST","5000 Kilometers", "Distance")   # this is the reported distance
    #arcpy.SpatialJoin_analysis(wells_sj, orderCentreSHP, wells_sja, "JOIN_ONE_TO_MANY", "KEEP_ALL","#", "CLOSEST_GEODESIC","5000 Kilometers", "Dist_cent")  # this is used for mapkey calculation
    arcpy.SpatialJoin_analysis(wells_sj, orderGeometryPR, wells_sja, "JOIN_ONE_TO_MANY", "KEEP_ALL","#", "CLOSEST","5000 Kilometers", "Dist_cent")  # this is used for mapkey calculation

    if int(arcpy.GetCount_management(os.path.join(wells_merge)).getOutput(0)) != 0:
        print "--- WaterWells section, exists water wells "

        wells_sja = getElevation(wells_sja,["X","Y","ID"])#wells_sja = arcpy.inhouseElevation_ERIS(wells_sja).getOutput(0)
        mxd_wells = arcpy.mapping.MapDocument(PSR_config.mxdfile_wells)
        elevationArray=[]
        Call_Google = ''
        rows = arcpy.SearchCursor(wells_sja)
        for row in rows:
            # print row.Elevation
            if row.Elevation == -999:
                Call_Google = 'YES'
                break
        del rows

        if Call_Google == 'YES':
            arcpy.ImportToolbox(PSR_config.tbx)
            wells_sja = arcpy.googleElevation_ERIS(wells_sja).getOutput(0)

        wells_sja_final= os.path.join(scratch_folder,"wells_sja_PR.shp")
        arcpy.Project_management(wells_sja,wells_sja_final,out_coordinate_system)
        wells_sja = wells_sja_final
        # Add mapkey with script from ERIS toolbox
        arcpy.AddField_management(wells_sja, "MapKeyNo", "SHORT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
         # Process: Add Field for mapkey rank storage based on location and total number of keys at one location
        arcpy.AddField_management(wells_sja, "MapKeyLoc", "SHORT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
        arcpy.AddField_management(wells_sja, "MapKeyTot", "SHORT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
         # Process: Mapkey- to create mapkeys
        arcpy.ImportToolbox(PSR_config.tbx)
        arcpy.mapKey_ERIS(wells_sja)

        # Add Direction to ERIS sites
        arcpy.AddField_management(wells_sja, "Direction", "TEXT", "", "", "3", "", "NULLABLE", "NON_REQUIRED", "")
        desc = arcpy.Describe(wells_sja)
        shapefieldName = desc.ShapeFieldName
        rows = arcpy.UpdateCursor(wells_sja)
        for row in rows:
            if(row.Distance<0.001):         # give onsite, give "-" in Direction field
                directionText = '-'
            else:
                ref_x = row.xCenUTM         # field is directly accessible
                ref_y = row.yCenUTM
                feat = row.getValue(shapefieldName)
                pnt = feat.getPart()
                directionText = getDirectionText.getDirectionText(ref_x,ref_y,pnt.X,pnt.Y)

            row.Direction = directionText   # field is directly accessible
            rows.updateRow(row)
        del rows

        wells_fin= os.path.join(scratch_folder,"wells_fin.shp")
        arcpy.Select_analysis(wells_sja, wells_fin, '"MapKeyTot" = 1')
        wells_disp= os.path.join(scratch_folder,"wells_disp.shp")
        arcpy.Sort_management(wells_fin, wells_disp, [["MapKeyLoc", "ASCENDING"]])

        arcpy.AddField_management(wells_disp, "Ele_diff", "DOUBLE", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")
        arcpy.CalculateField_management(wells_disp, 'Ele_diff', '!Elevation!-!Elevatio_1!', "PYTHON_9.3", "")
        arcpy.AddField_management(wells_disp, "eleRank", "SHORT", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")
        arcpy.ImportToolbox(PSR_config.tbx)
        arcpy.symbol_ERIS(wells_disp)
         ## create a map with water wells and ogw wells

        df_wells = arcpy.mapping.ListDataFrames(mxd_wells,"*")[0]
        df_wells.spatialReference = spatialRef

        lyr = arcpy.mapping.ListLayers(mxd_wells, "wells", df_wells)[0]
        lyr.replaceDataSource(scratch_folder,"SHAPEFILE_WORKSPACE", "wells_disp")
    else:
        print "--- WaterWells section, no water wells exists "
        mxd_wells = arcpy.mapping.MapDocument(PSR_config.mxdfile_wells)
        df_wells = arcpy.mapping.ListDataFrames(mxd_wells,"*")[0]
        df_wells.spatialReference = spatialRef

    for item in dsoid_wells:
        addBuffertoMxd("buffer_"+item, df_wells)
    df_wells.extent = arcpy.Describe(os.path.join(scratch_folder,"buffer_"+dsoid_wells_maxradius+'.shp')).extent
    df_wells.scale = df_wells.scale * 1.1
    addOrdergeomtoMxd("ordergeoNamePR", df_wells)

    if multipage_wells == False or int(arcpy.GetCount_management(wells_sja).getOutput(0))== 0:
        mxd_wells.saveACopy(os.path.join(scratch_folder, "mxd_wells.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_wells, outputjpg_wells, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
            os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
        shutil.copy(outputjpg_wells, os.path.join(report_path, 'PSRmaps', OrderNumText))
        del mxd_wells
        del df_wells
    else:
        gridlr = "gridlr_wells"   #gdb feature class doesn't work, could be a bug. So use .shp
        gridlrshp = os.path.join(scratch_gdb, gridlr)
        arcpy.GridIndexFeatures_cartography(gridlrshp, os.path.join(scratch_folder,"buffer_"+dsoid_wells_maxradius+'.shp'), "", "", "", gridsize, gridsize)  #note the tool takes featureclass name only, not the full path
        # part 1: the overview map
        #add grid layer
        gridLayer = arcpy.mapping.Layer(gridlyrfile)
        gridLayer.replaceDataSource(scratch_gdb,"FILEGDB_WORKSPACE","gridlr_wells")
        arcpy.mapping.AddLayer(df_wells,gridLayer,"Top")
        # turn the site label off
        well_lyr = arcpy.mapping.ListLayers(mxd_wells, "wells", df_wells)[0]
        well_lyr.showLabels = False
        df_wells.extent = gridLayer.getExtent()
        df_wells.scale = df_wells.scale * 1.1
        mxd_wells.saveACopy(os.path.join(scratch_folder, "mxd_wells.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_wells, outputjpg_wells, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
            os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
        shutil.copy(outputjpg_wells, os.path.join(report_path, 'PSRmaps', OrderNumText))
        del mxd_wells
        del df_wells

        # part 2: the data driven pages
        page = 1
        page = int(arcpy.GetCount_management(gridlrshp).getOutput(0))  + page
        mxdMM_wells = arcpy.mapping.MapDocument(mxdMMfile_wells)
        dfMM_wells = arcpy.mapping.ListDataFrames(mxdMM_wells)[0]
        dfMM_wells.spatialReference = spatialRef
        for item in dsoid_wells:
            addBuffertoMxd("buffer_"+item, dfMM_wells)

        #addBuffertoMxd("buffer_"+dsoid_wells_maxradius,dfMM_wells)
        addOrdergeomtoMxd("ordergeoNamePR", dfMM_wells)
        gridlayerMM = arcpy.mapping.ListLayers(mxdMM_wells,"Grid" ,dfMM_wells)[0]
        gridlayerMM.replaceDataSource(scratch_gdb, "FILEGDB_WORKSPACE","gridlr_wells")
        arcpy.CalculateAdjacentFields_cartography(gridlrshp, 'PageNumber')
        lyr = arcpy.mapping.ListLayers(mxdMM_wells, "wells", dfMM_wells)[0]   #"wells" or "Wells" doesn't seem to matter
        lyr.replaceDataSource(scratch_folder,"SHAPEFILE_WORKSPACE", "wells_disp")

        for i in range(1,int(arcpy.GetCount_management(gridlrshp).getOutput(0))+1):
            arcpy.SelectLayerByAttribute_management(gridlayerMM, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
            dfMM_wells.extent = gridlayerMM.getSelectedExtent(True)
            dfMM_wells.scale = dfMM_wells.scale * 1.1
            arcpy.SelectLayerByAttribute_management(gridlayerMM, "CLEAR_SELECTION")
            titleTextE = arcpy.mapping.ListLayoutElements(mxdMM_wells, "TEXT_ELEMENT", "MainTitleText")[0]
            titleTextE.text = "Wells & Additional Sources - Page " + str(i)
            titleTextE.elementPositionX = 0.6438
            arcpy.RefreshTOC()
            arcpy.mapping.ExportToJPEG(mxdMM_wells, outputjpg_wells[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            if not os.path.exists(os.path.join(report_path, 'PSRmaps', OrderNumText)):
                os.mkdir(os.path.join(report_path, 'PSRmaps', OrderNumText))
            shutil.copy(outputjpg_wells[0:-4]+str(i)+".jpg", os.path.join(report_path, 'PSRmaps', OrderNumText))
        del mxdMM_wells
        del dfMM_wells

    # send wells data to Oracle
    if (int(arcpy.GetCount_management(wells_sja).getOutput(0))== 0):
        # no records selected....
        print 'No well records are selected....'
        try:
            con = cx_Oracle.connect(connectionString)
            cur = con.cursor()
            query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'WELLS', OrderNumText+'_US_WELLS.jpg', 1))          #note type 'SOIL' or 'GEOL' is used internally
        finally:
            cur.close()
            con.close()
    else:
        try:
            con = cx_Oracle.connect(connectionString)
            cur = con.cursor()

            in_rows = arcpy.SearchCursor(wells_sja)
            for in_row in in_rows:
                erisid = str(int(in_row.ID))
                DS_OID=str(int(in_row.DS_OID))
                Distance = str(float(in_row.Distance))
                Direction = str(in_row.Direction)
                Elevation = str(float(in_row.Elevation))
                Elevatio_1 = str(float(in_row.Elevation) - float(in_row.Elevatio_1))
                MapKeyLoc = str(int(in_row.MapKeyLoc))
                MapKeyNo = str(int(in_row.MapKeyNo))

                cur.callproc('eris_psr.InsertOrderDetail', (OrderIDText, erisid,DS_OID,'',Distance,Direction,Elevation,Elevatio_1,MapKeyLoc,MapKeyNo))
            del in_row
            del in_rows

            query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'WELLS', OrderNumText+'_US_WELLS.jpg', 1))          #note type 'SOIL' or 'GEOL' is used internally
            if multipage_wells == True:
                for i in range(1,page):
                    query = cur.callproc('eris_psr.InsertMap', (OrderIDText, 'WELLS', OrderNumText+'_US_WELLS'+str(i)+'.jpg', i+1))
        finally:
            cur.close()
            con.close()

# Radon -----------------------------------------------------------------------------------------------------------
    print "--- starting Radon section " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    bufferSHP_radon = os.path.join(scratch_folder,"buffer_radon.shp")
    arcpy.Buffer_analysis(orderGeometryPR, bufferSHP_radon, bufferDist_radon)

    states_clip =os.path.join(scratch_folder,'states.shp')
    arcpy.Clip_analysis(masterlyr_states, bufferSHP_radon, states_clip)

    counties_clip =os.path.join(scratch_folder,'counties.shp')
    arcpy.Clip_analysis(masterlyr_counties, bufferSHP_radon, counties_clip)

    cities_clip =os.path.join(scratch_folder,'cities.shp')
    arcpy.Clip_analysis(masterlyr_cities, bufferSHP_radon, cities_clip)

    zipcodes_clip =os.path.join(scratch_folder,'zipcodes.shp')
    arcpy.Clip_analysis(masterlyr_zipcodes, bufferSHP_radon, zipcodes_clip)

    statelist = ''
    in_rows = arcpy.SearchCursor(states_clip)
    for in_row in in_rows:
        print in_row.STUSPS
        statelist = statelist+ ',' + in_row.STUSPS
    statelist = statelist.strip(',')        #two letter state
    statelist_str = str(statelist)
    del in_rows
    del in_row

    countylist = ''
    in_rows = arcpy.SearchCursor(counties_clip)
    for in_row in in_rows:
        #print in_row.NAME
        countylist = countylist + ','+in_row.NAME
    countylist = countylist.strip(',')
    countylist_str = str(countylist.replace(u'\xed','i').replace(u'\xe1','a').replace(u'\xf1','n'))
    del in_rows
    if 'in_row' in locals():     #sometimes returns no city
        del in_row

    citylist = ''
    in_rows = arcpy.SearchCursor(cities_clip)
    for in_row in in_rows:
        # print in_row.NAME
        citylist = citylist + ','+in_row.NAME
    citylist = citylist.strip(',')
    del in_rows
    if 'in_row' in locals():     #sometimes returns no city
        del in_row

    if 'NH' in statelist:
        towns_clip =os.path.join(scratch_folder,'towns.shp')
        arcpy.Clip_analysis(masterlyr_NHTowns, bufferSHP_radon, towns_clip)
        in_rows = arcpy.SearchCursor(towns_clip)
        for in_row in in_rows:
            print in_row.NAME
            citylist = citylist + ','+in_row.NAME
        citylist = citylist.strip(',')
        del in_rows
        if 'in_row' in locals():     #sometimes returns no city
            del in_row
    citylist_str = str(citylist.replace(u'\xed','i').replace(u'\xe1','a').replace(u'\xf1','n'))

    ziplist = ''
    in_rows = arcpy.SearchCursor(zipcodes_clip)
    for in_row in in_rows:
        print in_row.ZIP
        ziplist = ziplist + ',' + in_row.ZIP
    ziplist = ziplist.strip(',')
    ziplist_str = str(ziplist)
    del in_rows
    if 'in_row' in locals():
        del in_row

    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()
        ###cur.callproc('eris_psr.ClearOrder', (OrderIDText,))

        cur.callproc('eris_psr.GetRadon', (OrderIDText, statelist_str, ziplist_str, countylist_str, citylist_str))

    finally:
        cur.close()
        con.close()

# aspect calculation ---------------------------------------------------------------------------------------------------
    i=0
    imgs = []
    masterLayer_dem = arcpy.mapping.Layer(masterlyr_dem)
    bufferDistance = '0.25 MILES'
    arcpy.AddField_management(orderCentreSHP, "Aspect",  "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
    arcpy.AddXY_management(orderCentreSHP)
    outBufferSHP = os.path.join(scratch_folder, "siteMarker_Buffer.shp")
    arcpy.Buffer_analysis(orderCentreSHP, outBufferSHP, bufferDistance)
    arcpy.DefineProjection_management(outBufferSHP, out_coordinate_system)
    arcpy.SelectLayerByLocation_management(masterLayer_dem, 'intersect', outBufferSHP)

    if (int((arcpy.GetCount_management(masterLayer_dem).getOutput(0)))== 0):
        print "NO records selected for US"
        columns = arcpy.UpdateCursor(orderCentreSHP)
        for column in columns:
            column.Aspect = "Not Available"
        del column
        del columns
        masterLayer_buffer = None

    else:
        # loop through the relevant records, locate the selected cell IDs
        columns = arcpy.SearchCursor(masterLayer_dem)
        for column in columns:
            img = column.getValue("image_name")
            if img ==" ":
                print "no image found"
            else:
                imgs.append(os.path.join(imgdir_dem,img))
                i = i+1
                print "found img " + img
        del column
        del columns
    if i==0:
##        imgdir_demCA = r"\\cabcvan1fpr009\US_DEM\DEM1"
##        masterlyr_demCA = r"\\cabcvan1fpr009\US_DEM\Canada_DEM_edited.shp"
        masterLayer_dem = arcpy.mapping.Layer(masterlyr_demCA)
        arcpy.SelectLayerByLocation_management(masterLayer_dem, 'intersect', outBufferSHP)
        if int((arcpy.GetCount_management(masterLayer_dem).getOutput(0))) != 0:

            columns = arcpy.SearchCursor(masterLayer_dem)
            for column in columns:
                img = column.getValue("image_name")
                if img.strip() !="":
                    imgs.append(os.path.join(imgdir_demCA,img))
                    print "found img " + img
                    i = i+1
            del column
            del columns

    if i >=1:
            if i>1:
                clipped_img=''
                n = 1
                for im in imgs:
                    clip_name ="clip_img_"+str(n)+".img"
                    arcpy.Clip_management(im, "#",os.path.join(scratch_folder, clip_name),outBufferSHP,"#","NONE", "MAINTAIN_EXTENT")
                    clipped_img = clipped_img + os.path.join(scratch_folder, clip_name)+ ";"
                    n =n +1

                img = "img.img"
                arcpy.MosaicToNewRaster_management(clipped_img[0:-1],scratch_folder, img,out_coordinate_system, "32_BIT_FLOAT", "#","1", "FIRST", "#")
            elif i ==1:
                im = imgs[0]
                img = "img.img"
                arcpy.Clip_management(im, "#",os.path.join(scratch_folder,img),outBufferSHP,"#","NONE", "MAINTAIN_EXTENT")
            arr =  arcpy.RasterToNumPyArray(os.path.join(scratch_folder,img))

            x,y = gradient(arr)
            slope = 57.29578*arctan(sqrt(x*x + y*y))
            aspect = 57.29578*arctan2(-x,y)

            for i in range(len(aspect)):
                    for j in range(len(aspect[i])):
                        if -180 <=aspect[i][j] <= -90:
                            aspect[i][j] = -90-aspect[i][j]
                        else :
                            aspect[i][j] = 270 - aspect[i][j]
                        if slope[i][j] ==0:
                            aspect[i][j] = -1

            # gather some information on the original file
            spatialref = arcpy.Describe(os.path.join(scratch_folder,img)).spatialReference
            cellsize1  = arcpy.Describe(os.path.join(scratch_folder,img)).meanCellHeight
            cellsize2  = arcpy.Describe(os.path.join(scratch_folder,img)).meanCellWidth
            extent     = arcpy.Describe(os.path.join(scratch_folder,img)).Extent
            pnt        = arcpy.Point(extent.XMin,extent.YMin)

            # save the raster
            aspect_tif = os.path.join(scratch_folder,"aspect.tif")
            aspect_ras = arcpy.NumPyArrayToRaster(aspect,pnt,cellsize1,cellsize2)
            arcpy.CopyRaster_management(aspect_ras,aspect_tif)
            arcpy.DefineProjection_management(aspect_tif, spatialref)

            slope_tif = os.path.join(scratch_folder,"slope.tif")
            slope_ras = arcpy.NumPyArrayToRaster(slope,pnt,cellsize1,cellsize2)
            arcpy.CopyRaster_management(slope_ras,slope_tif)
            arcpy.DefineProjection_management(slope_tif, spatialref)

            aspect_tif_prj = os.path.join(scratch_folder,"aspect_prj.tif")
            arcpy.ProjectRaster_management(aspect_tif,aspect_tif_prj, out_coordinate_system)

            rows = arcpy.da.UpdateCursor(orderCentreSHP,["POINT_X","POINT_Y","Aspect"])
            for row in rows:
                pointX = row[0]
                pointY = row[1]
                location = str(pointX)+" "+str(pointY)
                asp = arcpy.GetCellValue_management(aspect_tif_prj,location)

                if asp.getOutput((0)) != "NoData":
                    asp_text = getDirectionText.dgrDir2txt(float(asp.getOutput((0))))
                    if float(asp.getOutput((0))) == -1:
                        asp_text = r'N/A'
                    row[2] = asp_text
                    print "assign "+asp_text
                    rows.updateRow(row)
                else:
                    print "fail to use point XY to retrieve"
                    row[2] =-9999
                    print "assign -9999"
                    rows.updateRow(row)
                    raise ValueError('No aspect retrieved CHECK data spatial reference')
            del row
            del rows

    in_rows = arcpy.SearchCursor(orderCentreSHP)
    for in_row in in_rows:
        # there is only one line
        site_elev =  in_row.Elevation
        UTM_X = in_row.POINT_X
        UTM_Y = in_row.POINT_Y
        Aspect = in_row.Aspect
    del in_row
    del in_rows

    in_rows = arcpy.SearchCursor(orderGeometryPR)
    for in_row in in_rows:
        # there is only one line
        UTM_Zone = str(in_row.UTM)[32:44]
    del in_row
    del in_rows

    need_viewer = 'N'
    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        cur.execute("select psr_viewer from order_viewer where order_id =" + str(OrderIDText))
        t = cur.fetchone()
        if t != None:
            need_viewer = t[0]

    finally:
        cur.close()
        con.close()

    if need_viewer == 'Y':
        #clip wetland, flood, geology, soil and covnert .lyr to kml
        #for now, use clipFrame_topo to clip
        #added clip current topo
      
        viewerdir_topo = os.path.join(scratch_folder,OrderNumText+'_psrtopo')
        if not os.path.exists(viewerdir_topo):
            os.mkdir(viewerdir_topo)
        viewertemp =os.path.join(scratch_folder,'viewertemp')
        if not os.path.exists(viewertemp):
            os.mkdir(viewertemp)

        viewerdir_relief = os.path.join(scratch_folder,OrderNumText+'_psrrelief')
        if not os.path.exists(viewerdir_relief):
            os.mkdir(viewerdir_relief)

        datalyr_wetland = PSR_config.datalyr_wetland#r"E:\GISData\PSR\python\mxd\wetland_kml.lyr"
        datalyr_flood = PSR_config.datalyr_flood#r"E:\GISData\PSR\python\mxd\flood.lyr"
        datalyr_geology = PSR_config.datalyr_geology#r"E:\GISData\PSR\python\mxd\geology.lyr"
        masterfilesoil = os.path.join(datapath_soil,'MUPOLYGON')
        

# wetland ----------------------------------------------------------------------------------------------------
        wetlandclip = os.path.join(scratch_gdb, "wetlandclip")
        mxdname = glob.glob(os.path.join(scratch_folder,'mxd_wetland.mxd'))[0]
        mxd = arcpy.mapping.MapDocument(mxdname)
        df = arcpy.mapping.ListDataFrames(mxd,"big")[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
        df.spatialReference = srWGS84
        if siteState == 'AK':
            df.spatialReference = srGoogle
        #re-focus using Buffer layer for multipage
        if multipage_wetland == True:
            bufferLayer = arcpy.mapping.ListLayers(mxd, "Buffer", df)[0]
            df.extent = bufferLayer.getSelectedExtent(False)
            df.scale = df.scale * 1.1

        dfAsFeature = arcpy.Polygon(arcpy.Array([df.extent.lowerLeft, df.extent.lowerRight, df.extent.upperRight, df.extent.upperLeft]),
                            df.spatialReference)    #df.spatialReference is currently UTM. dfAsFeature is a feature, not even a layer
        del df, mxd
        wetland_boudnary = os.path.join(scratch_gdb,"Extent_wetland_WGS84")
        arcpy.Project_management(dfAsFeature, wetland_boudnary, srWGS84)
        arcpy.Clip_analysis(datalyr_wetland, wetland_boudnary, wetlandclip)
        del dfAsFeature

        if int(arcpy.GetCount_management(wetlandclip).getOutput(0)) != 0:
            arcpy.AddField_management(wetland_boudnary,"WETLAND_TYPE", "TEXT", "", "", "15", "", "NULLABLE", "NON_REQUIRED", "")
            wetlandclip1 = os.path.join(scratch_gdb, "wetlandclip1")
            arcpy.Union_analysis([wetlandclip,wetland_boudnary],wetlandclip1)

            keepFieldList = ("WETLAND_TYPE")
            fieldInfo = ""
            fieldList = arcpy.ListFields(wetlandclip1)
            for field in fieldList:
                if field.name in keepFieldList:
                    if field.name == 'WETLAND_TYPE':
                        fieldInfo = fieldInfo + field.name + " " + "Wetland Type" + " VISIBLE;"
                    else:
                        pass
                else:
                    fieldInfo = fieldInfo + field.name + " " + field.name + " HIDDEN;"
            # print fieldInfo

            arcpy.MakeFeatureLayer_management(wetlandclip1, 'wetlandclip_lyr', "", "", fieldInfo[:-1])
            arcpy.ApplySymbologyFromLayer_management('wetlandclip_lyr', datalyr_wetland)
            arcpy.LayerToKML_conversion('wetlandclip_lyr', os.path.join(viewerdir_kml,"wetlandclip.kmz"))
            arcpy.Delete_management('wetlandclip_lyr')
        else:
            print "no wetland data, no kml to folder"
            arcpy.MakeFeatureLayer_management(wetlandclip, 'wetlandclip_lyr')
            arcpy.LayerToKML_conversion('wetlandclip_lyr', os.path.join(viewerdir_kml,"wetlandclip_nodata.kmz"))
            arcpy.Delete_management('wetlandclip_lyr')

# NY wetland --------------------------------------------------------------------------------------------------------
        if ProvStateText == 'NY':
            wetlandclipNY = os.path.join(scratch_gdb, "wetlandclipNY")
            mxdname = glob.glob(os.path.join(scratch_folder,'mxd_wetlandNY.mxd'))[0]
            mxd = arcpy.mapping.MapDocument(mxdname)
            df = arcpy.mapping.ListDataFrames(mxd,"big")[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
            df.spatialReference = srWGS84
            if multipage_wetland == True:
                bufferLayer = arcpy.mapping.ListLayers(mxd, "Buffer", df)[0]
                df.extent = bufferLayer.getSelectedExtent(False)
                df.scale = df.scale * 1.1

            del df, mxd
            wetland_boudnary = os.path.join(scratch_gdb,"Extent_wetland_WGS84")
            datalyr_wetlandNYkml = PSR_config.datalyr_wetlandNYkml
            arcpy.Clip_analysis(datalyr_wetlandNYkml, wetland_boudnary, wetlandclipNY)
            if int(arcpy.GetCount_management(wetlandclipNY).getOutput(0)) != 0:
                wetlandclip1NY = os.path.join(scratch_gdb, "wetlandclip1NY")
                arcpy.Union_analysis([wetlandclipNY,wetland_boudnary],wetlandclip1NY)

                keepFieldList = ("CLASS")
                fieldInfo = ""
                fieldList = arcpy.ListFields(wetlandclip1NY)
                for field in fieldList:
                    if field.name in keepFieldList:
                        if field.name == 'CLASS':
                            fieldInfo = fieldInfo + field.name + " " + "Wetland CLASS" + " VISIBLE;"
                        else:
                            pass
                    else:
                        fieldInfo = fieldInfo + field.name + " " + field.name + " HIDDEN;"

                arcpy.MakeFeatureLayer_management(wetlandclip1NY, 'wetlandclipNY_lyr', "", "", fieldInfo[:-1])
                arcpy.ApplySymbologyFromLayer_management('wetlandclipNY_lyr', datalyr_wetlandNYkml)
                arcpy.LayerToKML_conversion('wetlandclipNY_lyr', os.path.join(viewerdir_kml,"w_NYwetland.kmz"))
                #arcpy.SaveToLayerFile_management(r"wetlandclipNY_lyr",os.path.join(scratch_folder,"NYwetland.lyr"))
                arcpy.Delete_management('wetlandclipNY_lyr')
            else:
                print "no wetland data, no kml to folder"
                arcpy.MakeFeatureLayer_management(wetlandclipNY, 'wetlandclip_lyrNY')
                arcpy.LayerToKML_conversion('wetlandclip_lyrNY', os.path.join(viewerdir_kml,"w_NYwetland_nodata.kmz"))
                arcpy.Delete_management('wetlandclip_lyrNY')

#################################
            datalyr_wetlandNYAPAkml = PSR_config.datalyr_wetlandNYAPAkml
            wetlandclipNYAPA = os.path.join(scratch_gdb, "wetlandclipNYAPA")
            arcpy.Clip_analysis(datalyr_wetlandNYAPAkml, wetland_boudnary, wetlandclipNYAPA)
            if int(arcpy.GetCount_management(wetlandclipNYAPA).getOutput(0)) != 0:
                wetlandclipNYAPA1 = os.path.join(scratch_gdb, "wetlandclipNYAPA1")
                arcpy.Union_analysis([wetlandclipNYAPA,wetland_boudnary],wetlandclipNYAPA1)

                keepFieldList = ("ERIS_WTLD")
                fieldInfo = ""
                fieldList = arcpy.ListFields(wetlandclipNYAPA)
                for field in fieldList:
                    if field.name in keepFieldList:
                        if field.name == 'ERIS_WTLD':
                            fieldInfo = fieldInfo + field.name + " " + "Wetland CLASS" + " VISIBLE;"
                        else:
                            pass
                    else:
                        fieldInfo = fieldInfo + field.name + " " + field.name + " HIDDEN;"
                # print fieldInfo

                arcpy.MakeFeatureLayer_management(wetlandclipNYAPA1, 'wetlandclipNYAPA1_lyr', "", "", fieldInfo[:-1])
                arcpy.ApplySymbologyFromLayer_management('wetlandclipNYAPA1_lyr', datalyr_wetlandNYAPAkml)
                arcpy.LayerToKML_conversion('wetlandclipNYAPA1_lyr', os.path.join(viewerdir_kml,"w_APAwetland.kmz"))
                #arcpy.SaveToLayerFile_management(r"wetlandclipNYAPA1_lyr",os.path.join(scratch_folder,"APAwetland.lyr"))
                arcpy.Delete_management('wetlandclipNYAPA1_lyr')
            else:
                print "no wetland data, no kml to folder"
                arcpy.MakeFeatureLayer_management(wetlandclipNYAPA, 'wetlandclip_lyrNYAPA')
                arcpy.LayerToKML_conversion('wetlandclip_lyrNYAPA', os.path.join(viewerdir_kml,"w_APAwetland_nodata.kmz"))
                arcpy.Delete_management('wetlandclip_lyrNYAPA')

# flood -----------------------------------------------------------------------------------------------------------
        floodclip = os.path.join(scratch_gdb, "floodclip")
        mxdname = glob.glob(os.path.join(scratch_folder,'mxd_flood.mxd'))[0]
        mxd = arcpy.mapping.MapDocument(mxdname)
        df = arcpy.mapping.ListDataFrames(mxd,"Flood Hazard Zone")[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
        df.spatialReference = srWGS84
        if siteState == 'AK':
            df.spatialReference = srGoogle
        if multipage_flood == True:
            bufferLayer = arcpy.mapping.ListLayers(mxd, "Buffer", df)[0]
            df.extent = bufferLayer.getSelectedExtent(False)
            df.scale = df.scale * 1.1

        dfAsFeature = arcpy.Polygon(arcpy.Array([df.extent.lowerLeft, df.extent.lowerRight, df.extent.upperRight, df.extent.upperLeft]),
                            df.spatialReference)    #df.spatialReference is currently UTM. dfAsFeature is a feature, not even a layer
        del df, mxd
        arcpy.Project_management(dfAsFeature, os.path.join(viewertemp,"Extent_flood_WGS84.shp"), srWGS84)

        try:
            data_flood = PSR_config.data_flood
            arcpy.Clip_analysis(data_flood, os.path.join(viewertemp,"Extent_flood_WGS84.shp"), floodclip)
            #arcpy.Clip_analysis(os.path.join(scratch_gdb, "flood"), os.path.join(viewertemp,"Extent_flood_WGS84.shp"), floodclip)
        except arcpy.ExecuteError as e:
            print e
            arcpy.RepairGeometry_management(os.path.join(scratch_gdb, "flood"))
            arcpy.Clip_analysis(os.path.join(scratch_gdb, "flood"), os.path.join(viewertemp,"Extent_flood_WGS84.shp"), floodclip)
        except:
            print("Unexpected error:", sys.exc_info()[0])
            raise

        del dfAsFeature

        if int(arcpy.GetCount_management(floodclip).getOutput(0)) != 0:
            arcpy.AddField_management(floodclip, "CLASS", "TEXT", "", "", "15", "", "NULLABLE", "NON_REQUIRED", "")
            arcpy.AddField_management(floodclip,"ERISBIID", "TEXT", "", "", "15", "", "NULLABLE", "NON_REQUIRED", "")
            rows = arcpy.UpdateCursor(floodclip)
            for row in rows:
                row.CLASS = row.ERIS_CLASS
                ID = [id[1] for id in flood_IDs if row.ERIS_CLASS==id[0]]
                if ID !=[]:
                    row.ERISBIID = ID[0]
                    rows.updateRow(row)
                rows.updateRow(row)
            del rows
            keepFieldList = ("ERISBIID","CLASS", "FLD_ZONE","ZONE_SUBTY")
            fieldInfo = ""
            fieldList = arcpy.ListFields(floodclip)
            for field in fieldList:
                if field.name in keepFieldList:
                    if field.name =='ERISBIID':
                        fieldInfo = fieldInfo + field.name + " " + "ERISBIID" + " VISIBLE;"
                    elif field.name == 'CLASS':
                        fieldInfo = fieldInfo + field.name + " " + "Flood Zone Label" + " VISIBLE;"
                    elif field.name == 'FLD_ZONE':
                        fieldInfo = fieldInfo + field.name + " " + "Flood Zone" + " VISIBLE;"
                    elif field.name == 'ZONE_SUBTY':
                        fieldInfo = fieldInfo + field.name + " " + "Zone Subtype" + " VISIBLE;"
                    else:
                        pass
                else:
                    fieldInfo = fieldInfo + field.name + " " + field.name + " HIDDEN;"
            # print fieldInfo
            arcpy.MakeFeatureLayer_management(floodclip, 'floodclip_lyr', "", "", fieldInfo[:-1])
            arcpy.ApplySymbologyFromLayer_management('floodclip_lyr', datalyr_flood)
            arcpy.LayerToKML_conversion('floodclip_lyr', os.path.join(viewerdir_kml,"floodclip.kmz"))
            arcpy.Delete_management('floodclip_lyr')
        else:
            print "no flood data to kml"
            arcpy.MakeFeatureLayer_management(floodclip, 'floodclip_lyr')
            arcpy.LayerToKML_conversion('floodclip_lyr', os.path.join(viewerdir_kml,"floodclip_nodata.kmz"))
            arcpy.Delete_management('floodclip_lyr')

# geology ------------------------------------------------------------------------------------------------------------
        geologyclip = os.path.join(scratch_gdb, "geologyclip")
        mxdname = glob.glob(os.path.join(scratch_folder,'mxd_geol.mxd'))[0]
        mxd = arcpy.mapping.MapDocument(mxdname)
        df = arcpy.mapping.ListDataFrames(mxd,"*")[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
        df.spatialReference = srWGS84
        if siteState == 'AK':
            df.spatialReference = srGoogle
        if multipage_geology == True:
            bufferLayer = arcpy.mapping.ListLayers(mxd, "Buffer", df)[0]
            df.extent = bufferLayer.getSelectedExtent(False)
            df.scale = df.scale * 1.1

        dfAsFeature = arcpy.Polygon(arcpy.Array([df.extent.lowerLeft, df.extent.lowerRight, df.extent.upperRight, df.extent.upperLeft]),
                            df.spatialReference)    #df.spatialReference is currently UTM. dfAsFeature is a feature, not even a layer
        del df, mxd
        arcpy.Project_management(dfAsFeature, os.path.join(viewertemp,"Extent_geol_WGS84.shp"), srWGS84)
        arcpy.Clip_analysis(datalyr_geology, os.path.join(viewertemp,"Extent_geol_WGS84.shp"), geologyclip)
        del dfAsFeature

        if int(arcpy.GetCount_management(geologyclip).getOutput(0)) != 0:
            arcpy.AddField_management(geologyclip,"ERISBIID", "TEXT", "", "", "15", "", "NULLABLE", "NON_REQUIRED", "")
            rows = arcpy.UpdateCursor(geologyclip)
            for row in rows:
                ID = [id[1] for id in geology_IDs if row.ERIS_KEY==id[0]]
                if ID !=[]:
                    row.ERISBIID = ID[0]
                    rows.updateRow(row)
            del rows
            keepFieldList = ("ERISBIID","ORIG_LABEL", "UNIT_NAME", "UNIT_AGE","ROCKTYPE1", "ROCKTYPE2", "UNITDESC")
            fieldInfo = ""
            fieldList = arcpy.ListFields(geologyclip)
            for field in fieldList:
                if field.name in keepFieldList:
                    if field.name =='ERISBIID':
                        fieldInfo = fieldInfo + field.name + " " + "ERISBIID" + " VISIBLE;"
                    elif field.name == 'ORIG_LABEL':
                        fieldInfo = fieldInfo + field.name + " " + "Geologic_Unit" + " VISIBLE;"
                    elif field.name == 'UNIT_NAME':
                        fieldInfo = fieldInfo + field.name + " " + "Name" + " VISIBLE;"
                    elif field.name == 'UNIT_AGE':
                        fieldInfo = fieldInfo + field.name + " " + "Age" + " VISIBLE;"
                    elif field.name == 'ROCKTYPE1':
                        fieldInfo = fieldInfo + field.name + " " + "Primary_Rock_Type" + " VISIBLE;"
                    elif field.name == 'ROCKTYPE2':
                        fieldInfo = fieldInfo + field.name + " " + "Secondary_Rock_Type" + " VISIBLE;"
                    elif field.name == 'UNITDESC':
                        fieldInfo = fieldInfo + field.name + " " + "Unit_Description" + " VISIBLE;"
                    else:
                        pass
                else:
                    fieldInfo = fieldInfo + field.name + " " + field.name + " HIDDEN;"
            # print fieldInfo
            arcpy.MakeFeatureLayer_management(geologyclip, 'geologyclip_lyr', "", "", fieldInfo[:-1])
            arcpy.ApplySymbologyFromLayer_management('geologyclip_lyr', datalyr_geology)
            arcpy.LayerToKML_conversion('geologyclip_lyr', os.path.join(viewerdir_kml,"geologyclip.kmz"))
            arcpy.Delete_management('geologyclip_lyr')
        else:
            # print "no geology data to kml"
            arcpy.MakeFeatureLayer_management(geologyclip, 'geologyclip_lyr')
            arcpy.LayerToKML_conversion('geologyclip_lyr', os.path.join(viewerdir_kml,"geologyclip_nodata.kmz"))
            arcpy.Delete_management('geologyclip_lyr')

# soil ----------------------------------------------------------------------------------------------------------
        if os.path.exists((os.path.join(scratch_folder,"mxd_soil.mxd"))):
            soilclip = os.path.join(scratch_gdb,"soilclip")
            mxdname = glob.glob(os.path.join(scratch_folder,'mxd_soil.mxd'))[0]
            mxd = arcpy.mapping.MapDocument(mxdname)
            df = arcpy.mapping.ListDataFrames(mxd,"*")[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
            df.spatialReference = srWGS84
            if siteState == 'AK':
                df.spatialReference = srGoogle
            if multipage_soil == True:
                bufferLayer = arcpy.mapping.ListLayers(mxd, "Buffer", df)[0]
                df.extent = bufferLayer.getSelectedExtent(False)
                df.scale = df.scale * 1.1

            dfAsFeature = arcpy.Polygon(arcpy.Array([df.extent.lowerLeft, df.extent.lowerRight, df.extent.upperRight, df.extent.upperLeft]),
                                df.spatialReference)    #df.spatialReference is currently UTM. dfAsFeature is a feature, not even a layer
            del df, mxd
            arcpy.Project_management(dfAsFeature, os.path.join(viewertemp,"Extent_soil_WGS84.shp"), srWGS84)
            arcpy.Clip_analysis(masterfilesoil, os.path.join(viewertemp,"Extent_soil_WGS84.shp"),soilclip)
            del dfAsFeature
            if int(arcpy.GetCount_management(soilclip).getOutput(0)) != 0:
                arcpy.AddField_management(soilclip, "Map_Unit", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
                arcpy.AddField_management(soilclip, "Map_Unit_Name", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
                arcpy.AddField_management(soilclip, "Dominant_Drainage_Class", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
                arcpy.AddField_management(soilclip, "Dominant_Hydrologic_Group", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
                arcpy.AddField_management(soilclip, "Presence_Hydric_Classification", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
                arcpy.AddField_management(soilclip, "Min_Bedrock_Depth", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
                arcpy.AddField_management(soilclip, "Annual_Min_Watertable_Depth", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
                arcpy.AddField_management(soilclip, "component", "TEXT", "", "", "2500", "", "NULLABLE", "NON_REQUIRED", "")
                arcpy.AddField_management(soilclip,"ERISBIID", "TEXT", "", "", "15", "", "NULLABLE", "NON_REQUIRED", "")
                rows = arcpy.UpdateCursor(soilclip)
                for row in rows:
                    for mapunit in reportdata:
                        if row.musym == mapunit["Musym"]:
                            ID = [id[1] for id in soil_IDs if row.MUSYM==id[0]]
                            if ID !=[]:
                                row.ERISBIID = ID[0]
                                rows.updateRow(row)
                            for key in mapunit.keys():
                                if key =="Musym":
                                    row.Map_Unit = mapunit[key]
                                elif key == "Map Unit Name":
                                    row.Map_Unit_Name = mapunit[key]
                                elif key == "Bedrock Depth - Min":
                                    row.Min_Bedrock_Depth = mapunit[key]
                                elif key =="Drainage Class - Dominant":
                                    row.Dominant_Drainage_Class = mapunit[key]
                                elif key =="Hydric Classification - Presence":
                                    row.Presence_Hydric_Classification = mapunit[key]
                                elif key =="Hydrologic Group - Dominant":
                                    row.Dominant_Hydrologic_Group = mapunit[key]
                                elif key =="Watertable Depth - Annual Min":
                                    row.Annual_Min_Watertable_Depth = mapunit[key]
                                elif key =="component":
                                    new = ''
                                    component = mapunit[key]
                                    for i in range(len(component)):
                                        for j in range(len(component[i])):
                                            for k in range(len(component[i][j])):
                                                new = new+component[i][j][k]+" "
                                    row.component = new
                                else:
                                    pass
                                rows.updateRow(row)
                del rows

                keepFieldList = ("ERISBIID","Map_Unit", "Map_Unit_Name", "Dominant_Drainage_Class","Dominant_Hydrologic_Group", "Presence_Hydric_Classification", "Min_Bedrock_Depth","Annual_Min_Watertable_Depth","component")
                fieldInfo = ""
                fieldList = arcpy.ListFields(soilclip)
                for field in fieldList:
                    if field.name in keepFieldList:
                        if field.name =='ERISBIID':
                            fieldInfo = fieldInfo + field.name + " " + "ERISBIID" + " VISIBLE;"
                        elif field.name == 'Map_Unit':
                            fieldInfo = fieldInfo + field.name + " " + "Map_Unit" + " VISIBLE;"
                        elif field.name == 'Map_Unit_Name':
                            fieldInfo = fieldInfo + field.name + " " + "Map_Unit_Name" + " VISIBLE;"
                        elif field.name == 'Dominant_Drainage_Class':
                            fieldInfo = fieldInfo + field.name + " " + "Dominant_Drainage_Class" + " VISIBLE;"
                        elif field.name == 'Dominant_Hydrologic_Group':
                            fieldInfo = fieldInfo + field.name + " " + "Dominant_Hydrologic_Group" + " VISIBLE;"
                        elif field.name == 'Presence_Hydric_Classification':
                            fieldInfo = fieldInfo + field.name + " " + "Presence_Hydric_Classification" + " VISIBLE;"
                        elif field.name == 'Min_Bedrock_Depth':
                            fieldInfo = fieldInfo + field.name + " " + "Min_Bedrock_Depth" + " VISIBLE;"
                        elif field.name == 'Annual_Min_Watertable_Depth':
                            fieldInfo = fieldInfo + field.name + " " + "Annual_Min_Watertable_Depth" + " VISIBLE;"
                        elif field.name == 'component':
                            fieldInfo = fieldInfo + field.name + " " + "component" + " VISIBLE;"
                        else:
                            pass
                    else:
                        fieldInfo = fieldInfo + field.name + " " + field.name + " HIDDEN;"
                # print fieldInfo

                arcpy.MakeFeatureLayer_management(soilclip, 'soilclip_lyr',"", "", fieldInfo[:-1])
                soilsymbol_copy = os.path.join(scratch_folder,"soillyr_copy.lyr")
                arcpy.SaveToLayerFile_management(soillyr,soilsymbol_copy[:-4])
                arcpy.ApplySymbologyFromLayer_management('soilclip_lyr', soilsymbol_copy)
                arcpy.LayerToKML_conversion('soilclip_lyr', os.path.join(viewerdir_kml,"soilclip.kmz"))
                arcpy.Delete_management('soilclip_lyr')
            else:
                print "no soil data to kml"
                arcpy.MakeFeatureLayer_management(soilclip, 'soilclip_lyr')
                arcpy.LayerToKML_conversion('soilclip_lyr', os.path.join(viewerdir_kml,"soilclip_nodata.kmz"))
                arcpy.Delete_management('soilclip_lyr')

# current topo clipping for Xplorer ------------------------------------------------------------------------------
        if os.path.exists((os.path.join(scratch_folder,"mxd_topo.mxd"))):
            mxdname = glob.glob(os.path.join(scratch_folder,'mxd_topo.mxd'))[0]
            mxd = arcpy.mapping.MapDocument(mxdname)
            df = arcpy.mapping.ListDataFrames(mxd,"*")[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
            df.spatialReference = srWGS84
            if siteState == 'AK':
                df.spatialReference = srGoogle
            if multipage_topo == True:
                bufferLayer = arcpy.mapping.ListLayers(mxd, "Buffer", df)[0]
                df.extent = bufferLayer.getSelectedExtent(False)
                df.scale = df.scale * 1.1

            dfAsFeature = arcpy.Polygon(arcpy.Array([df.extent.lowerLeft, df.extent.lowerRight, df.extent.upperRight, df.extent.upperLeft]),
                                df.spatialReference)
            del df, mxd
            arcpy.Project_management(dfAsFeature, os.path.join(viewertemp,"Extent_topo_WGS84.shp"), srWGS84)
            del dfAsFeature

            tomosaiclist = []
            n = 0
            for item in glob.glob(os.path.join(scratch_folder,'*_TM_geo.tif')):
                try:
                    arcpy.Clip_management(item,"",os.path.join(viewertemp, "topo"+str(n)+".jpg"),os.path.join(viewertemp,"Extent_topo_WGS84.shp"),"255","ClippingGeometry")
                    tomosaiclist.append(os.path.join(viewertemp, "topo"+str(n)+".jpg"))
                    n = n+1
                except Exception, e:
                    print str(e) + item     #possibly not in the clipframe

            imagename = str(year)+".jpg"
            if tomosaiclist !=[]:
                arcpy.MosaicToNewRaster_management(tomosaiclist, viewerdir_topo,imagename,srGoogle,"","","3","MINIMUM","MATCH")
                desc = arcpy.Describe(os.path.join(viewerdir_topo, imagename))
                featbound = arcpy.Polygon(arcpy.Array([desc.extent.lowerLeft, desc.extent.lowerRight, desc.extent.upperRight, desc.extent.upperLeft]),
                                    desc.spatialReference)
                del desc
                tempfeat = os.path.join(scratch_folder, "imgbnd_"+str(year)+ ".shp")

                arcpy.Project_management(featbound, tempfeat, srWGS84) #function requires output not be in_memory
                del featbound
                desc = arcpy.Describe(tempfeat)

                metaitem = {}
                metaitem['type'] = 'psrtopo'
                metaitem['imagename'] = imagename
                metaitem['lat_sw'] = desc.extent.YMin
                metaitem['long_sw'] = desc.extent.XMin
                metaitem['lat_ne'] = desc.extent.YMax
                metaitem['long_ne'] = desc.extent.XMax

                try:
                    con = cx_Oracle.connect(connectionString)
                    cur = con.cursor()

                    cur.execute("delete from overlay_image_info where  order_id = %s and (type = 'psrtopo')" % str(OrderIDText))

                    cur.execute("insert into overlay_image_info values (%s, %s, %s, %.5f, %.5f, %.5f, %.5f, %s, '', '')" % (str(OrderIDText), str(OrderNumText), "'" + metaitem['type']+"'", metaitem['lat_sw'], metaitem['long_sw'], metaitem['lat_ne'], metaitem['long_ne'],"'"+metaitem['imagename']+"'" ) )
                    con.commit()

                finally:
                    cur.close()
                    con.close()

        if os.path.exists(os.path.join(viewertemp,"Extent_topo_WGS84.shp")):
            topoframe = os.path.join(viewertemp,"Extent_topo_WGS84.shp")
        else:
            topoframe =clipFrame_topo

        # clip relief map
        tomosaiclist = []
        n = 0
        for item in glob.glob(os.path.join(scratch_folder,'*_hs.img')):
            try:
                arcpy.Clip_management(item,"",os.path.join(viewertemp, "relief"+str(n)+".jpg"),topoframe,"255","ClippingGeometry")
                tomosaiclist.append(os.path.join(viewertemp, "relief"+str(n)+".jpg"))
                n = n+1
            except Exception, e:
                print str(e) + item     #possibly not in the clipframe

        imagename = "relief.jpg"
        if tomosaiclist != []:
            arcpy.MosaicToNewRaster_management(tomosaiclist, viewerdir_relief,imagename,srGoogle,"","","1","MINIMUM","MATCH")
            desc = arcpy.Describe(os.path.join(viewerdir_relief, imagename))
            featbound = arcpy.Polygon(arcpy.Array([desc.extent.lowerLeft, desc.extent.lowerRight, desc.extent.upperRight, desc.extent.upperLeft]),
                                desc.spatialReference)
            del desc
            if 'year' not in locals():
                year = '0'
            tempfeat = os.path.join(scratch_folder, "imgbnd_"+str(year)+ ".shp")

            arcpy.Project_management(featbound, tempfeat, srWGS84) #function requires output not be in_memory
            del featbound
            desc = arcpy.Describe(tempfeat)
            metaitem = {}
            metaitem['type'] = 'psrrelief'
            metaitem['imagename'] = imagename

            metaitem['lat_sw'] = desc.extent.YMin
            metaitem['long_sw'] = desc.extent.XMin
            metaitem['lat_ne'] = desc.extent.YMax
            metaitem['long_ne'] = desc.extent.XMax

            try:
                con = cx_Oracle.connect(connectionString)
                cur = con.cursor()

                cur.execute("delete from overlay_image_info where  order_id = %s and (type = 'psrrelief')" % str(OrderIDText))

                cur.execute("insert into overlay_image_info values (%s, %s, %s, %.5f, %.5f, %.5f, %.5f, %s, '', '')" % (str(OrderIDText), str(OrderNumText), "'" + metaitem['type']+"'", metaitem['lat_sw'], metaitem['long_sw'], metaitem['lat_ne'], metaitem['long_ne'],"'"+metaitem['imagename']+"'" ) )
                con.commit()

            finally:
                cur.close()
                con.close()

        # clip contour lines
        contourclip = os.path.join(scratch_gdb, "contourclip")
        arcpy.Clip_analysis(datalyr_contour,topoframe, contourclip)

        if int(arcpy.GetCount_management(contourclip).getOutput(0)) != 0:

            keepFieldList = ("CONTOURELE")
            fieldInfo = ""
            fieldList = arcpy.ListFields(contourclip)
            for field in fieldList:
                if field.name in keepFieldList:
                    if field.name == 'CONTOURELE':
                        fieldInfo = fieldInfo + field.name + " " + "elevation" + " VISIBLE;"
                    else:
                        pass
                else:
                    fieldInfo = fieldInfo + field.name + " " + field.name + " HIDDEN;"
            # print fieldInfo

            arcpy.MakeFeatureLayer_management(contourclip, 'contourclip_lyr', "", "", fieldInfo[:-1])
            arcpy.ApplySymbologyFromLayer_management('contourclip_lyr', datalyr_contour)
            arcpy.LayerToKML_conversion('contourclip_lyr', os.path.join(viewerdir_relief,"contourclip.kmz"))
            arcpy.Delete_management('contourclip_lyr')
        else:
            print "no contour data, no kml to folder"
            arcpy.MakeFeatureLayer_management(contourclip, 'contourclip_lyr')
            arcpy.LayerToKML_conversion('contourclip_lyr', os.path.join(viewerdir_relief,"contourclip_nodata.kmz"))
            arcpy.Delete_management('contourclip_lyr')

        if os.path.exists(os.path.join(viewer_path, OrderNumText+"_psrkml")):
            shutil.rmtree(os.path.join(viewer_path, OrderNumText+"_psrkml"))
        shutil.copytree(os.path.join(scratch_folder, OrderNumText+"_psrkml"), os.path.join(viewer_path, OrderNumText+"_psrkml"))
        arcpy.AddMessage('KML files: %s' % os.path.join(viewer_path, OrderNumText+"_psrkml"))
        url = upload_link + "PSRKMLUpload?ordernumber=" + OrderNumText
        urllib.urlopen(url)
        arcpy.AddMessage('Upload KML into Xplorer: %s' % url)

        if os.path.exists(os.path.join(viewer_path, OrderNumText+"_psrtopo")):
            shutil.rmtree(os.path.join(viewer_path, OrderNumText+"_psrtopo"))
        shutil.copytree(os.path.join(scratch_folder, OrderNumText+"_psrtopo"), os.path.join(viewer_path, OrderNumText+"_psrtopo"))
        url = upload_link + "PSRTOPOUpload?ordernumber=" + OrderNumText
        urllib.urlopen(url)

        if os.path.exists(os.path.join(viewer_path, OrderNumText+"_psrrelief")):
            shutil.rmtree(os.path.join(viewer_path, OrderNumText+"_psrrelief"))
        shutil.copytree(os.path.join(scratch_folder, OrderNumText+"_psrrelief"), os.path.join(viewer_path, OrderNumText+"_psrrelief"))
        url = upload_link + "ReliefUpload?ordernumber=" + OrderNumText
        urllib.urlopen(url)

    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        cur.callproc('eris_psr.UpdateOrder', (OrderIDText, UTM_Y, UTM_X, UTM_Zone, site_elev,Aspect))


    finally:
        cur.close()
        con.close()

    print "Process completed " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())

##    shutil.copy(reportcheck_path + r'\\'+OrderNumText+'_US_PSR.pdf', scratch_folder)  # occasionally get permission denied issue here when running locally
    arcpy.SetParameterAsText(1, "success")

except:
    # Get the traceback object
    tb = sys.exc_info()[2]
    tbinfo = traceback.format_tb(tb)[0]

    # Concatenate information together concerning the error into a message string
    pymsg = "PYTHON ERRORS:\nTraceback info:\n" + tbinfo + "\nError Info:\n" + str(sys.exc_info()[1])
    msgs = "ArcPy ERRORS:\n" + arcpy.GetMessages(2) + "\n"

    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()
        ###cur.callproc('eris_psr.ClearOrder', (OrderIDText,))
        cur.callproc('eris_psr.InsertPSRAudit', (OrderIDText, 'python-Error Handling',pymsg))

    finally:
        cur.close()
        con.close()

    # Return python error messages for use in script tool or Python Window
    arcpy.AddError("hit CC's error code in except: OrderID %s "%OrderIDText)
    arcpy.AddError(pymsg)
    arcpy.AddError(msgs)

    # Print Python error messages for use in Python / Python Window
    print pymsg + "\n"
    print msgs
    raise    #raise the error again
end = timeit.default_timer()
arcpy.AddMessage(('End PSR report process. Duration:', round(end -start,4)))