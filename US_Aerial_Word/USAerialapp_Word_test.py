#-------------------------------------------------------------------------------
# Name:        US Aerial report for US_Word
# Purpose:     create US Aerial report in US_Word required Word format
#
# Author:      jliu
#
# Created:     06/03/2017
# Copyright:
# Licence:
#-------------------------------------------------------------------------------

#Use Texas provided Aerial files
#export geometry seperately in emf format
#save picture/vector to the location which links to the word template
#modify text in word template

import time,json
print "#0 " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())

import arcpy, os, win32com
import csv, cx_Oracle
import xml.etree.ElementTree as ET
import operator
import shutil, zipfile
from win32com import client
import logging,traceback
import ConfigParser
##import docxcompose.composer
##from docxcompose.composer import Composer
##import docx

##from docx import Document
##from docxcompose.composer import Composer
#from compdocx import merge_docx

from time import strftime

def server_loc_config(configpath,environment):
    configParser = ConfigParser.RawConfigParser()
    configParser.read(configpath)
    if environment == 'test':
        reportcheck_test = configParser.get('server-config','reportcheck_test')
        reportviewer_test = configParser.get('server-config','reportviewer_test')
        reportinstant_test = configParser.get('server-config','instant_test')
        reportnoninstant_test = configParser.get('server-config','noninstant_test')
        upload_viewer = configParser.get('url-config','uploadviewer')
        server_config = {'reportcheck':reportcheck_test,'viewer':reportviewer_test,'instant':reportinstant_test,'noninstant':reportnoninstant_test,'viewer_upload':upload_viewer}
        return server_config
    elif environment == 'prod':
        reportcheck_prod = configParser.get('server-config','reportcheck_prod')
        reportviewer_prod = configParser.get('server-config','reportviewer_prod')
        reportinstant_prod = configParser.get('server-config','instant_prod')
        reportnoninstant_prod = configParser.get('server-config','noninstant_prod')
        upload_viewer = configParser.get('url-config','uploadviewer_prod')
        server_config = {'reportcheck':reportcheck_prod,'viewer':reportviewer_prod,'instant':reportinstant_prod,'noninstant':reportnoninstant_prod,'viewer_upload':upload_viewer}
        return server_config
    else:
        return 'invalid server configuration'
def goCoverPage(coverInfo):
     global summary_mxdfile
     global scratch
     mxd = arcpy.mapping.MapDocument(Covermxdfile)
     SITENAME = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "siteText")[0]
     SITENAME.text = coverInfo["SITE_NAME"]
     ADDRESS = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "addressText")[0]
     ADDRESS.text = coverInfo["ADDRESS"]
     PROJECT_NUM = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "projectidText")[0]
     PROJECT_NUM.text = coverInfo["PROJECT_NUM"]
     COMPANY_NAME = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "companyText")[0]
     COMPANY_NAME.text = coverInfo["COMPANY_NAME"]
     ORDER_NUM = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "ordernumText")[0]
     ORDER_NUM.text = coverInfo["ORDER_NUM"]
     coverPDF = os.path.join(scratch, "coverpage.emf")
     arcpy.mapping.ExportToEMF(mxd, coverPDF, "PAGE_LAYOUT")
     mxd.saveACopy(os.path.join(scratch,"coverpage.mxd"))
     return coverPDF

#create WORD and also make a copy of the geotiff files if the scale is too small
def createWORD(app):
    infomatrix = []
    directory = os.path.join(viewer_path, OrderNumText,'fin')
    i = 0
    items = os.listdir(directory)
    for i in range(0,len(items)):
        item = items[i]
        if item.lower() == 'thumbs.db' or item.lower() == '_aux' or item.lower() == '_bndry' or item.lower() == '_del' or "jpg" not in item[-4:]:
            continue
        year = item.split('.')[0].split('_')[0]
        source = item.split('.')[0].split('_')[1]
        scale = int(item.split('.')[0].split('_')[2])*12
        if len(item.split('.')[0].split('_')) > 3:
            comment = item.split('.')[0].split('_')[3]
        else:
            comment = ''
        infomatrix.append((item, year, source, scale, comment))


    if OrderType.lower()== 'point':
        orderGeomlyrfile = orderGeomlyrfile_point
    elif OrderType.lower() =='polyline':
        orderGeomlyrfile = orderGeomlyrfile_polyline
    else:
        orderGeomlyrfile = orderGeomlyrfile_polygon

    orderGeomLayer = arcpy.mapping.Layer(orderGeomlyrfile)
    orderGeomLayer.replaceDataSource(scratch,"SHAPEFILE_WORKSPACE","orderGeometry")


    if os.path.exists(os.path.join(scratch,'tozip')):
        shutil.rmtree(os.path.join(scratch,'tozip'))
    shutil.copytree(directorytemplate,os.path.join(scratch,'tozip'))

    # add to map template, clip (but need to keep both metadata: year, grid size, quadrangle name(s) and present in order
    mxd = arcpy.mapping.MapDocument(mxdfile)
    df = arcpy.mapping.ListDataFrames(mxd,"*")[0]
    spatialRef = arcpy.SpatialReference(out_coordinate_system)
    df.spatialReference = spatialRef
    #if OrderType.lower() == "polyline" or OrderType.lower() == "polygon":
    #    arcpy.mapping.AddLayer(df,orderGeomLayer,"Top")
    if OrderType.lower() == "polyline" or OrderType.lower() == "polygon":
        if yesBoundary.lower() == 'y' or yesBoundary.lower() == 'yes':
            arcpy.mapping.AddLayer(df,orderGeomLayer,"Top")

    infomatrix_sorted = sorted(infomatrix,key=operator.itemgetter(1,0), reverse = True)
    for i in range(0,len(infomatrix_sorted)):
        imagename = infomatrix_sorted[i][0]
        year = infomatrix_sorted[i][1]
        source = infomatrix_sorted[i][2]
        scale = infomatrix_sorted[i][3]
        comment = infomatrix_sorted[i][4]

##        for lyr in arcpy.mapping.ListLayers(mxd, "Project Property", df):
##            if lyr.name == "Project Property":
##                if OrderType.lower() == "point":
##                    lyr.visible = False
##                else:
##                    lyr.visible = True
##                df.extent = lyr.getSelectedExtent(False)

        for lyr in arcpy.mapping.ListLayers(mxd, "Project Property", df):
            if lyr.name == "Project Property":
                if OrderType.lower() == "point":
                    lyr.visible = False
                else:
                    lyr.visible = True
                df.extent = lyr.getExtent(True)

        #centerlyr = arcpy.mapping.Layer(orderCenter)
        #centerlyr = arcpy.MakeFeatureLayer_management(orderCenter, "center_lyr")
        #arcpy.mapping.AddLayer(df,centerlyr,"Top")


        #center = arcpy.mapping.ListLayers(mxd, "*", df)[0]
        #df.extent = center.getSelectedExtent(False)
        #center.visible = False


        df.scale = scale
        arcpy.RefreshTOC()
        arcpy.RefreshActiveView()

        outputemf = os.path.join(scratch, year+".emf")
        print outputemf
        arcpy.mapping.ExportToEMF(mxd, outputemf, "PAGE_LAYOUT")
        mxd.saveACopy(os.path.join(scratch,year+"_emf.mxd"))

        shutil.copyfile(os.path.join(directory,imagename), os.path.join(scratch,"tozip\word\media\image2.jpeg"))
        shutil.copyfile(os.path.join(scratch,year+".emf"), os.path.join(scratch,"tozip\word\media\image1.emf"))
        zipdir_noroot(os.path.join(scratch,'tozip'),year+".docx")
        worddoclist.append(os.path.join(scratch,year+".docx"))

        #the word template has been copied, the image files have also been copied, need to refresh and replace the text fields, save
        doc = app.Documents.Open(os.path.join(scratch,year+".docx"))

        fileName = OrderNumText
        fileDate = time.strftime('%Y-%m-%d', time.localtime())

        #quads = 'AERIAL PHOTOGRAPHY FROM SOURCE ' + source + '(' +str(year) + ')'
        quads = ''

        if scale == 6000:
            scaletxt = '1":' + str(scale/12)+"'"
        #scaletxt = '1:' + str(scale)
        else:
            scaletxt = '1:' + str(scale)


        allShapes = doc.Shapes
        allShapes(3).TextFrame.TextRange.Text = 'AERIAL PHOTO (' + str(year) + '-' + source+')'    #AERIAL PHOTOGRAPH line

        txt = allShapes(4).TextFrame.TextRange.Text.replace('Site Name', siteName)
        #allShapes(4).TextFrame.TextRange.Text.replace('Site Address', siteAddress)
        #allShapes(4).TextFrame.TextRange.Text.replace('Site City, Site State', siteCityState)
        txt = txt.replace('Site Address', siteAddress)
        txt = txt.replace('Site City, Site State', siteCityState)
        if not custom_profile:
            allShapes(4).TextFrame.TextRange.Text = txt
            allShapes(9).TextFrame.TextRange.Text = quads
            allShapes(11).TextFrame.TextRange.Text = officeAddress
            allShapes(12).TextFrame.TextRange.Text = officeCity
            allShapes(13).TextFrame.TextRange.Text = proNo
    ##        allShapes(26).TextFrame.TextRange.Text = scaletxt
    ##        allShapes(27).TextFrame.TextRange.Text = fileName
    ##        allShapes(28).TextFrame.TextRange.Text = fileDate
    ##        allShapes(13).TextFrame.TextRange.Text = proNo
            allShapes(23).TextFrame.TextRange.Text = scaletxt#good
            allShapes(24).TextFrame.TextRange.Text = fileName#good
            allShapes(25).TextFrame.TextRange.Text = fileDate#good
        else:
            allShapes(4).TextFrame.TextRange.Text = txt
            allShapes(9).TextFrame.TextRange.Text = quads
            allShapes(11).TextFrame.TextRange.Text = officeAddress
            allShapes(12).TextFrame.TextRange.Text = officeCity
            allShapes(13).TextFrame.TextRange.Text = proNo
    ##        allShapes(26).TextFrame.TextRange.Text = scaletxt
    ##        allShapes(27).TextFrame.TextRange.Text = fileName
    ##        allShapes(28).TextFrame.TextRange.Text = fileDate
    ##        allShapes(13).TextFrame.TextRange.Text = proNo
            allShapes(23).TextFrame.TextRange.Text = scaletxt#good
            allShapes(24).TextFrame.TextRange.Text = fileName#good
            allShapes(25).TextFrame.TextRange.Text = fileDate#good
        doc.Save()
        doc.Close()
        doc = None

    del mxd
    return infomatrix_sorted


def zipdir_noroot(path, zipfilename):
    myZipFile = zipfile.ZipFile(os.path.join(scratch,zipfilename),"w")
    for root, dirs, files in os.walk(path):
        for afile in files:
            arcname = os.path.relpath(os.path.join(root, afile), path)
            myZipFile.write(os.path.join(root, afile), arcname)
    myZipFile.close()

def get_sourcename(source_acronym):
    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        cur.execute("select fullname from aerial_source where source = '%s'"%(source_acronym))
        t = cur.fetchone()
        source_fullname = str(t[0])
        return source_fullname
    except Exception as e:
        print e.message

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
handler = logging.FileHandler(r'\\cabcvan1gis006\GISData\Aerial_US_Word\python\log\USAerial_US_Word_Log.txt')
handler.setLevel(logging.DEBUG)
logger.addHandler(handler)

lookup_state = {
'AL': 'Alabama',
'AK': 'Alaska',
'AZ': 'Arizona',
'AR': 'Arkansas',
'CA': 'California',
'CO': 'Colorado',
'CT': 'Connecticut',
'DC': 'District of Columbia',
'DE': 'Delaware',
'FL': 'Florida',
'GA': 'Georgia',
'HI': 'Hawaii',
'ID': 'Idaho',
'IL': 'Illinois',
'IN': 'Indiana',
'IA': 'Iowa',
'KS': 'Kansas',
'KY': 'Kentucky',
'LA': 'Louisiana',
'ME': 'Maine',
'MD': 'Maryland',
'MA': 'Massachusetts',
'MI': 'Michigan',
'MN': 'Minnesota',
'MS': 'Mississippi',
'MO': 'Missouri',
'MT': 'Montana',
'NE': 'Nebraska',
'NV': 'Nevada',
'NH': 'New Hampshire',
'NJ': 'New Jersey',
'NM': 'New Mexico',
'NY': 'New York',
'NC': 'North Carolina',
'ND': 'North Dakota',
'OH': 'Ohio',
'OK': 'Oklahoma',
'OR': 'Oregon',
'PA': 'Pennsylvania',
'RI': 'Rhode Island',
'SC': 'South Carolina',
'SD': 'South Dakota',
'TN': 'Tennessee',
'TX': 'Texas',
'UT': 'Utah',
'VT': 'Vermont',
'VA': 'Virginia',
'WA': 'Washington',
'WV': 'West Virginia',
'WI': 'Wisconsin',
'WY': 'Wyoming',
'PR': 'Puerto Rico',
'VI': 'Virgin Islands',
'ON': 'Ontario',
'BC': 'British Columbia',
'AB': 'Alberta',
'MB': 'Manitoba',
'SK': 'Saskatchewan',
'QC': 'Quebec',
'NS': 'Nova Scotia',
'NB': 'New Brunswick',
'PE': 'Prince Edward Island',
'NL': 'Newfoundland and Labrador',
'NT': 'Northwest Territories',
'YK': 'Yukon',
'NU': 'Nunavut'
}
server_environment = 'test'
server_config_file = r'\\cabcvan1gis006\GISData\ERISServerConfig.ini'
server_config = server_loc_config(server_config_file,server_environment)
##deploy parameter to change
####################################################################
connectionString = 'ERIS_GIS/gis295@GMTESTC.glaciermedia.inc'
viewer_path = r'\\cabcvan1eap003\v2_usaerial\JobData\test'
#viewer_path = r'E:\GISData\Aerial_US_Word'
reportcheck_path = os.path.join(server_config['reportcheck'],'AerialsDigital')
####################################################################


##global parameters
connectionPath = r'\\cabcvan1gis006\GISData\Aerial_US_Word\python'
mxdfile = r'\\cabcvan1gis006\GISData\Aerial_US_Word\python\templates\template.mxd'
orderGeomlyrfile_point = os.path.join(r'\\cabcvan1gis006\GISData\Aerial_US_Word\python\templates',"SiteMaker.lyr")
orderGeomlyrfile_polyline = r'\\cabcvan1gis006\GISData\Aerial_US_Word\python\templates\orderLine.lyr'
orderGeomlyrfile_polygon = r'\\cabcvan1gis006\GISData\Aerial_US_Word\python\templates\orderPoly.lyr'
marginTemplate = r'\\cabcvan1gis006\GISData\Aerial_US_Word\python\templates\margin.docx'
directorytemplate = r'\\cabcvan1gis006\GISData\Aerial_US_Word\python\templates\Environmental-Portrait_AerialOnly_dev_fin'

Summarymxdfile = os.path.join(connectionPath, r'mxd\SummaryPage.mxd')
Covermxdfile = os.path.join(connectionPath, r'mxd\CoverPage.mxd')
coverTemplate = r'\\cabcvan1gis006\GISData\Aerial_US_Word\python\templates\CoverPage'
summaryTemplate = r'\\cabcvan1gis006\GISData\\Aerial_US_Word\python\templates\SummaryPage'

OrderIDText = '1058505'#arcpy.GetParameterAsText(0)
yesBoundary = 'yes'#arcpy.GetParameterAsText(1)
scratch = r'C:\Users\JLoucks\Documents\JL\test4'#arcpy.env.scratchWorkspace
custom_profile = False

#OrderIDText = '934292'
#scratch = r"C:\Users\JLoucks\Documents\JL\test1"
#yesBoundary = 'no'

if yesBoundary == 'arrow':
    yesBoundary = 'yes'
    custom_profile = True
else:
    print 'no custom profile set'

if not custom_profile:
    directorytemplate = r"\\cabcvan1gis006\GISData\Aerial_US_Word\python\templates\Environmental-Portrait_AerialOnly_noarrow_fin"
else:
    directorytemplate = r"\\cabcvan1gis006\GISData\Aerial_US_Word\python\templates\Environmental-Portrait_AerialOnly_dev_fin_20200717"

try:
    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        coverInfotext = json.loads(cur.callfunc('eris_gis.getCoverPageInfo', str, (str(OrderIDText),)))
        for key in coverInfotext.keys():
            if coverInfotext[key]=='':
                coverInfotext[key]=' '
        OrderNumText = str(coverInfotext["ORDER_NUM"])
        siteName =coverInfotext["SITE_NAME"]
        proNo = coverInfotext["PROJECT_NUM"]
        ProName = coverInfotext["COMPANY_NAME"]
        siteAddress =coverInfotext["ADDRESS"]
        coverState = lookup_state[coverInfotext["PROVSTATE"]]+" "+coverInfotext["POSTALZIP"]
        siteCityState=coverInfotext["CITY"]+", "+coverState

        coverInfotext["ADDRESS"] = '%s\n%s %s %s'%(coverInfotext["ADDRESS"],coverInfotext["CITY"],coverInfotext["PROVSTATE"],coverInfotext["POSTALZIP"])

##    OrderDetails = json.loads(cur.callfunc('eris_gis.getBufferDetails', str, (str(orderIDText),)))
##    OrderType = OrderDetails["ORDERTYPE"]
##    OrderCoord = eval(OrderDetails["ORDERCOOR"])
##    RadiusType = OrderDetails["RADIUSTYPE"]
        cur.execute("select geometry_type, geometry, radius_type  from eris_order_geometry where order_id =" + OrderIDText)
        t = cur.fetchone()
        OrderType = str(t[0])
        OrderCoord = eval(str(t[1]))
        RadiusType = str(t[2])

        cur.execute("select customer_id from orders where order_id =" + OrderIDText)
        t = cur.fetchone()
        customer_id = str(t[0])
        cur.execute("select address1, address2, city, provstate, postal_code  from customer where customer_id =" + customer_id)
        t = cur.fetchone()
        if t[1] == None:
            officeAddress = str(t[0])
        else:
            officeAddress = str(t[0])+", "+str(t[1])
        officeCity = str(t[2])+", "+lookup_state[str(t[3])]+" "+str(t[4])

        cur.execute("select SWLAT,SWLONG,NELAT,NELONG from aerial_image_info where order_id = "+OrderIDText)
        t = cur.fetchone()
        long_center = str(t[0])
        lat_center = str(t[1])

    except Exception,e:
        logger.error("Error to get flag from Oracle " + str(e))
        raise
    finally:
        cur.close()
        con.close()

    docName = OrderNumText+'_US_Aerial.docx'

    arcpy.env.overwriteOutput = True
    arcpy.env.OverWriteOutput = True


    srGCS83 = arcpy.SpatialReference(4269)


    #create the center point shapefile, for positioning
    point = arcpy.Point()
    array = arcpy.Array()
    sr = arcpy.SpatialReference()
    sr.factoryCode = 4269  # requires input geometry is in 4269
    sr.XYTolerance = .00000001
    sr.scaleFactor = 2000
    sr.create()
    featureList = []
    point.X = float(long_center)
    point.Y = float(lat_center)
    sr.setDomain (point.X, point.X, point.Y, point.Y)
    array.add(point)
    feat = arcpy.Multipoint(array, sr)
    # Append to the list of Polygon objects
    featureList.append(feat)

    orderCenter= os.path.join(scratch,"orderCenter.shp")
    arcpy.CopyFeatures_management(featureList, orderCenter)
    arcpy.DefineProjection_management(orderCenter, srGCS83)
    del point, array, feat, featureList


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

    orderGeometry= os.path.join(scratch,"orderGeometry.shp")
    arcpy.CopyFeatures_management(featureList, orderGeometry)
    arcpy.DefineProjection_management(orderGeometry, srGCS83)


    arcpy.AddField_management(orderGeometry, "UTM", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
    arcpy.CalculateUTMZone_cartography(orderGeometry, 'UTM')
    UT= arcpy.SearchCursor(orderGeometry)
    for row in UT:
        UTMvalue = str(row.getValue('UTM'))[41:43]
    del UT
    out_coordinate_system = os.path.join(connectionPath+'/', r"projections/NAD1983/NAD1983UTMZone"+UTMvalue+"N.prj")


    orderGeometryPR = os.path.join(scratch, "ordergeoNamePR.shp")
    arcpy.Project_management(orderGeometry, orderGeometryPR, out_coordinate_system)

    del point
    del array

    arcpy.AddField_management(orderGeometryPR, "xCentroid", "DOUBLE", 18, 11)
    arcpy.AddField_management(orderGeometryPR, "yCentroid", "DOUBLE", 18, 11)

    xExpression = '!SHAPE.CENTROID.X!'
    yExpression = '!SHAPE.CENTROID.Y!'

    arcpy.CalculateField_management(orderGeometryPR, 'xCentroid', xExpression, "PYTHON_9.3")
    arcpy.CalculateField_management(orderGeometryPR, 'yCentroid', yExpression, "PYTHON_9.3")

    in_rows = arcpy.SearchCursor(orderGeometryPR)
    for in_row in in_rows:
        xCentroid = in_row.xCentroid
        yCentroid = in_row.yCentroid
    del in_row
    del in_rows


    worddoclist = []
    #outputPdfname = "map_" + OrderNumText + ".pdf"
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = 0
    infomatrix_sorted = createWORD(app)

    #step 2:  create summary page
    mxdSummary = arcpy.mapping.MapDocument(Summarymxdfile)

    j=0
    for item in infomatrix_sorted:
        j=j+1
        i=str(j)
        year = item[1]
        source = get_sourcename(item[2])
        if item[3] == 6000:
            scale = '1":' + str(item[3]/12)+"'"
        else:
            scale = "1:"+str(item[3])
        comment = item[4]

        exec("e"+i+"1E = arcpy.mapping.ListLayoutElements(mxdSummary, 'TEXT_ELEMENT', 'e"+i+"1')[0]")
        exec("e"+i+"1E.text = year")
        exec("e"+i+"2E = arcpy.mapping.ListLayoutElements(mxdSummary, 'TEXT_ELEMENT', 'e"+i+"2')[0]")
        exec("e"+i+"2E.text = source")
        exec("e"+i+"3E = arcpy.mapping.ListLayoutElements(mxdSummary, 'TEXT_ELEMENT', 'e"+i+"3')[0]")
        exec("e"+i+"3E.text = scale")
        if comment <> '':
            exec("e"+i+"4E = arcpy.mapping.ListLayoutElements(mxdSummary, 'TEXT_ELEMENT', 'e"+i+"4')[0]")
            exec("e"+i+"4E.text = comment")

    summaryEmf = os.path.join(scratch, "summary.emf")
    arcpy.mapping.ExportToEMF(mxdSummary, summaryEmf, "PAGE_LAYOUT")
    mxdSummary.saveACopy(os.path.join(scratch, "summarypage.mxd"))
    mxdSummary = None

    zipCover= os.path.join(scratch,'tozip_cover')
    zipSummary = os.path.join(scratch,'tozip_summary')
    if not os.path.exists(zipCover):
        shutil.copytree(coverTemplate,os.path.join(scratch,'tozip_cover'))
    if not os.path.exists(zipSummary):
        shutil.copytree(summaryTemplate,os.path.join(scratch,'tozip_summary'))

    coverPage = goCoverPage(coverInfotext)
    shutil.copyfile(coverPage, os.path.join(zipCover,"word\media\image2.emf"))
    zipdir_noroot(zipCover,"cover.docx")
    shutil.copyfile(summaryEmf, os.path.join(zipSummary,"word\media\image2.emf"))
    zipdir_noroot(zipSummary,"summary.docx")

    #concatenate the word docs into a big final file
    print "#5-0 " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    shutil.copyfile(marginTemplate,os.path.join(scratch,docName))
    finaldoc = app.Documents.Open(os.path.join(scratch,docName))
    sel = finaldoc.ActiveWindow.Selection
    sel.InsertFile(os.path.join(scratch, "summary.docx"))
    sel.InsertBreak()
    npages = 0
    for aDoc in worddoclist:
        npages = npages + 1
        sel.InsertFile(aDoc)
        if npages < len(worddoclist):
           sel.InsertBreak()
    finaldoc.Save()
    finaldoc.Close()

    finaldoc = app.Documents.Open(os.path.join(scratch,docName))
    sel = finaldoc.ActiveWindow.Selection
    sel.InsertFile(os.path.join(scratch,'cover.docx'))
    finaldoc.Save()
    finaldoc.Close()
    finalDoc = None
    app.Application.Quit(-1)

##    shutil.copyfile(marginTemplate,os.path.join(scratch,docName))
##    finaldoc = app.Documents.Open(os.path.join(scratch,docName))
##    sel = finaldoc.ActiveWindow.Selection
##    npages = 0
##    sel.InsertFile(os.path.join(scratch,'cover.docx'))
##    sel.InsertBreak()
##    sel.InsertFile(os.path.join(scratch,'summary.docx'))
##    sel.InsertBreak()
##    for aDoc in worddoclist:
##        npages = npages + 1
##        sel.InsertFile(aDoc)
##        if npages < len(worddoclist):
##           sel.InsertBreak()
##    finaldoc.Save()
##    finaldoc.Close()
##    finalDoc = None
##    app.Application.Quit()

##    merged_document = docu()
##    index = len(worddoclist)
##    for aDoc in worddoclist:
##        sub_doc = docu(aDoc)
##
##        # Don't add a page break if you've reached the last file.
##        if index < len(worddoclist)-1:
##           sub_doc.add_page_break()
##
##        for element in sub_doc.element.body:
##            merged_document.element.body.append(element)
##
##    merged_document.save(os.path.join(scratch,docName))

##    shutil.copyfile(marginTemplate,os.path.join(scratch,docName))
##    finaldoc = Document(os.path.join(scratch,docName))
##
##    docx_files = [os.path.join(scratch,'cover.docx'),os.path.join(scratch,'summary.docx'),worddoclist]
##
##    composer = Composer(finaldoc) # finaldoc composer
##    for docx_file in docx_files[0:2]:
##        composer.append(Document(docx_file))
##    finaldoc.add_page_break() # add pg break only for summary pg
##
##    for docx_file in docx_files[2][:-1]:
##        composer.append(Document(docx_file))
##        finaldoc.add_page_break()
##    composer.append(Document(docx_files[2][-1])) # append without pg break to prevent blank pg's
##
##    finaldoc.save(os.path.join(scratch,docName))
##    del finaldoc
##    del Document
##    del Composer
##    finalDoc = None

##    shutil.copyfile(marginTemplate,os.path.join(scratch,docName))
##    finaldoc = app.Documents.Open(os.path.join(scratch,docName))
##    finaldoc = app.Documents.Add()
##    sel = finaldoc.ActiveWindow.Selection
##    sel.InsertFile(os.path.join(scratch, "summary.docx"))
##    sel.InsertBreak()
##    npages = 0
##    for aDoc in worddoclist:
##        npages = npages + 1
##        sel.InsertFile(aDoc)
##        if npages < len(worddoclist):
##           sel.InsertBreak()
##    finaldoc.Save()
##    finaldoc.Close()
##
##    finaldoc = app.Documents.Open(os.path.join(scratch,docName))
##    sel = finaldoc.ActiveWindow.Selection
##    sel.InsertFile(os.path.join(scratch,'cover.docx'))
##    finaldoc.SaveAs(os.path.join(scratch,docName))
##    finaldoc.Close()
##    finalDoc = None
##    app.Application.Quit(-1)

    print "#5-1 " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())

    del orderCenter
    del orderGeometry
    shutil.copy(os.path.join(scratch,docName), reportcheck_path)  # occasionally get permission denied issue here when running locally

    arcpy.SetParameterAsText(2, os.path.join(scratch,docName))



    logger.removeHandler(handler)
    handler.close()
except:
    # Get the traceback object
    #
    tb = sys.exc_info()[2]
    tbinfo = traceback.format_tb(tb)[0]
    pymsg = "PYTHON ERRORS:\nTraceback info:\n" + tbinfo + "\nError Info:\n" + str(sys.exc_info()[1])

    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()
        cur.callproc('eris_aerial.InsertAerialAudit', (OrderIDText, 'python-Error Handling',pymsg))

    finally:
        cur.close()
        con.close()

    raise    #raise the error again

