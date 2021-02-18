#-------------------------------------------------------------------------------
# Name:        USFIM
# Purpose:     query the US georeferenced FIM database,
#              pull up relevant images, create pdf and prepare GIS images
#
# Author:      LiuJ
#
# Created:     12/20/2016
# Copyright:   (c) LiuJ 2016
# Licence:     <your licence>
#-------------------------------------------------------------------------------

from xlrd import open_workbook
from PyPDF2 import PdfFileReader,PdfFileWriter
from PyPDF2.generic import NameObject, createStringObject, ArrayObject, FloatObject
import logging, time,json
import arcpy, os, sys, glob
import cx_Oracle, urllib, shutil
import traceback
import ConfigParser

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Frame,Table
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import portrait, letter
from reportlab.pdfgen import canvas
from time import strftime

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
handler = logging.FileHandler(r"\\cabcvan1gis005\GISData\FIMS_USA\temp\USFIM_Log.txt")
handler.setLevel(logging.INFO)
logger.addHandler(handler)

def server_loc_config(configpath,environment):
    configParser = ConfigParser.RawConfigParser()
    configParser.read(configpath)
    if environment == 'test':
        reportcheck = configParser.get('server-config','reportcheck_test')
        reportviewer = configParser.get('server-config','reportviewer_test')
        reportinstant = configParser.get('server-config','instant_test')
        reportnoninstant = configParser.get('server-config','noninstant_test')
        upload_viewer = configParser.get('url-config','uploadviewer')
        server_config = {'reportcheck':reportcheck,'viewer':reportviewer,'instant':reportinstant,'noninstant':reportnoninstant,'viewer_upload':upload_viewer}
        return server_config
    elif environment == 'prod':
        reportcheck = configParser.get('server-config','reportcheck_prod')
        reportviewer = configParser.get('server-config','reportviewer_prod')
        reportinstant = configParser.get('server-config','instant_prod')
        reportnoninstant = configParser.get('server-config','noninstant_prod')
        upload_viewer = configParser.get('url-config','uploadviewer_prod')
        server_config = {'reportcheck':reportcheck,'viewer':reportviewer,'instant':reportinstant,'noninstant':reportnoninstant,'viewer_upload':upload_viewer}
        return server_config
    else:
        return 'invalid server configuration'

def createAnnotPdf(geom_type, myShapePdf):
    #input variables
    #geom_type = 'POLYLINE'      #or POLYGON

    # part 1: read geometry pdf to get the vertices and rectangle to use
    source  = PdfFileReader(open(myShapePdf,'rb'))
    geomPage = source.getPage(0)
    mystr = geomPage.getObject()['/Contents'].getData()
    #to pinpoint the string part: 1.19997 791.75999 m 1.19997 0.19466 l 611.98627 0.19466 l 611.98627 791.75999 l 1.19997 791.75999 l
    #the format seems to follow x1 y1 m x2 y2 l x3 y3 l x4 y4 l x5 y5 l
    geomString = mystr.split('S\r\n')[0].split('M\r\n')[1]
    coordsString = [value for value in geomString.split(' ') if value not in ['m','l','']]

    # part 2: update geometry in the map
    if geom_type.upper() == 'POLYGON':
        pdf_geom = PdfFileReader(open(annot_poly,'rb'))
    elif geom_type.upper() == 'POLYLINE':
        pdf_geom = PdfFileReader(open(annot_line,'rb'))
    page_geom = pdf_geom.getPage(0)

    annot = page_geom['/Annots'][0]
    updateVertices = "annot.getObject().update({NameObject('/Vertices'):ArrayObject([FloatObject("+coordsString[0]+")"
    for item in coordsString[1:]:
        updateVertices = updateVertices + ',FloatObject('+item+')'
    updateVertices = updateVertices + "])})"
    exec(updateVertices)

    xcoords = []
    ycoords = []
    for i in range(0,len(coordsString)-1):
        if i%2 == 0:
            xcoords.append(float(coordsString[i]))
        else:
            ycoords.append(float(coordsString[i]))

    #below rect seems to be geom bounding box coordinates: xmin, ymin, xmax,ymax
    annot.getObject().update({NameObject('/Rect'):ArrayObject([FloatObject(min(xcoords)), FloatObject(min(ycoords)), FloatObject(max(xcoords)), FloatObject(max(ycoords))])})
    annot.getObject().pop('/AP')  # this is to get rid of the ghost shape

    annot.getObject().update({NameObject('/T'):createStringObject(u'ERIS')})

    output = PdfFileWriter()
    output.addPage(page_geom)
    annotPdf = os.path.join(scratch, "annot.pdf")
    outputStream = open(annotPdf,"wb")
    #output.setPageMode('/UseOutlines')
    output.write(outputStream)
    outputStream.close()
    output = None
    return annotPdf

def annotatePdf(mapPdf, myAnnotPdf):
    pdf = PdfFileReader(open(mapPdf,'rb'))
    FIMpage = pdf.getPage(0)

    pdf_intermediate = PdfFileReader(open(myAnnotPdf,'rb'))
    page= pdf_intermediate.getPage(0)
    page.mergePage(FIMpage)

    output = PdfFileWriter()
    output.addPage(page)

    annotatedPdf = mapPdf[:-4]+'_a.pdf'
    outputStream = open(annotatedPdf,"wb")
    #output.setPageMode('/UseOutlines')
    output.write(outputStream)
    outputStream.close()
    output = None
    return annotatedPdf

def goCoverPage(coverPdf):
    from reportlab.pdfgen import canvas
    c = canvas.Canvas(coverPdf,pagesize = portrait(letter))
    from reportlab.lib.units import inch
    c.drawImage(coverPic,0,0, PAGE_WIDTH,PAGE_HEIGHT)
    leftsw= 54
    heights = 400
    rightsw = 200
    space = 20
    c.setFont('Helvetica-Bold', 13)
    c.drawString(leftsw, heights, "Project Property:")
    c.drawString(leftsw, heights-3*space,"Project No:")
    c.drawString(leftsw, heights-4*space,"Requested By:")
    c.drawString(leftsw, heights-5*space,"Order No:")
    c.drawString(leftsw, heights-6*space,"Date Completed:")
    c.setFont('Helvetica', 13)
    c.drawString(rightsw,heights-0*space, coverInfotext["SITE_NAME"])
    c.drawString(rightsw, heights-1*space,coverInfotext["ADDRESS"].split("\n")[0])
    c.drawString(rightsw, heights-2*space,coverInfotext["ADDRESS"].split("\n")[1])
    c.drawString(rightsw, heights-3*space,coverInfotext["PROJECT_NUM"])
    c.drawString(rightsw, heights-4*space,coverInfotext["COMPANY_NAME"])
    c.drawString(rightsw, heights-5*space,coverInfotext["ORDER_NUM"])
    c.drawString(rightsw, heights-6*space,time.strftime('%B %d, %Y', time.localtime()))
    if NRF=='Y':
        c.setStrokeColorRGB(0.67,0.8,0.4)
        c.line(50,180,PAGE_WIDTH-60,180)
        c.setFont('Helvetica-Bold', 12)
        c.drawString(70,160,"Please note that no information was found for your site or adjacent properties.")
    p=None
    Disclaimer = None
    style = None
    c.showPage()
    c.save()

def myFirstSummaryPage(canvas,doc):
    canvas.saveState()
    canvas.setFont('Helvetica', 9)
    canvas.drawImage(secondPic,0,0, PAGE_WIDTH,PAGE_HEIGHT)
    canvas.drawString(54, 690, "Listed below, please find the results of our search for historic fire insurance maps from our in-house collection, performed in")
    canvas.drawString(54, 678,"conjuction with your ERIS report.")
    canvas.drawString(54, 110, "Individual Fire Insurance Maps for the subject property and/or adjacent sites are included with the ERIS environmental database ")
    canvas.drawString(54, 98,"report to be used for research purposes only and cannot be resold for any other commercial uses other than for use in a Phase I")
    canvas.drawString(54, 86,"environmental assessment.")
    canvas.restoreState()
    del canvas

def myLaterSummaryPage(canvas,doc):
    canvas.saveState()
    canvas.drawImage(secondPic,0,0, PAGE_WIDTH,PAGE_HEIGHT)
    canvas.setFont('Helvetica', 9)
    canvas.drawString(56, 690, "continued")
    canvas.drawString(54, 110, "Individual Fire Insurance Maps for the subject property and/or adjacent sites are included with the ERIS environmental database ")
    canvas.drawString(54, 98,"report to be used for research purposes only and cannot be resold for any other commercial uses other than for use in a Phase I")
    canvas.drawString(54, 86,"environmental assessment.")
    canvas.saveState()
    del canvas

def goSummaryPage(summaryfile,data,years):
    doc = SimpleDocTemplate(summaryfile, pagesize = letter, topMargin=130,bottomMargin=123)
    Story = []

    newdata = []
    newdata.append(['Date','City','State','Volume','Sheet Number(s)'])
    style = ParagraphStyle("cover",parent=styles['Normal'],fontName="Helvetica",fontSize=9,leading=9)

    for key in years:
        for i in range(len(data[key])):
            [state,city,vol,year,sheets] =data[key][i][2:7]
            sheets = str([_ for _ in sheets]).replace("'","").replace("[","").replace("]","")
            newdata.append([Paragraph(('<para alignment="left">%s</para>')%(_), style) for _ in [year,city,state,vol, sheets]])

    table = Table(newdata,colWidths = [80,80,80,80, PAGE_WIDTH-420])
    table.setStyle([('FONT',(0,0),(4,0),'Helvetica-Bold'),
             ('VALIGN',(0,1),(-1,-1),'TOP'),
             ('ALIGN',(0,0),(4,0),'LEFT'),
             ('BOTTOMPADDING', [0,0], [-1, -1], 5),])
    Story.append(table)
    doc.build(Story, onFirstPage=myFirstSummaryPage, onLaterPages=myLaterSummaryPage)
    doc = None

def selectbyyear(selectionLayer):
    print '=================================================='
    print '########## in image site selection'
    years_s = []
    summaryList_s = []

    pathFieldName = "filepath"
    filenameFieldName = "filename"
    mxdFieldName = "mxd"

    filepathList_b = []     #mosaicked image path. Theoretically this list could be more than boundary path list, due to possibly boundary-overlapping multiple mosaicked images
    filenameList_b = []
    filepathList_m = []     #boundary path.
    filenameList_m = []

    i = 0
    # loop through the relevant records, locate the mosaiced layer(s)
    rows = arcpy.SearchCursor(selectionLayer)    # loop through the selected records
    for row in rows:
        filename = (row.getValue(filenameFieldName)).strip()
        filenameL = filename.lower()
        if filenameL[:3] == 'vol':   # the mosaicked file (vol***)
            apath = row.getValue(pathFieldName)
            bpath = apath.replace(r"W:\FIM_DATA_USA",r"\\cabcvan1fpr009\FIM_DATA_USA")
            filepathList_m.append(bpath)            #path list for mosaicked images
            filenameList_m.append(filename)
            i = i+1
        if 'image_boundary' in filenameL:
            apath = row.getValue(pathFieldName)
            bpath = apath.replace(r"W:\FIM_DATA_USA",r"\\cabcvan1fpr009\FIM_DATA_USA")
            filepathList_b.append(bpath)             #path list for boundary file
            filenameList_b.append(filename)

    del row
    del rows

    if len(filepathList_b) != len(filepathList_m):
        logger.info("order " + OrderIDText + ":        Data ERROR: different number image_boundary files and vol### files. Pay attention!!")
        print "Data Issue: different number image_boundary files and vol### files. Pay attention!" #could happen due to multiple mosaic, and site falls in gap
    
    print "VOLUME NAMES SELECTED: " + str(filenameList_m)
    print "VOLUMES SELECTED: " + str(filepathList_m)
    print "IMAGE BOUNDARIES SELECTED: " + str(filenameList_b)

    # one boundary file could correspond to one or more mosaicked image files

    for i in range(0, len(filepathList_b)):

        mosaickedimagepathList= []   #this is a clip and mosaic of images from the same volume
        mosaickedfilenameList = []

        boundarypath = filepathList_b[i]
        boundaryfilename = filenameList_b[i]
        print "___" + str(i) + ". " + str(boundarypath).split('\\')[5]
        # further check with the boundary file to see if intersected.

        logger.info("order " + OrderIDText + ":        boundarypath is " + boundarypath)
        # print "order " + OrderIDText + ":        boundarypath is " + boundarypath

        boundarySHP = os.path.join(boundarypath, boundaryfilename + '.shp')
        arcpy.MakeFeatureLayer_management(boundarySHP, "boundary_lyr")
        arcpy.SelectLayerByLocation_management("boundary_lyr","INTERSECT", orderGeometryPR,selectionDist)
        print "NUMBER OF FEATURES SELECTED: " + str(int((arcpy.GetCount_management("boundary_lyr").getOutput(0))))

        if(int((arcpy.GetCount_management("boundary_lyr").getOutput(0))) ==0): #GET TABLE ROW COUNT
            logger.info("order " + OrderIDText + ":            NO records selected in the boundary layer file in path " + boundarypath)    # this may eleminate the false positives
            print "order " + OrderIDText + ":            NO records selected in the boundary layer file in path " + boundarypath

        else:

            # find the same path in filepathList_b, to locate the corresponding vol### file(s)
            # note the image_boudary path is missing last "/" compared to other paths
            if boundarypath+"\\" not in filepathList_m:   #path is also path for the vol###, always at the same level. Note difference by an ending '/'

                logger.info("order " + OrderIDText + ":            ERROR: the boundary file doesn't have a matching vol### file!!! in " + boundarypath)
                raise Exception("order " + OrderIDText + ":            ERROR: the boundary file doesn't have a matching vol### file!!! in " + boundarypath)

            else:

                print "##### MXD addlayer and exporting #" + str(i) + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
                # need to consider multi-part images, their boundaries could overlap. cannot just randomly pick one
                # if both of the multi-part images are selected, it means both intersect with the buffer, even though the intersection could be blank.
                # may need to add both to mxd,
                # and clip (and mosaicked) both
                # particularly for huge custom orders
                # e.g. VOL3225_PART1, VOL3225_PART2, VOL3225_PART3
                # not all the 3 will be in the list, only whose boundary actually intersects with the buffer
                volumeNum = boundarypath.split("\\")[-1]

                for ind,x in enumerate(filepathList_m):
                    if x==boundarypath+"\\":
                        mosaickedimagepathList.append(x)
                        print "" + x + "" + filenameList_m[ind].replace(" ","")                    
                        imagefile = getImagefile(x,filenameList_m[ind])
                        mosaickedfilenameList.append(imagefile)

                book = open_workbook(excelfile)              #[vol123, los angeles 1983 vol 2, los angeles@@@CA@los angeles@1983@2, ... ]
                sheet = book.sheet_by_name('Sheet1')
                num_rows = sheet.nrows-1
                curr_row = -1

                city = ' '
                state = ' '
                year = ' '
                volseq = ' '
                while curr_row < num_rows:
                    curr_row += 1
                    row = sheet.row(curr_row)
                    volnum1 = str(sheet.cell_value(curr_row,0)).strip()     # e.g. vol112
                    duplicate = 0 if str(sheet.cell_value(curr_row,8)).strip() == '' else 1     # 1 if duplicate flag

                    if volnum1.lower() == volumeNum.lower(): #volumeNum is always like vol123
                        print "VOLUMENUM FOUND: " + volumeNum
                        volumeName_new = str(sheet.cell_value(curr_row,3)).strip()     #e.g. Long Beach 1914 - Sep 1963; Volume 3@@@California@Long Beach@1963@3
                        state = volumeName_new.split('@@@')[1].split('@')[0]          #e.g. California
                        city = volumeName_new.split('@@@')[1].split('@')[1]           #e.g. los angeles
                        year = int(volumeName_new.split('@@@')[1].split('@')[2].split('(')[0].strip())             #e.g. 1980
                        volseq = volumeName_new.split('@@@')[1].split('@')[3]         #e.g. volume 2 of a particular series
                        break
                    
                book.unload_sheet('Sheet1')

                # loop through relevant records in boundary file, locate the sheet numbers for this volume
                rows = arcpy.SearchCursor("boundary_lyr")
                sheetNos = []
                jpgfield = "IMAGE_NO"
                for row in rows:
                     sheetNos.append(str(row.getValue(jpgfield)).strip().lstrip('0'))
                del row
                del rows
                sheetNos.sort()

                sheetNos_noLetter = []
                for sheet in sheetNos:
                    sheet1 = sheet.split('_')[0].split('-')[0].strip('A').strip('B').strip('C').strip('D').strip('E').strip('F')
                    sheetNos_noLetter.append(sheet1)
                sheetNos_set = set(sheetNos_noLetter)
                sheetNos_noLetter = list(sheetNos_set)
                sheetNos_noLetter.sort()

                #t = (volumeNum, volumeName_new, state, city, volseq, year, sheetNos, mosaickedimagepathList, mosaickedfilenameList, nFIP, duplicate)    # note nFIP is an indicator of the individual map pdf
                t = (volumeNum, volumeName_new, state, city, volseq, year, sheetNos_noLetter, mosaickedimagepathList, mosaickedfilenameList, nFIP, duplicate)
                if duplicate == 0:
                    summaryList_s.append(t)       #only include non-duplicate, usable volumes
                    if year not in years_s:
                        years_s.append(year)
                del book

    print '########## finish image site selection'
    print '=================================================='

    return (summaryList_s, years_s)

def reorgbyyear(summaryList):
    dictsummary = {}    #with duplicate years (of overlapping coverage)
    dictsummary_dedup = {}     #without duplicate years (of overlapping coverage)

    for i in range(0,len(summaryList)):
        if summaryList[i][5] not in dictsummary.keys():
            dictsummary[summaryList[i][5]] = [summaryList[i]]    #the keys are arrays
            if summaryList[i][10] == 0:              #duplicate folder won't be on dedup list
                dictsummary_dedup[summaryList[i][5]] = [summaryList[i]]
            else:
                print "duplicate folder removed " + summaryList[i][0]
        else:
            dictsummary[summaryList[i][5]].append(summaryList[i])
            if summaryList[i][10] == 0:                                     #the same year, but supplementary
                if summaryList[i][5] not in dictsummary_dedup.keys():
                    dictsummary_dedup[summaryList[i][5]] = [summaryList[i]]
                else:
                    dictsummary_dedup[summaryList[i][5]].append(summaryList[i])
            else:
                print "duplicate folder removed " + summaryList[i][0]
    return (dictsummary, dictsummary_dedup)

def getImagefile(mosaicpath, mosaicname):
    imagefile = ''
    if os.path.exists(os.path.join(mosaicpath, mosaicname+'.jpg')):
        imagefile = mosaicname + '.jpg'
    elif os.path.exists(os.path.join(mosaicpath, mosaicname+'.jp2')):
        imagefile = mosaicname + '.jp2'
    elif os.path.exists(os.path.join(mosaicpath, mosaicname+'.tif')):
        imagefile = mosaicname + '.tif'
    else:
        print "mosaic path is " + mosaicpath
        print "mosaicname is " + mosaicname
        print ''+1
        #sys.exit("Unrecognized image file extension for ") + mosaicpath + mosaicname

    return imagefile

def centreFromPolygon(polygonSHP,coordinate_system):
    arcpy.AddField_management(polygonSHP, "xCentroid", "DOUBLE", 18, 11)
    arcpy.AddField_management(polygonSHP, "yCentroid", "DOUBLE", 18, 11)

    xExpression = '!SHAPE.CENTROID.X!'
    yExpression = '!SHAPE.CENTROID.Y!'

    arcpy.CalculateField_management(polygonSHP, "xCentroid", xExpression, "PYTHON_9.3")
    arcpy.CalculateField_management(polygonSHP, "yCentroid", yExpression, "PYTHON_9.3")

    in_rows = arcpy.SearchCursor(polygonSHP)
    outPointFileName = "polygonCentre.shp"
    centreSHP = os.path.join(scratch, outPointFileName)
    point1 = arcpy.Point()
    array1 = arcpy.Array()

    featureList = []
    arcpy.CreateFeatureclass_management(scratch, outPointFileName, "POINT", "", "DISABLED", "DISABLED", coordinate_system)
    cursor = arcpy.InsertCursor(centreSHP)
    feat = cursor.newRow()

    for in_row in in_rows:
        # Set X and Y for start and end points
        point1.X = in_row.xCentroid
        point1.Y = in_row.yCentroid
        array1.add(point1)

        centerpoint = arcpy.Multipoint(array1)
        array1.removeAll()
        featureList.append(centerpoint)
        feat.shape = point1
        cursor.insertRow(feat)
    del feat
    del in_rows
    del cursor
    del point1
    del array1
    arcpy.AddXY_management(centreSHP)
    return centreSHP

# -------------------------------------------------------------------------------------------------------------------------------------
##deployment variables
server_environment = 'test'
server_config_file = r"\\cabcvan1gis006\GISData\ERISServerConfig.ini"
server_config = server_loc_config(server_config_file,server_environment)
connectionString = 'eris_gis/gis295@cabcvan1ora006.glaciermedia.inc:1521/GMTESTC'
reportcheckFolder = server_config["reportcheck"]
viewerFolder = server_config["viewer"]
uploadlink =  server_config["viewer_upload"] + r"/ErisInt/BIPublisherPortal/Viewer.svc/FIMUpload?ordernumber="

##global variables
connectionPath = r"\\cabcvan1gis005\GISData\FIMS_USA"
masterlyr = os.path.join(connectionPath,"master\Master.shp")
shapefilepath = os.path.join(connectionPath,"master","mastershps")#g_ESRI_variable_15
excelfile = os.path.join(connectionPath,"master\MASTER_ALL_STATES.xlsx")
orderGeomlyrfile_point =  os.path.join(connectionPath,r"python\layer\SiteMaker.lyr")
orderGeomlyrfile_polyline =  os.path.join(connectionPath,r"python\layer\orderLine.lyr")
orderGeomlyrfile_polygon =  os.path.join(connectionPath,r"python\layer\orderPoly.lyr")

FIMmxdfile = os.path.join(connectionPath, r"python\mxd\FIMLayout.mxd")
##Covermxdfile = os.path.join(connectionPath, r"python\mxd\CoverPage.mxd")
Summarymxdfile = os.path.join(connectionPath, r"python\mxd\SummaryPage.mxd")
NRFmxdfile = os.path.join(connectionPath, r"python\mxd\NRF.mxd")
imagelyr =  os.path.join(connectionPath,r"python\layer\mosaic_jpg_255.lyr")
boundlyrfile =  os.path.join(connectionPath,r"python\layer\boundary.lyr")
logopath =  os.path.join(connectionPath,r"python\mxd\logos")
annot_poly =  os.path.join(connectionPath,r"python\mxd\annot_poly.pdf")
annot_line =  os.path.join(connectionPath,r"python\mxd\annot_line.pdf")
srGCS83 = arcpy.SpatialReference(4269)
srWGS84 = arcpy.SpatialReference(4326)
selectionDist = '150 FEET'
coverPic =  os.path.join(connectionPath,"python\ERIS_2018_ReportCover_Fire Insurance Maps_F.jpg")
secondPic =  os.path.join(connectionPath,"python\ERIS_2018_ReportCover_Second Page_F.jpg")

try:
    OrderIDText = r"" #arcpy.GetParameterAsText(0)
    OrderNumText = r"20200727135"
    BufsizeText ='0.17'#arcpy.GetParameterAsText(1) # '0.17'#
    yesBoundary = ""#arcpy.GetParameterAsText(2) #yes/no/fixed
    scratch = os.path.join(r"W:\Data Analysts\Alison\_GIS\FIM_SCRATCHY", OrderNumText)
    emgOrder= 'N'
    resolution = '50'
# -------------------------------------------------------------------------------------------------------------------------------------

    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        # GET ORDER_ID AND BOUNDARY FROM ORDER_NUM
        if OrderIDText == "":
            cur.execute("SELECT * FROM ERIS.FIM_AUDIT WHERE ORDER_ID IN (select order_id from orders where order_num = '" + str(OrderNumText) + "')")
            result = cur.fetchall()
            OrderIDText = str(result[0][0]).strip()
            yesBoundaryqry = str([row[3] for row in result if row[2]== "URL"][0])
            yesBoundary = re.search('(yesBoundary=)(\w+)(&)', yesBoundaryqry).group(2).strip()
            print("Order ID: " + OrderIDText)
            print("Yes Boundary: " + yesBoundary)

        cur.execute("select decode(c.company_id, 35, 'Y', 'N') is_emg from orders o, customer c where o.customer_id = c.customer_id and o.order_id=" + str(OrderIDText))
        t = cur.fetchone()
        emgOrder = t[0]

    finally:
        cur.close()
        con.close()

    is_newLogofile = 'N'

    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()
        newlogofile = cur.callfunc('ERIS_CUSTOMER.IsCustomLogo', str, (str(OrderIDText),))

        if newlogofile <> None:
            is_newLogofile = 'Y'
            if newlogofile =='RPS_RGB.gif':
                newlogofile='RPS.png'
            elif newlogofile == 'G2consulting.png':
                newlogofile = None
                is_newLogofile = 'N'

    finally:
        cur.close()
        con.close()

    is_aei = 'N'

    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        is_aei = cur.callfunc('ERIS_CUSTOMER.IsProductChron', str, (str(OrderIDText),))

    finally:
        cur.close()
        con.close()

    AddressText =''
    Sitename =''
    summaryList = []
    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        coverInfotext = json.loads(cur.callfunc('eris_gis.getCoverPageInfo', str, (str(OrderIDText),)))

        OrderNumText = str(coverInfotext["ORDER_NUM"])
        siteName =coverInfotext["SITE_NAME"]
        proNo = coverInfotext["PROJECT_NUM"]
        ProName = coverInfotext["COMPANY_NAME"]

        coverInfotext["ADDRESS"] = '%s\n%s %s %s'%(coverInfotext["ADDRESS"].replace("\t"," "),coverInfotext["CITY"],coverInfotext["PROVSTATE"],coverInfotext["POSTALZIP"])
        AddressText=coverInfotext["ADDRESS"].replace("\n"," ").replace("\t"," ")

        cur.execute("select geometry_type, geometry, radius_type  from eris_order_geometry where order_id =" + OrderIDText)
        t = cur.fetchone()
        OrderType = str(t[0])
        OrderCoord = eval(str(t[1]))
        RadiusType = str(t[2])

    except Exception,e:
        logger.error("Error to get flag from Oracle " + str(e))
        raise
    finally:
        cur.close()
        con.close()

    logger.info("order " + OrderIDText + " starting at: "+time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    #arcpy.env.workspace = r"C:\JIAN\FIPs\testData\temp"
    arcpy.env.overwriteOutput = True
    arcpy.env.OverWriteOutput = True

    deliverfolder = OrderNumText
    pdfreport_name = OrderNumText+"_US_FIM.pdf"
    if coverInfotext["COUNTRY"]=='MEX':
        pdfreport_name =  OrderNumText+"_MEX_FIM.pdf"
    pdfreport = os.path.join(scratch, pdfreport_name)

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
    arcpy.CalculateUTMZone_cartography(orderGeometry, "UTM")
    UT= arcpy.SearchCursor(orderGeometry)
    for row in UT:
        UTMvalue = str(row.getValue('UTM'))[41:43]
    del UT
    if UTMvalue[0]=='0':
        UTMvalue=' '+UTMvalue[1:]
    out_coordinate_system = arcpy.SpatialReference('NAD 1983 UTM Zone %sN'%UTMvalue)

    orderGeometryPR = os.path.join(scratch, "ordergeoNamePR.shp")

    arcpy.Project_management(orderGeometry, orderGeometryPR, out_coordinate_system)

    del point
    del array

    if OrderType.lower() == 'polygon' and float(BufsizeText)==0 :
        BufsizeText = "0.001"   # for polygon no buffer orders, buffer set to 1m, to avoid buffer clipping error
    elif float(BufsizeText) < 0.01:   # for site orders, usually has a radius of 0.001
        BufsizeText = "0.25"        # set the FIP search radius to 250m

    bufferDistance = BufsizeText + " KILOMETERS"
    outBufferSHP = os.path.join(scratch, "buffer.shp")
    logger.info("order " + OrderIDText + " OrderType and RadiusType are " + OrderType + " " + RadiusType)
    if OrderType.lower() == "polygon" and RadiusType.lower() == "centre":
        # polygon order, buffer from Centre instead of edge
        # completely change the order geometry to the center point
        centreSHP = centreFromPolygon(orderGeometryPR,arcpy.Describe(orderGeometryPR).spatialReference)
        if bufferDistance > 0.001:    # cause a fail for (polygon from center + no buffer) orders
            arcpy.Buffer_analysis(centreSHP, outBufferSHP, bufferDistance)
        logger.info("order " + OrderIDText + ' polygon centre: yes '+time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    else:
        arcpy.Buffer_analysis(orderGeometryPR, outBufferSHP, bufferDistance)
        logger.info("order " + OrderIDText + ' polygon centre: no '+time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))

    ##bufferDistance_e = str(float(BufsizeText)*15) + " KILOMETERS"
    ##if bufferDistance_e < 0.1:
    ##    bufferDistance_e = "0.5 KILOMETERS"    # for polygon edge orders, arbitrarily assign 500m, thsi will be used for mask clipping
    ##extentBufferSHP = os.path.join(scratch,"buffer_extent.shp")
    ##arcpy.Buffer_analysis(orderGeometryPR, extentBufferSHP, bufferDistance_e)

    arcpy.env.workspace = shapefilepath
    shplist = arcpy.ListFeatureClasses()
    outBufferSHP_GCS = os.path.join(scratch, "buffer_gcs.shp")
    arcpy.Project_management(outBufferSHP,outBufferSHP_GCS,srGCS83)
    desc = arcpy.Describe(outBufferSHP_GCS)
    extent = desc.extent
    xMax = extent.XMax
    xMin = extent.XMin
    yMax = extent.YMax
    yMin = extent.YMin

    selectedlist = []
    selectedFIPs = os.path.join(scratch, "selected.shp")
    if arcpy.Exists(selectedFIPs):
        arcpy.Delete_management(selectedFIPs)
    i = 0
    j = 0
    presentedlist = []
    presentedFIPs = os.path.join(scratch, "presented.shp")
    if arcpy.Exists(presentedFIPs):
        arcpy.Delete_management(presentedFIPs)

    for shp in shplist:
        desc = arcpy.Describe(shp)
        extent = desc.extent
        # print extent.XMax, extent.XMin, extent.YMax, extent.YMin
        if (xMax < extent.XMax and xMin > extent.XMin and yMax < extent.YMax and yMin > extent.YMin):        # algorithm optimization
            if len(shp) == 13 or len(shp) == 14:
                print "in extent: " + str(shp)
                shpLayer = arcpy.mapping.Layer(shp)

                arcpy.SelectLayerByLocation_management(shpLayer,'intersect', orderGeometryPR,selectionDist)
                if(int((arcpy.GetCount_management(shpLayer).getOutput(0))) >0):
                    i = i + 1
                    selected = "in_memory/selected" + str(i)
                    arcpy.CopyFeatures_management(shpLayer, selected)
                    logger.debug("i = "+str(i))
                    selectedlist.append(selected)

                arcpy.SelectLayerByLocation_management(shpLayer,'intersect', orderGeometryPR,bufferDistance, 'NEW_SELECTION')
                if(int((arcpy.GetCount_management(shpLayer).getOutput(0))) >0):
                    j = j + 1
                    presented = "in_memory/presented" + str(j)
                    arcpy.CopyFeatures_management(shpLayer, presented)
                    logger.debug("j = "+str(j))
                    presentedlist.append(presented)

    if i>0:
        logger.debug("right before merge")
        arcpy.Merge_management(selectedlist,selectedFIPs)
        arcpy.Merge_management(presentedlist,presentedFIPs)
        logger.debug("after merge")
        arcpy.Delete_management("in_memory")

    nFIP = 0
    years_s = []
    if(i ==0):
        print "NO records selected"

    else:
        years_s = []
        summaryList_s = []
        selectedLayer = arcpy.MakeFeatureLayer_management(selectedFIPs)
        (summaryList_s, years_s) = selectbyyear(selectedLayer)         #to get the selected years

        # need to clear selection on the layer
        presentedLayer = arcpy.MakeFeatureLayer_management(presentedFIPs)
        arcpy.SelectLayerByLocation_management(presentedLayer,'intersect', outBufferSHP, '','NEW_SELECTION')    #this is to work on a new selection

        pathFieldName = "filepath"
        filenameFieldName = "filename"
        mxdFieldName = "mxd"

        filepathList_b = []     #mosaicked image path. Theoretically this list could be more than boundary path list, due to possibly boundary-overlapping multiple mosaicked images
        filenameList_b = []
        filepathList_m = []     #boundary path.
        filenameList_m = []

        i = 0
        # loop through the relevant records, locate the mosaiced layer(s)
        #temp = arcpy.CopyFeatures_management(masterLayer, 'in_memory/temp1')
        rows = arcpy.SearchCursor(presentedLayer)    # loop through the selected records
        for row in rows:
            filename = (row.getValue(filenameFieldName)).strip()
            filenameL = filename.lower()
            # print "filenameL is " + filenameL
            if filenameL[:3] == 'vol':   # the mosaicked file (vol***)
                apath = row.getValue(pathFieldName)
                bpath = apath.replace("W:\FIM_DATA_USA",r"\\cabcvan1fpr009\FIM_DATA_USA")

                # print "append to _m list: " + bpath
                filepathList_m.append(bpath)            #path list for mosaicked images
                filenameList_m.append(filename)
                i = i+ 1
            if 'image_boundary' in filenameL:
                apath = row.getValue(pathFieldName)
                bpath = apath.replace("W:\FIM_DATA_USA",r"\\cabcvan1fpr009\FIM_DATA_USA")

                # print "append to _b list: " + bpath
                filepathList_b.append(bpath)             #path list for boundary file
                filenameList_b.append(filename)
        del row
        del rows

        if len(filepathList_b) != len(filepathList_m):
            logger.info("order " + OrderIDText + ":        Data ERROR: different number image_boundary files and vol### files. Pay attention!!")
            print "Data Issue: different number image_boundary files and vol### files. Pay attention!" #could happen due to multiple mosaic, and site falls in gap

        # one boundary file could correspond to one or more mosaicked image files

        for i in range(0, len(filepathList_b)):
            mosaickedimagepathList= []   #this is a clip and mosaic of images from the same volume
            mosaickedfilenameList = []

            boundarypath = filepathList_b[i]
            boundaryfilename = filenameList_b[i]
            # further check with the boundary file to see if intersected.

            logger.info("order " + OrderIDText + ":        boundarypath is " + boundarypath)
            print "order " + OrderIDText + ":        boundarypath is " + boundarypath

            boundarySHP = os.path.join(boundarypath, boundaryfilename + '.shp')
            arcpy.MakeFeatureLayer_management(boundarySHP, "boundary_lyr")
            arcpy.SelectLayerByLocation_management("boundary_lyr",'intersect', outBufferSHP)

            if(int((arcpy.GetCount_management("boundary_lyr").getOutput(0))) ==0):
                logger.info("order " + OrderIDText + ":            NO records selected in the boundary layer file in path " + boundarypath)    # this may eleminate the false positives
                print "order " + OrderIDText + ":            NO records selected in the boundary layer file in path " + boundarypath

            else:
                # find the same path in filepathList_b, to locate the corresponding vol### file(s)
                # note the image_boudary path is missing last "/" compared to other paths
                if boundarypath+"\\" not in filepathList_m:   #path is also path for the vol###, always at the same level. Note difference by an ending '/'
                    logger.info("order " + OrderIDText + ":            ERROR: the boundary file doesn't have a matching vol### file!!! in " + boundarypath)
                    #raise Exception("order " + OrderIDText + ":            ERROR: the boundary file doesn't have a matching vol### file!!! in " + boundarypath)

                else:
                    print "##### MXD addlayer and exporting #" + str(i) + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
                    # need to consider multi-part images, their boudnaries could overlap. cannot just randomly pick one
                    # if both of the multi-part images are selected, it means both intersect with the buffer, even though the intersection could be blank.
                    # may need to add both to mxd,
                    # and clip (and mosaicked) both
                    # particularly for huge custom orders
                    # e.g. VOL3225_PART1, VOL3225_PART2, VOL3225_PART3
                    # not all the 3 will be in the list, only whose boundary actually intersects with the buffer
                    volumeNum = boundarypath.split("\\")[-1]

                    for ind,x in enumerate(filepathList_m):
                        if x==boundarypath+"\\":
                            mosaickedimagepathList.append(x)
                            print "" + x + "" + filenameList_m[ind].replace(" ","")
                            imagefile = getImagefile(x,filenameList_m[ind].replace(" ",""))
                            mosaickedfilenameList.append(imagefile)

                    book = open_workbook(excelfile)              #[vol123, los angeles 1983 vol 2, los angeles@@@CA@los angeles@1983@2, ... ]
                    sheet = book.sheet_by_name('Sheet1')
                    num_rows = sheet.nrows-1
                    curr_row = -1

                    city = ' '
                    state = ' '
                    year = ' '
                    volseq = ' '
                    while curr_row < num_rows:
                        curr_row += 1
                        row = sheet.row(curr_row)
                        volnum1 = str(sheet.cell_value(curr_row,0)).strip()     # e.g. vol112
                        duplicate = 0 if str(sheet.cell_value(curr_row,8)).strip() == '' else 1     # 1 if duplicate flag

                        if volnum1.lower() == volumeNum.lower(): #volumeNum is always like vol123
                            volumeName_new = str(sheet.cell_value(curr_row,3)).strip()     #e.g. Long Beach 1914 - Sep 1963; Volume 3@@@California@Long Beach@1963@3
                            state = volumeName_new.split('@@@')[1].split('@')[0]          #e.g. California
                            city = volumeName_new.split('@@@')[1].split('@')[1]           #e.g. los angeles
                            year = int(volumeName_new.split('@@@')[1].split('@')[2].split('(')[0].strip())           #e.g. 1980
                            volseq = volumeName_new.split('@@@')[1].split('@')[3]         #e.g. volume 2 of a particular series

                            break

                    book.unload_sheet('Sheet1')

                    # loop through relevant records in boundary file, locate the sheet numbers for this volume
                    rows = arcpy.SearchCursor("boundary_lyr")
                    sheetNos = []
                    jpgfield = "IMAGE_NO"
                    for row in rows:
                         sheetNos.append(str(row.getValue(jpgfield)).strip().lstrip('0'))
                    del row
                    del rows
                    sheetNos.sort()

                    sheetNos_noLetter = []
                    for sheet in sheetNos:
                        sheet1 = sheet.split('_')[0].split('-')[0].strip('A').strip('B').strip('C').strip('D').strip('E').strip('F')
                        sheetNos_noLetter.append(sheet1)
                    sheetNos_set = set(sheetNos_noLetter)
                    sheetNos_noLetter = list(sheetNos_set)
                    sheetNos_noLetter.sort()

                    #t = (volumeNum, volumeName_new, state, city, volseq, year, sheetNos, mosaickedimagepathList, mosaickedfilenameList, nFIP, duplicate)    # note nFIP is an indicator of the individual map pdf
                    t = (volumeNum, volumeName_new, state, city, volseq, year, sheetNos_noLetter, mosaickedimagepathList, mosaickedfilenameList, nFIP, duplicate)
                    summaryList.append(t)
                    del book

    outBufferSHP = None
    masterLayer = None
    dictSummary = {}
    dictSummary_dedup = {}
    (dictSummary, dictSummary_dedup) = reorgbyyear(summaryList)
    yearlookup = {}
    #years = dictSummary_dedup.keys()
    #years.sort(reverse = True)
    if is_aei == 'Y':
        years_s.sort(reverse = False)
    else:
        years_s.sort(reverse = True)
    pdflist = []

    if OrderType.lower()== 'point':
        orderGeomlyrfile = orderGeomlyrfile_point
    elif OrderType.lower() =='polyline':
        orderGeomlyrfile = orderGeomlyrfile_polyline
    else:
        orderGeomlyrfile = orderGeomlyrfile_polygon

    orderGeomLayer = arcpy.mapping.Layer(orderGeomlyrfile)
    orderGeomLayer.replaceDataSource(scratch,"SHAPEFILE_WORKSPACE","orderGeometry")

    firstTime = True

    for year in years_s:

        mxdFIP = arcpy.mapping.MapDocument(FIMmxdfile)
        dfFIP = arcpy.mapping.ListDataFrames(mxdFIP,"main")[0]
        queryLayer = arcpy.mapping.ListLayers(mxdFIP,"Buffer Outline",dfFIP)[0]
        queryLayer.replaceDataSource(scratch, "SHAPEFILE_WORKSPACE", "buffer")  # note buffer.shp won't work

        dfinset = arcpy.mapping.ListDataFrames(mxdFIP,"inset")[0]

        #sheetNos= []
        sheetnoText = ''
        volumeNums = ''
        imageLayer = arcpy.mapping.Layer(imagelyr)
        items = dictSummary_dedup[year]
        for item in items:                                                       #multiple folders need to be mosaicked together
            if item[4] == '':
                sheetnoText = sheetnoText + 'Volume NA: '
            else:
                sheetnoText = sheetnoText + 'Volume ' + str(item[4]) + ': '
            for i in range(0,len(item[7])):
                imageLayer.replaceDataSource(item[7][i], "RASTER_WORKSPACE", item[8][i])      #one folder could have multipe mosaicked images
                if year not in yearlookup.keys():
                    yearlookup[year] = [os.path.join(item[7][i], item[8][i])]
                else:
                    yearlookup[year].append(os.path.join(item[7][i], item[8][i]))
                arcpy.mapping.AddLayer(dfFIP,imageLayer,"Bottom")

                #sheetNos.extend(item[6])  #here could use append and more processing to seperate sheets from multiple volumes
                sheetnoText = sheetnoText + ','.join(item[6]) + ';' + '\r\n'
                volumeNums = volumeNums + item[0]
                logger.info("order " + OrderIDText + ":                add to mxd the image: " + imageLayer.dataSource)

            boundLayer = arcpy.mapping.Layer(boundlyrfile)
            boundLayer.replaceDataSource(item[7][0],"SHAPEFILE_WORKSPACE","IMAGE_BOUNDARY")
            arcpy.mapping.AddLayer(dfinset,boundLayer,"Bottom")

        # refresh the view to reflect the updated image
        # center and scale the image
        spatialRef = out_coordinate_system#arcpy.SpatialReference(out_coordinate_system)
        dfFIP.spatialReference = spatialRef
        dfFIP.extent = queryLayer.getSelectedExtent(False)
        scale = dfFIP.scale * 1.1
        dfFIP.scale = ((int(scale)/100)+1)*100

        # update map document with sheet numbers, orderID and address
        orderIDTextE = arcpy.mapping.ListLayoutElements(mxdFIP, "TEXT_ELEMENT", "MainTitleText")[0]
        orderIDTextE.text = str(year)
        mapsheetTextE = arcpy.mapping.ListLayoutElements(mxdFIP, "TEXT_ELEMENT", "mapsheetText")[0]
        mapsheetTextE.text = "Map sheet(s): " + '\r\n' + sheetnoText
        mapsheetTextE.elementPositionX = 2.0833
        mapsheetTextE.elementPositionY = 1.1195
        orderIDTextE = arcpy.mapping.ListLayoutElements(mxdFIP, "TEXT_ELEMENT", "ordernumText")[0]
        orderIDTextE.text = "Order Number " + OrderNumText
        AddressTextE= arcpy.mapping.ListLayoutElements(mxdFIP, "TEXT_ELEMENT", "AddressText")[0]
        AddressTextE.text = "Address: " + AddressText + ""
        AddressTextE.elementPositionX = 0.5833
        AddressTextE.elementPositionY = 1.2194
        #searchradiusTextE= arcpy.mapping.ListLayoutElements(mxdFIP, "TEXT_ELEMENT", "searchradiusText")[0]
        #searchradiusTextE.text = "The dashed line indicates the search radius around the site: " + str(float(searchRadius)/1000) + " miles"
        #searchradiusTextE.elementPositionX = 0.5972
        #searchradiusTextE.elementPositionY = 0.2371

##        if OrderType.lower() == "polyline" or OrderType.lower() == "polygon":
        if yesBoundary.lower() == 'fixed':
            arcpy.mapping.AddLayer(dfFIP,orderGeomLayer,"Top")

        if is_newLogofile == 'Y' and emgOrder == 'N':
            logoE = arcpy.mapping.ListLayoutElements(mxdFIP, "PICTURE_ELEMENT", "logo")[0]
            logoE.sourceImage = os.path.join(logopath, newlogofile)

        arcpy.RefreshTOC()
        mxdFIP.saveACopy(os.path.join(scratch, "test_"+volumeNums+"_"+str(year)+".mxd"))
        # print the map pdf
        FIPpdf = os.path.join(scratch, 'FIPExport_'+volumeNums+"_"+str(year)+'.pdf')
        # note by using "Page_layout", the two document size parameters are not effective
        arcpy.mapping.ExportToPDF(mxdFIP, FIPpdf, "PAGE_LAYOUT", 0, 0,resolution, "BEST", "RGB", True, "ADAPTIVE", "RASTERIZE_BITMAP", False, True, "NONE", True, 85)

        # note by using export dataframe (instead of "page_layout") below, the map quality does look better. But the problem is with the size of the page and non-appearing layout elements
        # arcpy.mapping.ExportToPDF(mxdFIP, FIPpdf, dfFIP, 6800, 8800, 400, "BEST", "RGB", True, "ADAPTIVE", "RASTERIZE_BITMAP", False, True, "LAYERS_AND_ATTRIBUTES", True, 90)
        if ( yesBoundary.lower() == 'yes' and (OrderType.lower() == "polyline" or OrderType.lower() == "polygon")):
            if firstTime:
                #remove all other layers
                scale2use = dfFIP.scale
                for lyr in arcpy.mapping.ListLayers(mxdFIP, "", dfFIP):
                    arcpy.mapping.RemoveLayer(dfFIP, lyr)
                arcpy.mapping.AddLayer(dfFIP,orderGeomLayer,"Top") #the layer is visible
                dfFIP.scale = scale2use
                shapePdf = os.path.join(scratch, 'shape.pdf')
                arcpy.mapping.ExportToPDF(mxdFIP, shapePdf, "PAGE_LAYOUT", 0, 0, 150, "BEST", "RGB", True, "ADAPTIVE", "RASTERIZE_BITMAP", False, True, "NONE", True, 85)
                #create the a pdf with annotation just once
                myAnnotPdf = createAnnotPdf(OrderType, shapePdf)
                firstTime = False

            #merge annotation pdf to the map
            FIPpdf = annotatePdf(FIPpdf, myAnnotPdf)
            arcpy.AddMessage("FIPpdf is " + FIPpdf)
        pdflist.append(FIPpdf)
        arcpy.Delete_management("in_memory")
        FIPpdf = None
        del queryLayer
        del imageLayer
        del dfFIP
        del mxdFIP
    NRF='N'
    pagesize = portrait(letter)
    [PAGE_WIDTH,PAGE_HEIGHT]=pagesize[:2]
    PAGE_WIDTH = int(PAGE_WIDTH)
    PAGE_HEIGHT = int(PAGE_HEIGHT)
    styles = getSampleStyleSheet()

    summaryfile=os.path.join(scratch,"summary.pdf")
    coverfile = os.path.join(scratch,"cover.pdf")
##
    if len(years_s) ==0 :
        logger.info("order " + OrderIDText + ":    search completed. Will print out a NRF letter. ")
        NRF='Y'
        goCoverPage(coverfile)
        os.rename(coverfile,pdfreport)
    else:
        goSummaryPage(summaryfile,dictSummary_dedup,years_s)
        goCoverPage(coverfile)

        output = PdfFileWriter()
        cover = open(coverfile,'rb')
        output.addPage(PdfFileReader(cover).getPage(0))
        output.addBookmark("Cover Page",0)

        summary = open(summaryfile,'rb')
        for j in range(PdfFileReader(summary).getNumPages()):
            output.addPage(PdfFileReader(summary).getPage(j))
            output.addBookmark("Summary Page",j+1)

        for i in range(0, len(years_s)):
            pdf = pdflist[i]
            page = open(pdf,'rb')
            output.addPage(PdfFileReader(page).getPage(0))
            output.addBookmark(str(years_s[i]), i+j+2)

        outputStream = open(pdfreport,"wb")
        output.setPageMode('/UseOutlines')
        output.write(outputStream)
        outputStream.close()
        cover.close()
        page.close()
        summary.close()
        output = None
        summaryfile = None

    needViewer = 'N'
    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        cur.execute("select fim_viewer from order_viewer where order_id =" + str(OrderIDText))
        t = cur.fetchone()
        if t != None:
            needViewer = t[0]

    finally:
        cur.close()
        con.close()

    if needViewer == 'Y':
        metadata = []
        srGoogle = arcpy.SpatialReference(3857)
        arcpy.AddMessage("Viewer is needed")#, multipage. Need to copy data to obi002")
        viewerdir = os.path.join(scratch,OrderNumText+'_fim')
        if not os.path.exists(viewerdir):
            os.mkdir(viewerdir)
        tempdir = os.path.join(scratch,'viewertemp')
        if not os.path.exists(tempdir):
            os.mkdir(tempdir)
        #need to reorganize deliver directory

        arcpy.env.outputCoordinateSystem = srGoogle
        #to do: get the right year for each FIM
        for year in years_s:
            mxdname = glob.glob(os.path.join(scratch,'test*'+str(year)+'.mxd'))[0]
            mxd = arcpy.mapping.MapDocument(mxdname)
            df = arcpy.mapping.ListDataFrames(mxd,"main")[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
            df.spatialReference = srGoogle

            queryLayer = arcpy.mapping.ListLayers(mxd,"Buffer Outline",df)[0]
            df.extent = queryLayer.getSelectedExtent(False)

            imagename = str(year)+".jpg"
            #arcpy.mapping.ExportToJPEG(mxd, os.path.join(scratch, viewerdir, imagename), df,df_export_width= 14290,df_export_height=16000, color_mode='8-BIT_GRAYSCALE',world_file = True) #by default, the jpeg quality is 100
            arcpy.mapping.ExportToJPEG(mxd, os.path.join(scratch, viewerdir, imagename), df,5800,6100, color_mode='8-BIT_GRAYSCALE',world_file = True, jpeg_quality=50)

            desc = arcpy.Describe(os.path.join(viewerdir, imagename))
            featbound = arcpy.Polygon(arcpy.Array([desc.extent.lowerLeft, desc.extent.lowerRight, desc.extent.upperRight, desc.extent.upperLeft]),
                                srGoogle)
            del desc

            tempfeat = os.path.join(tempdir, "tilebnd_"+str(year)+ ".shp")
            arcpy.Project_management(featbound, tempfeat, srWGS84) #function requires output not be in_memory
            del featbound
            desc = arcpy.Describe(tempfeat)

            metaitem = {}
            metaitem['type'] = 'fim'
            metaitem['imagename'] = imagename[:-4]+'.jpg'

            metaitem['lat_sw'] = desc.extent.YMin
            metaitem['long_sw'] = desc.extent.XMin
            metaitem['lat_ne'] = desc.extent.YMax
            metaitem['long_ne'] = desc.extent.XMax

            metadata.append(metaitem)
            del mxd, df

        arcpy.env.outputCoordinateSystem = None

        if os.path.exists(os.path.join(viewerFolder, OrderNumText+"_fim")):
            shutil.rmtree(os.path.join(viewerFolder, OrderNumText+"_fim"))
        shutil.copytree(os.path.join(scratch, OrderNumText+"_fim"), os.path.join(viewerFolder, OrderNumText+"_fim"))
        url = uploadlink + OrderNumText
        urllib.urlopen(url)

        try:
            con = cx_Oracle.connect(connectionString)
            cur = con.cursor()

            cur.execute("delete from overlay_image_info where  order_id = %s and (type = 'fim')" % str(OrderIDText))

            for item in metadata:
                cur.execute("insert into overlay_image_info values (%s, %s, %s, %.5f, %.5f, %.5f, %.5f, %s, '', '')" % (str(OrderIDText), str(OrderNumText), "'" + item['type']+"'", item['lat_sw'], item['long_sw'], item['lat_ne'], item['long_ne'],"'"+item['imagename']+"'" ) )
            con.commit()
        finally:
            cur.close()
            con.close()

    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()
        cur.callproc('eris_fim.processFim', (int(OrderIDText),))
    finally:
        cur.close()
        con.close()
    
    if os.path.exists(os.path.join(reportcheckFolder,"FIM", pdfreport_name)):
        os.remove(os.path.join(reportcheckFolder,"FIM", pdfreport_name))
    shutil.copyfile(pdfreport,os.path.join(reportcheckFolder,"FIM", pdfreport_name))
    arcpy.SetParameterAsText(3, pdfreport)
    logger.info("order " + OrderIDText + ": pdf exported, with " + str(nFIP) + " maps"+time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    logger.removeHandler(handler)
    handler.close()

except:
    # Get the traceback object
    tb = sys.exc_info()[2]
    tbinfo = traceback.format_tb(tb)[0]
    pymsg = "Order ID: %s PYTHON ERRORS:\nTraceback info:\n"%OrderIDText + tbinfo + "\nError Info:\n" + str(sys.exc_info()[1])

    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()
        cur.callproc('eris_fim.InsertFIMAudit', (OrderIDText, 'python-Error Handling',pymsg))
    finally:
        cur.close()
        con.close()
    raise    #raise the error again

print ("Final FIM report directory: " + (str(os.path.join(reportcheckFolder,"FIM", pdfreport_name))))
print ("__________DONE")