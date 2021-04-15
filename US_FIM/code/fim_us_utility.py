import logging
import time
import json
import arcpy
import os
import glob
import urllib
import shutil
import re
import sys
import traceback
import textwrap
import cx_Oracle
import fim_us_config as cfg

from xlrd import open_workbook
from PyPDF2 import PdfFileReader,PdfFileWriter
from PyPDF2.generic import NameObject, createStringObject, ArrayObject, FloatObject
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Frame, Table
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import portrait, letter
from reportlab.pdfgen import canvas
from time import strftime

class oracle(object):    
    def __init__(self, connectionString):
        self.connectionString = connectionString
        try:
            self.con = cx_Oracle.connect(self.connectionString)
            self.cur = self.con.cursor()
        except cx_Oracle.Error as e:
            arcpy.AddError(e)
            arcpy.AddError("### oracle connection failed.")
    
    def query(self, expression):
        try:
            self.cur.execute(expression)
            t = self.cur.fetchall()
            return t
        except cx_Oracle.DatabaseError as e:
            arcpy.AddError(e)
            arcpy.AddError("### database error.")

    def exe(self, expression):
        try:
            self.cur.execute(expression)
            self.con.commit()
        except cx_Oracle.DatabaseError as e:
            arcpy.AddError(e)
            arcpy.AddError("### database error.")

    def proc(self, procedure, args):
        try:
            t = self.cur.callproc(procedure, args)
            return t
        except cx_Oracle.DatabaseError as e:
            arcpy.AddError(e)
            arcpy.AddError("### database error.")

    def func(self, function, type, args):
        try:
            t = self.cur.callfunc(function, type, args)
            return t
        except cx_Oracle.DatabaseError as e:
            arcpy.AddError(e)
            arcpy.AddError("### database error.")

    def close(self):
        try:
            self.cur.close()
            self.con.close()
        except cx_Oracle.Error as e:
            arcpy.AddError(e)
            arcpy.AddError("### oracle failed to close.")

class fim_us_rpt(object):
    def __init__(self,order_obj, oracle):
        self.order_obj = order_obj
        self.oracle = oracle
    
    def findpath(a,b):
        a = a.upper().replace(r"\\CABCVAN1FPR009", r"W:")
        path= lambda a,b:os.path.join(a[0:[m.start() for m in re.finditer(r'\\', a)][2]].replace(r"W:",r"\\CABCVAN1FPR009"),b)
        return path(a,b)

    def createAnnotPdf(self, myShapePdf):
        # input variables
        # self.order_obj.geometry.type = 'POLYLINE'      # or POLYGON

        # part 1: read geometry pdf to get the vertices and rectangle to use
        source  = PdfFileReader(open(myShapePdf,'rb'))
        geomPage = source.getPage(0)
        mystr = geomPage.getObject()['/Contents'].getData()
        # to pinpoint the string part: 1.19997 791.75999 m 1.19997 0.19466 l 611.98627 0.19466 l 611.98627 791.75999 l 1.19997 791.75999 l
        # the format seems to follow x1 y1 m x2 y2 l x3 y3 l x4 y4 l x5 y5 l
        geomString = mystr.split('S\r\n')[0].split('M\r\n')[1].replace("rg\r\n", "").replace("h\r\n", "")
        coordsString = [value for value in geomString.split(' ') if value not in ['m','l','']]

        # part 2: update geometry in the map
        if self.order_obj.geometry.type.upper() == 'POLYGON':
            pdf_geom = PdfFileReader(open(cfg.annot_poly,'rb'))
        elif self.order_obj.geometry.type.upper() == 'POLYLINE':
            pdf_geom = PdfFileReader(open(cfg.annot_line,'rb'))
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

        # below rect seems to be geom bounding box coordinates: xmin, ymin, xmax,ymax
        annot.getObject().update({NameObject('/Rect'):ArrayObject([FloatObject(min(xcoords)), FloatObject(min(ycoords)), FloatObject(max(xcoords)), FloatObject(max(ycoords))])})
        annot.getObject().pop('/AP')  # this is to get rid of the ghost shape

        annot.getObject().update({NameObject('/T'):createStringObject(u'ERIS')})

        output = PdfFileWriter()
        output.addPage(page_geom)
        annotPdf = os.path.join(cfg.scratch, "annot.pdf")
        outputStream = open(annotPdf,"wb")
        # output.setPageMode('/UseOutlines')
        output.write(outputStream)
        outputStream.close()
        output = None
        return annotPdf

    def annotatePdf(self, mapPdf, myAnnotPdf):
        pdf = PdfFileReader(open(mapPdf,'rb'))
        FIMpage = pdf.getPage(0)

        pdf_intermediate = PdfFileReader(open(myAnnotPdf,'rb'))
        page= pdf_intermediate.getPage(0)
        page.mergePage(FIMpage)

        output = PdfFileWriter()
        output.addPage(page)

        annotatedPdf = mapPdf[:-4]+'_a.pdf'
        outputStream = open(annotatedPdf,"wb")
        # output.setPageMode('/UseOutlines')
        output.write(outputStream)
        outputStream.close()
        output = None
        return annotatedPdf

    def goCoverPage(self, coverPdf, nrf):
        pagesize = portrait(letter)
        [PAGE_WIDTH,PAGE_HEIGHT]=pagesize[:2]
        PAGE_WIDTH = int(PAGE_WIDTH)
        PAGE_HEIGHT = int(PAGE_HEIGHT)

        AddressText = '%s\n%s %s %s'%(self.order_obj.address, self.order_obj.city, self.order_obj.province, self.order_obj.postal_code)

        c = canvas.Canvas(coverPdf,pagesize = portrait(letter))
        c.drawImage(cfg.coverPic,0,0, PAGE_WIDTH,PAGE_HEIGHT)

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
        c.drawString(rightsw,heights-0*space, self.order_obj.site_name)
        c.drawString(rightsw, heights-1*space,AddressText.split("\n")[0])
        c.drawString(rightsw, heights-2*space,AddressText.split("\n")[1])
        c.drawString(rightsw, heights-3*space,self.order_obj.project_num)
        c.drawString(rightsw, heights-4*space,self.order_obj.company_desc)
        c.drawString(rightsw, heights-5*space,self.order_obj.number)
        c.drawString(rightsw, heights-6*space,time.strftime('%B %d, %Y', time.localtime()))
        if nrf == 'Y':
            c.setStrokeColorRGB(0.67,0.8,0.4)
            c.line(50,180,PAGE_WIDTH-60,180)
            c.setFont('Helvetica-Bold', 12)
            c.drawString(70,160,"Please note that no information was found for your site or adjacent properties.")
        p=None
        Disclaimer = None
        style = None
        c.showPage()
        c.save()

    def myFirstSummaryPage(self, canvas,doc):
        pagesize = portrait(letter)
        [PAGE_WIDTH,PAGE_HEIGHT]=pagesize[:2]
        PAGE_WIDTH = int(PAGE_WIDTH)
        PAGE_HEIGHT = int(PAGE_HEIGHT)

        canvas.saveState()
        canvas.setFont('Helvetica', 9)
        canvas.drawImage(cfg.secondPic,0,0, PAGE_WIDTH,PAGE_HEIGHT)
        canvas.drawString(54, 690, "Listed below, please find the results of our search for historic fire insurance maps from our in-house collection, performed in")
        canvas.drawString(54, 678,"conjuction with your ERIS report.")
        canvas.drawString(54, 110, "Individual Fire Insurance Maps for the subject property and/or adjacent sites are included with the ERIS environmental database ")
        canvas.drawString(54, 98,"report to be used for research purposes only and cannot be resold for any other commercial uses other than for use in a Phase I")
        canvas.drawString(54, 86,"environmental assessment.")
        canvas.restoreState()
        del canvas

    def myLaterSummaryPage(self, canvas,doc):
        pagesize = portrait(letter)
        [PAGE_WIDTH,PAGE_HEIGHT]=pagesize[:2]
        PAGE_WIDTH = int(PAGE_WIDTH)
        PAGE_HEIGHT = int(PAGE_HEIGHT)

        canvas.saveState()
        canvas.drawImage(cfg.secondPic,0,0, PAGE_WIDTH,PAGE_HEIGHT)
        canvas.setFont('Helvetica', 9)
        canvas.drawString(56, 690, "continued")
        canvas.drawString(54, 110, "Individual Fire Insurance Maps for the subject property and/or adjacent sites are included with the ERIS environmental database ")
        canvas.drawString(54, 98,"report to be used for research purposes only and cannot be resold for any other commercial uses other than for use in a Phase I")
        canvas.drawString(54, 86,"environmental assessment.")
        canvas.saveState()
        del canvas

    def goSummaryPage(self, summaryfile, mainList):
        pagesize = portrait(letter)
        [PAGE_WIDTH,PAGE_HEIGHT]=pagesize[:2]
        PAGE_WIDTH = int(PAGE_WIDTH)
        PAGE_HEIGHT = int(PAGE_HEIGHT)
        styles = getSampleStyleSheet()

        doc = SimpleDocTemplate(summaryfile, pagesize = letter, topMargin=130,bottomMargin=123)
        Story = []

        newdata = []
        newdata.append(['Date','City','State','Volume','Sheet Number(s)'])
        style = ParagraphStyle("cover",parent=styles['Normal'],fontName="Helvetica",fontSize=9,leading=9)

        years = mainList.keys()        
        if self.is_aei == 'Y':
            years.sort(reverse = False)
        else:
            years.sort(reverse = True)

        for key in years:
            volumes = mainList[key]
            for v in volumes.values():
                [state,city,vol,year,sheets] = v[2:7]

                if vol == 'None':
                    vol = ""

                sheetnoText = ""
                sheets.sort()
                for sht in sheets:
                    sht = sht.lstrip("0").split('_')[0].split('-')[0].strip('A').strip('B').strip('C').strip('D').strip('E').strip('F') 
                    sheetnoText = sheetnoText + sht + ', '
                sheetnoText = sheetnoText.rstrip(", ")

                newdata.append([Paragraph(('<para alignment="left">%s</para>')%(_), style) for _ in [year,city,state,vol, sheetnoText]])
        table = Table(newdata,colWidths = [80,80,80,80, PAGE_WIDTH-420])
        table.setStyle([('FONT',(0,0),(4,0),'Helvetica-Bold'),
                ('VALIGN',(0,1),(-1,-1),'TOP'),
                ('ALIGN',(0,0),(4,0),'LEFT'),
                ('BOTTOMPADDING', [0,0], [-1, -1], 5),])
        Story.append(table)
        doc.build(Story, onFirstPage=self.myFirstSummaryPage, onLaterPages=self.myLaterSummaryPage)
        doc = None

    def centreFromPolygon(self, polygonSHP,coordinate_system):
        arcpy.AddField_management(polygonSHP, "xCentroid", "DOUBLE", 18, 11)
        arcpy.AddField_management(polygonSHP, "yCentroid", "DOUBLE", 18, 11)

        xExpression = '!SHAPE.CENTROID.X!'
        yExpression = '!SHAPE.CENTROID.Y!'

        arcpy.CalculateField_management(polygonSHP, "xCentroid", xExpression, "PYTHON_9.3")
        arcpy.CalculateField_management(polygonSHP, "yCentroid", yExpression, "PYTHON_9.3")

        in_rows = arcpy.SearchCursor(polygonSHP)
        outPointFileName = "polygonCentre"
        centreSHP = os.path.join(cfg.scratch, cfg.scratchgdb, outPointFileName)
        point1 = arcpy.Point()
        array1 = arcpy.Array()

        featureList = []
        arcpy.CreateFeatureclass_management(os.path.join(cfg.scratch, cfg.scratchgdb), outPointFileName, "POINT", "", "DISABLED", "DISABLED", coordinate_system)
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

    def projlist(self, order_obj):
        self.srGCS83 = arcpy.SpatialReference(4269)     # GCS_North_American_1983
        self.srWGS84 = arcpy.SpatialReference(4326)     # GCS_WGS_1984
        self.srGoogle = arcpy.SpatialReference(3857)    # WGS_1984_Web_Mercator_Auxiliary_Sphere
        self.srUTM = order_obj.spatial_ref_pcs
        arcpy.AddMessage(self.srUTM.name)
        return self.srGCS83, self.srWGS84, self.srGoogle, self.srUTM

    def customrpt(self, order_obj):
        self.is_newLogo = 'N'
        try:
            function = 'ERIS_CUSTOMER.IsCustomLogo'
            self.newlogofile = self.oracle.func(function, str, (str(order_obj.id),))

            if self.newlogofile != None:
                self.is_newLogo = 'Y'
                if self.newlogofile =='RPS_RGB.gif':
                    self.newlogofile='RPS.png'
                elif self.newlogofile == 'G2consulting.png':
                    self.newlogofile = None
                    self.is_newLogo = 'N'
        except Exception as e:
            arcpy.AddMessage(e)
            arcpy.AddMessage("### ERIS_CUSTOMER.IsCustomLogo failed...")
            pass
        arcpy.AddMessage("is_newLogo = " + self.is_newLogo)

        self.is_aei = 'N'
        try:
            function = 'ERIS_CUSTOMER.IsProductChron'
            self.is_aei = self.oracle.func(function, str, (str(order_obj.id),))
        except Exception as e:
            arcpy.AddMessage(e)
            arcpy.AddMessage("### ERIS_CUSTOMER.IsProductChron failed...")
            pass
        arcpy.AddMessage("is_aei = " + self.is_aei)        
        
        self.is_emg= 'N'
        try:
            statement = "select decode(c.company_id, 35, 'Y', 'N') is_emg from orders o, customer c where o.customer_id = c.customer_id and o.order_id=" + str(order_obj.id)
            self.is_emg = self.oracle.query(statement)[0][0]
        except Exception as e:
            arcpy.AddMessage(e)
            arcpy.AddMessage("### ERIS_CUSTOMER.IsProductChron failed...")
            pass
        arcpy.AddMessage("is_emg = " + self.is_emg)

        return self.is_newLogo, self.is_aei, self.is_emg

    def createOrderGeometry(self, order_obj, projection, buffsize):
        arcpy.env.overwriteOutput = True

        point = arcpy.Point()
        array = arcpy.Array()
        sr = arcpy.SpatialReference()
        sr.factoryCode = 4269  # requires input geometry is in 4269
        sr.XYTolerance = .00000001
        sr.scaleFactor = 2000
        sr.create()
        featureList = []
        for feature in json.loads(order_obj.geometry.JSON).values()[0]:     # order coordinates
            # For each coordinate pair, set the x,y properties and add to the Array object.
            for coordPair in feature:
                try:
                    point.X = coordPair[0]
                    point.Y = coordPair[1]
                except:
                    point.X = feature[0]
                    point.Y = feature[1]
                sr.setDomain (point.X, point.X, point.Y, point.Y)
                array.add(point)
            if order_obj.geometry.type.lower() == 'point' or order_obj.geometry.type.lower() == 'multipoint':
                feat = arcpy.Multipoint(array, sr)
            elif order_obj.geometry.type.lower() =='polyline':
                feat  = arcpy.Polyline(array, sr)
            else :
                feat = arcpy.Polygon(array,sr)
            array.removeAll()

            # Append to the list of Polygon objects
            featureList.append(feat)

        arcpy.CopyFeatures_management(featureList, cfg.orderGeometry)        
        arcpy.Project_management(cfg.orderGeometry, cfg.orderGeometryPR, projection)

        del point
        del array

        # create buffer
        if order_obj.geometry.type.lower() == 'polygon' and float(buffsize) == 0 :
            buffsize = "0.001"           # for polygon no buffer orders, buffer set to 1m, to avoid buffer clipping error
        elif float(buffsize) < 0.01:     # for site orders, usually has a radius of 0.001
            buffsize = "0.25"            # set the FIP search radius to 250m

        bufferDistance = buffsize + " KILOMETERS"    
        if order_obj.geometry.type.lower() == "polygon" and order_obj.radius_type.lower() == "centre":
            # polygon order, buffer from Centre instead of edge
            # completely change the order geometry to the center point
            centreSHP = self.centreFromPolygon(cfg.orderGeometryPR,arcpy.Describe(cfg.orderGeometryPR).spatialReference)
            if bufferDistance > 0.001:    # cause a fail for (polygon from center + no buffer) orders
                arcpy.Buffer_analysis(centreSHP, cfg.outBufferSHP, bufferDistance)
        else:
            arcpy.Buffer_analysis(cfg.orderGeometryPR, cfg.outBufferSHP, bufferDistance)            

    def mapDocument(self, projection):
        arcpy.env.overwriteOutput = True
        arcpy.env.outputCoordinateSystem = projection

        mxd = arcpy.mapping.MapDocument(cfg.FIMmxdfile)
        dfmain = arcpy.mapping.ListDataFrames(mxd,"main")[0]
        dfinset = arcpy.mapping.ListDataFrames(mxd,"inset")[0]
        dfmain.spatialReference = projection
        dfinset.spatialReference = projection

        if self.order_obj.geometry.type.lower() == 'point' or self.order_obj.geometry.type.lower() == 'multipoint':
            orderGeomlyrfile = cfg.orderGeomlyrfile_point
        elif self.order_obj.geometry.type.lower() =='polyline':
            orderGeomlyrfile = cfg.orderGeomlyrfile_polyline
        else:
            orderGeomlyrfile = cfg.orderGeomlyrfile_polygon
            
        orderGeomLayer = arcpy.mapping.Layer(orderGeomlyrfile)
        orderGeomLayer.replaceDataSource(os.path.join(cfg.scratch, cfg.scratchgdb),"FILEGDB_WORKSPACE","orderGeometryPR")
        arcpy.mapping.AddLayer(dfmain,orderGeomLayer,"TOP")

        buffer = arcpy.mapping.ListLayers(mxd,"Buffer Outline",dfmain)[0]
        buffer.replaceDataSource(os.path.join(cfg.scratch, cfg.scratchgdb),"FILEGDB_WORKSPACE", "buffer")

        return mxd, dfmain, dfinset

    def mapExtent(self, mxd, dfmain, dfinset, multipage):
        # note buffer.shp won't work
        dfmain.extent = arcpy.mapping.ListLayers(mxd,"Buffer Outline",dfmain)[0].getSelectedExtent(False)
        arcpy.RefreshActiveView()
        if multipage == True:
            scale = dfmain.scale * 1.75
        else:
            scale = dfmain.scale * 1.1
        dfmain.scale = ((int(scale)/100)+1)*100

        # create df extent polygon
        extent = dfmain.extent
        XMAX = extent.XMax
        XMIN = extent.XMin
        YMAX = extent.YMax
        YMIN = extent.YMin
        pnt1 = arcpy.Point(XMIN, YMIN)
        pnt2 = arcpy.Point(XMIN, YMAX)
        pnt3 = arcpy.Point(XMAX, YMAX)
        pnt4 = arcpy.Point(XMAX, YMIN)
        array = arcpy.Array()
        array.add(pnt1)
        array.add(pnt2)
        array.add(pnt3)
        array.add(pnt4)
        array.add(pnt1)
        polygon = arcpy.Polygon(array)
        arcpy.CopyFeatures_management(polygon, cfg.extent)

    def setBoundary(self, mxd, df, yesboundary):
        # get yesboundary flag
        arcpy.AddMessage("yesboundary = " + yesboundary)

        if yesboundary.lower() == 'fixed':
            for lyr in arcpy.mapping.ListLayers(mxd, "", df):
                if lyr.name == "Project Property":
                    lyr.visible = True
                else:
                    lyr.visible = False

        elif yesboundary.lower() == 'yes':
            if not os.path.exists(cfg.shapePdf) and not os.path.exists(cfg.annotPdf):
                if self.order_obj.geometry.type.lower() == "polyline" or self.order_obj.geometry.type.lower() == "polygon":
                    for lyr in arcpy.mapping.ListLayers(mxd, "", df):
                        if lyr.name == "Project Property":
                            lyr.visible = True
                        else:
                            lyr.visible = False
                    arcpy.mapping.ExportToPDF(mxd, cfg.shapePdf, "PAGE_LAYOUT", 640, 480, 250, "BEST", "RGB", True, "ADAPTIVE", "RASTERIZE_BITMAP", False, True, "LAYERS_AND_ATTRIBUTES", True, 90)
                    for lyr in arcpy.mapping.ListLayers(mxd, "", df):
                        lyr.visible = False

                    # create the map_a.pdf with annotation just once
                    self.createAnnotPdf(cfg.shapePdf)                   # creates annot.pdf

                elif self.order_obj.geometry.type.lower() == "point" or self.order_obj.geometry.type.lower() == "multipoint":
                    shutil.copy(cfg.annot_point, os.path.join(cfg.scratch, "annot.pdf"))
                    for lyr in arcpy.mapping.ListLayers(mxd, "", df):
                        lyr.visible = False

        elif yesboundary.lower() == 'no':
            for lyr in arcpy.mapping.ListLayers(mxd, "", df):
                lyr.visible = False

        return mxd, df, yesboundary

    def selectFim(self, mastergdb, projection):
        arcpy.env.overwriteOutput = True
        arcpy.env.outputCoordinateSystem = projection
        arcpy.env.workspace = mastergdb

        shplist = arcpy.ListFeatureClasses()
        desc = arcpy.Describe(cfg.orderGeometry)                    # must use orderGeometry in srGCS83 or can't select FIM mastershp boundaries as they are in srGWGS84
        extent = desc.extent
        xMax = extent.XMax
        xMin = extent.XMin
        yMax = extent.YMax
        yMin = extent.YMin

        j = 0
        i = 0
        mainlist = []
        adjlist = []

        for shp in shplist:
            desc = arcpy.Describe(shp)
            extent = desc.extent
            if (xMax < extent.XMax and xMin > extent.XMin and yMax < extent.YMax and yMin > extent.YMin):        # algorithm optimization
                arcpy.AddMessage("in extent: " + shp)
                shpLayer = arcpy.mapping.Layer(shp)

                # select main
                arcpy.SelectLayerByLocation_management(shpLayer,'intersect', cfg.orderGeometry, None, 'NEW_SELECTION')
                if(int((arcpy.GetCount_management(shpLayer).getOutput(0))) >0):
                    j = j + 1
                    selectedmain = "in_memory/selectedmain" + str(j)
                    arcpy.CopyFeatures_management(shpLayer, selectedmain)
                    mainlist.append(selectedmain)

                # select adjacent
                arcpy.SelectLayerByLocation_management(shpLayer,'intersect', cfg.extent, None, 'NEW_SELECTION')
                if(int((arcpy.GetCount_management(shpLayer).getOutput(0))) >0):
                    i = i + 1
                    selectedadj = "in_memory/selectedadj" + str(i)
                    arcpy.CopyFeatures_management(shpLayer, selectedadj)
                    adjlist.append(selectedadj)

        if j > 0:
            arcpy.Merge_management(mainlist,cfg.selectedmain)

        if i > 0:
            arcpy.Merge_management(adjlist,cfg.selectedadj)

        arcpy.Delete_management("in_memory")

        return mainlist, adjlist

    def getFimRecords(self, selectedmainSHP, selectedadjSHP):
        mainList={}
        adjacentList={}
        selMainLayer = arcpy.MakeFeatureLayer_management(selectedmainSHP)
        selAdjLayer = arcpy.MakeFeatureLayer_management(selectedadjSHP)
        
        # get main records
        for row in arcpy.SearchCursor(selMainLayer):
            voldict = {}
            volumeNum = row.getValue('VOLUMENO').encode("utf-8")
            volumeName  = row.getValue('VOLUMENAME')
            state = row.getValue('STATE').title()
            city = row.getValue('CITY').title()
            volseq = str(row.getValue('VOLUME')).title()
            year = row.getValue('YEAR')
            sheetNo = row.getValue('IMAGE_NO').encode("utf-8")
            imagepath = os.path.join(row.getValue('VOLUMEPATH'), volumeNum, "INDIVIDUAL_GEOREFERENCE", row.getValue('IMAGE_NO'))
            volpath = os.path.join(row.getValue('VOLUMEPATH'), volumeNum)

            t = [volumeNum, volumeName, state, city, volseq, year, [sheetNo], [imagepath],volpath]
            if year in mainList and volumeNum in mainList[year]:
                mainList[year][volumeNum][6].append(sheetNo)
                mainList[year][volumeNum][7].append(imagepath)
            elif year in mainList and volumeNum not in mainList[year]:
                mainList[year][volumeNum] = t
            else:
                voldict[volumeNum] = t
                mainList[year] = voldict

        # get adjacent records
        for row in arcpy.SearchCursor(selAdjLayer):
            voldict = {}
            volumeNum = row.getValue('VOLUMENO')
            volumeName  = row.getValue('VOLUMENAME')
            state = row.getValue('STATE').title()
            city = row.getValue('CITY').title()
            volseq = str(row.getValue('VOLUME')).title()
            year = row.getValue('YEAR')
            sheetNo = row.getValue('IMAGE_NO').encode("utf-8")
            imagepath = os.path.join(row.getValue('VOLUMEPATH'), volumeNum, "INDIVIDUAL_GEOREFERENCE", row.getValue('IMAGE_NO'))
            volpath = os.path.join(row.getValue('VOLUMEPATH'), volumeNum)
            
            # skip sheets that already exist in main
            if volumeNum in [item for sublist in mainList.values() for item in sublist] and sheetNo in [v[6] for sublist in mainList.values() for k,v in sublist.items() if k == volumeNum][0]:
                continue

            t = [volumeNum, volumeName, state, city, volseq, year, [sheetNo], [imagepath],volpath]
            if year in adjacentList and volumeNum in adjacentList[year]:
                adjacentList[year][volumeNum][6].append(sheetNo)
                adjacentList[year][volumeNum][7].append(imagepath)
            elif year in adjacentList and volumeNum not in adjacentList[year]:
                adjacentList[year][volumeNum] = t
            else:
                voldict[volumeNum] = t
                adjacentList[year] = voldict

        return mainList, adjacentList

    def delyear(self, delyear, mainList):
        for year in delyear:
            if any(year in key for key in mainList.keys()):
                del mainList[year]
        return mainList

    def createPDF(self, mainList, adjacentList, is_aei, mxd, dfmain, dfinset, yesboundary, multipage, gridsize, resolution):
        years = mainList.keys()

        if is_aei == 'Y':
            years.sort(reverse = False)
        else:
            years.sort(reverse = True)

        count = 0
        for year in years:
            sheetnoText = ''

            # add main images to mxd
            for item in mainList.get(year).values():
                (volumeNum, volumeName, state, city, volseq, year, sheetNos, imagepaths, volpath) = item
                for lyr in imagepaths:
                    if os.path.exists(lyr+".tif"):
                        lyr+=".tif"
                    elif  os.path.exists(lyr+".jpg"):
                        lyr+=".jpg"

                    count+=1
                    image_lyr_name = "%s"%(lyr.replace("\\","").replace(".","_"))
                    image = arcpy.MakeRasterLayer_management(lyr,image_lyr_name)
                    arcpy.ApplySymbologyFromLayer_management(image_lyr_name, cfg.sheetLayer)

                    layer_temp = os.path.join(cfg.scratch,"image_%s.lyr"%(count))
                    arcpy.SaveToLayerFile_management(image_lyr_name,layer_temp)

                    layer_temp = arcpy.mapping.Layer(layer_temp)
                    arcpy.mapping.AddLayer(dfmain, layer_temp,"Bottom")

                # only display info from main records
                if volseq == '' or volseq == ' ' or volseq == 'None' or volseq == None:
                    sheetnoText = sheetnoText + 'Volume NA: '
                else:
                    sheetnoText = sheetnoText + 'Volume ' + str(volseq) + ': '

                sheetNos.sort()
                for sht in sheetNos:
                    sht = sht.lstrip("0").split('_')[0].split('-')[0].strip('A').strip('B').strip('C').strip('D').strip('E').strip('F') 
                    sheetnoText = sheetnoText + sht + ', '
                sheetnoText = sheetnoText.rstrip(", ") + '; ' + '\r\n'

                # add image_boundary.shp to inset dataframe
                boundLayer = arcpy.mapping.ListLayers(mxd, "IMAGE_BOUNDARY", dfinset)[0]
                boundLayer.replaceDataSource(volpath,"SHAPEFILE_WORKSPACE","IMAGE_BOUNDARY")

            # add adjacent images from the same year to mxd
            if year in adjacentList:
                for item in adjacentList.get(year).values():
                    (volumeNum, volumeName, state, city, volseq, year, sheetNos, imagepaths, volpath) = item
                    for lyr in imagepaths:
                        if os.path.exists(lyr+".tif"):
                            lyr+=".tif"
                        elif  os.path.exists(lyr+".jpg"):
                            lyr+=".jpg"

                        count+=1
                        image_lyr_name = "%s"%(lyr.replace("\\","").replace(".","_"))
                        image = arcpy.MakeRasterLayer_management(lyr,image_lyr_name)
                        arcpy.ApplySymbologyFromLayer_management(image_lyr_name, cfg.sheetLayer)

                        layer_temp = os.path.join(cfg.scratch,"image_%s.lyr"%(count))
                        arcpy.SaveToLayerFile_management(image_lyr_name,layer_temp)

                        layer_temp = arcpy.mapping.Layer(layer_temp)
                        arcpy.mapping.AddLayer(dfmain, layer_temp,"Bottom")

                    # add image_boundary to inset dataframe
                    boundLayer = arcpy.mapping.Layer(cfg.bndylyrfile)
                    if volpath not in [v[8] for v in mainList.get(year).values()]:
                        boundLayer.replaceDataSource(volpath,"SHAPEFILE_WORKSPACE","IMAGE_BOUNDARY")
                        arcpy.mapping.AddLayer(dfinset,boundLayer,"Bottom")

            # set map elements
            self.mapSetElement(mxd, year, sheetnoText)

            # set multipage
            if multipage == True:
                self.setMultipage(year, mxd, dfmain, gridsize, yesboundary)    

            # export pdf and save mxd
            FIPpdf = os.path.join(cfg.scratch, 'FIPExport_'+str(year)+'.pdf')
            arcpy.mapping.ExportToPDF(mxd, FIPpdf, "PAGE_LAYOUT", 640, 480, resolution, "BEST", "RGB", True, "ADAPTIVE", "RASTERIZE_BITMAP", False, True, "None", True, 90)
            mxd.saveACopy(os.path.join(cfg.scratch, "test_"+str(year)+".mxd"))

            # merge annotation pdf to the map if yesBoundary == Y
            if yesboundary == 'yes':
                self.annotatePdf(FIPpdf, cfg.annotPdf)

            # remove layers
            for lyr in arcpy.mapping.ListLayers(mxd, "", dfinset)[1:]:
                arcpy.mapping.RemoveLayer(dfinset, lyr)

            for lyr in arcpy.mapping.ListLayers(mxd, "", dfmain):
                if lyr.name not in ["Project Property", "Buffer Outline", "Grid"]:
                    arcpy.mapping.RemoveLayer(dfmain, lyr)

            arcpy.Delete_management("in_memory")

    def mapSetElement(self, mxd, year, sheetnoText):
        # update map document with sheet numbers, orderID and address
        yearTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "MainTitleText")[0]
        yearTextE.text = str(year)

        mapsheetTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "mapsheetText")[0]
        mapsheetTextE.text = "Map sheet(s): " + '\r\n' + sheetnoText

        orderIDTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "ordernumText")[0]
        orderIDTextE.text = "Order No. " + self.order_obj.number

        AddressText = '%s %s %s %s'%(self.order_obj.address, self.order_obj.city, self.order_obj.province, self.order_obj.postal_code)
        AddressTextE= arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "AddressText")[0]
        AddressTextE.text = "Address: " + AddressText + " "

        if self.is_newLogo == 'Y' and self.is_emg == 'N':
            logoE = arcpy.mapping.ListLayoutElements(mxd, "PICTURE_ELEMENT", "logo")[0]
            logoE.sourceImage = os.path.join(cfg.logopath, self.newlogofile)

        arcpy.RefreshTOC()

    def setMultipage(self, year, mxd, dfmain, gridsize, yesboundary):
        # reset boundary
        if yesboundary != "no":
            arcpy.mapping.ListLayers(mxd, "Project Property", dfmain)[0].visible = True

        # create grid based on sheets clipped to buffer
        expression = str('"YEAR" LIKE \'%' + str(year) + "%'")
        arcpy.MakeFeatureLayer_management(cfg.selectedmain, "selMainLayer")
        arcpy.SelectLayerByAttribute_management("selMainLayer",'NEW_SELECTION', expression)

        clipSelected = os.path.join(cfg.scratch, cfg.scratchgdb, "clip_"+str(year))
        arcpy.Clip_analysis("selMainLayer", cfg.outBufferSHP, clipSelected)
        
        Gridlrshp = os.path.join(cfg.scratch, cfg.scratchgdb, "gridlr_" + str(year))   
        arcpy.GridIndexFeatures_cartography(Gridlrshp, clipSelected, "INTERSECTFEATURE", "", "", gridsize, gridsize)

        gridLayer = arcpy.mapping.ListLayers(mxd,"Grid",dfmain)[0]
        gridLayer.replaceDataSource(os.path.join(cfg.scratch, cfg.scratchgdb),"FILEGDB_WORKSPACE", "gridlr" + "_" + str(year))
        gridLayer.visible = True

        # export multipages
        FIPpdfMM = os.path.join(cfg.scratch, 'FIPExport_'+str(year)+'_multipage.pdf')
        ddMMDDP = mxd.dataDrivenPages
        ddMMDDP.refresh()
        ddMMDDP.exportToPDF(out_pdf=FIPpdfMM, page_range_type="ALL",resolution=600, image_quality="BEST", georef_info = True)

        # reset view
        dfmain.extent = arcpy.mapping.ListLayers(mxd,"Buffer Outline",dfmain)[0].getSelectedExtent(False)
        scale = dfmain.scale * 1.75
        dfmain.scale = ((int(scale)/100)+1)*100

        # reset boundary
        if yesboundary != "fixed":
            arcpy.mapping.ListLayers(mxd, "Project Property", dfmain)[0].visible = False

        del FIPpdfMM
        del ddMMDDP
        return mxd, dfmain

    def appendMapPages(self, output,mainList, yesboundary, multipage):
        years = mainList.keys()
        n=0
        if self.is_aei == 'Y':
            years.sort(reverse = False)
        else:
            years.sort(reverse = True)
        
        for year in years:
            if yesboundary == "yes":
                pdf = PdfFileReader(open(os.path.join(cfg.scratch, 'FIPExport_'+str(year)+'_a.pdf'),'rb'))
            else:
                pdf = PdfFileReader(open(os.path.join(cfg.scratch, 'FIPExport_'+str(year)+'.pdf'),'rb'))

            if multipage == True:
                pdfmm = PdfFileReader(open(os.path.join(cfg.scratch, 'FIPExport_'+str(year)+'_multipage.pdf'),'rb'))
                
                output.addPage(pdf.getPage(0))
                output.addBookmark(str(year), n+2)   #n+1 to accommodate the summary page                        
                
                for x in range(pdfmm.getNumPages()):
                    output.addPage(pdfmm.getPage(x))
                    n = n + 1
                n = n + 1
            else:
                output.addPage(pdf.getPage(0))
                output.addBookmark(str(year), n+2)   #n+1 to accommodate the summary page
                n = n + 1

    def toXplorer(self, mainList, inprojection, outprojection):
        needViewer = 'N'
        try:
            expression = "select fim_viewer from order_viewer where order_id =" + str(self.order_obj.id)
            t = self.oracle.query(expression)
            if t:
                needViewer = t[0][0]
        except:
            arcpy.AddError("### order_viewer failed...")

        if needViewer == 'Y':
            arcpy.AddMessage("...Viewer is needed.")
            arcpy.env.outputCoordinateSystem = inprojection

            metadata = []

            viewerdir = os.path.join(cfg.scratch,self.order_obj.number+'_fim')
            if not os.path.exists(viewerdir):
                os.mkdir(viewerdir)

            # to do: get the right year for each FIM
            years = mainList.keys()
            for year in years:
                mxdname = glob.glob(os.path.join(cfg.scratch,'test*'+str(year)+'.mxd'))[0]
                mxd = arcpy.mapping.MapDocument(mxdname)
                df = arcpy.mapping.ListDataFrames(mxd,"main")[0]                                        # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
                df.spatialReference = inprojection

                for lyr in [x for x in arcpy.mapping.ListLayers(mxd, "", df) if x.name in ["Project Property","Buffer Outline", "Grid"]]:
                        lyr.visible = False

                imagename = str(year)+".jpg"
                arcpy.mapping.ExportToJPEG(mxd, os.path.join(cfg.scratch, viewerdir, imagename), df, df_export_width= 7650, df_export_height=9900, color_mode='8-BIT_GRAYSCALE', world_file = True, jpeg_quality=80)

                desc = arcpy.Describe(os.path.join(viewerdir, imagename))
                featbound = arcpy.Polygon(arcpy.Array([desc.extent.lowerLeft, desc.extent.lowerRight, desc.extent.upperRight, desc.extent.upperLeft]), inprojection)

                tempfeat = os.path.join(cfg.scratch, cfg.scratchgdb, "tilebnd_"+str(year))
                arcpy.Project_management(featbound, tempfeat, outprojection, None, inprojection)        # function requires output not be in_memory
                desc = arcpy.Describe(tempfeat)

                metaitem = {}
                metaitem['type'] = 'fim'
                metaitem['imagename'] = imagename[:-4]+'.jpg'
                metaitem['lat_sw'] = desc.extent.YMin
                metaitem['long_sw'] = desc.extent.XMin
                metaitem['lat_ne'] = desc.extent.YMax
                metaitem['long_ne'] = desc.extent.XMax

                metadata.append(metaitem)
                del desc
                del featbound
                del mxd, df

            arcpy.env.outputCoordinateSystem = None

            if os.path.exists(os.path.join(cfg.viewerFolder, self.order_obj.number+"_fim")):
                shutil.rmtree(os.path.join(cfg.viewerFolder, self.order_obj.number+"_fim"))
            shutil.copytree(os.path.join(cfg.scratch, self.order_obj.number+"_fim"), os.path.join(cfg.viewerFolder, self.order_obj.number+"_fim"))
            url = cfg.uploadlink + self.order_obj.number
            urllib.urlopen(url)

            # insert to oracle
            try:
                expression  = "delete from overlay_image_info where  order_id = %s and (type = 'fim')" % str(self.order_obj.id)
                self.oracle.exe(expression)

                for item in metadata:
                    expression = "insert into overlay_image_info values (%s, %s, %s, %.5f, %.5f, %.5f, %.5f, %s, '', '')" % (str(self.order_obj.id), str(self.order_obj.number), "'" + item['type']+"'", item['lat_sw'], item['long_sw'], item['lat_ne'], item['long_ne'],"'"+item['imagename']+"'") 
                    self.oracle.exe(expression)
            except Exception as e:
                arcpy.AddError(e)
                arcpy.AddError("### overlay_image_info failed...")

        else:
            arcpy.AddMessage("No viewer is needed. Do nothing.")

    def toReportCheck(self, pdfreport_name):
        pdfreport = os.path.join(cfg.scratch, pdfreport_name)

        if os.path.exists(os.path.join(cfg.reportcheckFolder,"FIM", pdfreport_name)):
            os.remove(os.path.join(cfg.reportcheckFolder, "FIM", pdfreport_name))
        shutil.copyfile(pdfreport,os.path.join(cfg.reportcheckFolder, "FIM", pdfreport_name))
        arcpy.SetParameterAsText(5, pdfreport)

        try:
            procedure = 'eris_fim.processFim'
            self.oracle.proc(procedure, (int(self.order_obj.id),))
        except Exception as e:
            arcpy.AddError(e)
            arcpy.AddError("### eris_fim.processFim failed...")

    def log(self, logfile, logname):
        logger = logging.getLogger(logname)
        logger.setLevel(logging.INFO)
        handler = logging.FileHandler(logfile)
        handler.setLevel(logging.INFO)
        logger.addHandler(handler)
        return logger,handler