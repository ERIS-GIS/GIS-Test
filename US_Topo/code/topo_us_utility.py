import json
import arcpy
import os
import csv
import operator
import shutil
import zipfile
import logging
import traceback
import urllib
import re
import time
import cx_Oracle
import xml.etree.ElementTree as ET
import topo_us_config as cfg

from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.generic import NameObject, createStringObject, ArrayObject, FloatObject
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Frame,Table
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import portrait, letter
from reportlab.pdfgen import canvas

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
        
class topo_us_rpt(object):
    def __init__(self, order_obj, oracle):
        self.order_obj = order_obj
        self.oracle = oracle

    def createAnnotPdf(self, myShapePdf):   # inputs cfg.shapePdf, outputs cfg.annotPdf
        '''there is a limitation on the number of vertices available to PDF annotations. 
        This can cause order geometry to get cut off.
        If this happens, change yesBoundary = 'fixed'
        '''
        # input variables
        # part 1: read geometry pdf to get the vertices and rectangle to use
        source  = PdfFileReader(open(myShapePdf,'rb'))
        geomPage = source.getPage(0)
        mystr = geomPage.getObject()['/Contents'].getData()

        # to pinpoint the string part: 1.19997 791.75999 m 1.19997 0.19466 l 611.98627 0.19466 l 611.98627 791.75999 l 1.19997 791.75999 l
        # the format seems to follow x1 y1 m x2 y2 l x3 y3 l x4 y4 l x5 y5 l (i.e. 302.99519 433.46045 m 308.75555 433.74844 l 309.04357 423.09304 l 303.28321 423.09304 l 302.99519 433.46045 l S)
        previousline = ""
        geomString = ""

        # parse coordinates from pdf source code
        try:
            for line in mystr.split("\r\n"):
                if bool(re.match('^\d+\..+ m .+l S$', line)) == True and bool(re.match('.+M$', previousline)) == True:              # if line matches above format
                    geomString = line
                    break                           # get the first match only
                previousline = line
            if geomString == "":
                arcpy.AddWarning("### Failed to parse geomString...")
                raise ValueError
        except:
            try:
                for line in mystr.split("\r\n"):
                    if "l S" in line and bool(re.match('.+M$', previousline)) == True:
                        geomString = line
                        break                       # get the first match only
                    previousline = line
                if geomString == "":
                    arcpy.AddWarning("### Failed to parse geomString...")
                    raise
            except:
                geomString = mystr.split('S\r\n')[0].split('M\r\n')[1]
                if geomString == "":
                    arcpy.AddError("### Failed to parse geomString...")

        coordsString = [value for value in geomString.split(' ') if value not in ['m','l','', 'S']]

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
        outputStream = open(cfg.annotPdf,"wb")
        output.write(outputStream)
        outputStream.close()
        output = None

    def annotatePdf(self, mapPdf, myAnnotPdf):      # input cfg.annotPdf and outputpdf, outputs outputpdf_a
        pdf_intermediate = PdfFileReader(open(mapPdf,'rb'))
        page= pdf_intermediate.getPage(0)

        pdf = PdfFileReader(open(myAnnotPdf,'rb'))
        getpdf = pdf.getPage(0)
        page.mergePage(getpdf)

        output = PdfFileWriter()
        output.addPage(page)

        annotatedPdf = mapPdf[:-4]+'_a.pdf'
        outputStream = open(annotatedPdf,"wb")
        output.write(outputStream)
        outputStream.close()
        output = None
        return annotatedPdf

    def myFirstPage(self, canvas, doc):
        pagesize = portrait(letter)
        [PAGE_WIDTH,PAGE_HEIGHT]=pagesize[:2]
        PAGE_WIDTH=int(PAGE_WIDTH)
        PAGE_HEIGHT=int(PAGE_HEIGHT)

        canvas.saveState()
        canvas.drawImage(cfg.summarypage,0,0, int(PAGE_WIDTH),int(PAGE_HEIGHT))
        canvas.setStrokeColorRGB(0.67,0.8,0.4)
        canvas.line(50,100,int(PAGE_WIDTH-30),100)
        styles = getSampleStyleSheet()
        style = styles["Normal"]

        # set hyperlinks        https://www.usgs.gov/faqs/where-can-i-find-a-topographic-map-symbol-sheet?qt-news_science_products=0#qt-news_science_products
        canvas.linkURL(r"https://pubs.usgs.gov/unnumbered/70039569/report.pdf", (60,247,220,257), thickness=0, relative=1)
        canvas.linkURL(r"https://pubs.usgs.gov/bul/0788e/report.pdf", (60,237,220,247), thickness=0, relative=1)
        canvas.linkURL(r"https://pubs.usgs.gov/gip/TopographicMapSymbols/topomapsymbols.pdf", (60,217,220,227), thickness=0, relative=1)
        canvas.linkURL(r"US Topo Map Symbols.pdf", (60,197,280,207), thickness=0, relative=1)

        canvas.setFont('Helvetica-Bold', 8)
        canvas.drawString(54, 270, "Topographic Map Symbology for the maps may be available in the following documents:")

        canvas.setFont('Helvetica-Oblique', 8)
        canvas.drawString(54, 260, "Pre-1947")
        canvas.drawString(54, 230, "1947-2009")
        canvas.drawString(54, 210, "2009-present")

        canvas.setFont('Helvetica', 8)
        canvas.drawString(54, 180, "Topographic Maps included in this report are produced by the USGS and are to be used for research purposes including a phase I report.")
        canvas.drawString(54, 170, "Maps are not to be resold as commercial property.")
        canvas.drawString(54, 160, "No warranty of Accuracy or Liability for ERIS: The information contained in this report has been produced by ERIS Information Inc.(in the US)")
        canvas.drawString(54, 150, "and ERIS Information Limited Partnership (in Canada), both doing business as 'ERIS', using Topographic Maps produced by the USGS.")
        canvas.drawString(54, 140, "This maps contained herein does not purport to be and does not constitute a guarantee of the accuracy of the information contained herein.")
        canvas.drawString(54, 130, "Although ERIS has endeavored to present you with information that is accurate, ERIS disclaims, any and all liability for any errors, omissions, ")
        canvas.drawString(54, 120, "or inaccuracies in such information and data, whether attributable to inadvertence, negligence or otherwise, and for any consequences")
        canvas.drawString(54, 110, "arising therefrom. Liability on the part of ERIS is limited to the monetary value paid for this report.")
        
        canvas.setFillColorRGB(0,0,255)
        canvas.drawString(54, 250, "    Page 223 of 1918 Topographic Instructions")
        canvas.drawString(54, 240, "    Page 130 of 1928 Topographic Instructions")
        canvas.drawString(54, 220, "    Topographic Map Symbols")
        canvas.drawString(54, 200, "    US Topo Map Symbols (see attached document in this report)")
        
        canvas.restoreState()
        style = None
        del canvas

    def goSummaryPage(self, dictlist, summaryPdf):
        tocData = []
        for d in dictlist:
            years = d.keys()

            if self.is_aei == 'Y':
                years.sort(reverse = False)
            else:
                years.sort(reverse = True)

            for year in years:
                seriesText = d[year][1].replace("75", "7.5")
                if year != "":
                    tocData.append([year, seriesText])
            years = None

        doc = SimpleDocTemplate(summaryPdf, pagesize = letter)
        Story = [Spacer(1,0.5*inch)]
        styles = getSampleStyleSheet()
        style = styles["Normal"]

        p = None
        try:
            p = Paragraph('<para alignment="justify"><font name=Helvetica size = 11>We have searched USGS collections of current topographic maps and historical topographic maps for the project property. Below is a list of maps found for the project property and adjacent area. Maps are from 7.5 and 15 minute topographic map series, if available.</font></para>',style)
        except Exception as e:
            raise

        Story.append(p)
        Story.append(Spacer(1,0.28*inch))

        if len(tocData) < 26:
            tocData.insert(0,["  "])
            tocData.insert(0,['Year','Map Series'])
            table = Table(tocData, colWidths = 35, rowHeights = 14)
            table.setStyle([('FONT',(0,0),(1,0),'Helvetica-Bold'),
                            ('ALIGN',(0,1),(-1,-1),'CENTER'),
                            ('ALIGN',(0,0),(1,0),'LEFT'),])     # note the last comma
            Story.append(table)
        elif len(tocData) > 25 and len(tocData) < 51:           # break into 2 columns
            newdata = []
            newdata.append(['Year','Map Series','   ','Year','Map Series'])
            newdata.append([' ','   ',' '])
            i = 0
            while i < 25:
                row= tocData[i]
                row.append('    ')
                if (i+25) < len(tocData):
                    row.extend(tocData[i+25])
                else:
                    row.extend(['    ','  '])
                newdata.append(row)
                i = i + 1
            table = Table(newdata, colWidths = 35,rowHeights=12)
            table.setStyle([('ALIGN',(0,0),(4,0),'LEFT'),
                    ('FONT',(0,0),(4,0),'Helvetica-Bold'),
                    ('ALIGN',(0,1),(-1,-1),'CENTER'),])
            Story.append(table)
        elif len(tocData) > 50 and len(tocData) < 76:               # break into 3 columns
            newdata = []
            newdata.append(['Year','Map Series','   ','Year','Map Series','   ','Year','Map Series'])
            newdata.append([' ',' ','   ',' ',' ','   ',' ',' '])
            i = 0
            while i < 25:
                row= tocData[i]
                row.append('    ')
                row.extend(tocData[i+25])
                row.append('    ')
                if(i+50) < len(tocData):
                    row.extend(tocData[i+50])
                else:
                    row.append('    ')
                    row.append('  ')

                newdata.append(row)
                i = i + 1
            table = Table(newdata, colWidths = 35,rowHeights=12)
            table.setStyle([('FONT',(0,0),(7,0),'Helvetica-Bold'),
                    ('ALIGN',(0,1),(-1,-1),'CENTER'),
                    ('ALIGN',(0,0),(7,0),'LEFT'),])
            Story.append(table)

        doc.build(Story, onFirstPage=self.myFirstPage, onLaterPages=self.myFirstPage)
        doc = None

    def myCoverPage(self, canvas, doc):
        pagesize = portrait(letter)
        [PAGE_WIDTH,PAGE_HEIGHT]=pagesize[:2]
        PAGE_WIDTH=int(PAGE_WIDTH)
        PAGE_HEIGHT=int(PAGE_HEIGHT)

        if len(self.order_obj.site_name) > 40:
                self.order_obj.site_name = self.order_obj.site_name[0: self.order_obj.site_name[0:40].rfind(' ')] + '\n' + self.order_obj.site_name[self.order_obj.site_name[0:40].rfind(' ')+1:]

        AddressText = '%s\n%s %s %s'%(self.order_obj.address, self.order_obj.city, self.order_obj.province, self.order_obj.postal_code)

        canvas.drawImage(cfg.coverPic,0,0, PAGE_WIDTH,PAGE_HEIGHT)
        leftsw= 54
        heights = 400
        rightsw = 200
        space = 20

        canvas.setFont('Helvetica-Bold', 13)
        canvas.drawString(leftsw, heights, "Project Property:")
        canvas.drawString(leftsw, heights-3*space,"Project No:")
        canvas.drawString(leftsw, heights-4*space,"Requested By:")
        canvas.drawString(leftsw, heights-5*space,"Order No:")
        canvas.drawString(leftsw, heights-6*space,"Date Completed:")
        canvas.setFont('Helvetica', 13)
        canvas.drawString(rightsw,heights-0*space, self.order_obj.site_name)
        canvas.drawString(rightsw, heights-1*space,AddressText.split("\n")[0])
        canvas.drawString(rightsw, heights-2*space,AddressText.split("\n")[1])
        canvas.drawString(rightsw, heights-3*space,self.order_obj.project_num)
        canvas.drawString(rightsw, heights-4*space,self.order_obj.company_desc)
        canvas.drawString(rightsw, heights-5*space,self.order_obj.number)
        canvas.drawString(rightsw, heights-6*space,time.strftime('%B %d, %Y', time.localtime()))
        canvas.saveState()

        del canvas

    def goCoverPage(self,coverPdf):
        doc = SimpleDocTemplate(coverPdf, pagesize = letter)
        doc.build([Spacer(0,4*inch)], onFirstPage=self.myCoverPage, onLaterPages=self.myCoverPage)
        doc = None

    def dedupMaplist(self, mapslist):
        if mapslist != []:
        # remove duplicates (same cell and same year)
            if len(mapslist) > 1:   # if just 1, no need to do anything
                mapslist = sorted(mapslist,key=operator.itemgetter(3,0), reverse = True)  # sorted based on year then cell
                i=1
                remlist = []
                while i<len(mapslist):
                    row = mapslist[i]
                    if row[3] == mapslist[i-1][3] and row[0] == mapslist[i-1][0]:
                        remlist.append(i)
                    i = i+1

                for index in sorted(remlist,reverse = True):
                    del mapslist[index]
        return mapslist

    def reorgByYear(self, maplist):                         # [64818, 15X15 GRID,  LA_Zachary_335142_1963_62500_geo.pdf,  1963, 'htmc', "rowsMain"]
        # reorganize the pdf dictionary based on years
        # filter out irrelevant background years (which doesn't have a centre selected map)
        diction_pdf_inPresentationBuffer = {}

        for row in maplist:
            seriesText = ""
            if row[1] == "15X15 GRID":
                seriesText = "15"
            elif row[1] == "7.5X7.5 GRID":
                seriesText = "75"

            if row[3] in diction_pdf_inPresentationBuffer.keys():
                diction_pdf_inPresentationBuffer[row[3]][2].append([row[2],row[5]])
            else:                
                diction_pdf_inPresentationBuffer[row[3]] = [row[4], seriesText, [[row[2],row[5]]]]
        return diction_pdf_inPresentationBuffer             # {'1946': ['htmc', '15', [['NY_Newburgh_130830_1946_62500_geo.pdf', 'rowsMain']]], '1903': ['htmc', '15', [['NY_Rosendale_129217_1903_62500_geo.pdf', 'rowsAdj'], ['NY_Newburg_144216_1903_62500_geo.pdf', 'rowsMain']]]}

    def createPDF(self, diction, yearalldict, mxd, df, yesboundary, projection, multipage, gridsize, needtif):         # {'1946': ['htmc', '15', [['NY_Newburgh_130830_1946_62500_geo.pdf', 'rowsMain']]], '1903': ['htmc', '15', [['NY_Rosendale_129217_1903_62500_geo.pdf', 'rowsAdj'], ['NY_Newburg_144216_1903_62500_geo.pdf', 'rowsMain']]]}
        # create PDF and also make a copy of the geotiff files if the scale is too small
        years = diction.keys()

        for year in years:
            if year == "":
                years.remove("")
                continue

            seriesText = diction[year][1]
            
            # add topo images to mxd
            quaddict = self.addTopoImages(df, diction, year)

            # set mxd elements
            self.mapSetElements(mxd, year, yearalldict, seriesText, quaddict)

            # multipage
            if multipage == "Y":
                self.setMultipage(df, mxd, seriesText, year, projection, gridsize, needtif, yesboundary)

            # export mxd to pdf       
            outputpdf = os.path.join(cfg.scratch, "map_"+seriesText+"_"+year+".pdf")            
            arcpy.RefreshTOC()  
            arcpy.RefreshActiveView()         
            arcpy.mapping.ExportToPDF(mxd, outputpdf, "PAGE_LAYOUT", 640, 480, 350, "BEST", "RGB", True, "ADAPTIVE", "RASTERIZE_BITMAP", False, True, "LAYERS_AND_ATTRIBUTES", True, 90)
            
            mxd.saveACopy(os.path.join(cfg.scratch, seriesText + "_" + year + ".mxd"))

            # merge annotation pdf to the map if yesBoundary == Y
            if yesboundary == 'yes' and self.order_obj.geometry.type.lower() != 'point' and self.order_obj.geometry.type.lower() != 'multipoint':
                self.annotatePdf(outputpdf, cfg.annotPdf)

            for lyr in arcpy.mapping.ListLayers(mxd, "", df):
                if lyr.name not in ["Project Property", "Buffer Outline", "Grid"]:
                    arcpy.mapping.RemoveLayer(df, lyr)

    def zipDir(self, pdfreport):
        path = os.path.join(cfg.scratch,self.order_obj.number)

        shutil.copy(os.path.join(cfg.scratch,pdfreport), path)
        shutil.copy(cfg.readmefile, path)
        myZipFile = zipfile.ZipFile(os.path.join(path+"_US_Topo.zip"),"w")

        for root, dirs, files in os.walk(path):
            for file in files:
                arcname = os.path.relpath(os.path.join(root, file), os.path.join(path, '..'))
                myZipFile.write(os.path.join(root, file), arcname)

        myZipFile.close()

    def log(self, logfile, logname):
        logger = logging.getLogger(logname)
        handler = logging.FileHandler(logfile)
        handler.setLevel(logging.DEBUG)
        logger.setLevel(logging.DEBUG)    
        logger.addHandler(handler)
        return logger,handler

    def customrpt(self, order_obj):
        # Get company flag for custom colour border
        if order_obj.company_id == 109085: # Mid-Atlantic Associates, Inc.
            cfg.annot_poly = cfg.annot_poly_c
            cfg.annot_line = cfg.annot_line_c
            arcpy.AddMessage('...custom colour boundary set.')
        else:
            arcpy.AddMessage('...no custom colour boundary set.')

        self.is_nova = 'N'
        try:
            expression = "select decode(c.company_id, 40385, 'Y', 'N') is_nova from orders o, customer c where o.customer_id = c.customer_id and o.order_id=" + str(order_obj.id)
            self.is_nova = self.oracle.query(expression)[0][0]
        except Exception as e:
            arcpy.AddWarning(e)
            arcpy.AddWarning("### is_nova failed...")
        arcpy.AddMessage("is_nova = " + self.is_nova)

        self.is_aei = 'N'
        try:
            function = 'ERIS_CUSTOMER.IsProductChron'
            self.is_aei = self.oracle.func(function, str, (str(order_obj.id),))
        except Exception as e:
            arcpy.AddWarning(e)
            arcpy.AddWarning("### ERIS_CUSTOMER.IsProductChron failed...")
        arcpy.AddMessage("is_aei = " + self.is_aei)

        self.is_newLogo = 'N'
        try:
            function = 'ERIS_CUSTOMER.IsCustomLogo'
            self.newlogofile = self.oracle.func(function, str, (str(order_obj.id),))

            if self.newlogofile != None:
                self.is_newLogo = 'Y'
                if self.newlogofile =='RPS_RGB.gif':
                    self.newlogofile='RPS.png'
        except Exception as e:
            arcpy.AddWarning(e)
            arcpy.AddWarning("### ERIS_CUSTOMER.IsCustomLogo failed...")
        arcpy.AddMessage("is_newLogo = " + self.is_newLogo)

        return self.is_nova, self.is_aei, self.is_newLogo

    def delyear(self, yeardel7575, yeardel1515, dict7575, dict1515):
        if yeardel7575:
            for y in yeardel7575:
                del dict7575[y]
        if yeardel1515:
            for y in yeardel1515:
                del dict1515[y]
        return dict7575, dict1515

    def projlist(self, order_obj):
        self.srGCS83 = arcpy.SpatialReference(4269)     # GCS_North_American_1983
        self.srWGS84 = arcpy.SpatialReference(4326)     # GCS_WGS_1984
        self.srGoogle = arcpy.SpatialReference(3857)    # WGS_1984_Web_Mercator_Auxiliary_Sphere
        self.srUTM = order_obj.spatial_ref_pcs
        arcpy.AddMessage(self.srUTM.name)
        return self.srGCS83, self.srWGS84, self.srGoogle, self.srUTM

    def createordergeometry(self, order_obj, projection):
        arcpy.env.overwriteOutput = True
        point = arcpy.Point()
        array = arcpy.Array()

        sr = arcpy.SpatialReference()
        sr.factoryCode = 4269   # requires input geometry is in 4269
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
                sr.setDomain(point.X, point.X, point.Y, point.Y)
                array.add(point)
            if order_obj.geometry.type.lower() == 'point' or order_obj.geometry.type.lower() == 'multipoint':
                feat = arcpy.Multipoint(array, sr)
            elif order_obj.geometry.type.lower() =='polyline':
                feat  = arcpy.Polyline(array, sr)
            else:
                feat = arcpy.Polygon(array,sr)
            array.removeAll()

            # Append to the list of Polygon objects
            featureList.append(feat)
        
        arcpy.CopyFeatures_management(featureList, cfg.orderGeometry)
        arcpy.Project_management(cfg.orderGeometry, cfg.orderGeometryPR, projection)

        del point
        del array

    def mapDocument(self, is_nova, projection):
        arcpy.env.overwriteOutput = True
        arcpy.env.outputCoordinateSystem = projection

        if is_nova == 'Y':
            mxd = arcpy.mapping.MapDocument(cfg.mxdfile_nova)
        else:
            mxd = arcpy.mapping.MapDocument(cfg.mxdfile)

        df = arcpy.mapping.ListDataFrames(mxd,"*")[0]
        df.spatialReference = projection

        if self.order_obj.geometry.type.lower() == 'point' or self.order_obj.geometry.type.lower() == 'multipoint':
            orderGeomlyrfile = cfg.orderGeomlyrfile_point
        elif self.order_obj.geometry.type.lower() == 'polyline':
            orderGeomlyrfile = cfg.orderGeomlyrfile_polyline
        else:
            orderGeomlyrfile = cfg.orderGeomlyrfile_polygon

        arcpy.Buffer_analysis(cfg.orderGeometryPR, cfg.orderBuffer, '0.5 KILOMETERS')                     # has to be not smaller than the search radius to void white page

        orderGeomLayer = arcpy.mapping.Layer(orderGeomlyrfile)
        orderGeomLayer.replaceDataSource(os.path.join(cfg.scratch, cfg.scratchgdb),"FILEGDB_WORKSPACE","orderGeometryPR")

        bufferLayer = arcpy.mapping.Layer(cfg.bufferlyrfile)
        bufferLayer.replaceDataSource(os.path.join(cfg.scratch, cfg.scratchgdb),"FILEGDB_WORKSPACE","buffer")                       # change on 11/3/2016, fix all maps to the same scale
        
        arcpy.mapping.AddLayer(df,bufferLayer,"Top")
        arcpy.mapping.AddLayer(df,orderGeomLayer,"Top")

        return mxd, df

    def mapExtent(self, df, mxd, projection, multipage):
        df.extent = arcpy.mapping.ListLayers(mxd, "Buffer Outline", df)[0].getSelectedExtent(True)
        if multipage == "Y":
            df.scale = df.scale * 1.3               
        else:
            df.scale = df.scale * 1.1               # need to add 10% buffer or ordergeometry might touch dataframe boundary.

        needtif = False
        mscale = 24000      # change on 11/3/2016, to fix all maps to the same scale
        # add to map template, clip (but need to keep both metadata: year, grid size, quadrangle name(s) and present in order
        # 7.5x7.5 minute 24000 scale
        # 30x60 minute 100000 scale
        # 15x15 minute 62500 scale

        if df.scale < mscale:
            df.scale = mscale
            needtif = False
        else:
            # if df.scale > 2 * mscale:  # 2 is an empirical number
            if df.scale > 1.5 * mscale:
                arcpy.AddMessage("...need to provide geotiffs.")
                needtif = True
            else:
                arcpy.AddMessage("...scale is slightly bigger than the original map scale, use the standard topo map scale.")
                needtif = False

        extent = df.extent

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
        polygon = arcpy.Polygon(array, projection)
        arcpy.CopyFeatures_management(polygon, cfg.extent)
        
        return needtif

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
                    yesboundary = 'fixed'
                    for lyr in arcpy.mapping.ListLayers(mxd, "", df):
                        if lyr.name == "Project Property":
                            lyr.visible = True
                        else:
                            lyr.visible = False

        elif yesboundary.lower() == 'no':
            for lyr in arcpy.mapping.ListLayers(mxd, "", df):
                lyr.visible = False

        return mxd, df, yesboundary
        
    def selectTopo(self, orderGeometry, extent, projection):                
        arcpy.env.overwriteOutput = True
        arcpy.env.outputCoordinateSystem = projection
        arcpy.env.workspace = cfg.mastergdb

        arcpy.MakeFeatureLayer_management("Cell_PolygonAll", 'masterLayer')

        arcpy.SelectLayerByLocation_management("masterLayer",'INTERSECT', orderGeometry, None, 'NEW_SELECTION')     # select main topo images that intersect geometry
        rowsMain = [str(r.getValue("CELL_ID")) for r in arcpy.SearchCursor("masterLayer")]
        arcpy.AddMessage("...selected " + str(len(rowsMain)))

        arcpy.SelectLayerByLocation_management("masterLayer",'INTERSECT', extent, None, 'NEW_SELECTION')            # select adjacent topo images to main maps/intersect dataframe extent
        rowsAdj = [str(r.getValue("CELL_ID")) for r in arcpy.SearchCursor("masterLayer")]
        arcpy.AddMessage("...selected " + str(len(rowsAdj)))

        return rowsMain, rowsAdj

    def getTopoRecords(self, rowsMain, rowsAdj, csvfile_h, csvfile_c):
        # cellids are found, need to find corresponding map .pdf by reading the .csv file
        # also get the year info from the corresponding .xml
        infomatrix = []
        yearalldict = {}

        with open(csvfile_h, "rb") as f:
            arcpy.AddMessage("___All USGS HTMC Topo List.")
            reader = csv.DictReader(f)
            for row in reader:                                              # grab main topo records
                if row["Cell ID"] in rowsMain:
                    pdfname = row["Filename"].strip()
                    xmlname = pdfname[0:-3] + "xml"                         # read the year from .xml file
                    xmlpath = os.path.join(cfg.tifdir_h,xmlname)
                    tree = ET.parse(xmlpath)
                    root = tree.getroot()
                    procsteps = root.findall("./dataqual/lineage/procstep")

                    yeardict = {}                                           
                    for procstep in procsteps:
                        procdate = procstep.find("./procdate")
                        if procdate != None:
                            procdesc = procstep.find("./procdesc")
                            yeardict[procdesc.text.lower()] = procdate.text

                    year2use = yeardict.get("date on map")
                    if year2use == "":
                        year2use = row["SourceYear"].strip()
                        arcpy.AddMessage("### cannot determine year of the map from xml...get from csv instead..." + year2use)

                    yearalldict[pdfname] = yeardict                                                 # {'WA_Seattle_243651_1992_100000_geo.pdf': {'aerial photo year': '1988', 'edit year': '1992', 'date on map': '1992'}], ['WA_Seattle_243652_1992_100000_geo.pdf', {'aerial photo year': '1988', 'edit year': '1992', 'date on map': '1992'},...}
                    infomatrix.append([row["Cell ID"],row["Grid Size"],pdfname,year2use,"htmc", "rowsMain"])    # [64818, 15X15 GRID,  LA_Zachary_335142_1963_62500_geo.pdf,  1963]

            f.seek(0)       # resets to beginning of csv
            next(f)         # skip header
            for row in reader:              
                if row["Cell ID"] in rowsAdj:                               # grab surrounding topo records
                    pdfname = row["Filename"].strip()
                    xmlname = pdfname[0:-3] + "xml"                         # read the year from .xml file
                    xmlpath = os.path.join(cfg.tifdir_h,xmlname)
                    tree = ET.parse(xmlpath)
                    root = tree.getroot()
                    procsteps = root.findall("./dataqual/lineage/procstep")

                    yeardict = {}
                    for procstep in procsteps:
                        procdate = procstep.find("./procdate")
                        if procdate != None:
                            procdesc = procstep.find("./procdesc")
                            yeardict[procdesc.text.lower()] = procdate.text

                    year2use = yeardict.get("date on map")
                    if year2use == "":
                        year2use = row["SourceYear"].strip()
                        arcpy.AddMessage("### cannot determine year of the map from xml...get from csv instead..." + year2use)

                    if [row["Grid Size"],year2use,"htmc"] in [[y[1],y[3],y[4]] for y in infomatrix]:
                        if pdfname not in yearalldict:                                                  
                            yearalldict[pdfname] = yeardict                                             # {'WA_Seattle_243651_1992_100000_geo.pdf': {'aerial photo year': '1988', 'edit year': '1992', 'date on map': '1992'}], ['WA_Seattle_243652_1992_100000_geo.pdf', {'aerial photo year': '1988', 'edit year': '1992', 'date on map': '1992'},...}
                            infomatrix.append([row["Cell ID"],row["Grid Size"],pdfname,year2use,"htmc", "rowsAdj"])    # [64818, 15X15 GRID,  LA_Zachary_335142_1963_62500_geo.pdf,  1963, "htmc"]

        with open(csvfile_c, "rb") as f:
            arcpy.AddMessage("___All USGS Current Topo List.")
            reader = csv.DictReader(f)
            for row in reader:
                if row["Cell ID"] in rowsMain:
                    pdfname = row["Filename"].strip()
                    year2use = row["Filename"].split("_")[-3][:4]   # for current topos, read the year from the geopdf file name

                    if year2use == "" or year2use == None:
                        year2use = row["SourceYear"].strip()        # else get it from csv file
                        if year2use[0:2] != "20":
                            arcpy.AddWarning("### Error in the year of the map..." + year2use)
                    
                    yearalldict[pdfname] = {}
                    infomatrix.append([row["Cell ID"],row["Grid Size"],pdfname,year2use,"topo", "rowsMain"])

            f.seek(0)       # resets to beginning of csv
            next(f)         # skip header
            for row in reader:
                if row["Cell ID"] in rowsAdj:                               # grab surrounding topo records
                    pdfname = row["Filename"].strip()
                    year2use = row["Filename"].split("_")[-3][:4]

                    if year2use == "" or year2use == None:
                        year2use = row["SourceYear"].strip()
                        if year2use[0:2] != "20":
                            arcpy.AddWarning("### Error in the year of the map..." + year2use)

                    if [row["Grid Size"],year2use,"topo"] in [[y[1],y[3],y[4]] for y in infomatrix]:
                        if pdfname not in yearalldict:
                            yearalldict[pdfname] = {}
                            infomatrix.append([row["Cell ID"],row["Grid Size"],pdfname,year2use,"topo", "rowsAdj"])  # [64818, 15X15 GRID,  LA_Zachary_335142_1963_62500_geo.pdf,  1963, "topo"]

        maps7575 = []
        maps1515 = []
        for row in infomatrix:
            if row[3] =="":
                arcpy.AddWarning("...blank value in row: " + str(row))
            else:
                if row[1] == "7.5X7.5 GRID":
                    maps7575.append(row)
                elif row[1] == "15X15 GRID":
                    maps1515.append(row)

        return maps7575, maps1515, yearalldict
    
    def addTopoImages(self, df, diction, year): # {'1946': ['htmc', '15', [['NY_Newburgh_130830_1946_62500_geo.pdf', 'rowsMain']]], '1903': ['htmc', '15', [['NY_Rosendale_129217_1903_62500_geo.pdf', 'rowsAdj'], ['NY_Newburg_144216_1903_62500_geo.pdf', 'rowsMain']]]}
        topoType = diction[year][0]
        seriesText = diction[year][1]
        mscale = 24000
        if topoType == 'htmc':
            mscale = int(diction[year][2][0][0].split('_')[-2])   # assumption: WI_Ashland East_500066_1964_24000_geo.pdf, and all pdfs from the same year are of the same scale

        if topoType == "htmc":
            tifdir = cfg.tifdir_h
            if len(diction.keys()) > 1:             # year
                topolyrfile = cfg.topolyrfile_b
            else:
                topolyrfile = cfg.topolyrfile_none
        elif topoType == "topo":
            tifdir = cfg.tifdir_c
            if len(diction.keys()) > 1:             # year
                topolyrfile = cfg.topolyrfile_w
            else:
                topolyrfile = cfg.topolyrfile_none

        copydir = os.path.join(cfg.scratch,self.order_obj.number,str(year)+"_"+seriesText+"_"+str(mscale))
        os.makedirs(copydir)   # WI_Marengo_503367_1984_24000_geo.pdf -> 1984_7.5_24000

        pdfnames = diction[year][2]
        pdfnames.sort()

        quaddict = {}
        seq = 1
        for pdfname in pdfnames:
            tifname = pdfname[0][0:-4]   # note without .tif part
            tifMainAdj = pdfname[1]

            if os.path.exists(os.path.join(tifdir, tifname + "_t.tif")):
                if '.' in tifname:
                    tifname = tifname.replace('.','')

                # need to make a local copy of the tif file for fast data source replacement
                shutil.copyfile(os.path.join(tifdir,tifname+"_t.tif"),os.path.join(copydir,tifname+'.tif'))
                topoLayer = arcpy.mapping.Layer(topolyrfile)
                topoLayer.replaceDataSource(copydir, "RASTER_WORKSPACE", tifname)
                topoLayer.name = tifname
                arcpy.mapping.AddLayer(df, topoLayer, "BOTTOM")

                quad = tifname.split('_')
                if tifMainAdj == "rowsMain":                                            # we only want quadstext that intersect with the ordergeometry displayed
                    if topoType == "htmc":
                        quadname = quad[1] +", "+quad[0]
                    else:
                        quadname = " ".join(quad[1:len(quad)-3])+", "+quad[0]

                    quaddict[pdfname[0]] = quadname
            else:
                arcpy.AddWarning("### tif file doesn't exist " + tifname)
                if not os.path.exists(tifdir):
                    arcpy.AddWarning("tif dir does NOT exist " + tifdir)
                else:
                    arcpy.AddWarning("tif dir does exist " + tifdir)

            seq = seq + 1
        return quaddict

    def mapSetElements(self, mxd, year, yearalldict, seriesText, quaddict):
        # write photo and photo revision year
        quadText = ""
        yearlistText = ""                                           # must include blank space if blank to write text element        
        i = 1
        for key, value in quaddict.items():                         # {'NY_Clintondale_137764_1943_31680_geo.pdf': 'Clintondale, NY', 'NY_Rosendale_129220_1943_31680_geo.pdf': 'Clintondale, NY'}
            pdfname = key
            quadname = value

            pdfyears =  yearalldict.get(pdfname)                    # {'WA_Seattle_243651_1992_100000_geo.pdf': {'aerial photo year': '1988', 'edit year': '1992', 'date on map': '1992'}], ['WA_Seattle_243652_1992_100000_geo.pdf', {'aerial photo year': '1988', 'edit year': '1992', 'date on map': '1992'},...}
            filteryears = {k: v for k, v in pdfyears.items() if k in ["aerial photo year","photo revision year"]}
            if filteryears:
                seqno = "<SUB>(" + str(i) + ")</SUB>" 
                yearlistText = yearlistText + seqno + "\r\n"
                for k,v in filteryears.items():                     # for now we only want to include "aerial photo year","photo revision year" out of 8 year types
                    if len(filteryears) == 2:
                        x = k.title() + ": " + v + "\r\n"
                    else:                                           # if there is only 1 year available
                        x = k.title() + ": " + v + "\r\n\r\n"
                    yearlistText = yearlistText + x
                quadText = quadText + quadname + seqno + "; "
                i+=1
            else:
                yearlistText = " "
                quadText = quadText + quadname + "; "

        yearTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "year")[0]
        yearTextE.text = year

        yearlist = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "yearlist")[0]
        yearlist.text = yearlistText

        quadrangleTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "quadrangle")[0]
        quadrangleTextE.text = quadText.rstrip('; ')

        sourceTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "source")[0]
        sourceTextE.text = "Source: USGS " + seriesText.replace("75", "7.5") + " Minute Topographic Map"

        ordernoTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "orderno")[0]
        ordernoTextE.text = "Order No. " + self.order_obj.number

        if self.is_nova == 'Y':
            projNoTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "projno")[0]
            projNoTextE.text = "Project No: " + self.order_obj.project_num

            siteNameTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "sitename")[0]
            siteNameTextE.text = "Site Name: " + self.order_obj.site_name + ', ' + self.order_obj.address

        if self.is_newLogo == 'Y':     # custom logs
            logoE = arcpy.mapping.ListLayoutElements(mxd, "PICTURE_ELEMENT", "logo")[0]
            logoE.sourceImage = os.path.join(cfg.logopath, self.newlogofile)

    def oracleSummary(self, dictlist, pdfreport):
        summarydata = []
        topoSource = 'USGS'
        for d in dictlist:
            tempyears = d.keys()
            tempyears.sort(reverse = True)
            for year in tempyears:
                seriesText = d[year][1].replace("75", "7.5")
                if year != "":
                    summarydata.append([year, seriesText, topoSource])
            tempyears = None
                
        summarylist = {"ORDER_ID":self.order_obj.id,"FILENAME":pdfreport,"SUMMARY":summarydata}
        topassorc = json.dumps(summarylist,ensure_ascii=False)

        try:
            function = 'eris_gis.AddTopoSummary'
            orc_return = self.oracle.func(function, str, (str(topassorc),))

            if orc_return == 'Success':
                arcpy.AddMessage("...Summary successfully populated to Oracle.")
            else:
                arcpy.AddWarning("...Summary failed to populate to Oracle, check DB admin.")
        except Exception as e:
            arcpy.AddError(e)
            arcpy.AddError("### Oracle eris_gis.AddTopoSummary failed... " + str(e))

    def appendMapPages(self, dictlist, output, yesboundary, multipage):
        n=0
        for d in dictlist:
            if len(d) > 0:
                years = d.keys()
                years = filter(None, years)         # removes empty strings

                if self.is_aei == 'Y':
                    years.sort(reverse = False)
                else:
                    years.sort(reverse = True)
                
                for year in years:
                    seriesText = d[year][1]
                    if yesboundary == "yes" and self.order_obj.geometry.type.lower() != 'point' and self.order_obj.geometry.type.lower() != 'multipoint':
                        pdf = PdfFileReader(open(os.path.join(cfg.scratch,"map_" + seriesText + "_" + year + "_a.pdf"),'rb'))
                    else:
                        pdf = PdfFileReader(open(os.path.join(cfg.scratch,"map_" + seriesText + "_" + year + ".pdf"),'rb'))

                    if multipage == "Y":
                        pdfmm = PdfFileReader(open(os.path.join(cfg.scratch,"map_" + seriesText + "_" + year + "_mm.pdf"),'rb'))
                        
                        output.addPage(pdf.getPage(0))
                        output.addBookmark(year + "_" + seriesText.replace("75", "7.5"), n+2)   #n+1 to accommodate the summary page                        
                        
                        for x in range(pdfmm.getNumPages()):
                            output.addPage(pdfmm.getPage(x))
                            n = n + 1
                        n = n + 1
                    else:
                        output.addPage(pdf.getPage(0))
                        output.addBookmark(year + "_" + seriesText.replace("75", "7.5"), n+2)   #n+1 to accommodate the summary page
                        n = n + 1

    def toXplorer(self, needtif, dictlist, inprojection, outprojection):
        # check if need to copy data to Topo viewer
        needViewer = 'N'
        try:
            expression = "select topo_viewer from order_viewer where order_id =" + str(self.order_obj.id)
            t = self.oracle.query(expression)
            if t:
                needViewer = t[0][0]
        except:
            raise
        
        if needViewer == 'Y':                
            arcpy.AddMessage("...Viewer is needed.")

            # need to reorganize deliver directory
            metadata = []
            for d in dictlist:
                metadata = self.getXplorerImage(d, inprojection, outprojection, metadata, needtif)

            # insert to oracle
            try:
                expression  = "delete from overlay_image_info where order_id = %s and (type = 'topo75' or type = 'topo150' or type = 'topo15')" % str(self.order_obj.id)
                self.oracle.exe(expression)

                for item in metadata:
                    expression = "insert into overlay_image_info values (%s, %s, %s, %.5f, %.5f, %.5f, %.5f, %s, '', '')" % (str(self.order_obj.id), str(self.order_obj.number), "'" + item['type']+"'", item['lat_sw'], item['long_sw'], item['lat_ne'], item['long_ne'],"'"+item['imagename']+"'" )
                    self.oracle.exe(expression)
            except Exception as e:
                arcpy.AddError(e)
                arcpy.AddError("### overlay_image_info failed...")

            if os.path.exists(os.path.join(cfg.viewerFolder, self.order_obj.number+"_topo")):
                shutil.rmtree(os.path.join(cfg.viewerFolder, self.order_obj.number+"_topo"))
            shutil.copytree(os.path.join(cfg.scratch, self.order_obj.number+"_topo"), os.path.join(cfg.viewerFolder, self.order_obj.number+"_topo"))
            url = cfg.topouploadurl + self.order_obj.number
            urllib.urlopen(url)

        else:
            arcpy.AddMessage("No viewer is needed. Do nothing.")

    def getXplorerImage(self, diction, inprojection, outprojection, metadata, needtif):
        arcpy.env.outputCoordinateSystem = inprojection

        viewerdir = os.path.join(cfg.scratch, self.order_obj.number+'_topo')
        if not os.path.exists(viewerdir):
            os.mkdir(viewerdir)        
        if not os.path.exists(os.path.join(viewerdir,"75")):
            os.mkdir(os.path.join(viewerdir,"75"))
        if not os.path.exists(os.path.join(viewerdir,"150")):
            os.mkdir(os.path.join(viewerdir,"150"))

        for year in diction.keys():
            arcpy.AddMessage(year)
            seriesText = diction[year][1]
            mxdname = os.path.join(cfg.scratch, seriesText+'_'+str(year)+'.mxd')

            mxd = arcpy.mapping.MapDocument(mxdname)
            df = arcpy.mapping.ListDataFrames(mxd)[0]                   # the spatial reference here was UTM zone #, need to change to WGS84 Web Mercator
            df.spatialReference = inprojection                          # need to use srGoogle because xplorer uses Google

            for lyr in [x for x in arcpy.mapping.ListLayers(mxd, "", df) if x.name in ["Project Property","Buffer Outline", "Grid"]]:
                    lyr.visible = False

            if needtif == True:
                df.extent = arcpy.mapping.ListLayers(mxd, "Buffer Outline", df)[0].getSelectedExtent(True)
                df.scale = df.scale * 1.1                               # need to add 10% buffer or ordergeometry might touch dataframe boundary.            
            else:
                df.scale = 24000

            imagename = str(year)+".jpg"
            imagepath = os.path.join(viewerdir, seriesText.replace("15", "150"), imagename)
            arcpy.mapping.ExportToJPEG(mxd, imagepath, df, df_export_width=3573, df_export_height=4000, color_mode='24-BIT_TRUE_COLOR', world_file = True, jpeg_quality=50)

            desc = arcpy.Describe(imagepath)
            featbound = arcpy.Polygon(arcpy.Array([desc.extent.lowerLeft, desc.extent.lowerRight, desc.extent.upperRight, desc.extent.upperLeft]), inprojection)                            

            tempfeat = os.path.join(cfg.scratch, cfg.scratchgdb, "tilebnd_"+str(year))
            arcpy.Project_management(featbound, tempfeat, outprojection, None, inprojection)  # function requires output not be in_memory
            desc = arcpy.Describe(tempfeat)

            metaitem = {}
            metaitem['type'] = 'topo' + seriesText.replace("15", "150")
            metaitem['imagename'] = imagename
            metaitem['lat_sw'] = desc.extent.YMin
            metaitem['long_sw'] = desc.extent.XMin
            metaitem['lat_ne'] = desc.extent.YMax
            metaitem['long_ne'] = desc.extent.XMax

            metadata.append(metaitem)
            arcpy.env.outputCoordinateSystem = None

        del desc
        del featbound
        del mxd, df
        return metadata

    def toReportCheck(self, needtif, pdfreport):
        scratchpdf = os.path.join(cfg.scratch,pdfreport)
        reportcheckpdf = os.path.join(cfg.reportcheckFolder,"TopographicMaps",pdfreport)
        scratchzip = os.path.join(cfg.scratch,pdfreport[:-3] + "zip")
        reportcheckzip = os.path.join(cfg.reportcheckFolder,"TopographicMaps",pdfreport[:-3] + "zip")

        if needtif == "Y":
            if os.path.exists(reportcheckzip):
                os.remove(reportcheckzip)
            shutil.copyfile(scratchzip,reportcheckzip)
            arcpy.SetParameterAsText(5, scratchzip)
        else:
            if os.path.exists(reportcheckpdf):
                os.remove(reportcheckpdf)
            shutil.copyfile(scratchpdf,reportcheckpdf)
            arcpy.SetParameterAsText(5, scratchpdf)

        try:
            procedure = 'eris_topo.processTopo'
            self.oracle.proc(procedure, (int(self.order_obj.id),))
        except Exception as e:
            arcpy.AddError(e)
            arcpy.AddError("### eris_topo.processTopo failed...")

    def setMultipage(self, df, mxd, seriestext, year, projection, gridsize, needtif, yesboundary):
        gridname = seriestext + "_" + str(year)

        # clip topo image extents to buffer
        rasterExt = arcpy.CreateFeatureclass_management(os.path.join(cfg.scratch, cfg.scratchgdb), "ext_" + gridname, "POLYGON")
        for lyr in arcpy.mapping.ListLayers(mxd, "", df):
            if lyr.name not in ["Project Property","Buffer Outline", "Grid"]:
                rasterExtent = arcpy.Describe(lyr)
                polygon = arcpy.Polygon(arcpy.Array([rasterExtent.extent.lowerLeft, rasterExtent.extent.lowerRight, rasterExtent.extent.upperRight, rasterExtent.extent.upperLeft]), projection)
                arcpy.Append_management(polygon, rasterExt, "NO_TEST")

            if lyr.name == "Grid":
                lyr.visible = True
                gridLayer = lyr

            if lyr.name == "Project Property" and yesboundary != "no":
                lyr.visible = True

        rasterExtClip = os.path.join(cfg.scratch, cfg.scratchgdb, "extclip_" + gridname) 
        arcpy.Clip_analysis(rasterExt, cfg.orderBuffer, rasterExtClip)

        # create grid
        grid = os.path.join(cfg.scratch, cfg.scratchgdb, "grid_" + gridname)
        arcpy.GridIndexFeatures_cartography(grid, rasterExtClip, "", "", "", gridsize, gridsize)

        gridLayer = arcpy.mapping.ListLayers(mxd, "Grid", df)[0]
        gridLayer.replaceDataSource(os.path.join(cfg.scratch,cfg.scratchgdb),"FILEGDB_WORKSPACE","grid_" + gridname)

        # export mmpage
        pdfmm = os.path.join(cfg.scratch, "map_" + gridname + "_mm.pdf")
        mxdmm = mxd.dataDrivenPages             # data driven pages must be enabled on ArcMap first in the template.mxd.
        mxdmm.refresh()
        mxdmm.exportToPDF(pdfmm, page_range_type="ALL", resolution=100, georef_info = True)

        # reset extent
        if needtif == True:
            df.extent = arcpy.mapping.ListLayers(mxd, "Buffer Outline", df)[0].getSelectedExtent(True)
            df.scale = df.scale * 1.3                               # need to add 10% buffer or ordergeometry might touch dataframe boundary.            
        else:
            df.scale = 24000

        # reset boundary
        if yesboundary != "fixed":
            arcpy.mapping.ListLayers(mxd, "Project Property", df)[0].visible = False

        del mxdmm