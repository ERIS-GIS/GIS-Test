import time
import json
import arcpy
import os
import cx_Oracle
import csv
import urllib
import operator
import shutil
import zipfile
import logging

import xml.etree.ElementTree as ET
import topo_us_word_config as cfg

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
        
class topo_us_word_rpt(object):
    def __init__(self, order_obj, oracle):
        self.order_obj = order_obj
        self.oracle = oracle

    def goCoverPage(self, order_obj):
        if not os.path.exists(cfg.zipCover):
            shutil.copytree(cfg.coverTemplate,cfg.zipCover)

        mxd = arcpy.mapping.MapDocument(cfg.Covermxdfile)

        SITENAME = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "siteText")[0]
        SITENAME.text = order_obj.site_name
        ADDRESS = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "addressText")[0]
        ADDRESS.text = order_obj.address + "\r\n" + order_obj.city + " " + order_obj.province + " " + str(order_obj.postal_code).replace("None", "")
        PROJECT_NUM = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "projectidText")[0]
        PROJECT_NUM.text = order_obj.project_num
        COMPANY_NAME = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "companyText")[0]
        COMPANY_NAME.text = order_obj.company_desc
        ORDER_NUM = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "ordernumText")[0]
        ORDER_NUM.text = order_obj.number
        
        arcpy.mapping.ExportToEMF(mxd, cfg.coverEMF, "PAGE_LAYOUT")
        mxd.saveACopy(cfg.coverMXD)

        shutil.copyfile(cfg.coverEMF, os.path.join(cfg.zipCover,"word\media\image2.emf"))
        self.zipdir_noroot(cfg.zipCover,"cover.docx")

    def goSummaryPage(self, dictlist):
        mxdSummary = arcpy.mapping.MapDocument(cfg.Summarymxdfile)

        # update summary table
        j = 0
        for d in dictlist:
            years = d.keys()
            years.sort(reverse=True)

            for year in years:
                mapseries = d[year][1].replace("75", "7.5")
                j=j+1
                i = str(j)
                exec("e"+i+"1E = arcpy.mapping.ListLayoutElements(mxdSummary, 'TEXT_ELEMENT', 'e"+i+"1')[0]")
                exec("e"+i+"1E.text = year")
                exec("e"+i+"2E = arcpy.mapping.ListLayoutElements(mxdSummary, 'TEXT_ELEMENT', 'e"+i+"2')[0]")
                exec("e"+i+"2E.text = mapseries")

        arcpy.mapping.ExportToEMF(mxdSummary, cfg.summaryEMF, "PAGE_LAYOUT")
        mxdSummary.saveACopy(cfg.summaryMXD)
        mxdSummary = None

        if not os.path.exists(cfg.zipSummary):
            shutil.copytree(cfg.summaryTemplate,cfg.zipSummary)
            
        shutil.copyfile(cfg.summaryEMF, os.path.join(cfg.zipSummary,"word\media\image2.emf"))
        self.zipdir_noroot(cfg.zipSummary,"summary.docx")

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

    def reorgByYear(self, maplist):                      # [64818, 15X15 GRID,  LA_Zachary_335142_1963_62500_geo.pdf,  1963]
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

    def mapDocument(self, projection):
        arcpy.env.overwriteOutput = True
        arcpy.env.outputCoordinateSystem = projection

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

    def mapExtent(self, df, mxd, projection):
        df.extent = arcpy.mapping.ListLayers(mxd, "Buffer Outline", df)[0].getSelectedExtent(True)
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

    # create PDF and also make a copy of the geotiff files if the scale is too small
    def createDOCX(self, yesboundary, dictlist, yearalldict, mxd, df, custom_profile, app):
        worddoclist = []

        if not custom_profile:
            directorytemplate = cfg.customdocxTemplate
        else:
            directorytemplate = cfg.docxTemplate

        if not os.path.exists(cfg.zipMap):
            shutil.copytree(directorytemplate, cfg.zipMap)

        for diction in dictlist:
            years = diction.keys()
            years.sort(reverse = True)
            for year in years:
                if year == "":
                    years.remove("")
                    continue

                seriesText = diction[year][1]

                # add topo images
                quaddict = self.addTopoImages(df, diction, year)

                # export jpg (map image)
                outputjpg = os.path.join(cfg.scratch, "map_"+seriesText+"_"+year+".jpg")
                if diction[year][0] == "htmc":
                    arcpy.mapping.ExportToJPEG(mxd, outputjpg, "PAGE_LAYOUT", 480, 640, 125, "False", "24-BIT_TRUE_COLOR", 90)    #note: because of exporting PAGE_LAYOUT, the wideth and height parameters are ignored.
                else:
                    arcpy.mapping.ExportToJPEG(mxd, outputjpg, "PAGE_LAYOUT", 480, 640, 250, "False", "24-BIT_TRUE_COLOR", 90)
                mxd.saveACopy(os.path.join(cfg.scratch,seriesText+"_"+year+".mxd"))

                # remove all the raster layers
                for lyr in arcpy.mapping.ListLayers(mxd, "", df):
                    if lyr.name not in ["Project Property","Buffer Outline", "Grid"]:
                        arcpy.mapping.RemoveLayer(df, lyr)      # remove the clipFrame layer

                # copy over the unzipped directory
                # change the image files
                # zip up to .docx and change text

                # copy the images to the docx template folder and zip as docx
                imageemf = os.path.join(cfg.zipMap,"word\media\image1.emf")
                topojpg = os.path.join(cfg.zipMap,"word\media\image3.jpg")

                shutil.copyfile(os.path.join(cfg.scratch,"map_"+seriesText+"_"+year+".jpg"), topojpg)              # replace topo image in template folder

                if yesboundary.lower() == "yes" or yesboundary.lower() == "y":
                    shutil.copyfile(cfg.boundaryEMF, imageemf)                                                          # replace boundary image in template folder

                self.zipdir_noroot(os.path.join(cfg.scratch,'tozip'),seriesText+"_"+year+".docx")
                worddoclist.append(os.path.join(cfg.scratch,seriesText+"_"+year+".docx"))

                # set text elements in word
                if quaddict:
                    self.setDocxText(app, quaddict, yearalldict, seriesText, year, yesboundary)

        del mxd
        return worddoclist, app

    def setDocxText(self, app, quaddict, yearalldict, seriesText, year, yesboundary):
        # the word template has been copied, the image files have also been copied, need to refresh and replace the text fields, save

        # get quad and years info
        quadText = ""
        for key, value in quaddict.items():                         # {'NY_Clintondale_137764_1943_31680_geo.pdf': 'Clintondale, NY', 'NY_Rosendale_129220_1943_31680_geo.pdf': 'Clintondale, NY'}
            pdfname = key
            quadname = value.upper()

            yearlist = "("
            pdfyears =  yearalldict.get(pdfname)                    # {'WA_Seattle_243651_1992_100000_geo.pdf': {'aerial photo year': '1988', 'edit year': '1992', 'date on map': '1992'}], ['WA_Seattle_243652_1992_100000_geo.pdf', {'aerial photo year': '1988', 'edit year': '1992', 'date on map': '1992'},...}
            filteryears = {k: v for k, v in pdfyears.items() if k in ["aerial photo year","photo revision year"]}
            if filteryears:
                for k,v in filteryears.items():                     # for now we only want to include "aerial photo year","photo revision year" out of 8 year types
                    x = k.title() + " " + v + ", "
                    yearlist = yearlist + x
                quadText = quadText + quadname + " " + yearlist.rstrip(", ") + "); "
            else:
                quadText = quadText + quadname + "; "
        quads = 'TOPOGRAPHIC MAP IMAGE COURTESY OF THE U.S. GEOLOGICAL SURVEY\r' + seriesText.replace("75", "7.5") + ' Minute Series - '+ str(year) + '\r' + "QUADRANGLES INCLUDE: " + quadText.rstrip("; ")
        
        siteCityState = self.order_obj.city + ", " + self.lookupstate(self.order_obj.province) + " " + str(self.order_obj.postal_code).replace("None", "")
        
        expression = "select address1, address2, city, provstate, postal_code from customer where customer_id = (select customer_id from orders where order_id = " + self.order_obj.id + ")"
        t = self.oracle.query(expression)[0]
        if t[1] == None:
            officeAddress = str(t[0])
        else:
            officeAddress = str(t[0]) + ", " + str(t[1])
        officeCity = str(t[2]) + ", " + self.lookupstate(str(t[3])) + " " + str(t[4])            

        doc = app.Documents.Open(os.path.join(cfg.scratch,seriesText+"_"+year+".docx"))
        allShapes = doc.Shapes

        for item in allShapes:
            if item.Title == "mainYear":
                item.TextFrame.TextRange.Text = 'TOPOGRAPHIC MAP (' + str(year) + ')'    # TOPOGRAPHIC MAP title line
            elif item.Title == "siteInfo":
                item.TextFrame.TextRange.Text = self.order_obj.site_name + "\r" + self.order_obj.address + "\r" + siteCityState
            elif item.Title == "quads":
                item.TextFrame.TextRange.Text = quads
            elif item.Title == "officeAddress":
                item.TextFrame.TextRange.Text = officeAddress
            elif item.Title == "officeCity":
                item.TextFrame.TextRange.Text = officeCity
            elif item.Title == "projectNum":
                item.TextFrame.TextRange.Text = self.order_obj.project_num
            elif item.Title == "orderNum":
                item.TextFrame.TextRange.Text = self.order_obj.number
            elif item.Title == "date":
                item.TextFrame.TextRange.Text = time.strftime('%Y-%m-%d', time.localtime())
            elif item.Title == "boundaryEMF" and (yesboundary == "fixed" or yesboundary == "no" or yesboundary == "n"):
                item.Delete()

        doc.Save()
        doc.Close()
        doc = None

    def addTopoImages(self, df, diction, year): # {'1946': ['htmc', '15', [['NY_Newburgh_130830_1946_62500_geo.pdf', 'rowsMain']]], '1903': ['htmc', '15', [['NY_Rosendale_129217_1903_62500_geo.pdf', 'rowsAdj'], ['NY_Newburg_144216_1903_62500_geo.pdf', 'rowsMain']]]}
        topoType = diction[year][0]
        seriesText = diction[year][1]
        mscale = 24000

        if topoType == "htmc":
            mscale = int(diction[year][2][0][0].split('_')[-2])   # assumption: WI_Ashland East_500066_1964_24000_geo.pdf, and all pdfs from the same year are of the same scale
            tifdir = cfg.tifdir_h
            
            if len(diction.keys()) > 1:             # year
                topolyrfile = cfg.topolyrfile_b
            else:
                topolyrfile = cfg.topolyrfile_none
        elif topoType == str("topo"):
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
                    arcpy.AddWarning("\ttif dir does NOT exist " + tifdir)
                else:
                    arcpy.AddWarning("\ttif dir does exist " + tifdir)

            seq = seq + 1
        return quaddict

    def appendPages(self, app, worddoclist, order_obj):
        # concatenate the word docs into a big final file
        shutil.copyfile(cfg.marginTemplate,os.path.join(cfg.scratch, order_obj.number+"_US_Topo.docx"))
        finaldoc = app.Documents.Open(os.path.join(cfg.scratch, order_obj.number+"_US_Topo.docx"))
        sel = finaldoc.ActiveWindow.Selection
        npages = 0
        sel.InsertFile(cfg.coverDOCX)
        sel.InsertBreak()
        sel.InsertFile(cfg.summaryDOCX)
        sel.InsertBreak()
        for aDoc in worddoclist:
            npages = npages + 1
            sel.InsertFile(aDoc)
            if npages < len(worddoclist):
                sel.InsertBreak()
        
        finaldoc.Save()
        finaldoc.Close()
        app.Application.Quit()

    def unzipDocx(self, file):          # adds password protected string (copied from map template directory settings.xml to final report docx to prevent editing)
        settingsfile = os.path.join(file[:-5], r"word\settings.xml")
        docProtStr = r'<w:documentProtection w:edit="readOnly" w:enforcement="1" w:cryptProviderType="rsaFull" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="4" w:cryptSpinCount="100000" w:hash="EMDyQW9Uyt2eymkd3v1mSeHQxRU=" w:salt="pVAanRWeb0Qx/V/eRR5AIA=="/>'

        # rename docx to zip and extract files
        os.rename(file, file[:-4] + "zip")
        with zipfile.ZipFile(file[:-4] + "zip", "r") as zipobj:
            os.mkdir(file[:-4])
            zipobj.extractall(file[:-4])

        # read data from settings.xml
        with open(settingsfile, "r") as f:
            filedata = f.read()

        # insert the string near the end but inside parent tag
        with open(settingsfile, 'w') as f:
            f.write(filedata.replace("</w:settings>", docProtStr + "</w:settings>"))

        self.zipdir_noroot(file[:-4], file.split("\\")[-1])

        os.remove(file[:-4] + "zip")

    def zipdir_noroot(self, path, zipfilename):
        myZipFile = zipfile.ZipFile(os.path.join(cfg.scratch,zipfilename),"w")
        for root, dirs, files in os.walk(path):
            for afile in files:
                arcname = os.path.relpath(os.path.join(root, afile), path)
                myZipFile.write(os.path.join(root, afile), arcname)
        myZipFile.close()

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

    def projlist(self, order_obj):
        self.srGCS83 = arcpy.SpatialReference(4269)     # GCS_North_American_1983
        self.srWGS84 = arcpy.SpatialReference(4326)     # GCS_WGS_1984
        self.srGoogle = arcpy.SpatialReference(3857)    # WGS_1984_Web_Mercator_Auxiliary_Sphere
        self.srUTM = order_obj.spatial_ref_pcs
        arcpy.AddMessage(self.srUTM.name)
        return self.srGCS83, self.srWGS84, self.srGoogle, self.srUTM

    def setBoundary(self, mxd, df, yesboundary, custom_profile):        
        # get yesboundary flag
        arcpy.AddMessage("yesboundary = " + yesboundary)
        if yesboundary.lower() == 'arrow':
            yesboundary = 'yes'
            custom_profile = True
            arcpy.AddMessage('...custom profile set.')
        else:
            arcpy.AddMessage('...no custom profile set.')

        if yesboundary.lower() == 'fixed':
            for lyr in arcpy.mapping.ListLayers(mxd, "", df):
                if lyr.name == "Project Property":
                    lyr.visible = True
                else:
                    lyr.visible = False

        elif  yesboundary.lower() == 'yes' or yesboundary.lower() == 'y':
            for lyr in arcpy.mapping.ListLayers(mxd, "", df):
                if lyr.name == "Project Property":
                    lyr.visible = True
                else:
                    lyr.visible = False

            # export emf first (just the boundary)
            arcpy.mapping.ExportToEMF(mxd, cfg.boundaryEMF, "PAGE_LAYOUT")
            arcpy.mapping.ListLayers(mxd, "Project Property", df)[0].visible = False

        elif yesboundary.lower() == 'no':
            for lyr in arcpy.mapping.ListLayers(mxd, "", df):
                lyr.visible = False

        return mxd, df, yesboundary, custom_profile

    def lookupstate(self, provstate):
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
        
        return lookup_state[provstate]

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
       
    def selectTopo(self, orderGeometry, extent, projection):                
        arcpy.env.overwriteOutput = True
        arcpy.env.outputCoordinateSystem = projection

        arcpy.MakeFeatureLayer_management(os.path.join(cfg.mastergdb, "Cell_PolygonAll"), 'masterLayer')

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
                    if not os.path.exists(os.path.join(cfg.tifdir_h, pdfname.replace(".pdf", "_t.tif"))):
                        continue

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
                    if not os.path.exists(os.path.join(cfg.tifdir_h, pdfname.replace(".pdf", "_t.tif"))):
                        continue

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
                    if not os.path.exists(os.path.join(cfg.tifdir_c, pdfname.replace(".pdf", "_t.tif"))):
                        continue

                    year2use = row["Filename"].split("_")[-3][:4]   # for current topos, read the year from the geopdf file name

                    if year2use == "" or year2use == None:
                        year2use = row["SourceYear"].strip()        # else get it from csv file
                        if year2use[0:2] != "20":
                            arcpy.AddWarning("### Error in the year of the map..." + year2use)
                    
                    yearalldict[pdfname] = {}
                    infomatrix.append([row["Cell ID"],row["Grid Size"],pdfname,year2use,str("topo"), "rowsMain"])                       # need to 'str' 'topo' because arcgis server analyze was giving error 00068

            f.seek(0)       # resets to beginning of csv
            next(f)         # skip header
            for row in reader:
                if row["Cell ID"] in rowsAdj:                               # grab surrounding topo records
                    pdfname = row["Filename"].strip()
                    if not os.path.exists(os.path.join(cfg.tifdir_c, pdfname.replace(".pdf", "_t.tif"))):
                        continue

                    year2use = row["Filename"].split("_")[-3][:4]

                    if year2use == "" or year2use == None:
                        year2use = row["SourceYear"].strip()
                        if year2use[0:2] != "20":
                            arcpy.AddWarning("### Error in the year of the map..." + year2use)

                    if [row["Grid Size"],year2use,str("topo")] in [[y[1],y[3],y[4]] for y in infomatrix]:                               # need to 'str' 'topo' because arcgis server analyze was giving error 00068
                        if pdfname not in yearalldict:
                            yearalldict[pdfname] = {}
                            infomatrix.append([row["Cell ID"],row["Grid Size"],pdfname,year2use,str("topo"), "rowsAdj"])  # [64818, 15X15 GRID,  LA_Zachary_335142_1963_62500_geo.pdf,  1963, str("topo")]

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
            metaitem['type'] = str('topo' + seriesText.replace("15", "150"))
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
            arcpy.SetParameterAsText(2, scratchzip)
        else:
            if os.path.exists(reportcheckpdf):
                os.remove(reportcheckpdf)
            shutil.copyfile(scratchpdf,reportcheckpdf)
            arcpy.SetParameterAsText(2, scratchpdf)

        try:
            procedure = 'eris_topo.processTopo'
            self.oracle.proc(procedure, (int(self.order_obj.id),))
        except Exception as e:
            arcpy.AddError(e)
            arcpy.AddError("### eris_topo.processTopo failed...")

    def delyear(self, yeardel7575, yeardel1515, dict7575, dict1515):
        if yeardel7575:
            for y in yeardel7575:
                del dict7575[y]
        if yeardel1515:
            for y in yeardel1515:
                del dict1515[y]
        return dict7575, dict1515