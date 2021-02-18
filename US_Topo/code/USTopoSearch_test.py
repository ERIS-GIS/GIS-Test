#-------------------------------------------------------------------------------
# Name:        USGS Topo retrieval
# Purpose:
#
# Author:      LiuJ
#
# Created:     14/10/2014
# Copyright:   (c) LiuJ 2014
# Licence:     <your licence>
#-------------------------------------------------------------------------------

# changes 11/13/2016: change all maps to fixed scale 1:24000
import time,json
import arcpy, os, sys
import csv
import xml.etree.ElementTree as ET
import operator
import shutil, zipfile
import logging
import traceback
import cx_Oracle, glob, urllib
import re
import ConfigParser

from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.generic import NameObject, createStringObject, ArrayObject, FloatObject
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Frame,Table
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import portrait, letter
from reportlab.pdfgen import canvas
from time import strftime

print ("#0 " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
arcpy.env.overwriteOutput = True

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
    # input variables
    # geom_type = 'POLYLINE'      # or POLYGON

    # part 1: read geometry pdf to get the vertices and rectangle to use
    source  = PdfFileReader(open(myShapePdf,'rb'))
    geomPage = source.getPage(0)
    mystr = geomPage.getObject()['/Contents'].getData()
    # to pinpoint the string part: 1.19997 791.75999 m 1.19997 0.19466 l 611.98627 0.19466 l 611.98627 791.75999 l 1.19997 791.75999 l
    # the format seems to follow x1 y1 m x2 y2 l x3 y3 l x4 y4 l x5 y5 l
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

    # below rect seems to be geom bounding box coordinates: xmin, ymin, xmax,ymax
    annot.getObject().update({NameObject('/Rect'):ArrayObject([FloatObject(min(xcoords)), FloatObject(min(ycoords)), FloatObject(max(xcoords)), FloatObject(max(ycoords))])})
    annot.getObject().pop('/AP')  # this is to get rid of the ghost shape

    annot.getObject().update({NameObject('/T'):createStringObject(u'ERIS')})

    output = PdfFileWriter()
    output.addPage(page_geom)
    annotPdf = os.path.join(scratch, "annot.pdf")
    outputStream = open(annotPdf,"wb")
    # output.setPageMode('/UseOutlines')
    output.write(outputStream)
    outputStream.close()
    output = None
    return annotPdf

def annotatePdf(mapPdf, myAnnotPdf):
    pdf_intermediate = PdfFileReader(open(mapPdf,'rb'))
    page= pdf_intermediate.getPage(0)

    pdf = PdfFileReader(open(myAnnotPdf,'rb'))
    FIMpage = pdf.getPage(0)
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

def myFirstPage(canvas, doc):
    canvas.saveState()
    canvas.drawImage(r"\\cabcvan1gis005\GISData\Topo_USA\python\mxd\ERIS_2018_ReportCover_Second Page_F.jpg",0,0, int(PAGE_WIDTH),int(PAGE_HEIGHT))
    canvas.setStrokeColorRGB(0.67,0.8,0.4)
    canvas.line(50,100,int(PAGE_WIDTH-30),100)
    Footer = []
    style = styles["Normal"]
    Disclaimer = []
    # style = styles["Italic"]
    # p1 = Paragraph('<para alignment="justify"><font name=Helvetica size = 8>Topographic Maps included in this report are produced by the USGS and are to be used for research purposes including a phase I report.  Maps are not to be resold as commercial property.</font></para>', style)
    # Disclaimer.append(p1)
    # #p2 = Paragraph('<para alignment="justify"><font name=Helvetica-Bold size=8>No warranty of Accuracy or Liability for ERIS: </font><font name=Helvetica size=8>The information contained in this report has been produced by EcoLog Environmental Risk Information Services Ltd ("ERIS") using Topographic Maps produced by the USGS. This maps contained herein does not purport to be and does not constitute a guarantee of the accuracy of the information contained herein. Although ERIS has endeavored to present you with information that is accurate, EcoLog ERIS disclaims, any and all liability for any errors, omissions, or inaccuracies in such information and data, whether attributable to inadvertence, negligence or otherwise, and for any consequences arising therefrom. Liability on the part of EcoLog ERIS is limited to the monetary value paid for this report.</font></para>',style)
    # p2 = Paragraph("<para alignment='justify'><font name=Helvetica-Bold size=8>No warranty of Accuracy or Liability for ERIS: </font><font name=Helvetica size=8>The information contained in this report has been produced by ERIS Information Inc. (in the US) and ERIS Information Limited Partnership (in Canada), both doing business as 'ERIS', using Topographic Maps produced by the USGS. This maps contained herein does not purport to be and does not constitute a guarantee of the accuracy of the information contained herein. Although ERIS has endeavored to present you with information that is accurate, ERIS disclaims, any and all liability for any errors, omissions, or inaccuracies in such information and data, whether attributable to inadvertence, negligence or otherwise, and for any consequences arising therefrom. Liability on the part of ERIS is limited to the monetary value paid for this report.</font></para>",style)
    # Disclaimer.append(p2)
    # Frame(65,70,int(PAGE_WIDTH-130),155).addFromList(Disclaimer,canvas)
    canvas.setFont('Helvetica', 8)
    canvas.drawString(54, 180, "Topographic Maps included in this report are produced by the USGS and are to be used for research purposes including a phase I report.")
    canvas.drawString(54, 170,"Maps are not to be resold as commercial property.")
    canvas.drawString(54, 160,"No warranty of Accuracy or Liability for ERIS: The information contained in this report has been produced by ERIS Information Inc.(in the US)")
    canvas.drawString(54, 150,"and ERIS Information Limited Partnership (in Canada), both doing business as 'ERIS', using Topographic Maps produced by the USGS.")
    canvas.drawString(54, 140,"This maps contained herein does not purport to be and does not constitute a guarantee of the accuracy of the information contained herein.")
    canvas.drawString(54,130,"Although ERIS has endeavored to present you with information that is accurate, ERIS disclaims, any and all liability for any errors, omissions, ")
    canvas.drawString(54,120,"or inaccuracies in such information and data, whether attributable to inadvertence, negligence or otherwise, and for any consequences")
    canvas.drawString(54,110,"arising therefrom. Liability on the part of ERIS is limited to the monetary value paid for this report.")
    canvas.restoreState()
    p=None
    Footer = None
    Disclaimer = None
    style = None
    del canvas

def goSummaryPage(summaryPdf, data):
    logger.debug("Coming into go(summaryPDF)")
    doc = SimpleDocTemplate(summaryPdf, pagesize = letter)
    # doc.pagesize = portrait(letter)

    logger.debug("#1")
    Story = [Spacer(1,0.5*inch)]
    logger.debug("#2")
    style = styles["Normal"]
    logger.debug("#2-1")

    p = None
    try:
        p = Paragraph('<para alignment="justify"><font name=Helvetica size = 11>We have searched USGS collections of current topographic maps and historical topographic maps for the project property. Below is a list of maps found for the project property and adjacent area. Maps are from 7.5 and 15 minute topographic map series, if available.</font></para>',style)
    except Exception as e:
        logger.error(e)
        logger.error(style)
        logger.error(p)
    logger.debug("#3")

    Story.append(p)
    Story.append(Spacer(1,0.28*inch))

    logger.debug("#####len of data is " + str(len(data)))
    if len(data) < 31:
        data.insert(0,["  "])
        data.insert(0,['Year','Map Series'])
        table = Table(data, colWidths = 35,rowHeights=14)
        table.setStyle([('FONT',(0,0),(1,0),'Helvetica-Bold'),
                 ('ALIGN',(0,1),(-1,-1),'CENTER'),
                 ('ALIGN',(0,0),(1,0),'LEFT'),])   #note the last comma
        Story.append(table)
    elif len(data) > 30 and len(data) < 61: #break into 2 columns
        logger.debug("####len(data) > 30 and len(data) < 61")
        newdata = []
        newdata.append(['Year','Map Series','   ','Year','Map Series'])
        newdata.append([' ','   ',' '])
        i = 0
        while i < 30:
            row= data[i]
            row.append('    ')
            if (i+30) < len(data):
                row.extend(data[i+30])
            else:
                row.extend(['    ','  '])
            newdata.append(row)
            i = i + 1
        table = Table(newdata, colWidths = 35,rowHeights=12)
        table.setStyle([('ALIGN',(0,0),(4,0),'LEFT'),
                 ('FONT',(0,0),(4,0),'Helvetica-Bold'),
                 ('ALIGN',(0,1),(-1,-1),'CENTER'),])
        Story.append(table)
    elif len(data) > 60 and len(data) < 91:   #break into 3 columns
        logger.debug("####len(data) > 90")
        newdata = []
        newdata.append(['Year','Map Series','   ','Year','Map Series','   ','Year','Map Series'])
        newdata.append([' ',' ','   ',' ',' ','   ',' ',' '])
        i = 0
        while i < 30:
            row= data[i]
            row.append('    ')
            row.extend(data[i+30])
            row.append('    ')
            if(i+60) < len(data):
                row.extend(data[i+60])
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

    doc.build(Story, onFirstPage=myFirstPage, onLaterPages=myFirstPage)
    doc = None

def myCoverPage(canvas, doc):
    # canvas.saveState()
    canvas.drawImage(coverPic,0,0, PAGE_WIDTH,PAGE_HEIGHT)
    # elements = []
    # style = ParagraphStyle("cover",parent=styles['Normal'],fontName="Helvetica",fontSize=13,leading=11)
    # if len(coverInfotext["ADDRESS"].split("\n")[0]) >65:
    #     style1 = ParagraphStyle("cover",parent=styles['Normal'],fontName="Helvetica",fontSize=13,leading=17)
    #     print len(coverInfotext["ADDRESS"].split("\n")[0])
    # else:
    #     style1=style
    # data =[[Paragraph('<para alignment="left"><strong> Project Property:</strong></para>',style),Paragraph(('<para alignment="left"><i>%s</i></para>')%(coverInfotext["SITE_NAME"]), style)],
    # [" ",Paragraph(('<para alignment="left"><i>%s</i></para>')%(coverInfotext["ADDRESS"].split("\n")[0]), style1)],
    # [" ",Paragraph(('<para alignment="left"><i>%s</i></para>')%(coverInfotext["ADDRESS"].split("\n")[1]), style1)],
    # [Paragraph(('<para alignment="left"><strong> Project No:</strong></para>'), style), Paragraph(('<para alignment="left"><i>%s</i></para>')%(coverInfotext["PROJECT_NUM"]), style)],
    # [Paragraph(('<para alignment="left"><strong> Requested By:</strong></para>'), style),Paragraph(('<para alignment="left"><i>%s</i></para>')%(coverInfotext["COMPANY_NAME"]), style)],
    # [Paragraph(('<para alignment="left"><strong> Order No:</strong></para>'), style),Paragraph(('<para alignment="left"><i>%s</i></para>')%(coverInfotext["ORDER_NUM"]), style)],
    # [Paragraph(('<para alignment="left"><strong> Date Completed:</strong></para>'), style),Paragraph(('<para alignment="left"><i>%s</i></para>')%(time.strftime('%B %d, %Y', time.localtime())), style)]]

    # t=Table(data,colWidths = [160, int(PAGE_WIDTH-200)])
    # t.wrapOn(canvas,int(PAGE_WIDTH-100),int(PAGE_HEIGHT))
    # t.drawOn(canvas,50,270)
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
    canvas.drawString(rightsw,heights-0*space, coverInfotext["SITE_NAME"])
    canvas.drawString(rightsw, heights-1*space,coverInfotext["ADDRESS"].split("\n")[0])
    canvas.drawString(rightsw, heights-2*space,coverInfotext["ADDRESS"].split("\n")[1])
    canvas.drawString(rightsw, heights-3*space,coverInfotext["PROJECT_NUM"])
    canvas.drawString(rightsw, heights-4*space,coverInfotext["COMPANY_NAME"])
    canvas.drawString(rightsw, heights-5*space,coverInfotext["ORDER_NUM"])
    canvas.drawString(rightsw, heights-6*space,time.strftime('%B %d, %Y', time.localtime()))
    canvas.saveState()

    # canvas.restoreState()
    # p=None
    # Disclaimer = None
    # style = None
    del canvas

def goCoverPage(coverPdf):#, data):
    doc = SimpleDocTemplate(coverPdf, pagesize = letter)
    doc.build([Spacer(0,4*inch)],onFirstPage=myCoverPage, onLaterPages=myCoverPage)
    doc = None

def dedupMaplist(mapslist):
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

def countSheets(mapslist):
    if len(mapslist) == 0:
        count = []
    elif len(mapslist) == 1:
        count = [1]
    else:
        count = [1]
        i = 1
        while i < len(mapslist):
            if mapslist[i][3] == mapslist[i-1][3]:
                count.append(count[i-1]+1)
            else:
                count.append(1)
            i = i + 1
    return count

# reorgnize the pdf dictionary based on years
# filter out irrelevant background years (which doesn't have a centre selected map)
def reorgByYear(mapslist):      # [64818, 15X15 GRID,  LA_Zachary_335142_1963_62500_geo.pdf,  1963]
    diction_pdf_inPresentationBuffer = {}    #{1975: [geopdf1.pdf, geopdf2.pdf...], 1968: [geopdf11.pdf, ...]}
    diction_pdf_inSearchBuffer = {}
    diction_cellids = {}        # {1975:[cellid1,cellid2...], 1968:[cellid11,cellid12,...]}
    for row in mapslist:
        if row[3] in diction_pdf_inPresentationBuffer.keys():  #{1963:LA_Zachary_335142_1963_62500_geo.pdf, 1975:....}
            diction_pdf_inPresentationBuffer[row[3]].append(row[2])
            diction_cellids[row[3]].append(row[0])
        else:
            diction_pdf_inPresentationBuffer[row[3]] = [row[2]]
            diction_cellids[row[3]] = [row[0]]
    for key in diction_cellids:    # key is the year
        hasSelectedMap = False
        for (cellid,pdfname) in zip(diction_cellids[key],diction_pdf_inPresentationBuffer[key]):
            if cellid in cellids_selected:
                if key in diction_pdf_inSearchBuffer.keys():
                    diction_pdf_inSearchBuffer[key].append(pdfname)
                else:
                    diction_pdf_inSearchBuffer[key] = [pdfname]
                hasSelectedMap = True
                # break;
        if not hasSelectedMap:
            diction_pdf_inPresentationBuffer.pop(key,None)
    return (diction_pdf_inPresentationBuffer,diction_pdf_inSearchBuffer)

# create PDF and also make a copy of the geotiff files if the scale is too small
def createPDF(seriesText,diction,diction_s,outpdfname):

    if OrderType.lower()== 'point':
        orderGeomlyrfile = orderGeomlyrfile_point
    elif OrderType.lower() =='polyline':
        orderGeomlyrfile = orderGeomlyrfile_polyline
    else:
        orderGeomlyrfile = orderGeomlyrfile_polygon

    logger.debug("#4-1")
    orderGeomLayer = arcpy.mapping.Layer(orderGeomlyrfile)
    orderGeomLayer.replaceDataSource(scratch,"SHAPEFILE_WORKSPACE","orderGeometry")
    logger.debug("#4-2")

    extentBufferLayer = arcpy.mapping.Layer(bufferlyrfile)
    extentBufferLayer.replaceDataSource(scratch,"SHAPEFILE_WORKSPACE","buffer_extent75")   #change on 11/3/2016, fix all maps to the same scale

    # if seriesText == "7.5":
    #     extentBufferLayer.replaceDataSource(scratch,"SHAPEFILE_WORKSPACE","buffer_extent75")
    # else:   #by default "15"
    #     extentBufferLayer.replaceDataSource(scratch,"SHAPEFILE_WORKSPACE","buffer_extent15")

    outputPDF = arcpy.mapping.PDFDocumentCreate(os.path.join(scratch, outpdfname))

    years = diction.keys()
    if is_aei == 'Y':
        years.sort(reverse = False)
    else:
        years.sort(reverse = True)

    for year in years:
        if year == "":
            years.remove("")
    print(years)

    for year in years:
        if int(year) < 2008:
            tifdir = tifdir_h
            if len(years) > 1:
                topofile = topolyrfile_b
            else:
                topofile = topolyrfile_none
            mscale = int(diction[year][0].split('_')[-2])   #assumption: WI_Ashland East_500066_1964_24000_geo.pdf, and all pdfs from the same year are of the same scale
            print ("########" + str(mscale))
            if is_aei == 'Y' and mscale in [24000,31680]:
                seriesText = '7.5'
            elif is_aei == 'Y' and mscale == 62500:
                seriesText = '15'
            elif is_aei == 'Y':
                seriesText = '7.5'
            else:
                pass
        else:
            tifdir = tifdir_c
            if len(years) > 1:
                topofile = topolyrfile_w
            else:
                topofile = topolyrfile_none
            mscale = 24000
        mscale = 24000      # change on 11/3/2016, to fix all maps to the same scale

    # years = diction.keys()
    # if is_aei == 'Y':
    #     years.sort(reverse = False)
    #     for year in years:
    #         if int(year) < 2008:
    #             tifdir = tifdir_h
    #             if len(years) > 1:
    #                 topofile = topolyrfile_b
    #             else:
    #                 topofile = topolyrfile_none
    #             mscale = int(diction[year][0].split('_')[-2])   #assumption: WI_Ashland East_500066_1964_24000_geo.pdf, and all pdfs from the same year are of the same scale
    #             print "########" + str(mscale)
    #             if mscale == 24000:
    #                 seriesText = '7.5'
    #             else:
    #                 seriesText = '15'
    #         else:
    #             tifdir = tifdir_c
    #             if len(years) > 1:
    #                 topofile = topolyrfile_w
    #             else:
    #                 topofile = topolyrfile_none
    #             mscale = 24000
    #         mscale = 24000
    # else:
    #     years.sort(reverse = True)
    #     for year in years:
    #         if int(year) < 2008:
    #             tifdir = tifdir_h
    #             if len(years) > 1:
    #                 topofile = topolyrfile_b
    #             else:
    #                 topofile = topolyrfile_none
    #             mscale = int(diction[year][0].split('_')[-2])   #assumption: WI_Ashland East_500066_1964_24000_geo.pdf, and all pdfs from the same year are of the same scale
    #             print "########" + str(mscale)
    #         else:
    #             tifdir = tifdir_c
    #             if len(years) > 1:
    #                 topofile = topolyrfile_w
    #             else:
    #                 topofile = topolyrfile_none
    #             mscale = 24000
    #         mscale = 24000   #change on 11/3/2016, to fix all maps to the same scale

        # add to map template, clip (but need to keep both metadata: year, grid size, quadrangle name(s) and present in order

        if is_nova == 'Y':
            mxd = arcpy.mapping.MapDocument(mxdfile_nova)
        else:
            mxd = arcpy.mapping.MapDocument(mxdfile)
        df = arcpy.mapping.ListDataFrames(mxd,"*")[0]
        spatialRef = out_coordinate_system
        df.spatialReference = spatialRef

        arcpy.mapping.AddLayer(df,extentBufferLayer,"Top")
        # if yesBoundary.lower() == "yes" or yesBoundary.lower() == "y":
        #     if emgOrder =='N':        #add geometry for non-emg orders
        #         arcpy.mapping.AddLayer(df,orderGeomLayer,"Top")
        # else:
        #     if is_nova == 'Y':   #nova wants boundary
        #         arcpy.mapping.AddLayer(df,orderGeomLayer,"Top")
        if yesBoundary.lower() == 'fixed':
            arcpy.mapping.AddLayer(df,orderGeomLayer,"Top")

        # change scale, modify map elements, export
        needtif = False
        df.extent = extentBufferLayer.getSelectedExtent(False)
        logger.debug('');
        if df.scale < mscale:
            scale = mscale
            needtif = False
        else:
            # if df.scale > 2 * mscale:  # 2 is an empirical number
            if df.scale > 1.5 * mscale:
                print ("***** need to provide geotiffs")
                scale = df.scale
                needtif = True
            else:
                print ("scale is slightly bigger than the original map scale, use the standard topo map scale")
                scale = df.scale
                needtif = False

        copydir = os.path.join(scratch,deliverfolder,str(year)+"_"+seriesText+"_"+str(mscale))
        os.makedirs(copydir)   # WI_Marengo_503367_1984_24000_geo.pdf -> 1984_7.5_24000
        if needtif == True:
            copydirs.append(copydir)

        pdfnames = diction[year]
        pdfnames.sort()

        quadrangles = ""
        seq = 0
        firstTime = True
        for pdfname in pdfnames:
            tifname = pdfname[0:-4]   # note without .tif part
            tifname_bk = tifname
            if os.path.exists(os.path.join(tifdir,tifname+ "_t.tif")):
                if '.' in tifname:
                    tifname = tifname.replace('.','')

                # need to make a local copy of the tif file for fast data source replacement
                namecomps = tifname.split('_')
                namecomps.insert(-2,year)
                newtifname = '_'.join(namecomps)

                shutil.copyfile(os.path.join(tifdir,tifname_bk+"_t.tif"),os.path.join(copydir,newtifname+'.tif'))
                logger.debug(os.path.join(tifdir,tifname+"_t.tif"))
                topoLayer = arcpy.mapping.Layer(topofile)
                topoLayer.replaceDataSource(copydir, "RASTER_WORKSPACE", newtifname)
                topoLayer.name = newtifname
                arcpy.mapping.AddLayer(df, topoLayer, "BOTTOM")

                if pdfname in diction_s[year]:
                    comps = diction[year][seq].split('_')
                    if int(year)<2008:
                        quadname = comps[1] +", "+comps[0]
                    else:
                        quadname = " ".join(comps[1:len(comps)-3])+", "+comps[0]

                    if quadrangles =="":
                        quadrangles = quadname
                    else:
                        quadrangles = quadrangles + "; " + quadname

            else:
                print ("tif file doesn't exist " + tifname)
                logger.debug("tif file doesn't exist " + tifname)
                if not os.path.exists(tifdir):
                    logger.debug("tif dir doesn't exist " + tifdir)
                else:
                    logger.debug("tif dir does exist " + tifdir)
            seq = seq + 1

        df.extent = extentBufferLayer.getSelectedExtent(False) # this helps centre the map
        df.scale = scale
        for lyr in arcpy.mapping.ListLayers(mxd, "", df):
            if lyr.name == "Buffer Outline":
                arcpy.mapping.RemoveLayer(df, lyr)

        if is_nova == 'Y':
            yearTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "year")[0]
            yearTextE.text = year

            quadrangleTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "quadrangle")[0]
            quadrangleTextE.text = "Quadrangle(s): " + quadrangles

            sourceTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "source")[0]
            sourceTextE.text = "Source: USGS " + seriesText + " Minute Topographic Map"

            projNoTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "projno")[0]
            projNoTextE.text = "Project No. "+ProjNo

            siteNameTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "sitename")[0]
            siteNameTextE.text = "Site Name: "+Sitename+','+AddressText

            ordernoTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "orderno")[0]
            ordernoTextE.text = "Order No. "+ OrderNumText
        else:
            yearTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "year")[0]
            yearTextE.text = year
            yearTextE.elementPositionX = 0.4959

            quadrangleTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "quadrangle")[0]
            quadrangleTextE.text = "Quadrangle(s): " + quadrangles
            quadrangleTextE.elementPositionX = 0.44
            quadrangleTextE.elementPositionY = 0.3875

            sourceTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "source")[0]
            # sourceTextE.text = "Source: USGS " + seriesText + " Minute Topographic Map"  + "   -- scale is " + str(df.scale)
            sourceTextE.text = "Source: USGS " + seriesText + " Minute Topographic Map"
            sourceTextE.elementPositionX = 0.4959

            ordernoTextE = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "orderno")[0]
            ordernoTextE.text = "Order No. "+ OrderNumText

        if is_newLogofile == 'Y':     # need to change logo for emg
            logoE = arcpy.mapping.ListLayoutElements(mxd, "PICTURE_ELEMENT", "logo")[0]
            logoE.sourceImage = os.path.join(logopath, newlogofile)

        # arcpy.mapping.RemoveLayer(df,extentBufferLayer) # this is not working
        arcpy.RefreshTOC()
        outputpdf = os.path.join(scratch, "map_"+seriesText+"_"+year+".pdf")
        if int(year)<2008:
            arcpy.mapping.ExportToPDF(mxd, outputpdf, "PAGE_LAYOUT", 640, 480, 250, "BEST", "RGB", True, "ADAPTIVE", "RASTERIZE_BITMAP", False, True, "LAYERS_AND_ATTRIBUTES", True, 90)
        else:
            arcpy.mapping.ExportToPDF(mxd, outputpdf, "PAGE_LAYOUT", 640, 480, 350, "BEST", "RGB", True, "ADAPTIVE", "RASTERIZE_BITMAP", False, True, "LAYERS_AND_ATTRIBUTES", True, 90)

        # outputPDF.appendPages(outputpdf)
        if seriesText == "7.5":
            # mxd.relativePaths = True
            mxd.saveACopy(os.path.join(scratch,"75_"+year+".mxd"))
        else:
            # mxd.relativePaths = True
            mxd.saveACopy(os.path.join(scratch,"15_"+year+".mxd"))

        if (yesBoundary.lower() == 'yes' and (OrderType.lower() == "polyline" or OrderType.lower() == "polygon")):

            if firstTime:
                # remove all other layers
                scale2use = df.scale
                for lyr in arcpy.mapping.ListLayers(mxd, "", df):
                    arcpy.mapping.RemoveLayer(df, lyr)
                arcpy.mapping.AddLayer(df,orderGeomLayer,"Top") #the layer is visible
                df.scale = scale2use
                shapePdf = os.path.join(scratch, 'shape.pdf')
                arcpy.mapping.ExportToPDF(mxd, shapePdf, "PAGE_LAYOUT", 640, 480, 250, "BEST", "RGB", True, "ADAPTIVE", "RASTERIZE_BITMAP", False, True, "LAYERS_AND_ATTRIBUTES", True, 90)
                # create the a pdf with annotation just once
                myAnnotPdf = createAnnotPdf(OrderType, shapePdf)
                firstTime = False

            # merge annotation pdf to the map
            Topopdf = annotatePdf(outputpdf, myAnnotPdf)
            outputpdf = Topopdf

        outputPDF.appendPages(outputpdf)

    outputPDF.saveAndClose()
    return "Success! :)"

def zipdir(path, zip):
    for root, dirs, files in os.walk(path):
        for file in files:
            # print file + " " + root
            arcname = os.path.relpath(os.path.join(root, file), os.path.join(path, '..'))
            zip.write(os.path.join(root, file), arcname)

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
handler = logging.FileHandler(r'\\cabcvan1gis005\GISData\Topo_USA\python\log\USTopoSearch_Log.txt')
handler.setLevel(logging.DEBUG)
logger.addHandler(handler)

# -----------------------------------------------------------------------------------------------------------------------------------------------
# OrderIDText = arcpy.GetParameterAsText(0)#'734618'#
# BufsizeText = arcpy.GetParameterAsText(1)#'2.4'
# yesBoundary = arcpy.GetParameterAsText(2)#'no'##
# scratch = arcpy.env.scratchWorkspace#r"C:\Users\JLoucks\Documents\JL\Topo_USA_scratch8"#

OrderIDText = ""
OrderNumText = r"20200626202"
BufsizeText = "2.4"
yesBoundary = ""
scratch = os.path.join(r"W:\Data Analysts\Alison\_GIS\TOPO_SCRATCHY", OrderNumText)

# Deployment parameters
server_environment = 'test'
server_config_file = r"\\cabcvan1gis006\GISData\ERISServerConfig.ini"
server_config = server_loc_config(server_config_file,server_environment)
connectionString = 'eris_gis/gis295@cabcvan1ora006.glaciermedia.inc:1521/GMTESTC'
reportcheckFolder = server_config["reportcheck"]
viewerFolder = server_config["viewer"]
topouploadurl =  server_config["viewer_upload"] + r"/ErisInt/BIPublisherPortal_prod/Viewer.svc/TopoUpload?ordernumber="

filepath = r'\\cabcvan1gis005\GISData\Topo_USA\python\mxd'
connectionPath = r'\\cabcvan1gis005\GISData\Topo_USA\python'
masterlyr = r'\\cabcvan1gis005\GISData\Topo_USA\masterfile\Cell_PolygonAll.shp'
csvfile_h = r'\\cabcvan1gis005\GISData\Topo_USA\masterfile\All_HTMC_all_all_gda_results.csv'
csvfile_c = r'\\cabcvan1gis005\GISData\Topo_USA\masterfile\All_USTopo_T_7.5_gda_results.csv'
tifdir_h = r'\\cabcvan1fpr009\USGS_Topo\USGS_HTMC_Geotiff'
# tifdir_c = r"Y:\TOPO_DATA_USA\USGS_currentTopo_Geotiff"
tifdir_c = r'\\cabcvan1fpr009\USGS_Topo\USGS_currentTopo_Geotiff'
mxdfile = os.path.join(filepath,"template.mxd")
mxdfile_nova = os.path.join(filepath,'template_nova_t.mxd')
topolyrfile_none = os.path.join(filepath,"topo.lyr")
topolyrfile_b = os.path.join(filepath,"topo_black.lyr")
topolyrfile_w = os.path.join(filepath,"topo_white.lyr")
bufferlyrfile = os.path.join(filepath,"buffer_extent.lyr")
orderGeomlyrfile_point = os.path.join(filepath,"SiteMaker.lyr")
orderGeomlyrfile_polyline = os.path.join(filepath,"orderLine.lyr")
orderGeomlyrfile_polygon = os.path.join(filepath,"orderPoly.lyr")
readmefile = os.path.join(filepath,"readme.txt")
# erislogopdf = r"\\cabcvan1gis005\GISData\Topo_USA\python\mxd\erislogo.pdf"
erislogojpg = os.path.join(filepath,"ERIS agency logo.jpg")
# emglogopdf = r"\\cabcvan1gis005\GISData\Topo_USA\python\mxd\emglogo.pdf"
emglogopng = os.path.join(filepath,"emglogo.png")
annot_poly = os.path.join(filepath,"annot_poly.pdf")
annot_line = os.path.join(filepath,"annot_line.pdf")
logopath = os.path.join(filepath,"logos")
coverPic = os.path.join(filepath,"ERIS_2018_ReportCover_Topographic Maps_F.jpg")
arcpy.env.overwriteOutput = True
arcpy.env.OverWriteOutput = True
# -----------------------------------------------------------------------------------------------------------------------------------------------

try:
    try:
        # OrderNum = str(raw_input("Enter Order Number:")).strip()
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        # GET ORDER_ID AND BOUNDARY FROM ORDER_NUM
        if OrderIDText == "":
            cur.execute("SELECT * FROM ERIS.TOPO_AUDIT WHERE ORDER_ID IN (select order_id from orders where order_num = '" + str(OrderNumText) + "')")
            result = cur.fetchall()
            OrderIDText = str(result[0][0]).strip()
            yesBoundaryqry = str([row[3] for row in result if row[2]== "URL"][0])
            yesBoundary = re.search('(yesBoundary=)(\w+)(&)', yesBoundaryqry).group(2).strip()
            print("Order ID: " + OrderIDText)
            print("Yes Boundary: " + yesBoundary)

        cur.execute("select order_num from orders where order_id='%s'"%(OrderIDText))
        t = cur.fetchone()
        OrderNum = str(t[0])

    finally:
        cur.close()
        con.close()

    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        coverInfotext = json.loads(cur.callfunc('eris_gis.getCoverPageInfo', str, (str(OrderIDText),)))

        OrderNumText = str(coverInfotext["ORDER_NUM"])
        Sitename =coverInfotext["SITE_NAME"]
        if len(Sitename) > 40:
                Sitename = Sitename[0: Sitename[0:40].rfind(' ')] + '\n' + Sitename[Sitename[0:40].rfind(' ')+1:]
        ProjNo = coverInfotext["PROJECT_NUM"]
        ProName = coverInfotext["COMPANY_NAME"]
        AddressText= coverInfotext["ADDRESS"]

        coverInfotext["ADDRESS"] = '%s\n%s %s %s'%(coverInfotext["ADDRESS"],coverInfotext["CITY"],coverInfotext["PROVSTATE"],coverInfotext["POSTALZIP"])

        # OrderDetails = json.loads(cur.callfunc('eris_gis.getBufferDetails', str, (str(OrderIDText),)))
        # OrderType = OrderDetails["ORDERTYPE"]
        # OrderCoord = eval(OrderDetails["ORDERCOOR"])
        # RadiusType = OrderDetails["RADIUSTYPE"]

        cur.execute("select geometry_type, geometry, radius_type  from eris_order_geometry where order_id =" + OrderIDText)
        t = cur.fetchone()
        OrderType = str(t[0])
        OrderCoord = eval(str(t[1]))
        RadiusType = str(t[2])

    except Exception as e:
        logger.error("Error to get flag from Oracle " + str(e))
        raise
    finally:
        cur.close()
        con.close()

    is_nova = 'N'
    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        cur.execute("select decode(c.company_id, 40385, 'Y', 'N') is_nova from orders o, customer c where o.customer_id = c.customer_id and o.order_id=" + str(OrderIDText))
        t = cur.fetchone()
        is_nova = t[0]

    finally:
        cur.close()
        con.close()

    # Get company flag for custom colour border
    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()

        cur.execute("select customer_id from orders where order_id = '"+ OrderIDText+"'")
        a = cur.fetchone()
        CustIDText = str(a[0])

        cur.execute("select company_id from customer where customer_id = '"+CustIDText+"'")
        b = cur.fetchone()
        CompIDText = str(b[0])

        print (CompIDText)
    except Exception as e:
        logger.error("Error to get company flag from Oracle "+ str(e))
    finally:
        cur.close()
        con.close()

    if CompIDText == '109085':
        annot_poly = os.path.join(filepath,"annot_poly_red.pdf")
        annot_line = os.path.join(filepath,"annot_line_red.pdf")
        print ('custom colour boundary set')
    else:
        print ('no custom colour boundary set')

    is_newLogofile = 'N'
    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()
        newlogofile = cur.callfunc('ERIS_CUSTOMER.IsCustomLogo', str, (str(OrderIDText),))

        if newlogofile != None:
            is_newLogofile = 'Y'
            if newlogofile =='RPS_RGB.gif':
                newlogofile='RPS.png'
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

    copydirs = []    # will keep directories to be zipped
    deliverfolder = OrderNumText
    pdfreport = OrderNumText+"_US_Topo.pdf"

    srGCS83 = arcpy.SpatialReference(os.path.join(connectionPath, u'projections\\GCSNorthAmerican1983.prj'))
    srWGS84 = arcpy.SpatialReference(os.path.join(connectionPath, u'projections\\WGS1984.prj'))

    point = arcpy.Point()
    array = arcpy.Array()
    sr = arcpy.SpatialReference()
    sr.factoryCode = 4269   # requires input geometry is in 4269
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
    if UTMvalue[0]=='0':
        UTMvalue=' '+UTMvalue[1:]
    out_coordinate_system = arcpy.SpatialReference('WGS 1984 UTM Zone %sN'%UTMvalue)#arcpy.SpatialReference('NAD 1983 UTM Zone %sN'%UTMvalue)

    orderGeometryPR = os.path.join(scratch, "ordergeoNamePR.shp")
    arcpy.Project_management(orderGeometry, orderGeometryPR, out_coordinate_system)

    del point
    del array

    logger.debug("#1")
    bufferDistance_e75 = '2 KILOMETERS'         #has to be not smaller than the search radius to void white page
    extentBuffer75SHP = os.path.join(scratch,"buffer_extent75.shp")
    arcpy.Buffer_analysis(orderGeometryPR, extentBuffer75SHP, bufferDistance_e75)

    # bufferDistance_e15 = "5 KILOMETERS"   #note 5km may result in 48K scale maps to produce .zip file
    # extentBuffer15SHP = os.path.join(scratch,"buffer_extent15.shp")
    # arcpy.Buffer_analysis(orderGeometryPR, extentBuffer15SHP, bufferDistance_e15)

    masterLayer = arcpy.mapping.Layer(masterlyr)
    arcpy.SelectLayerByLocation_management(masterLayer,'intersect', orderGeometryPR,'0.25 KILOMETERS')  #it doesn't seem to work without the distance

    logger.debug("#2")
    if(int((arcpy.GetCount_management(masterLayer).getOutput(0))) ==0):
        print ("NO records selected")
        masterLayer = None
    else:
        cellids_selected = []
        # loop through the relevant records, locate the selected cell IDs
        rows = arcpy.SearchCursor(masterLayer)    # loop through the selected records
        for row in rows:
            cellid = str(int(row.getValue("CELL_ID")))
            cellids_selected.append(cellid)
        del row
        del rows

        arcpy.SelectLayerByLocation_management(masterLayer,'intersect', orderGeometryPR,'7 KILOMETERS','NEW_SELECTION')
        cellids = []
        cellsizes = []
        # loop through the relevant records, locate the selected cell IDs
        rows = arcpy.SearchCursor(masterLayer)    # loop through the selected records
        for row in rows:
            cellid = str(int(row.getValue("CELL_ID")))
            cellsize = str(int(row.getValue("CELL_SIZE")))
            cellids.append(cellid)
            cellsizes.append(cellsize)
        del row
        del rows

        masterLayer = None
        logger.debug(cellids)

        infomatrix = []
        # cellids are found, need to find corresponding map .pdf by reading the .csv file
        # also get the year info from the corresponding .xml
        print ("#1 " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
        with open(csvfile_h, "rb") as f:
            print("___All USGS HTMC Topo List.")
            reader = csv.reader(f)
            for row in reader:
                if row[9] in cellids:
                    # print "#2 " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
                    pdfname = row[15].strip()
                    # read the year from .xml file
                    xmlname = pdfname[0:-3] + "xml"
                    xmlpath = os.path.join(tifdir_h,xmlname)
                    tree = ET.parse(xmlpath)
                    root = tree.getroot()
                    procsteps = root.findall("./dataqual/lineage/procstep")
                    yeardict = {}
                    for procstep in procsteps:
                        procdate = procstep.find("./procdate")
                        if procdate != None:
                            procdesc = procstep.find("./procdesc")
                            yeardict[procdesc.text.lower()] = procdate.text
                    # print yeardict
                    year2use = ""
                    yearcandidates = []
                    if "edit year" in yeardict.keys():
                        yearcandidates.append(int(yeardict["edit year"]))

                    if "aerial photo year" in yeardict.keys():
                        yearcandidates.append(int(yeardict["aerial photo year"]))

                    if "photo revision year" in yeardict.keys():
                        yearcandidates.append(int(yeardict["photo revision year"]))

                    if "field check year" in yeardict.keys():
                        yearcandidates.append(int(yeardict["field check year"]))

                    if "photo inspection year" in yeardict.keys():
                        # print "photo inspection year is " + yeardict["photo inspection year"]
                        yearcandidates.append(int(yeardict["photo inspection year"]))

                    if "date on map" in yeardict.keys():
                        # print "date on  map " + yeardict["date on map"]
                        yearcandidates.append(int(yeardict["date on map"]))

                    if len(yearcandidates) > 0:
                        # print "***** length of yearcnadidates is " + str(len(yearcandidates))
                        year2use = str(max(yearcandidates))

                    if year2use == "":
                        print ("################### cannot determine the year of the map!!")

                    # logger.debug(row[9] + " " + row[5] + "  " + row[15] + "  " + year2use)
                    infomatrix.append([row[9],row[5],row[15],year2use])  # [64818, 15X15 GRID,  LA_Zachary_335142_1963_62500_geo.pdf,  1963]

        with open(csvfile_c, "rb") as f:
            print("___All USGS Current Topo List.")
            reader = csv.reader(f)
            for row in reader:
                if row[9] in cellids:
                    # print "#2 " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
                    pdfname = row[15].strip()

                    # for current topos, read the year from the geopdf file name
                    templist = pdfname.split("_")
                    year2use = templist[len(templist)-3][0:4]

                    if year2use[0:2] != "20":
                        print ("################### Error in the year of the map!!")

                    # print (row[9] + " " + row[5] + "  " + row[15] + "  " + year2use)
                    infomatrix.append([row[9],row[5],row[15],year2use])

        logger.debug("#3")
        print ("#3 " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
        # locate the geopdf and find the exact year to report, only use one from the same year per cell
        maps7575 = []
        maps1515 = []
        maps3060 =[]
        maps12 = []
        
        for row in infomatrix:
            if row[3] =="":
                print("BLANK YEAR VALUE IN ROW EXISTS: " + str(row))
            else:
                if row[1] == "7.5X7.5 GRID":
                    maps7575.append(row)
                    print(row)
                elif row[1] == "15X15 GRID":
                    maps1515.append(row)
                    print(row)
                elif row[1] == "30X60 GRID":
                    maps3060.append(row)
                elif row[1] == "1X2 GRID":
                    maps12.append(row)
        # print([item[3] for item in maps7575] )
        # # for debugging
        # for amap in maps7575:
        #     tifname = amap[2][0:-4] + "_t.tif"
        #     xmlname = amap[2][0:-4] + ".xml"
        #     shutil.copy(os.path.join(tifdir,tifname),scratch)
        #     shutil.copy(os.path.join(tifdir,xmlname),scratch)

        # dedup the duplicated years
        maps7575 = dedupMaplist(maps7575)
        maps1515 = dedupMaplist(maps1515)
        maps3060 = dedupMaplist(maps3060)
        maps12 = dedupMaplist(maps12)

        logger.debug("#4")
        print ("#4 " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
        # # count for the same year, how many cells (adjacent sheets)
        # count7575 = countSheets(maps7575)
        # count1515 = countSheets(maps1515)
        # count3060 = countSheets(maps3060)
        # count12 = countSheets(maps12)

        # reorganize data structure
        (dict7575,dict7575_s) = reorgByYear(maps7575)  # {1975: geopdf.pdf, 1973: ...}
        (dict1515,dict1515_s) = reorgByYear(maps1515)
        (dict3060,dict3060_s) = reorgByYear(maps3060)
        (dict12,dict12_s) = reorgByYear(maps12)

        # -----------------------------------------------------------------------------------------------------------
        # UNCOMMENT TO REMOVE BLANK MAPS
        # del dict7575['1975']
        # del dict7575_s['1975']
        # del dict1515['1938']
        # del dict1515_s['1938']
        # del dict1515['1930']
        # del dict1515_s['1930']
        # del dict1515['1929']
        # del dict1515_s['1929']
        # -----------------------------------------------------------------------------------------------------------

        logger.debug("#5")
        # outputPdfname = "map_" + OrderNumText + ".pdf"
        print (is_aei)
        if is_aei == 'Y':
            comb7515 = {}
            comb7515_s = {}
            comb7515.update(dict1515)
            comb7515.update(dict7575)
            comb7515_s.update(dict1515_s)
            comb7515_s.update(dict7575_s)

            createPDF("15-7.5",comb7515,comb7515_s,"map_" + OrderNumText + "_7515.pdf")
        else:
            print("dict7575: " + str(dict7575.keys()))
            print("dict7575_s: " + str(dict7575_s.keys()))
            createPDF("7.5",dict7575,dict7575_s,"map_" + OrderNumText + "_75.pdf")
            createPDF("15",dict1515,dict1515_s,"map_" + OrderNumText + "_15.pdf")

        summarypdf = os.path.join(scratch,'summary.pdf')
        tabledata = []
        summarydata = []
        topoSource = 'USGS'

        if is_aei == 'Y':
            tempyears = comb7515_s.keys()
            tempyears.sort(reverse = False)
            for year in tempyears:
                if year != "":
                    if comb7515_s[year][0].split('_')[-2] == '24000':
                        tabledata.append([year,'7.5'])
                        summarydata.append([year,'7.5',topoSource])
                    elif comb7515_s[year][0].split('_')[-2] == '62500':
                        tabledata.append([year,'15'])
                        summarydata.append([year,'15',topoSource])
                    else:
                        tabledata.append([year,'7.5'])
                        summarydata.append([year,'7.5',topoSource])
            tempyears = None

        else:
            tempyears = dict7575.keys()
            tempyears.sort(reverse = True)
            for year in tempyears:
                if year != "":
                    tabledata.append([year,'7.5'])
                    summarydata.append([year,'7.5',topoSource])
            tempyears = None

            tempyears = dict1515.keys()
            tempyears.sort(reverse = True)
            for year in tempyears:
                if year != "":
                    tabledata.append([year,'15'])
                    summarydata.append([year,'15',topoSource])
            tempyears = None

        pagesize = portrait(letter)
        [PAGE_WIDTH,PAGE_HEIGHT]=pagesize[:2]
        PAGE_WIDTH=int(PAGE_WIDTH)
        PAGE_HEIGHT=int(PAGE_HEIGHT)

        styles = getSampleStyleSheet()

        coverPDF = os.path.join(scratch,"cover.pdf")
        goCoverPage(coverPDF)
        goSummaryPage(summarypdf,tabledata)

        output = PdfFileWriter()
        coverPages = PdfFileReader(open(coverPDF,'rb'))
        summaryPages = PdfFileReader(open(os.path.join(scratch,summarypdf),'rb'))
        output.addPage(coverPages.getPage(0))
        coverPages= None
        output.addPage(summaryPages.getPage(0))
        summaryPages=None
        output.addBookmark("Cover Page",0)
        output.addBookmark("Summary",1)

# Save Summary ##########################################################
        summarylist = {"ORDER_ID":OrderIDText,"FILENAME":pdfreport,"SUMMARY":summarydata}
        # print (summarylist)
        topassorc = json.dumps(summarylist,ensure_ascii=False)
        try:
            con = cx_Oracle.connect(connectionString)
            cur = con.cursor()

            orc_return = cur.callfunc('eris_gis.AddTopoSummary', str, (str(topassorc),))
            if orc_return == 'Success':
                print ("Summary successfully populated to Oracle")
            else:
                print ("Summary failed to populate to Oracle, check DB admin")
        except Exception as e:
            logger.error("Function error, " + str(e))
            raise
        finally:
            cur.close()
            con.close()

        if is_aei == 'Y':

            j = 0
            if len(comb7515) > 0:
                map1575Pages = PdfFileReader(open(os.path.join(scratch,"map_" + OrderNumText+"_7515.pdf"),'rb'))
                years = comb7515.keys()
                years.sort(reverse = False)
                print("==========years 7515")
                for year in years:
                    if year == "":
                        years.remove("")
                print(years)

                for year in years:
                    page = map1575Pages.getPage(j)
                    output.addPage(page)
                    if comb7515_s[year][0].split('_')[-2] in ['24000','31680'] and year < 2008:
                        seriesbkm = '7.5'
                    elif comb7515_s[year][0].split('_')[-2] == '62500':
                        seriesbkm = '15'
                    else:
                        seriesbkm = '7.5'
                    output.addBookmark(year+"_"+seriesbkm,j+2)
                    j=j+1
                    page = None

        else:
            i=0
            if len(dict7575) > 0:
                map7575Pages = PdfFileReader(open(os.path.join(scratch,"map_" + OrderNumText+"_75.pdf"),'rb'))
                years = dict7575.keys()
                years.sort(reverse = True)
                print("==========years 7575")
                for year in years:
                    if year == "":
                        years.remove("")
                print(years)

                for year in years:
                    page = map7575Pages.getPage(i)
                    # if is_nova == 'N':
                    #    page.mergePage(logopage)
                    output.addPage(page)
                    # output.addPage(map7575Pages.getPage(i))
                    output.addBookmark(year+"_7.5",i+2)   #i+1 to accomodate the summary page
                    i = i + 1
                    page = None

            j=0
            if len(dict1515) > 0:
                map1515Pages = PdfFileReader(open(os.path.join(scratch,"map_" + OrderNumText+"_15.pdf"),'rb'))
                years = dict1515.keys()
                years.sort(reverse = True)
                print("==========years 1515")
                for year in years:
                    if year == "":
                        years.remove("")
                print(years)

                for year in years:
                    page = map1515Pages.getPage(j)
                    # page.mergePage(logopage)
                    output.addPage(page)
                    # output.addPage(map1515Pages.getPage(j))
                    output.addBookmark(year+"_15",i+j+2)  # +1 to accomodate the summary page
                    j = j + 1
                    page = None

        outputStream = open(os.path.join(scratch,pdfreport),"wb")
        output.setPageMode('/UseOutlines')
        output.write(outputStream)
        outputStream.close()
        output = None
        summaryPages = None
        map7575Pages = None
        map1515Pages = None

        print ("#5 " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))

        if len(copydirs) > 0:
            shutil.copy(os.path.join(scratch,pdfreport),os.path.join(scratch,deliverfolder))
            shutil.copy(readmefile,os.path.join(scratch,deliverfolder))
            myZipFile = zipfile.ZipFile(os.path.join(scratch,OrderNumText+"_US_Topo.zip"),"w")
            zipdir(os.path.join(scratch,deliverfolder),myZipFile)
            myZipFile.close()

        needViewer = 'N'
        # check if need to copy data for Topo viewer
        try:
            con = cx_Oracle.connect(connectionString)
            cur = con.cursor()

            cur.execute("select topo_viewer from order_viewer where order_id =" + str(OrderIDText))
            t = cur.fetchone()
            if t != None:
                needViewer = t[0]

        finally:
            cur.close()
            con.close()

        if needViewer == 'Y':
            metadata = []
            srGoogle = arcpy.SpatialReference(3857)   # web mercator
            arcpy.AddMessage("Viewer is needed. Need to copy data to obi002")
            viewerdir = os.path.join(scratch,OrderNumText+'_topo')
            if not os.path.exists(viewerdir):
                os.mkdir(viewerdir)
            tempdir = os.path.join(scratch,'viewertemp')
            if not os.path.exists(tempdir):
                os.mkdir(tempdir)
            # need to reorganize deliver directory

            dirs = filter(os.path.isdir, glob.glob(os.path.join(scratch,deliverfolder)+'\*_7.5_*'))
            if len(dirs) > 0:
                if not os.path.exists(os.path.join(viewerdir,"75")):
                    os.mkdir(os.path.join(viewerdir,"75"))
                # get the extent to use. use one uniform for now
                year = dirs[0].split('_7.5_')[0][-4:]
                mxdname = '75_'+year+'.mxd'
                mxd = arcpy.mapping.MapDocument(os.path.join(scratch,mxdname))
                df = arcpy.mapping.ListDataFrames(mxd,"*")[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
                df.spatialReference = srGoogle
                extent = df.extent

                del df, mxd
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
                arcpy.CopyFeatures_management(polygon, os.path.join(tempdir, "Extent75.shp"))
                arcpy.DefineProjection_management(os.path.join(tempdir, "Extent75.shp"), srGoogle)

                arcpy.Project_management(os.path.join(tempdir, "Extent75.shp"), os.path.join(tempdir,"Extent75_WGS84.shp"), srWGS84)
                desc = arcpy.Describe(os.path.join(tempdir, "Extent75_WGS84.shp"))
                lat_sw = desc.extent.YMin
                long_sw = desc.extent.XMin
                lat_ne = desc.extent.YMax
                long_ne = desc.extent.XMax
                # clip_envelope = str(extent75.XMin) + " " + str(extent75.YMin) + " " + str(extent75.XMax) + " " + str(extent75.YMax)
            # arcpy.env.compression = "JPEG 85"

            arcpy.env.outputCoordinateSystem = srGoogle
            # arcpy.env.extent = extent
            if is_aei == 'Y':
                for year in comb7515.keys():
                    if os.path.exists(os.path.join(scratch,'75_'+str(year)+'.mxd')):
                        mxdname = os.path.join(scratch,'75_'+str(year)+'.mxd')
                        mxd = arcpy.mapping.MapDocument(mxdname)
                        df = arcpy.mapping.ListDataFrames(mxd)[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
                        df.spatialReference = srGoogle

                        # queryLayer = arcpy.mapping.ListLayers(mxd,"Buffer Outline",df)[0]
                        # df.extent = queryLayer.getSelectedExtent(False)

                        imagename = str(year)+".jpg"
                        # arcpy.mapping.ExportToJPEG(mxd, os.path.join(scratch, viewerdir, imagename), df,df_export_width= 14290,df_export_height=16000, color_mode='8-BIT_GRAYSCALE',world_file = True) #by default, the jpeg quality is 100
                        arcpy.mapping.ExportToJPEG(mxd, os.path.join(scratch, viewerdir,"75", imagename), df,df_export_width= 3573,df_export_height=4000, color_mode='24-BIT_TRUE_COLOR',world_file = True, jpeg_quality=50)#,df_export_width= 14290,df_export_height=16000, color_mode='8-BIT_GRAYSCALE',world_file = True, jpeg_quality=100)

                        desc = arcpy.Describe(os.path.join(viewerdir,"75",imagename))
                        featbound = arcpy.Polygon(arcpy.Array([desc.extent.lowerLeft, desc.extent.lowerRight, desc.extent.upperRight, desc.extent.upperLeft]),
                                            srGoogle)
                        del desc

                        tempfeat = os.path.join(tempdir, "tilebnd_"+str(year)+ ".shp")

                        arcpy.Project_management(featbound, tempfeat, srWGS84) #function requires output not be in_memory
                        del featbound
                        desc = arcpy.Describe(tempfeat)

                        metaitem = {}
                        metaitem['type'] = 'topo75'
                        metaitem['imagename'] = imagename[:-4]+'.jpg'
                        metaitem['lat_sw'] = desc.extent.YMin
                        metaitem['long_sw'] = desc.extent.XMin
                        metaitem['lat_ne'] = desc.extent.YMax
                        metaitem['long_ne'] = desc.extent.XMax

                        metadata.append(metaitem)
                        del mxd, df
                    elif os.path.exists(os.path.join(scratch,'15_'+str(year)+'.mxd')):
                        if not os.path.exists(os.path.join(viewerdir,"150")):
                            os.mkdir(os.path.join(viewerdir,"150"))
                        mxdname = os.path.join(scratch,'15_'+str(year)+'.mxd')
                        mxd = arcpy.mapping.MapDocument(mxdname)
                        df = arcpy.mapping.ListDataFrames(mxd)[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
                        df.spatialReference = srGoogle

                        imagename = str(year)+".jpg"
                        # arcpy.mapping.ExportToJPEG(mxd, os.path.join(scratch, viewerdir, imagename), df,df_export_width= 14290,df_export_height=16000, color_mode='8-BIT_GRAYSCALE',world_file = True) #by default, the jpeg quality is 100
                        arcpy.mapping.ExportToJPEG(mxd, os.path.join(scratch, viewerdir,"150", imagename), df,df_export_width= 3573,df_export_height=4000, color_mode='24-BIT_TRUE_COLOR',world_file = True, jpeg_quality=50)#,df_export_width= 14290,df_export_height=16000, color_mode='8-BIT_GRAYSCALE',world_file = True, jpeg_quality=100)

                        desc = arcpy.Describe(os.path.join(viewerdir,"150",imagename))
                        featbound = arcpy.Polygon(arcpy.Array([desc.extent.lowerLeft, desc.extent.lowerRight, desc.extent.upperRight, desc.extent.upperLeft]),srGoogle)
                        del desc

                        tempfeat = os.path.join(tempdir, "tilebnd_"+str(year)+ ".shp")

                        arcpy.Project_management(featbound, tempfeat, srWGS84) #function requires output not be in_memory
                        del featbound
                        desc = arcpy.Describe(tempfeat)

                        metaitem = {}
                        metaitem['type'] = 'topo150'
                        metaitem['imagename'] = imagename[:-4]+'.jpg'
                        metaitem['lat_sw'] = desc.extent.YMin
                        metaitem['long_sw'] = desc.extent.XMin
                        metaitem['lat_ne'] = desc.extent.YMax
                        metaitem['long_ne'] = desc.extent.XMax

                        metadata.append(metaitem)
                        del mxd, df
                    arcpy.env.outputCoordinateSystem = None

            else:
                for year in dict7575.keys():
                    print(year)
                    mxdname = glob.glob(os.path.join(scratch,'75_'+str(year)+'.mxd'))[0]
                    mxd = arcpy.mapping.MapDocument(mxdname)
                    df = arcpy.mapping.ListDataFrames(mxd)[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
                    df.spatialReference = srGoogle

                    # queryLayer = arcpy.mapping.ListLayers(mxd,"Buffer Outline",df)[0]
                    # df.extent = queryLayer.getSelectedExtent(False)

                    imagename = str(year)+".jpg"
                    # arcpy.mapping.ExportToJPEG(mxd, os.path.join(scratch, viewerdir, imagename), df,df_export_width= 14290,df_export_height=16000, color_mode='8-BIT_GRAYSCALE',world_file = True) #by default, the jpeg quality is 100
                    arcpy.mapping.ExportToJPEG(mxd, os.path.join(scratch, viewerdir,"75", imagename), df,df_export_width= 3573,df_export_height=4000, color_mode='24-BIT_TRUE_COLOR',world_file = True, jpeg_quality=50)#,df_export_width= 14290,df_export_height=16000, color_mode='8-BIT_GRAYSCALE',world_file = True, jpeg_quality=100)

                    desc = arcpy.Describe(os.path.join(viewerdir,"75",imagename))
                    featbound = arcpy.Polygon(arcpy.Array([desc.extent.lowerLeft, desc.extent.lowerRight, desc.extent.upperRight, desc.extent.upperLeft]),
                                        srGoogle)
                    del desc

                    tempfeat = os.path.join(tempdir, "tilebnd_"+str(year)+ ".shp")

                    arcpy.Project_management(featbound, tempfeat, srWGS84) #function requires output not be in_memory
                    del featbound
                    desc = arcpy.Describe(tempfeat)

                    metaitem = {}
                    metaitem['type'] = 'topo75'
                    metaitem['imagename'] = imagename[:-4]+'.jpg'
                    metaitem['lat_sw'] = desc.extent.YMin
                    metaitem['long_sw'] = desc.extent.XMin
                    metaitem['lat_ne'] = desc.extent.YMax
                    metaitem['long_ne'] = desc.extent.XMax

                    metadata.append(metaitem)
                    del mxd, df
                arcpy.env.outputCoordinateSystem = None

                for year in dict1515.keys():
                    if not os.path.exists(os.path.join(viewerdir,"150")):
                        os.mkdir(os.path.join(viewerdir,"150"))
                    mxdname = glob.glob(os.path.join(scratch,'15_'+str(year)+'.mxd'))[0]
                    mxd = arcpy.mapping.MapDocument(mxdname)
                    df = arcpy.mapping.ListDataFrames(mxd)[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
                    df.spatialReference = srGoogle

                    imagename = str(year)+".jpg"
                    # arcpy.mapping.ExportToJPEG(mxd, os.path.join(scratch, viewerdir, imagename), df,df_export_width= 14290,df_export_height=16000, color_mode='8-BIT_GRAYSCALE',world_file = True) #by default, the jpeg quality is 100
                    arcpy.mapping.ExportToJPEG(mxd, os.path.join(scratch, viewerdir,"150", imagename), df,df_export_width= 3573,df_export_height=4000, color_mode='24-BIT_TRUE_COLOR',world_file = True, jpeg_quality=50)#,df_export_width= 14290,df_export_height=16000, color_mode='8-BIT_GRAYSCALE',world_file = True, jpeg_quality=100)

                    desc = arcpy.Describe(os.path.join(viewerdir,"150",imagename))
                    featbound = arcpy.Polygon(arcpy.Array([desc.extent.lowerLeft, desc.extent.lowerRight, desc.extent.upperRight, desc.extent.upperLeft]),srGoogle)
                    del desc

                    tempfeat = os.path.join(tempdir, "tilebnd_"+str(year)+ ".shp")

                    arcpy.Project_management(featbound, tempfeat, srWGS84) #function requires output not be in_memory
                    del featbound
                    desc = arcpy.Describe(tempfeat)

                    metaitem = {}
                    metaitem['type'] = 'topo150'
                    metaitem['imagename'] = imagename[:-4]+'.jpg'
                    metaitem['lat_sw'] = desc.extent.YMin
                    metaitem['long_sw'] = desc.extent.XMin
                    metaitem['lat_ne'] = desc.extent.YMax
                    metaitem['long_ne'] = desc.extent.XMax

                    metadata.append(metaitem)
                    del mxd, df
                arcpy.env.outputCoordinateSystem = None

            if os.path.exists(os.path.join(viewerFolder, OrderNumText+"_topo")):
                shutil.rmtree(os.path.join(viewerFolder, OrderNumText+"_topo"))
            shutil.copytree(os.path.join(scratch, OrderNumText+"_topo"), os.path.join(viewerFolder, OrderNumText+"_topo"))
            url = topouploadurl + OrderNumText
            urllib.urlopen(url)

        else:
            arcpy.AddMessage("No viewer is needed. Do nothing")

        try:
            con = cx_Oracle.connect(connectionString)
            cur = con.cursor()

            cur.execute("delete from overlay_image_info where  order_id = %s and (type = 'topo75' or type = 'topo150')" % str(OrderIDText))

            if needViewer == 'Y':
                for item in metadata:
                    cur.execute("insert into overlay_image_info values (%s, %s, %s, %.5f, %.5f, %.5f, %.5f, %s, '', '')" % (str(OrderIDText), str(OrderNumText), "'" + item['type']+"'", item['lat_sw'], item['long_sw'], item['lat_ne'], item['long_ne'],"'"+item['imagename']+"'" ) )
                con.commit()

        finally:
            cur.close()
            con.close()
        # see if need to provide the tiffs too
        if len(copydirs) > 0:
            if os.path.exists(os.path.join(reportcheckFolder,"TopographicMaps",OrderNumText+"_US_Topo.zip")):
                os.remove(os.path.join(reportcheckFolder,"TopographicMaps",OrderNumText+"_US_Topo.zip"))
            shutil.copyfile(os.path.join(scratch,OrderNumText+"_US_Topo.zip"),os.path.join(reportcheckFolder,"TopographicMaps",OrderNumText+"_US_Topo.zip"))
            arcpy.SetParameterAsText(3, os.path.join(scratch,OrderNumText+"_US_Topo.zip"))
        else:
            if os.path.exists(os.path.join(reportcheckFolder,"TopographicMaps",pdfreport)):
                os.remove(os.path.join(reportcheckFolder,"TopographicMaps",pdfreport))
            shutil.copyfile(os.path.join(scratch, pdfreport),os.path.join(reportcheckFolder,"TopographicMaps",pdfreport))
            arcpy.SetParameterAsText(3, os.path.join(scratch,pdfreport))

        try:
            con = cx_Oracle.connect(connectionString)
            cur = con.cursor()

            cur.callproc('eris_topo.processTopo', (int(OrderIDText),))

        finally:
            cur.close()
            con.close()

    logger.removeHandler(handler)
    handler.close()

except:
    # Get the traceback object
    tb = sys.exc_info()[2]
    tbinfo = traceback.format_tb(tb)[0]
    pymsg = "Order ID: %s PYTHON ERRORS:\nTraceback info:\n"%OrderIDText + tbinfo + "\nError Info:\n" + str(sys.exc_info()[1])
    arcpy.AddError("hit CC's error code in except: ")
    arcpy.AddError(pymsg)
    try:
        con = cx_Oracle.connect(connectionString)
        cur = con.cursor()
        cur.callproc('eris_topo.InsertTopoAudit', (OrderIDText, 'python-Error Handling',pymsg))

    finally:
        cur.close()
        con.close()

    raise       # raise the error again

print(reportcheckFolder + "\\TopographicMaps\\" + str(OrderNumText) + "_US_Topo.pdf")
print("DONE!")