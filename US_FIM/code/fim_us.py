#-------------------------------------------------------------------------------
# Name:        USFIM rework
# Purpose:
#
# Author:      cchen
#
# Created:     01/12/2019
# Copyright:   (c) cchen 2019
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import arcpy, os, sys
import traceback
import time
import fim_us_config as cfg
import fim_us_utility as fp

from PyPDF2 import PdfFileReader,PdfFileWriter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import portrait, letter

'''
select 
a.order_id, a.order_num, a.site_name,
b.customer_id, 
c.company_id, c.company_desc,
d.radius_type, d.geometry_type, d.geometry, length(d.geometry),
e.fim_viewer
from orders a, customer b, company c, eris_order_geometry d, order_viewer e
where 
a.customer_id = b.customer_id and
b.company_id = c.company_id and
a.order_id = d.order_id and
a.order_id = e.order_id
and upper(a.site_name) not like '%TEST%'
and upper(a.site_name) not like '%DEMO%'
--and upper(c.company_desc) like '%MID-ATLANTIC%'
and d.geometry_type = 'POLYGON'
and e.fim_viewer = 'Y'
order by length(d.geometry),a.order_num  desc;
'''

if __name__ == '__main__':
    arcpy.AddMessage("...starting..." + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    start = time.clock()
    
    order_obj = cfg.order_obj
    oc = fp.oracle(cfg.connectionString)
    ff = fp.fim_us_rpt(cfg.order_obj, oc)

    arcpy.AddMessage("Order: " + str(order_obj.id) + ", " + str(order_obj.number))
    arcpy.AddMessage("multipage = " + str(cfg.multipage))

    try:        
        # declare report name
        if order_obj.country =='MX':
            pdfreport_name =  cfg.order_obj.number+"_MEX_FIM.pdf"
        else:
            pdfreport_name = cfg.order_obj.number+"_US_FIM.pdf"

        # logger
        logger,handler = ff.log(cfg.logfile, cfg.logname)
        
        # get spatial references
        srGCS83,srWGS84,srGoogle,srUTM = ff.projlist(order_obj)

        # create order geometry
        ff.createOrderGeometry(order_obj, srUTM, cfg.BufsizeText)

        # open mxd and create map extent
        mxd, dfmain, dfinset = ff.mapDocument(srUTM)
        ff.mapExtent(mxd, dfmain, dfinset, cfg.multipage)

        # select FIM
        mainlist, adjlist = ff.selectFim(cfg.mastergdb, srUTM)

        if len(mainlist) == 0 or cfg.nrf == 'Y':
            logger.info("order " + order_obj.number + ":    search completed. Will print out a NRF letter. ")
            arcpy.AddMessage("...NO records selected, will print out NRF letter.")
            cfg.nrf = 'Y'
            ff.goCoverPage(cfg.coverfile, cfg.nrf)
            os.rename(cfg.coverfile, os.path.join(cfg.scratch, pdfreport_name))
        else:
            # get FIM records
            mainList, adjacentList = ff.getFimRecords(cfg.selectedmain, cfg.selectedadj)

            # get custom order flags
            is_newLogo, is_aei, is_emg = ff.customrpt(order_obj)

            # set boundary
            mxd, df, yesBoundary = ff.setBoundary(mxd, dfmain, cfg.yesBoundary)

            # # remove blank maps flag, for internal use
            # if cfg.delyearFlag == 'Y':
            #     delyear = filter(None, str(raw_input("Years you want to delete (comma-delimited):\n>>> ")).replace(" ", "").strip().split(","))        # No quotes
            #     ff.delyear(delyear, mainList)
            
            # create map page
            ff.createPDF(mainList, adjacentList, is_aei, mxd, dfmain, dfinset, yesBoundary, cfg.multipage, cfg.gridsize, cfg.resolution)

            # create cover and summary pages
            ff.goSummaryPage(cfg.summaryfile,mainList)
            ff.goCoverPage(cfg.coverfile, cfg.nrf)

            # append pages to blank pdf
            output = PdfFileWriter()
            output.addPage(PdfFileReader(open(cfg.coverfile,'rb')).getPage(0))
            output.addBookmark("Cover Page",0)

            for j in range(PdfFileReader(open(cfg.summaryfile,'rb')).getNumPages()):
                output.addPage(PdfFileReader(open(cfg.summaryfile,'rb')).getPage(j))
                output.addBookmark("Summary Page",j+1)

            ff.appendMapPages(output,mainList, yesBoundary, cfg.multipage)

            outputStream = open(os.path.join(cfg.scratch, pdfreport_name),"wb")
            output.setPageMode('/UseOutlines')
            output.write(outputStream)
            outputStream.close()

            # upload to xplorer
            ff.toXplorer(mainList, srGoogle, srWGS84)

            # copy to report check
            ff.toReportCheck(pdfreport_name)

    except:
        # Get the traceback object
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "Order ID: %s PYTHON ERRORS:\nTraceback info:\n"%order_obj.id + tbinfo + "\nError Info:\n" + str(sys.exc_info()[1])
        traceback.print_exc()

        procedure = 'eris_fim.InsertFIMAudit'
        oc.proc(procedure, (order_obj.id, 'python-Error Handling',pymsg))

        raise                                   # raise the error again

    finally:
        oc.close()
        logger.removeHandler(handler)
        handler.close()

    finish = time.clock()
    # arcpy.AddMessage("Final FIM report directory: " + (str(os.path.join(cfg.reportcheckFolder,"FIM", pdfreport_name))))
    arcpy.AddMessage("Finished in " + str(round(finish-start, 2)) + " secs")