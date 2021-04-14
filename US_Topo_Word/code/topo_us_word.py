#-------------------------------------------------------------------------------
# Name:        US TOPO report for Terracon
# Purpose:     create US TOPO report in Terracon required Word format
#
# Author:      jliu
#
# Created:     23/10/2015
# Copyright:   (c) jliu 2015
# Licence:     <your licence>
#-------------------------------------------------------------------------------

# search and create topo maps as usual
# export geometry seperately in emf format
# save picture/vector to the location which links to the word template
# modify text in word template

import time
import traceback
import sys
import arcpy
import os
import time
import win32com
from win32com import client

import topo_us_word_utility as tp
import topo_us_word_config as cfg

# ----------------------------------------------------------------------------------------------------------------------------------------
'''for quick grab of order_id(s) for testing:
select 
a.order_id, a.order_num, a.site_name,
b.customer_id, 
c.company_id, c.company_desc,
d.radius_type, d.geometry_type, d.geometry, length(d.geometry),
e.topo_viewer
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
and e.topo_viewer = 'Y'
order by length(d.geometry),a.order_num  desc;
'''

if __name__ == '__main__':
    arcpy.AddMessage("...starting..." + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    start = time.clock()

    order_obj = cfg.order_obj
    oc = tp.oracle(cfg.connectionString)
    tf = tp.topo_us_word_rpt(order_obj, oc)

    arcpy.AddMessage("Order: " + str(order_obj.id) + ", " + str(order_obj.number))

    try:
        logger,handler = tf.log(cfg.logfile, cfg.logname)
        
        # get spatial references
        srGCS83,srWGS84,srGoogle,srUTM = tf.projlist(order_obj)

        # create order geometry
        tf.createordergeometry(order_obj, srUTM)
        
        # open mxd and create map extent
        logger.debug("#1")
        mxd, df = tf.mapDocument(srUTM)
        needtif = tf.mapExtent(df, mxd, srUTM)

        # set boundary
        mxd, df, yesboundary, custom_profile = tf.setBoundary(mxd, df, cfg.yesBoundary, cfg.custom_profile)

        # select topo records
        rowsMain, rowsAdj = tf.selectTopo(cfg.orderGeometryPR, cfg.extent, srUTM)
        
        # get topo records
        maps7575, maps1515, yearalldict = tf.getTopoRecords(rowsMain, rowsAdj, cfg.csvfile_h, cfg.csvfile_c)

        logger.debug("#2")
        if int(len(maps7575)) == 0 and int(len(maps1515)) == 0:
            arcpy.AddMessage("...NO records selected.")
        else:
            # remove duplicated years
            maps7575 = tf.dedupMaplist(maps7575)
            maps1515 = tf.dedupMaplist(maps1515)

            # reorganize data structure by year
            logger.debug("#4")
            dict7575 = tf.reorgByYear(maps7575)  # {1975: geopdf.pdf, 1973: ...}
            dict1515 = tf.reorgByYear(maps1515)
            arcpy.AddMessage("dict7575: " + str(dict7575.keys()))
            arcpy.AddMessage("dict1515: " + str(dict1515.keys()))

            dictlist = []
            if dict7575:
                dictlist.append(dict7575)
            if dict1515:
                dictlist.append(dict1515)
            
            # # remove blank maps flag
            # if cfg.delyearFlag == 'Y':
            #     delyear75 = filter(None, str(raw_input("Years you want to delete in the 7.5min series (comma-delimited):\n>>> ")).replace(" ", "").strip().split(","))
            #     delyear15 = filter(None, str(raw_input("Years you want to delete in the 15min series (comma-delimited):\n>>> ")).replace(" ", "").strip().split(","))
            #     tf.delyear(delyear75, delyear15, dict7575, dict1515)

            # run word application
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = 0

            # create cover and summary page
            tf.goCoverPage(order_obj)
            tf.goSummaryPage(dictlist)

            # create map page
            worddoclist, app = tf.createDOCX(yesboundary, dictlist, yearalldict, mxd, df, custom_profile, app)

            # append all pages to final docx
            tf.appendPages(app, worddoclist, order_obj)

            # convert docx to zip and add edit restriction
            tf.unzipDocx(os.path.join(cfg.scratch, order_obj.number + "_US_Topo.docx"))

            # save summary data to oracle
            tf.oracleSummary(dictlist, order_obj.number + "_US_Topo.docx")

            # zip tiffs if needtif  
            copydirs = [os.path.join(os.path.join(cfg.scratch,order_obj.number), name) for name in os.listdir(os.path.join(cfg.scratch,order_obj.number))]
            if len(copydirs) > 0 and needtif == True:
                tf.zipDir(order_obj.number + "_US_Topo.docx")

            # export to xplorer
            tf.toXplorer(needtif, dictlist, srGoogle, srWGS84)

            # copy files to report check
            tf.toReportCheck(needtif, order_obj.number + "_US_Topo.docx")
            
    except Exception:
        # Get the traceback object
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "Order ID: %s PYTHON ERRORS:\nTraceback info:\n"%order_obj.id + tbinfo + "\nError Info:\n" + str(sys.exc_info()[1])

        procedure = 'eris_topo.InsertTopoAudit'
        oc.proc(procedure, (order_obj.id, 'python-Error Handling',pymsg))

        raise                                   # raise the error again

    finally:
        oc.close()                              # close oracle connection
        logger.removeHandler(handler)
        handler.close()
                
        tasklist = os.popen('tasklist /v').read().strip()
        if "winword.exe" in tasklist:
            print("...need to close winword.exe.")
            os.system('TASKKILL /F /IM winword.exe')

    finish = time.clock()
    # arcpy.AddMessage(cfg.reportcheckFolder + "\\TopographicMaps\\" + order_obj.number + "_US_Topo.docx")
    arcpy.AddMessage("Finished in " + str(round(finish-start, 2)) + " secs")