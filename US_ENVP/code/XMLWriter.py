import xml.etree.ElementTree as xml
import arcpy, os, re, sys, string, httplib
import xml.dom.minidom as minidom
from collections import defaultdict
import cx_Oracle
import datetime

class ERISXmlWriter(object):

    def __init__(self):
        pass

    def prettify(self, elem):
        """Return a pretty-printed XML string for the element"""
        rough_string = xml.tostring(elem, 'utf-8')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ")




    def create_config_xml(self, name, params):
        """Write out elements to XML config file"""
        ERIS = xml.Element("Order_detail")
        # WTG elements
        order= xml.SubElement(ERIS, "Order")
        wtgs = params["order"]

        for wtg in wtgs:
        	wtg_id_element =  xml.SubElement(order, "OrderID")
        	wtg_id_element.text = str(wtg[0])
        	order_id= wtg_id_element.text
        	wtg_n_element =  xml.SubElement(order, "Easting_m")
        	wtg_n_element.text = str(wtg[1])
        	wtg_e_element = xml.SubElement(order, "Northing_m")
        	wtg_e_element.text = str(wtg[2])
        	wtg_lon_element = xml.SubElement(order, "Longitude")
        	wtg_lon_element.text = str(wtg[3])
        	wtg_lat_element = xml.SubElement(order, "Latitude")
        	wtg_lat_element.text = str(wtg[4])
        	wtg_utm_element = xml.SubElement(order, "UTM")
        	wtg_utm_element.text = str(wtg[5])
        	wtg_ele_element = xml.SubElement(order, "Elevation_m")
        	wtg_ele_element.text = str(wtg[6])
        	wtg_pg_element = xml.SubElement(order, "page")
        	wtg_pg_element.text = str(wtg[7])

        sitesdetail= xml.SubElement(ERIS, "sites")
        sites = params["sites"]

        for site in sites:
      		site_element =  xml.SubElement(sitesdetail, "site_detail")
        	id_element =  xml.SubElement(site_element, "ERISID")
        	id_element.text = str(site[0])
        	dist_element =  xml.SubElement(site_element, "Distance")
        	dist_element.text = str(site[1])
        	ele_element = xml.SubElement(site_element, "Elevation")
        	ele_element.text = str(site[2])
        	direction_element = xml.SubElement(site_element, "Direction")
        	direction_element.text = str(site[3])
           	maploc_element = xml.SubElement(site_element, "MapKeyloc")
        	maploc_element.text = str(site[4])
        	mapno_element = xml.SubElement(site_element, "MapKeyNo")
        	mapno_element.text = str(site[5])
        	buffer_element = xml.SubElement(site_element, "Buffer")
        	buffer_element.text = str(site[6])
##        	onsite_element = xml.SubElement(site_element, "Onsite")
##         	onsite_element.text = str(site[6])
##        largedetail= xml.SubElement(ERIS, "LRsites")
##        largera = params["LRsites"]
##        for lar in largera:
##      		lar_element =  xml.SubElement(largedetail, "ERISID")
##        	#lar_element =  xml.SubElement(lar_element, "ERISID")
##        	lar_element.text = str(lar[0])
####        	lar_sou_element =  xml.SubElement(lar_element, "Source")
####        	lar_sou_element.text = str(lar[1])
####        	lar_ID_element = xml.SubElement(lar_element, "FID")
####        	lar_ID_element.text = str(lar[2])

    # Get rid of extra line returns in prettified XML
        uglyXml = self.prettify(ERIS)
        text_re = re.compile('>\n\s+([^<>\s].*?)\n\s+</', re.DOTALL)
        prettyXml = text_re.sub('>\g<1></', uglyXml)
        body1 = """<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <AuthenticationHeader xmlns="http://www.bizinfogroup.ca/">
      <Username>%s</Username>
      <Password>%s</Password>
    </AuthenticationHeader>
  </soap:Header>
  <soap:Body>
     <AcceptSites xmlns="http://www.bizinfogroup.ca/">
      <items><![CDATA[%s]]></items>
      <environment>%s</environment>
    </AcceptSites>
  </soap:Body>
</soap:Envelope> """ % ("43north", "D1ff1cultP@ssw0r6", prettyXml, "prod")

        do_request1(prettyXml)


        return prettyXml


def do_request1(body):
       out_web = open(os.path.join(log ), "w")
       print >> out_web, "oracle start" , datetime.datetime.now()
       try:
             #get database connection
           con = cx_Oracle.connect('eris_gis/gis295@GMPRODC.glaciermedia.inc')
           cur=con.cursor()
           #body1="""%s""" %(body)
           print >> out_web, type(body)
           cv= cur.var(cx_Oracle.CLOB)
           cv.setvalue(0,str(body))
           print >> out_web, body,"body1"
           #execute the stored procedure or function in Oracle
           r=cur.callfunc('ERIS_WSERVICE.AcceptSites',cx_Oracle.STRING,(cv,))
           print >> out_web, r, "oracle finish", datetime.datetime.now()
           if r == 'D':
           #soemthing wrong with AcceptSites, should fail.
               arcpy.AddError('AcceptSites returns D');
               #print 'this is ' + 1
               raise ValueError('AcceptSites returns D')
               #return 'N'

##won't catch any error, will let the code fail purposely.
##       except Exception as e:
##           #error, = e.args
##           print >> out_web,"error"
##           print >> out_web,e

       finally:
           cur.close()
           con.close()
           out_web.close()




try:
   order = arcpy.GetParameter(0)
   locations = arcpy.GetParameter(1)
   #larRadius = arcpy.GetParameter(2)
   file_name = arcpy.GetParameter(2)
   log = arcpy.GetParameter(3)


   fclass= arcpy.SearchCursor(order)
   sites= arcpy.SearchCursor(locations,"" ,"","", 'MapKeyLoc A ; MapKeyNo A')
   #lars= arcpy.SearchCursor(larRadius,"" ,"","", 'FID A')
   params = defaultdict(list)


   order=[]
   site=[]
   #lar=[]
   for row in fclass:
      order.append(str(row.getValue('ID')))
      order.append(str(row.getValue('POINT_X')))
      order.append(str(row.getValue('POINT_Y')))
      order.append(str(row.getValue('Lon_X')))
      order.append(str(row.getValue('Lat_Y')))
      order.append(str(row.getValue('UTM'))[32:44])
      order.append(str(row.getValue('Elevation')))
      order.append(str(row.getValue('page')))
   del fclass
   del row


   params['order'].append(order)

   i = 0
   a= []
   result = int(arcpy.GetCount_management(locations).getOutput(0))
   if result!= 0:
     for row in sites:
        a= []
        if row.getValue('ERISID')!='':
          a.append(str(row.getValue('ERISID')))
          a.append(str(row.getValue('Distance')))
          a.append(str(row.getValue('Elevation')))
          a.append(str(row.getValue('Direction')))
          a.append(str(row.getValue('MapkeyLoc')))
          a.append(str(row.getValue('MapKeyNo')))
          a.append(str(row.getValue('Buffer')))
##          a.append(str(row.getValue('Onsite')))
          params['sites'].append(a)
     del sites, row
   else:
       if a== []:
        a.append('-1')
        a.append('-1')
        a.append('-1')
        a.append('-1')
        a.append('-1')
        a.append('-1')
        a.append('-1')
##        a.append('-1')
        params['sites'].append(a)


##   b =[]
##   for row in lars:
##      b= []
##      if row.getValue('ERISID')!='':
##         b.append(str(row.getValue('ERISID')))
####         b.append(str(row.getValue('SOURCE')))
##         params['LRsites'].append(b)
##
##   if b ==[]:
##        b.append('-1')
####        b.append('-1')
##        params['LRsites'].append(b)
##   del row
##   del lars
   kk= ERISXmlWriter()


   a= kk.create_config_xml(file_name, params)
   f = open(os.path.join(file_name ), "w")

   f.write(a)
   f.close()

##   order_id='12345'
##   connection = cx_Oracle.connect("eris/eris@unix_test")
##   cursor = connection.cursor()
##   cursor.execute("INSERT INTO ERIS_SITES_XML (order_id,site_data) VALUES (:order_id,:site_data)", dict(order_id= order_id, site_data=a))
##   connection.commit()
##   cursor.execute("select * from ERIS_SITES_XML where order_id =:arg_1", arg_1= order_id)
##   out_web = open(os.path.join(log ), "w")
##   for column_1, column_2 in cursor.fetchall():
##	    print "Values from DB:", column_1, column_2 >>out_web
##   connection.commit()
##   out_web.close()
##   cursor.close()
##   connection.commit()
##   connection.close()
except:
   # If an error occurred, print the message to the screen
   arcpy.AddMessage(arcpy.GetMessages())
   #arcpy.SetParameterAsText(3,'0')
