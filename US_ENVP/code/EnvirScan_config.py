#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      cchen
#
# Created:     27/04/2017
# Copyright:   (c) cchen 2017
# Licence:     <your licence>
#-------------------------------------------------------------------------------
from collections import OrderedDict
import os

# TEST #
connectionString = 'eris_gis/gis295@cabcvan1ora006.glaciermedia.inc:1521/GMTESTC'
report_path = r"\\cabcvan1eap006\ErisData\Reports\test\noninstant_reports"
reportcheck_path = r'\\cabcvan1eap006\ErisData\Reports\test\reportcheck'
connectionPath = r"\\cabcvan1gis005\GISData\EnvironmentScanReport"
viewer_path = r"\\CABCVAN1EAP006\ErisData\Reports\test\viewer"
upload_link = r"http://CABCVAN1EAP006/ErisInt/BIPublisherPortal/Viewer.svc"

# ORDER SETTING #
orderGeomlyrfile_point = os.path.join(connectionPath,r"layer","SiteMaker.lyr")
orderGeomlyrfile_polyline = os.path.join(connectionPath,r"layer","orderLine.lyr")
orderGeomlyrfile_polygon = os.path.join(connectionPath,r"layer","orderPoly.lyr")
bufferlyrfile = os.path.join(connectionPath,r"layer","buffer.lyr")
ERIScanIncident = os.path.join(connectionPath,r"layer","ERIScanIncident.lyr")
ERIScanPermit = os.path.join(connectionPath,r"layer","ERIScanPermit.lyr")
ERIScanPoint = os.path.join(connectionPath,r"layer","ERIScanPoint.lyr")
geom = os.path.join(connectionPath,r"layer","ERISVaporPoints.lyr")
mxd = os.path.join(connectionPath,r"mxd","ESRimagery.mxd")