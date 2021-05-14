import arcpy, os,sys,timeit
from imp import reload
file_path =os.path.dirname(os.path.abspath(__file__))
if 'arcgisserver' in file_path:
    model_path = os.path.join('D:/arcgisserver/directories/arcgissystem/arcgisinput/GPtools/DB_Framework')
else:
    model_path = os.path.join(os.path.dirname(file_path),'DB_Framework')
    
sys.path.insert(1,model_path)
import models
import psr_utility as utility
import psr_config as config
import flood_plain
import topo
import relief
import wetland
import geology
import soil
import ogw
import radon
import aspect
import kml
reload(sys)
sys.setdefaultencoding('utf8')

if __name__ == "__main__":
    ### input parameters
    id = arcpy.GetParameterAsText(0)
    config.if_relief_report = bool(arcpy.GetParameterAsText(1))
    config.if_topo_report = bool(arcpy.GetParameterAsText(2))
    config.if_wetland_report = bool(arcpy.GetParameterAsText(3))
    config.if_flood_report = bool(arcpy.GetParameterAsText(4))
    config.if_geology_report = bool(arcpy.GetParameterAsText(5))
    config.if_soil_report = bool(arcpy.GetParameterAsText(6))
    config.if_ogw_report = bool(arcpy.GetParameterAsText(7))
    config.if_radon_report = bool(arcpy.GetParameterAsText(8))
    config.if_aspect_map = bool(arcpy.GetParameterAsText(9))
    config.if_kml_output =bool(arcpy.GetParameterAsText(10))
    
    id = 1026370
    arcpy.AddMessage('Start PSR report...')
    start = timeit.default_timer()
    # ### set workspace
    arcpy.env.workspace = config.scratch_folder
    arcpy.AddMessage('  -- scratch folder: %s' % config.scratch_folder)
    arcpy.env.overwriteOutput = True
    if not os.path.exists(config.temp_gdb):
        arcpy.CreateFileGDB_management(config.scratch_folder,r"temp")

    ### isntantiate of order class and set order geometry and buffering
    order_obj = models.Order().get_order(int(id))
    if order_obj is not None:
        utility.set_order_geometry(order_obj)
        config.if_multi_page = utility.if_multipage(config.order_geometry_pcs_shp)
        arcpy.AddMessage('  -- multiple pages: %s' % str(config.if_multi_page))
        ### Populate radius list of PSR for this Order object
        order_obj.get_search_radius() # populate search radius
        if len(order_obj.psr.search_radius) > 0:
            
            # config.if_relief_report = True
            # config.if_topo_report = True
            config.if_wetland_report = True
            # config.if_flood_report = True
            # config.if_geology_report = True
            # config.if_soil_report = True
            # config.if_ogw_report = True
            # config.if_radon_report = True
            # config.if_aspect_map = True
            config.if_kml_output = True

            # shaded releif map report
            if config.if_relief_report:
                relief.generate_relief_report(order_obj)
            # topo map report
            if config.if_topo_report:
                topo.generate_topo_report(order_obj)
            # Wetland report
            if config.if_wetland_report:
                wetland.generate_wetland_report(order_obj)
            # flood report
            if config.if_flood_report:
                flood_plain.generate_flood_report(order_obj)
            # geology report
            if config.if_geology_report:
                geology.generate_geology_report(order_obj)
            # soil report
            if config.if_soil_report:
                soil.generate_soil_report(order_obj)
            # oil, gas and water wells report
            if config.if_ogw_report:
                ogw.generate_ogw_report(order_obj)
            # radon report
            if config.if_radon_report:
                radon.generate_radon_report(order_obj)
            # multi_proc_test.generate_flood_report(order_obj)
            if config.if_aspect_map:
                aspect.generate_aspect_map(order_obj)
            # convert to kml for viewer
            if config.if_kml_output:
                kml.convert_to_kml(order_obj)
        else:
            arcpy.AddMessage('No PSR is availabe for this order')
    else:
        arcpy.AddMessage('This order is not availabe')
    end = timeit.default_timer()
    arcpy.AddMessage(('End PSR report process. Duration:', round(end -start,4)))