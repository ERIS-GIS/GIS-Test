## Loop through each image, apply spatial reference & produce .xml that ArcGiS can read. Temporarily save output to new location
import arcpy, os
##source files
def image_properties(image_path):
    ## Check .TAB before converting to WGS84
    tab_info = open(os.path.join(orig_image_file,image_path),'r')
    print tab_info
"""     sr = arcpy.SpatialReference(4326)
    arcpy.DefineProjection_management(os.path.join(orig_image_file,image_path),sr)
    arcpy.Copy_management(os.path.join(orig_image_file,image_path), os.path.join(arc_image_output, image_path)) """
    return 'success'

orig_image_file = r'C:\Users\JLoucks\Desktop\imagetest_US'
arc_image_output = r'C:\Users\JLoucks\Desktop\output_test'
file_extensions = ('.tif','.tiff','.png','.jpg','.jpeg','.sid')
for aerial_image in os.listdir(orig_image_file):
    if aerial_image.endswith(file_extensions):
        image_properties(aerial_image)
    else:
        print aerial_image

        