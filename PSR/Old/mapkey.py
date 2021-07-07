###Libraries###
import sys, os, string, arcpy, logging
from arcpy import env, mapping
import time


try:
  ###Constants and Global Variables###
  arcpy.env.OverWriteOutput = True

  fclass = arcpy.GetParameter(0)

  temp = arcpy.CopyFeatures_management(fclass, "in_memory/temp1")
  cur = arcpy.UpdateCursor(temp,"" ,"","Dist_cent; MapKeyLoc; MapKeyNo", 'Dist_cent A; Source A')

  row = cur.next()
  last = row.getValue('Dist_cent') # the last value in field A
  row.setValue('mapkeyloc', 1)
  row.setValue('mapkeyno', 1)


  cur.updateRow(row)
  run = 1 # how many values in this run
  count = 1 # how many runs so far, including the current one

# the for loop should begin from row 2, since
# cur.next() has already been called once.
  for row in cur:
    current = row.getValue('Dist_cent')
    if current == last:
        run += 1
    else:
        run = 1
        count += 1
    row.setValue('mapkeyloc', count)
    row.setValue('mapkeyno', run)
    cur.updateRow(row)

    last = current

# release the layer from locks
  del row, cur

  cur = arcpy.UpdateCursor(temp, "", "", 'MapKeyLoc; mapKeyNo; MapkeyTot', 'MapKeyLoc D; mapKeyNo D')

  row = cur.next()
  last = row.getValue('mapkeyloc') # the last value in field A
  max= 1
  row.setValue('mapkeytot', max)
  cur.updateRow(row)

  for row in cur:
    current = row.getValue('mapkeyloc')

    if current < last:
        max= 1
    else:
        max= 0
    row.setValue('mapkeytot', max)
    cur.updateRow(row)

    last = current

# release the layer from locks
  del row, cur

  arcpy.SetParameter(1,fclass)
  arcpy.CopyFeatures_management(temp, fclass)

except:
  # If an error occurred, print the message to the screen
  arcpy.AddMessage(arcpy.GetMessages())
finally:
  arcpy.Delete_management("in_memory")
  #logger.removeHandler(handler)

