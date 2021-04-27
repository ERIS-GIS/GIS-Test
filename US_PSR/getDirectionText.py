#-------------------------------------------------------------------------------
# Name:        getDirectionText
# Purpose:     return a letter direction between two points
#              -- calculate "bearing" angle between two points with projected coordinates
#              -- convert the angle to a letter direction (eg. "ENE")
#
# Author:      LiuJ
#
# Created:     11/09/2013
# Copyright:   (c) ERIS 2013
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import math
import arcpy


def main():
    print getDirectionText(0,0,1,1)
    return

#takes coordinates of two points, return a letter direction
#(x1,y1) is the reference point
def getDirectionText(x1,y1,x2,y2):

    angle = calAngle(x1,y1,x2,y2)

    return dgrDir2txt(angle)




#(x1,y1) is the reference point.
#return an angle in degrees [0,360]
def calAngle(x1,y1,x2,y2):
    a = y2-y1
    b = x2-x1
    angle180 = math.degrees(math.atan2(a,b))

    angle360= (450-angle180) % 360

    return angle360


#convert the degree angle into a letter direction
def dgrDir2txt(dgr360):
    if (dgr360 >= 348.75) or (dgr360 < 11.25):
        dirText = "N"
    elif (dgr360 >= 11.25) and (dgr360 < 33.75):
        dirText = "NNE"
    elif (dgr360 >= 33.75) and (dgr360 < 56.25):
        dirText = "NE"
    elif (dgr360 >= 56.25) and (dgr360 < 78.75):
        dirText = "ENE"
    elif (dgr360 >= 78.75) and (dgr360 < 101.25):
        dirText = "E"
    elif (dgr360 >= 101.25) and (dgr360 < 123.75):
        dirText = "ESE"
    elif (dgr360 >= 123.75) and (dgr360 < 146.25):
        dirText = "SE"
    elif (dgr360 >= 146.25) and (dgr360 < 168.75):
        dirText = "SSE"
    elif (dgr360 >= 168.75) and (dgr360 < 191.25):
        dirText = "S"
    elif (dgr360 >= 191.25) and (dgr360 < 213.75):
        dirText = "SSW"
    elif (dgr360 >= 213.75) and (dgr360 < 236.25):
        dirText = "SW"
    elif (dgr360 >= 236.25) and (dgr360 < 258.75):
        dirText = "WSW"
    elif (dgr360 >= 258.75) and (dgr360 < 281.25):
        dirText = "W"
    elif (dgr360 >= 281.25) and (dgr360 < 303.75):
        dirText = "WNW"
    elif (dgr360 >= 303.75) and (dgr360 < 326.25):
        dirText = "NW"
    elif (dgr360 >= 326.25) and (dgr360 < 348.75):
        dirText = "NNW"
    else:
        dirText = "NUL"    # this line should not be reached

    return dirText

if __name__ == '__main__':
    main()