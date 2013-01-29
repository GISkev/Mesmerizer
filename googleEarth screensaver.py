
import random
import time
import win32com.client
ge = win32com.client.Dispatch("GoogleEarth.ApplicationGE")
while not ge.IsInitialized():
    print "waiting for GE to start up"


def setCam(lat, lon, geRange, geTilt, azimuth, geTime):
    ge.SetCameraParams( lat, lon, 0.0, 1, 
                        geRange, 
                        geTilt, 
                        azimuth, 
                        0.1)
    time.sleep(geTime)
    

#startLat = 38.544359
#startLon = -110.907446
startLat = 38.469007
startLon = -110.923046

i = 0
for x in xrange(100):  # number of times to shift the view
    
    azimuth = 90
    if i == 0:
        lat = startLat
        lon = startLon

    # how far to wander from your starting point
    latshift = random.randrange(-200, 200)
    lonshift = random.randrange(-200, 200)
    
    # adjusting the decimal degrees here could likely
    # be cleaned up/refactored.  Not very pythonic here.
    if len(str(abs(latshift))) == 1:
        lat += (latshift * 0.0000001) 
    if len(str(abs(latshift))) == 2:
        lat += (latshift * 0.000001) 
    if len(str(abs(latshift))) == 3:
        lat += (latshift * 0.00001)    
    
    if len(str(abs(lonshift))) == 1:
        lon += (lonshift * 0.0000001) 
    if len(str(abs(lonshift))) == 2:
        lon += (lonshift * 0.000001) 
    if len(str(abs(lonshift))) == 3:
        lon += (lonshift * 0.00001)        
    
    # how close to the ground you'll be
    geRange = random.randrange(300, 500)
    # yup, the tilt range...
    geTilt = random.randrange(0, 45)
    # yup, the amount of rotation...
    rotation = random.randrange(-30, 30)
    azimuth += rotation
    # 10-15 seconds between target stops.  Doesn't
    # include the flying to/from each stop.
    geTime = random.randrange(10,15)

    if i == 0:
        geTime = 25 # this is to allow GE to get started up
        
    setCam(lat, lon, geRange, geTilt, azimuth, geTime)   
    i += 1
    
   
