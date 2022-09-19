from datetime import datetime

import sys
sys.path.append(r"C:\Users\omami\Desktop\Code\Python_Code\Weather")
from user_preference.UserPreference import requirements, preference



def indexdt(data, dt):
    for i in range(len(data)):
        
        if data[i]["dt"] == dt:
            return i

class ProcessReq:

    def processMaxAndMin(self, weatherParam, min, max):

        if (min != None) and (max != None):
            return weatherParam >= min and weatherParam <= max
        elif (min != None) and (max == None):
            return weatherParam >= min
        elif (min == None) and (max != None):
            return weatherParam <= max
        else:
            return True

    def getUserMaxAndMin(self, group):

        min = requirements[group]["min"]
        max = requirements[group]["max"]

        return min, max
        
    def meetReq(self, data):

        passReq = []

        for time in data["list"]:

            passedReq = self.processTimeEvent(time)

            #python lazy 'and' processing to avoid having to run 
            #process event methods when item already has failed
            passedReq = passedReq and self.processTempEvent(time)
            passedReq = passedReq and self.processHumidityEvent(time)
            passedReq = passedReq and self.processPressureEvent(time)
            passedReq = passedReq and self.processWindSpeedEvent(time)
            passedReq = passedReq and self.processPrecipitationEvent(time)
            passedReq = passedReq and self.processVisibilityEvent(time)

            if passedReq:
                passReq.append(time)

        return passReq

    def processTimeEvent(self, weather):

        time = weather["dt_txt"]
        time = datetime.strptime(time, "%Y-%m-%d %H:%M:%S").time()

        start = requirements["time"]["start"]
        start = datetime.strptime(start, "%I:%M%p").time()

        end = requirements["time"]["end"]
        end = datetime.strptime(end, "%I:%M%p").time()
       
        return (time >= start) and (time <= end)

    def processTempEvent(self, weather):

        temp = weather["main"]["temp"]
        min, max = self.getUserMaxAndMin("temp")
        return self.processMaxAndMin(temp, min, max)

    def processHumidityEvent(self, weather):

        humid = weather["main"]["humidity"]
        min, max = self.getUserMaxAndMin("humidity")
        return self.processMaxAndMin(humid, min, max)
    
    def processPressureEvent(self, weather):

        pressure = weather["main"]["pressure"]
        min, max = self.getUserMaxAndMin("pressure")
        return self.processMaxAndMin(pressure, min, max)
    
    def processWindSpeedEvent(self, weather):

        speed = weather["wind"]["speed"]
        min, max = self.getUserMaxAndMin("wind_speed")
        return self.processMaxAndMin(speed, min, max)

    def processPrecipitationEvent(self, weather):

        precp = weather["pop"]
        min, max = self.getUserMaxAndMin("precipitation%")
        return self.processMaxAndMin(precp, min, max)

    def processVisibilityEvent(self, weather):

        vis = weather["visibility"]
        min, max = self.getUserMaxAndMin("visibility")
        return self.processMaxAndMin(vis, min, max)

class ProcessPref:

    def __init__(self, weather):
        self.weatherData = weather
        self.prefCounter = 0

        self.dataweights = {}
        for w in self.weatherData:
            self.dataweights[w["dt"]] = 0

    def getPercentFit(self, x, rate, v):
        e = 2.71828
        y = 1 / (1 + (e**((-rate)*x))**(1/v))

        return 1 - 2 * abs(0.5 - y)

    def meetExtra(self):
        
        self.processTempEvent()
        self.processHumidityEvent()
        self.processPressureEvent()
        self.processWindSpeedEvent()
        self.processPrecipitationEvent()
        self.processVisibilityEvent()

        self.processPercents()

        return sorted(self.dataweights.items(), key=lambda x: x[1])

    def processTempEvent(self):

        min, max = self.getUserMaxAndMin("temp")


        if (min == None) and (max == None):
            return
        else:
            self.prefCounter +=  1
        
        for w in self.weatherData:

            x = w["main"]["temp"]
        
            #if data is within the bounds, set fit to 1
            if self.processMaxAndMin(x, min, max):
                fit = 1.0
            else:
                if (min != None) and (max != None):
                    if x > max:
                        diff = abs(max - x)
                    else:
                        diff = abs(min - x)
                elif (min != None) and (max == None):
                    diff = abs(min - x)
                elif (min == None) and (max != None):
                    diff = abs(max - x)
                
                fit = self.getPercentFit(diff, 0.7, 4.8)

            self.dataweights[w["dt"]] += fit

    def processHumidityEvent(self):

        min, max = self.getUserMaxAndMin("humidity")


        if (min == None) and (max == None):
            return
        else:
            self.prefCounter +=  1
        
        for w in self.weatherData:

            x = w["main"]["humidity"]
        
            #if data is within the bounds, set fit to 1
            if self.processMaxAndMin(x, min, max):
                fit = 1.0
            else:
                if (min != None) and (max != None):
                    if x > max:
                        diff = abs(max - x)
                    else:
                        diff = abs(min - x)
                elif (min != None) and (max == None):
                    diff = abs(min - x)
                elif (min == None) and (max != None):
                    diff = abs(max - x)
                
                fit = self.getPercentFit(diff, 0.7, 4.8)

            self.dataweights[w["dt"]] += fit

    def processPressureEvent(self):

        min, max = self.getUserMaxAndMin("pressure")


        if (min == None) and (max == None):
            return
        else:
            self.prefCounter +=  1
        
        for w in self.weatherData:

            x = w["main"]["pressure"]
        
            #if data is within the bounds, set fit to 1
            if self.processMaxAndMin(x, min, max):
                fit = 1.0
            else:
                if (min != None) and (max != None):
                    if x > max:
                        diff = abs(max - x)
                    else:
                        diff = abs(min - x)
                elif (min != None) and (max == None):
                    diff = abs(min - x)
                elif (min == None) and (max != None):
                    diff = abs(max - x)
                
                fit = self.getPercentFit(diff, 0.7, 4.8)

            self.dataweights[w["dt"]] += fit

    def processWindSpeedEvent(self):
        
        min, max = self.getUserMaxAndMin("wind_speed")


        if (min == None) and (max == None):
            return
        else:
            self.prefCounter +=  1
        
        for w in self.weatherData:

            x = w["wind"]["speed"]
        
            #if data is within the bounds, set fit to 1
            if self.processMaxAndMin(x, min, max):
                fit = 1.0
            else:
                if (min != None) and (max != None):
                    if x > max:
                        diff = abs(max - x)
                    else:
                        diff = abs(min - x)
                elif (min != None) and (max == None):
                    diff = abs(min - x)
                elif (min == None) and (max != None):
                    diff = abs(max - x)
                
                fit = self.getPercentFit(diff, 0.7, 4.8)

            self.dataweights[w["dt"]] += fit
    
    def processPrecipitationEvent(self):
        
        min, max = self.getUserMaxAndMin("precipitation%")


        if (min == None) and (max == None):
            return
        else:
            self.prefCounter +=  1
        
        for w in self.weatherData:

            x = w["pop"]
        
            #if data is within the bounds, set fit to 1
            if self.processMaxAndMin(x, min, max):
                fit = 1.0
            else:
                if (min != None) and (max != None):
                    if x > max:
                        diff = abs(max - x)
                    else:
                        diff = abs(min - x)
                elif (min != None) and (max == None):
                    diff = abs(min - x)
                elif (min == None) and (max != None):
                    diff = abs(max - x)
                
                fit = self.getPercentFit(diff, 0.7, 4.8)

            self.dataweights[w["dt"]] += fit

    def processVisibilityEvent(self):
        
        min, max = self.getUserMaxAndMin("visibility")


        if (min == None) and (max == None):
            return
        else:
            self.prefCounter +=  1
        
        for w in self.weatherData:

            x = w["visibility"]
        
            #if data is within the bounds, set fit to 1
            if self.processMaxAndMin(x, min, max):
                fit = 1.0
            else:
                if (min != None) and (max != None):
                    if x > max:
                        diff = abs(max - x)
                    else:
                        diff = abs(min - x)
                elif (min != None) and (max == None):
                    diff = abs(min - x)
                elif (min == None) and (max != None):
                    diff = abs(max - x)
                
                fit = self.getPercentFit(diff, 0.7, 4.8)

            self.dataweights[w["dt"]] += fit

    def getUserMaxAndMin(self, group):

        min = preference[group]["min"]
        max = preference[group]["max"]

        return min, max

    def processMaxAndMin(self, weatherParam, min, max):
        if (min != None) and (max != None):
            return weatherParam >= min and weatherParam <= max
        elif (min != None) and (max == None):
            return weatherParam >= min
        elif (min == None) and (max != None):
            return weatherParam <= max
        else:
            return True
   
    def processPercents(self):

        for i in self.dataweights:
            if self.prefCounter != 0:
                
                self.dataweights[i] /= self.prefCounter
            else:
                self.dataweights = 1




       
