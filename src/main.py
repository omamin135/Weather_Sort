import requests
from Process import ProcessReq, ProcessPref, indexdt

import ExcelExport as ex
import os

BASE_URL = "http://api.openweathermap.org/data/2.5/forecast?units=imperial&"
API_KEY = "1f06f2dceba5c5f425f89d7af06a3b09"
CITY = "Hilliard"

def getWeatherData():  
    url = BASE_URL + "appid=" + API_KEY + "&q=" + CITY
    return requests.get(url).json()


if __name__ == "__main__":

    weatherData = getWeatherData()
    
    ProcReq = ProcessReq()
    filteredData = ProcReq.meetReq(weatherData)
    
    ProcPref = ProcessPref(filteredData)
    sortedData = ProcPref.meetExtra()
  

    for i in range(len(sortedData)-1, -1, -1):
        l = len(sortedData) - i

        index = indexdt(filteredData, sortedData[i][0])
        dt = filteredData[index]["dt_txt"]
        percent = sortedData[i][1] * 100
        
        print("[{0}] {1}: {2:.2f}%".format(l, dt, percent))


    
    userInput = input("Do you want export data to Excel sheet [y/n]: ")

    if userInput == "y":

        fileName = input("Enter file name: ")
        file = "Export_Excel\\" + fileName

        x = ex.ExportWorkbook(file)
        userData = x.pullPreferenceData()
        x.exportPreferenceData(userData)
        x.exportDtData(filteredData, sortedData)
        x.formatSpreadSheet(len(sortedData))
        x.saveWorkbook()

        print("Export Successful!")

        userInput = input("Launch Excel sheet [y/n]: ")
        if userInput == "y":
            os.startfile(file)
        