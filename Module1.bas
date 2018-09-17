Attribute VB_Name = "Module1"
Public Sub ReadNoaaPrecip()
    ' initialize req'd variables and sheet
    Dim http As Object 'object used for the setting up the http request
    Dim webServiceURL As String 'variable string used to hold in the URL of the website
    ActiveSheet.Range("A12:ZZ99999").ClearContents 'clear worksheet
    
    origSelectedCell = ActiveCell.Address
    ActiveSheet.Range("A11").Select
    wscount = WorksheetFunction.CountA(Range("11:11")) / 2 'creates a count for the number of stations you will be looking through
    Application.ScreenUpdating = False
    wscounter = 1     'set another counter to 1 to count up to the number of stations you have to through
    While ActiveCell.Value <> ""
    
        'create a status bar
        Application.StatusBar = Format(wscounter, "000") & " of " & Format(wscount, "000") & " -- " & WorksheetFunction.Floor(100 * (wscounter - 1) / wscount, 1) & "% Complete | Reading data for " & ActiveCell.Offset(-1, 1).Value
        Application.Wait Now + #12:00:01 AM#
        
        'build api url for series api endpoint
        'https://www.ncdc.noaa.gov/cdo-web/api/v2/data?stationid=GHCND:USW00023158&datasetid=GSOM&datatypeid=PRCP&startdate=2015-05-01&enddate=2018-05-31
        ' datasetid: GHCND = Daily Summaries| GSOM = Global Summary of the Month
        ' datatype: PRCP = Precipitation
        
        webServiceURL = "https://www.ncdc.noaa.gov/cdo-web/api/v2/data?stationid=" & _
            ActiveCell.Offset(0, 1).Value & "&datasetid=GSOM&datatypeid=PRCP&startdate=" & _
            Format(Range("B3").Value, "yyyy-dd-mm") & "&enddate=" & Format(Range("B4").Value, "yyyy-dd-mm") & "&limit=1000"
            
        Set http = CreateObject("msxml2.xmlhttp")
            http.Open "GET", webServiceURL, False
            http.setRequestHeader "token", "QuZXECfykjiAqmGtODQADCAXqBwJrsvb"
            http.Send
            
        ' Example of the JSON file being received
        '{
        '   "metadata": {
        '       "resultset": {
        '           "offset": 1,
        '           "count": 37,
        '           "limit": 25
        '       }
        '   },
        '   "results": [
        '       {
        '           "date": "2015-05-01T00:00:00",
        '           "datatype": "PRCP",
        '           "station": "GHCND:USW00023158",
        '           "attributes": ",,,W",
        '           "value": 8.9
        '       }
        '   ]
        '}
        '
        
    On Error GoTo ErrCatcher
                    
        ' PARSE JSON RESPONSE
        JsonText = http.ResponseText
        Set JSON = JsonConverter.ParseJson(JsonText)
        
        ' GET TS DATA INTO VBA DICT
        Dim Values As Variant
        ReDim Values(JSON("results").Count, 1)
        Dim Value As Dictionary
        Dim i As Long
        i = 0
        For Each Value In JSON("results")   'sorting through the results after the request
            Values(i, 0) = Value("date")    'putting in the date into the values array
            Values(i, 1) = Value("value")   'putting in the values into the values array
            i = i + 1
        Next Value
        
        ' WRITE TS DATA
        a = UBound(Values)
        If UBound(Values) < 1 Then 'Catch case for no data
            ActiveCell.Offset(1, 0).Value = "No Data..."
        Else 'data is being printed into the spreadsheet
            ActiveSheet.Range(Cells(ActiveCell.Offset(1, 0).Row, ActiveCell.Column), Cells(ActiveCell.Offset(1, 0).Row + JSON("results").Count, ActiveCell.Column + 1)) = Values
        End If
        
       
    
ErrCatcher:
        Set http = Nothing
        ' MOVE TO NEXT COLUMN OF DATA
        ActiveCell.Offset(0, 2).Select
        wscounter = wscounter + 1
        
    Wend
    Application.ScreenUpdating = True
    ActiveSheet.Range(origSelectedCell).Select
    Application.StatusBar = "Done!"

End Sub
