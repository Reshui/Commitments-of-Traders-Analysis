Attribute VB_Name = "Query_Tables"
Public Sub RefreshTimeZoneTable(Optional eventErrors As Collection, Optional profiler As TimedTask)
'===================================================================================================================
    'Summary: Queries an external time source to find the current time.
    'Outputs: Stores the current time on the Variable_Sheet along with the local time on the running environment.
'===================================================================================================================

    Dim localDateTime As Date, jp As New JsonParserB, savedState As Boolean, _
    easternTimeAndLocalDT(1 To 2, 1 To 1) As Date, apiResponse$
    
    Const retrievalTask$ = "Time Zone Retrieval"
    On Error GoTo Catch_GetRequest_Failed
    
    If Not profiler Is Nothing Then profiler.StartSubTask retrievalTask
        
    If TryGetRequest("http://worldtimeapi.org/api/timezone/America/New_York", apiResponse) Then
    
        easternTimeAndLocalDT(2, 1) = Now
        easternTimeAndLocalDT(1, 1) = jp.Deserialize(apiResponse, True, False, False).Item("datetime")
        
        savedState = ThisWorkbook.Saved
        Variable_Sheet.ListObjects("Time_Zones").DataBodyRange.columns(2).Resize(2).Value2 = easternTimeAndLocalDT
        
        localDateTime = easternTimeAndLocalDT(2, 1)
        On Error GoTo Exit_Sub
        
        If localDateTime > CFTC_Release_Dates(Find_Latest_Release:=False, convertToLocalTime:=True) Then
            'Update Release Schedule if the current Local time is greater than the
            '[ next ] Local Release Date and Time.
            Call Release_Schedule_Refresh
        Else
            Variable_Sheet.Range("Release_Schedule_Queried").Value2 = True
            ThisWorkbook.Saved = savedState
        End If
    End If
Exit_Sub:
    If Not profiler Is Nothing Then profiler.StopSubTask retrievalTask
    Exit Sub
Catch_GetRequest_Failed:
    
    AddParentToErrSource Err, "RefreshTimeZoneTable"
    
    If Not eventErrors Is Nothing Then
        With Err
            eventErrors.Add "Failed GET request." & vbNewLine & .Source & vbNewLine & .Description
        End With
    End If
    
    Resume Exit_Sub
    
End Sub
Private Sub Release_Schedule_Refresh()
'===================================================================================================================
    'Summary: Queries the CFTC website for the COT data release time table.
    'Outputs: Array of COT release dates.
'===================================================================================================================

    Dim ListOB_RNG As Range, result() As Variant, _
    FNL As Variant, x As Byte, L As Byte, Z As Byte, _
    Query_Exists As Boolean, QueryTable_Object As QueryTable  ',Query_Events As New ClassQTE,
    
    Dim ReleaseScheduleTimer As New TimedTask
    Const url$ = "https://docs.google.com/spreadsheets/d/1ubpPnoj7hQkMkwgLpFwOwmFftWI4yN3jMihEshVC89A/export?format=csv&id=1ubpPnoj7hQkMkwgLpFwOwmFftWI4yN3jMihEshVC89A&gid=266164582"
    
    ReleaseScheduleTimer.Start "CFTC Release Schedule Query"
    
    #If Mac Then
        Using_PQuery = False
    #Else
        If Application.Version < 16# Then 'IF excel version is prior to Excel 2016 then
            If IsPowerQueryAvailable Then Using_PQuery = True 'Check if Power Query is available
        Else
            Using_PQuery = True
        End If
    #End If

    If Not Using_PQuery Then 'If Power Query is unavailable
        
        'Application.EnableEvents = False
    
        For Each QueryTable_Object In QueryT.QueryTables
            If InStrB(1, QueryTable_Object.Name, "Release_S") <> 0 Then
                Query_Exists = True
                Exit For
            End If
        Next QueryTable_Object
        
        If Not Query_Exists Then 'Create Query
            
            Set QueryTable_Object = QueryT.QueryTables.Add(Connection:="TEXT;" & url, Destination:=QueryT.Range("A1"))
            
            With QueryTable_Object
                .TextFileCommaDelimiter = True
                .WorkbookConnection.Name = "Release_Schedule_Refresh"
                .Name = "Release_S"
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlOverwriteCells
                .AdjustColumnWidth = False
            End With
        End If
    Else
        Set QueryTable_Object = Variable_Sheet.ListObjects("Release_Schedule").QueryTable
    End If
    
    QueryTable_Object.Refresh BackgroundQuery:=False 'Use False to trap for errors
    
    If Not Using_PQuery Then
    
        Set ListOB_RNG = Variable_Sheet.ListObjects("Release_Schedule").DataBodyRange

        With QueryTable_Object.ResultRange
            result = .Value2
            .ClearContents
        End With
    
        For x = 1 To UBound(result, 1) 'skip blank rows
            If LenB(result(x, 1)) <> 0 Then L = L + 1
        Next x

        ReDim FNL(1 To L, 1 To UBound(result, 2))
        
        For x = 1 To UBound(result, 1) 'compile to array and edit if needed. remove * from column 1
            If LenB(result(x, 1)) <> 0 Then
                Z = Z + 1
                For L = 1 To UBound(result, 2)
                    If L = 1 Then
                       FNL(Z, L) = Replace$(result(x, L), "*", vbNullString)
                     Else
                        FNL(Z, L) = result(x, L)
                    End If
                Next L
            End If
        Next x

        ListOB_RNG.Cells(1, 1).Resize(UBound(FNL, 1), UBound(FNL, 2)).Value2 = FNL
                
    End If
    
    'If the procudure to run is the auto schedule Workbook data update and Workbook_Open Events
    'are currently being processed.
    
    Variable_Sheet.Range("Release_Schedule_Queried").Value2 = True
    
    ReleaseScheduleTimer.DPrint
    
    Exit Sub

RS_Refresh_Failed:
    
    On Error Resume Next
    
    If Not QueryTable_Object Is Nothing Then
        
        With QueryTable_Object
            .WorkbookConnection.Delete
            .Delete
        End With
        
    End If
    Err.Clear
    
End Sub





