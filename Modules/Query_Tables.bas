Attribute VB_Name = "Query_Tables"

Private Sub Time_Zones_Refresh()

    Dim ListOB_RNG As Range, Result As Variant, Query_Exists As Boolean, URL As String, ERR_STR As String, _
    QueryTable_Object As QueryTable, RefreshTimer As New TimedTask, datetimeResponse As Date
    
    Dim usingTimeApi As Boolean, dateTimeRNG As Range

'After Background Query has finished run this procedure again using an event but supply a QueryTable
'To skip the refresh portion, do additional parsing if needed and then start the next background Query

    On Error GoTo TZ_Refresh_Failed
    
    Const timeZoneRetrievalTimer = "Time Zone Retrieval"
    
    RefreshTimer.Start timeZoneRetrievalTimer
    
    usingTimeApi = True

Restart_Using_CSV:

    If usingTimeApi Then
        URL = "http://worldtimeapi.org/api/timezone/America/Toronto"
    Else
        URL = "https://docs.google.com/spreadsheets/d/1ubpPnoj7hQkMkwgLpFwOwmFftWI4yN3jMihEshVC89A/export?format=csv&id=1ubpPnoj7hQkMkwgLpFwOwmFftWI4yN3jMihEshVC89A&gid=0"
    End If
    
'    #If Mac Then
'
'        Using_PQuery = False
'
'    #Else
'
'        If Application.Version < 16# Then 'IF excel version is prior to Excel 2016 then
'
'            If Not IsPowerQueryAvailable Then Using_PQuery = True 'Check if Power Query is available
'        Else
'
'            Using_PQuery = True
'
'        End If
'
'    #End If
    
    Using_PQuery = False
    
    If Not Using_PQuery Then
        
        For Each QueryTable_Object In QueryT.QueryTables           'Determine if QueryTable Exists
        
            If InStr(1, QueryTable_Object.name, "Time_Z") > 0 Then 'Instr method used in case Excel appends a number to the
                Query_Exists = True                                'QueryTable Name
                Exit For
            End If
            
        Next QueryTable_Object
        
        If Not Query_Exists Then 'Create QueryTable
        
            Set QueryTable_Object = QueryT.QueryTables.Add(Connection:="TEXT;" & URL, Destination:=QueryT.Range("A1"))
            
            With QueryTable_Object
                .name = "Time_Z"
                .WorkbookConnection.name = "Time_Zone_Info"
                .TextFileCommaDelimiter = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlOverwriteCells
                .AdjustColumnWidth = False
                .TextFileColumnDataTypes = Array(xlDMYFormat)
            End With
            
        End If
        
    Else
        Set QueryTable_Object = Variable_Sheet.ListObjects("Time_Zones").QueryTable
    End If
    
    ReDim Result(1 To 2, 1 To 2)
    
    QueryTable_Object.Refresh False
    ' Destination range for retrieved time data.
    
    Result(2, 2) = Now
    Result(2, 1) = "Local Time"
    Result(1, 1) = "EST Time"
    
    Set ListOB_RNG = Variable_Sheet.ListObjects("Time_Zones").DataBodyRange
    
    If Not Using_PQuery Then
    
        With QueryTable_Object.ResultRange
            
            If usingTimeApi Then
                
                Set dateTimeRNG = .Find("datetime", , xlValues, xlPart, , , False)
                
                If dateTimeRNG Is Nothing Then
                    Set dateTimeRNG = .Cells(3)
                End If
                
                On Error GoTo Unable_To_Convert_DateTime
                datetimeResponse = CDate(Left(Split(Replace(dateTimeRNG.Value2, "T", " "), ":" & Chr(34))(1), 19))
            
                On Error GoTo 0
                
                Result(1, 2) = datetimeResponse

            Else
                Result(1, 2) = CDate(.Cells(1, 2))
            End If
            
            .ClearContents

        End With
        
    End If
    
    RefreshTimer.DPrint
    
    With ListOB_RNG
    
        If Not Using_PQuery Then
            .ClearContents
            .Resize(UBound(Result, 1), UBound(Result, 2)).Value2 = Result 'overwrite Query Range with values
        End If
        
        On Error GoTo 0
        
        If Result(2, 2) > CFTC_Release_Dates(Find_Latest_Release:=False) Then
            'Update Release Schedule if the current Local time is greater than the
            '[ next ] Local Release Date and Time.
            Call Release_Schedule_Refresh
        Else
            Variable_Sheet.Range("Release_Schedule_Queried").Value2 = True
        End If
        
    End With
    
    Exit Sub

TZ_Refresh_Failed:
    
    On Error Resume Next
    
    If Not QueryTable_Object Is Nothing Then
        
        With QueryTable_Object
            .WorkbookConnection.Delete
            .Delete
        End With
        
    End If

    ERR_STR = "Failed to connect to external Time source. Aborting Auto-Scheduling Procedures.."
    
    With ThisWorkbook.Event_Storage
        .Item("Event_Error").Add ERR_STR, "Event_Error_TimeZone_Refresh"
    End With
    
    Application.Run "'" & ThisWorkbook.name & "'!Schedule_Data_Update", True 'Check For new Data but skip scheduling
    
    Exit Sub
    
Unable_To_Convert_DateTime:

    With QueryTable_Object.ResultRange
        ThisWorkbook.Event_Storage.Item("Event_Error").Add "Unable to convert [ " & dateTimeRNG.value & " ] from API to date.", "API_Time_Retrieval_Error"
        .ClearContents
    End With
    
    If Not QueryTable_Object Is Nothing Then
        
        With QueryTable_Object
            .WorkbookConnection.Delete
            .Delete
        End With
        
    End If
    
    usingTimeApi = False
    Resume Restart_Using_CSV
    
End Sub
Private Sub Release_Schedule_Refresh()

Dim ListOB_RNG As Range, Result As Variant, _
FNL As Variant, X As Byte, L As Byte, Z As Byte, _
Query_Exists As Boolean, URL As String, QueryTable_Object As QueryTable ',Query_Events As New ClassQTE,

Dim ReleaseScheduleTimer As New TimedTask

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
            If InStr(1, QueryTable_Object.name, "Release_S") > 0 Then
                Query_Exists = True
                Exit For
            End If
        Next QueryTable_Object
        
        If Not Query_Exists Then 'Create Query
        
            URL = "https://docs.google.com/spreadsheets/d/1ubpPnoj7hQkMkwgLpFwOwmFftWI4yN3jMihEshVC89A/export?format=csv&id=1ubpPnoj7hQkMkwgLpFwOwmFftWI4yN3jMihEshVC89A&gid=266164582"
            
            Set QueryTable_Object = QueryT.QueryTables.Add(Connection:="TEXT;" & URL, Destination:=QueryT.Range("A1"))
            
            With QueryTable_Object
                .TextFileCommaDelimiter = True
                .WorkbookConnection.name = "Release_Schedule_Refresh"
                .name = "Release_S"
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlOverwriteCells
                .AdjustColumnWidth = False
            End With
        
        End If
        
    Else
    
        Set QueryTable_Object = Variable_Sheet.ListObjects("Release_Schedule").QueryTable
        
    End If
    
   ' Query_Events.HookUpQueryTable QueryTable_Object, "Release_Schedule_Refresh", ThisWorkbook, Variable_Sheet, Using_PQuery, Weekly
                                    
    '0          1                    2               3         4         5
    QueryTable_Object.Refresh BackgroundQuery:=False 'Use False to trap for errors
    
    If Not Using_PQuery Then
    
        Set ListOB_RNG = Variable_Sheet.ListObjects("Release_Schedule").DataBodyRange

        With QueryTable_Object.ResultRange
            Result = .Value2
            .ClearContents
        End With
    
        For X = 1 To UBound(Result, 1) 'skip blank rows
            If Result(X, 1) <> vbNullString Then L = L + 1
        Next X

        ReDim FNL(1 To L, 1 To UBound(Result, 2))
        
        For X = 1 To UBound(Result, 1) 'compile to array and edit if needed.. remove * from column 1
            If Result(X, 1) <> vbNullString Then
                Z = Z + 1
                For L = 1 To UBound(Result, 2)
                    If L = 1 Then
                       FNL(Z, L) = Replace(Result(X, L), "*", vbNullString)
                     Else
                        FNL(Z, L) = Result(X, L)
                    End If
                Next L
            End If
        Next X

        ListOB_RNG.Cells(1, 1).Resize(UBound(FNL, 1), UBound(FNL, 2)).Value2 = FNL
                
    End If
    
    'If the procudure to run is the auto schedule Workbook data update and Workbook_Open Events
    'are currently being processed.
    
    Variable_Sheet.Range("Release_Schedule_Queried").Value2 = True
    
    ReleaseScheduleTimer.DPrint
    
    Exit Sub

RS_Refresh_Failed:
    
    On Error GoTo -1
    
    On Error Resume Next
    
    If Not QueryTable_Object Is Nothing Then
        
        With QueryTable_Object
            .WorkbookConnection.Delete
            .Delete
        End With
        
    End If
    
    ERR_STR = "Failed to connect to Release Schedule Source. Aborting Auto-Scheduling Procedures."
    
    With ThisWorkbook.Event_Storage
    
        .Item("Event_Error").Add ERR_STR, "Event_Error_TimeZone_Refresh"
        
    End With
    
End Sub





