Attribute VB_Name = "Query_Tables"

Private Sub Time_Zones_Refresh()

Dim ListOB_RNG As Range, Result As Variant, Query_Exists As Boolean, URL As String, ERR_STR As String, _
QueryTable_Object As QueryTable ', Query_Events As New ClassQTE

'After Background Query has finished run this procedure again using an event but supply a QueryTable
'To skip the refresh portion, do additional parsing if needed and then start the next background Query

On Error GoTo TZ_Refresh_Failed

    #If Mac Then
    
        Using_PQuery = False
        
    #Else
    
        If Application.Version < 16# Then 'IF excel version is prior to Excel 2016 then
        
            If Not IsPowerQueryAvailable Then Using_PQuery = True 'Check if Power Query is available
        Else
        
            Using_PQuery = True
            
        End If
        
    #End If
    
    If Not Using_PQuery Then
        
        For Each QueryTable_Object In QueryT.QueryTables           'Determine if QueryTable Exists
        
            If InStr(1, QueryTable_Object.Name, "Time_Z") > 0 Then 'Instr method used in case Excel appends a number to the
                Query_Exists = True                                'QueryTable Name
                Exit For
            End If
            
        Next QueryTable_Object
        
        If Not Query_Exists Then 'Create QueryTable
        
            URL = "https://docs.google.com/spreadsheets/d/1ubpPnoj7hQkMkwgLpFwOwmFftWI4yN3jMihEshVC89A/export?format=csv&id=1ubpPnoj7hQkMkwgLpFwOwmFftWI4yN3jMihEshVC89A&gid=0"
            
            Set QueryTable_Object = QueryT.QueryTables.Add(Connection:="TEXT;" & URL, Destination:=QueryT.Range("A1"))
            
            With QueryTable_Object
                .Name = "Time_Z"
                .WorkbookConnection.Name = "Time_Zone_Info"
                .TextFileCommaDelimiter = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlOverwriteCells
                .AdjustColumnWidth = False
            End With
            
        End If
        
    Else
    
        Set QueryTable_Object = Variable_Sheet.ListObjects("Time_Zones").QueryTable
        
    End If
    '[  Query Object, Procedure to Call after refresh,  Workbook Object, Variable Worksheet ,
    '   Querytable ListObject boolean,Optional Worksheet ]
    
    'Query_Events.HookUpQueryTable QueryTable_Object, "Time_Zones_Refresh", ThisWorkbook, Variable_Sheet, Using_PQuery, Weekly
                                  '0                        1                    2             3                  4                    5
    QueryTable_Object.Refresh False

    Set ListOB_RNG = Variable_Sheet.ListObjects("Time_Zones").DataBodyRange 'Destination range
    
    If Not Using_PQuery Then 'If not using PowerQuery as boolean
    
        With QueryTable_Object.ResultRange
        
            Result = .Value2
            .ClearContents
            
        End With
        
    End If
    
    With ListOB_RNG
    
        If Not Using_PQuery Then .Resize(UBound(Result, 1), UBound(Result, 2)).Value2 = Result 'overwrite Query Range with values
        
        .Rows(3).Value2 = Array("Local Time", Now) 'overwrite  row below data with the current time on the user machine
        
        On Error GoTo 0
        
        If .Cells(3, 2) > CFTC_Release_Dates(False) Then 'Update Release Schedule if the current Local time
                                                         'is greater than the [ next ] Local Release Date and Time
            Call Release_Schedule_Refresh
            
        Else
        
            With Variable_Sheet.ListObjects("Saved_Variables").DataBodyRange.Columns(1)
                'Store that the macro has been triggered Release Schedule Query
                .Cells(WorksheetFunction.Match("Release Schedule Queried", .Value2, 0), 2).Value2 = True
                
            End With
            
            On Error Resume Next 'If the item isn't found then the condition will be executed
            
            If Not ThisWorkbook.Event_Storage.Item("Currently Scheduling") Then
                                'This item will only be in the collection if the scheduling macro is currently active
                On Error GoTo 0
                
                Application.Run "'" & ThisWorkbook.Name & "'!Schedule_Data_Update", True
            
            End If
            
        End If
        
    End With
    
Exit Sub

TZ_Refresh_Failed:
    
    On Error GoTo -1
    
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
    
    Application.Run "'" & ThisWorkbook.Name & "'!Schedule_Data_Update", True 'Check For new Data but skip scheduling
    
End Sub
Private Sub Release_Schedule_Refresh()

Dim ListOB_RNG As Range, Result As Variant, _
FNL As Variant, X As Long, L As Long, ATC As CheckBox, data As Variant, Z As Long, Y As Long, _
Query_Exists As Boolean, URL As String, QueryTable_Object As QueryTable ',Query_Events As New ClassQTE,


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
            If InStr(1, QueryTable_Object.Name, "Release_S") > 0 Then
                Query_Exists = True
                Exit For
            End If
        Next QueryTable_Object
        
        If Not Query_Exists Then 'Create Query
        
            URL = "https://docs.google.com/spreadsheets/d/1ubpPnoj7hQkMkwgLpFwOwmFftWI4yN3jMihEshVC89A/export?format=csv&id=1ubpPnoj7hQkMkwgLpFwOwmFftWI4yN3jMihEshVC89A&gid=266164582"
            
            Set QueryTable_Object = QueryT.QueryTables.Add(Connection:="TEXT;" & URL, Destination:=QueryT.Range("A1"))
            
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
    
   ' Query_Events.HookUpQueryTable QueryTable_Object, "Release_Schedule_Refresh", ThisWorkbook, Variable_Sheet, Using_PQuery, Weekly
                                        
    If Not HasKey(ThisWorkbook.Event_Storage, "Currently Scheduling") Then
    
        ThisWorkbook.Event_Storage.Add False, "Currently Scheduling"
        
    End If
    
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
    
    With Variable_Sheet.ListObjects("Saved_Variables").DataBodyRange.Columns(1)
        .Cells(WorksheetFunction.Match("Release Schedule Queried", .Value2, 0), 2).Value2 = True
        'Succesfully updated Release Schedule Query
    End With
    
    On Error Resume Next
    
    If Not ThisWorkbook.Event_Storage.Item("Currently Scheduling") Then 'this item will olnly be in the collection if the scheduling macro is currently active
                                              'If scheduling macro isn't activated ie: if there is an error then it will be executed
        On Error GoTo 0
        
        Application.Run "'" & ThisWorkbook.Name & "'!Schedule_Data_Update", True
    
    End If
    
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





