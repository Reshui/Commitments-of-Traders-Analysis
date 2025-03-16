Attribute VB_Name = "Query_Tables"
Public Sub RefreshTimeZoneTable(Optional eventErrors As Collection)
'===================================================================================================================
    'Summary: Queries an external time source to find the current time.
    'Outputs: Stores the current time on the Variable_Sheet along with the local time on the running environment.
'===================================================================================================================

    Dim savedState As Boolean, dateTimeRange As Range, _
    easternTimeAndLocalDT(1 To 3, 1 To 1) As Date, success As Boolean, currentLocalTime As Date, timeZoneTable As ListObject

    savedState = ThisWorkbook.Saved
    On Error GoTo Missing_Time_Table
    
    Set timeZoneTable = Variable_Sheet.ListObjects("Time_Zones")
    Set dateTimeRange = timeZoneTable.DataBodyRange.columns(2).Resize(UBound(easternTimeAndLocalDT, 1))
    
    On Error GoTo Catch_GetRequest_Failed
    #If Not Mac Then
        ' Local Time
        currentLocalTime = Now
        easternTimeAndLocalDT(2, 1) = currentLocalTime
        ' UTC
        easternTimeAndLocalDT(3, 1) = ConvertLocalToUTC(currentLocalTime)
        ' Eastern Time
        easternTimeAndLocalDT(1, 1) = ConvertGmtWithTimeZone(easternTimeAndLocalDT(3, 1), -5)
        success = True
    #Else
        
        Dim apiResponse$, jp As New JsonParserB
        
        If IsPowerQueryAvailable() Then
            With timeZoneTable
                .QueryTable.Refresh False
                currentLocalTime = WorksheetFunction.VLookup("Local Time", .DataBodyRange, 2, False)
                success = True
            End With
        ElseIf TryGetRequest("https://worldtimeapi.org/api/timezone/America/New_York", apiResponse) Then
            easternTimeAndLocalDT(2, 1) = Now
            currentLocalTime = easternTimeAndLocalDT(2, 1)
            
            On Error GoTo Catch_Json_Deserialization_Failure
            With jp.Deserialize(apiResponse, False, False, False)
                ' 86400 = 60s * 60min * 24hrs > Converting seconds to days.
                ' Eastern Time
                easternTimeAndLocalDT(1, 1) = ((.item("unixtime") + .item("raw_offset")) / 86400) + #1/1/1970#
                ' UTC
                easternTimeAndLocalDT(3, 1) = (.item("unixtime") / 86400) + #1/1/1970#
            End With
            
            success = True
        End If
    #End If

    On Error GoTo Exit_Sub

    If success Then
        If Not IsEmpty(easternTimeAndLocalDT(1, 1)) Then dateTimeRange.Value2 = easternTimeAndLocalDT
        If currentLocalTime > CFTC_Release_Dates(Find_Latest_Release:=False, convertToLocalTime:=True) Then
            'Update Release Schedule if the current Local time is greater than the
            '[ next ] Local Release Date and Time.
            Call Release_Schedule_Refresh
        Else
            Variable_Sheet.Range("Release_Schedule_Queried").Value2 = True
            ThisWorkbook.Saved = savedState
        End If
    End If

Exit_Sub:

    Exit Sub
        
Catch_GetRequest_Failed:
    AddParentToErrSource Err, "RefreshTimeZoneTable"
    If Not eventErrors Is Nothing Then
        With Err
            eventErrors.Add "Failed GET request." & vbNewLine & .Source & vbNewLine & .Description
        End With
    End If
    Resume Exit_Sub

Catch_Json_Deserialization_Failure:
    Dim datetimeLocation&
    AddParentToErrSource Err, "RefreshTimeZoneTable"
    'If key not found error.
    If Err.Number = 5 Then
        datetimeLocation = InStrB(apiResponse, """datetime""")
        If datetimeLocation <> 0 Then
            easternTimeAndLocalDT(1, 1) = DateValue(Mid$(apiResponse, datetimeLocation + 12, 10)) + TimeValue(Mid$(apiResponse, datetimeLocation + 23, 8))
            Resume Next
        Else
            AppendErrorDescription Err, "Key not found > 'datetime'"
        End If
    End If

    If Not eventErrors Is Nothing Then
        With Err
            eventErrors.Add .Source & vbNewLine & .Description
        End With
    End If
    Resume Exit_Sub
    
Missing_Time_Table:
    If Err.Number = 9 Then PropagateError Err, "RefreshTimeZoneTable", "Missing 'Time_Zones' list object on Variables worksheet."
    GoTo Propagate
Propagate:
    PropagateError Err, "RefreshTimeZoneTable"
End Sub
Private Sub Release_Schedule_Refresh()
'===================================================================================================================
    'Summary: Queries the CFTC website for the COT data release time table.
    'Outputs: Array of COT release dates.
'===================================================================================================================

    Dim ReleaseScheduleTimer As New TimedTask
    
    Const QueryTableName$ = "CFTC_Website_Schedule"
    ReleaseScheduleTimer.Start "CFTC Release Schedule Query"

    On Error GoTo RS_Refresh_Failed
    
    If Not (IsPowerQueryAvailable() And False) Then

        Dim result() As Variant, Query_Exists As Boolean, releaseScheduleQueryTable As QueryTable, _
        iRow&, iColumn&, cc As New Collection, dataRow() As Variant
        
        With QueryT
            For Each releaseScheduleQueryTable In .QueryTables
                If InStrB(1, releaseScheduleQueryTable.name, QueryTableName) <> 0 Then
                    Query_Exists = True
                    Exit For
                End If
            Next releaseScheduleQueryTable
        
            If Not Query_Exists Then
                Const url$ = "https://www.cftc.gov/MarketReports/CommitmentsofTraders/ReleaseSchedule/index.htm"
                Set releaseScheduleQueryTable = .QueryTables.Add(Connection:="URL;" & url, Destination:=.Range("A1"))
                
                With releaseScheduleQueryTable
                    .name = QueryTableName
                    .WorkbookConnection.name = QueryTableName
                    .RefreshOnFileOpen = False
                    .RefreshStyle = xlOverwriteCells
                    .AdjustColumnWidth = False
                    .WebTables = "1,2"
                End With
            End If
        End With
        
        With releaseScheduleQueryTable
            
            .Refresh False
            
            With .ResultRange
                result = .Value2
                ReDim dataRow(LBound(result, 2) To UBound(result, 2))
                With cc
                    For iRow = 1 To UBound(result, 1)
                        If Not (IsEmpty(result(iRow, 1)) Or result(iRow, 1) = "Month") Then
                            For iColumn = LBound(result, 2) To UBound(result, 2)
                                dataRow(iColumn) = result(iRow, iColumn)
                            Next iColumn
                            .Add dataRow
                        End If
                    Next iRow
                End With
                .ClearContents
            End With
            
        End With
        
        result = CombineArraysInCollection(cc, Multiple_1d)
        
        With Variable_Sheet.ListObjects("Release_Schedule")
            If .ListRows.Count <> UBound(result, 1) Then .Resize .Range.Resize(UBound(result, 1) + 1, UBound(result, 2))
            .DataBodyRange.Value2 = result
        End With
        
    Else
         Variable_Sheet.ListObjects("Release_Schedule").QueryTable.Refresh False
    End If

    Variable_Sheet.Range("Release_Schedule_Queried").Value2 = True

    ReleaseScheduleTimer.DPrint

    Exit Sub

RS_Refresh_Failed:

    With HoldError(Err)
        On Error GoTo -1: On Error Resume Next
    
        If Not releaseScheduleQueryTable Is Nothing Then
            With releaseScheduleQueryTable
                .WorkbookConnection.Delete
                .Delete
            End With
        End If
        On Error GoTo 0
        PropagateError .HeldError, "Release_Schedule_Refresh"
    End With
End Sub
Private Sub DeleteAllQueryTablesOnQueryTSheet()

    Dim QT As QueryTable

    For Each QT In QueryT.QueryTables
         'Debug.Print QT.name
         With QT
            'If Not .WorkbookConnection Is Nothing Then .WorkbookConnection.Delete
            '.Delete
            Debug.Print .name
        End With
    Next

'    For Each QT In ThisWorkbook.Connections
'         Debug.Print QT.name
'         'QT.Delete
'
'         If QT.name Like "jun7-fc8e*" Then QT.Delete
'    Next


End Sub

