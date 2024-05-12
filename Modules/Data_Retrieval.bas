Attribute VB_Name = "Data_Retrieval"
Public Const rawCftcDateColumn As Byte = 3
Public Retrieval_Halted_For_User_Interaction As Boolean
Public Data_Updated_Successfully As Boolean
Public Running_Weekly_Retrieval As Boolean

Public Enum ReportStatusCode
    NoUpdateAvailable = 0
    Updated = 1
    Failure = 2
    AttemptingRetrieval = 3
    AttemptingUpdate = 4
    NotInitialized = 5
    CheckingDataAvailability = 6
End Enum

Public Enum Append_Type
    Add_To_Old = 1
    Multiple_1d = 2
    Multiple_2d = 3
End Enum

Public Enum Data_Identifier
    Block_Data = 14
    Old_Data = 44
    Weekly_Data = 33
End Enum

Option Explicit
Sub New_Data_Query(Optional Scheduled_Retrieval As Boolean = False, Optional Overwrite_All_Data As Boolean = False)
'===================================================================================================================
    'Purpose: Retrieves CFTC data that hasn't been stored either on a worksheet or database.
    'Inputs: Scheduled_Retrieval - If true then error messages will not be displayed.
    '        Overwrite_All_Data - Reserved for non-Database version. If true then all data on a worksheet will be replaced.
    'Outputs:
'===================================================================================================================

    Dim Last_Update_CFTC As Date, CFTC_Incoming_Date As Date, ICE_Incoming_Date As Date, Last_Update_ICE As Date

    Dim Debug_Mode As Boolean, CFTC_Data() As Variant, report As Variant, ICE_Data() As Variant, Historical_Query() As Variant

    Dim debugWeeklyRetrieval As Boolean, DBM_Historical_Retrieval As Boolean, queryingCombinedData As Variant

    Dim newDataSuccessfullyHandled As Boolean, cftcDateRange As Range, iceDateRange As Range  ', All_Available_Contracts() As Variant
    
    Dim Download_CFTC As Boolean, Download_ICE As Boolean, Check_ICE As Boolean, uploadDataToDatabaseOrWorksheet As Boolean, Weekly_ICE_CLCTN As Collection, iceKey As String
    
    Dim DataBase_Not_Found_CLCTN As Collection, availableContractInfo As Collection, Data_CLCTN As Collection, Legacy_Combined_Data As Boolean, iceHasBeenQueried As Boolean
    
    Dim TestTimers As New Timers, reportDataTimer As String, individualCotTask As TimedTask, exitSubroutine As Boolean, CFTC_Retrieval_Error As Boolean
        
    Dim executeableReturn As Collection, cftcMappedFieldInfo As Collection
    
    Const dataRetrieval As String = "C.O.T data retrieval", uploadTime As String = "Upload Time", _
    totalTime As String = "Total runtime", databaseDateQuery As String = "Query database for latest date"
     
    Const legacy_initial As String = "L"
    
    Dim reportsToQuery() As String, combinedFutures() As Boolean, useAPIData As Boolean, cftcApiFailed As Boolean, databaseReturnedDate As Boolean
    
    'This function will either return all contracts for which price data is available or all contracts in the workbook.
    Set availableContractInfo = GetAvailableContractInfo
    
    Running_Weekly_Retrieval = True
    
    Call IncreasePerformance

    On Error GoTo Deny_Debug_Mode
    
    If Weekly.Shapes("Test_Toggle").OLEFormat.Object.value = xlOn Then Debug_Mode = True 'Determine if Debug status
    
    #If DatabaseFile Then
        
        #If Mac Then
            GateMacAccessToWorkbook
        #End If
        
        reportsToQuery = Split("L,D,T", ",")
        ReDim combinedFutures(1): combinedFutures(0) = True
        Set cftcDateRange = Variable_Sheet.Range("Most_Recently_Queried_Date")
        
        If Not Debug_Mode Then
            On Error GoTo ExecutableFailed
            Set executeableReturn = RunCSharpExtractor()
        End If
        
    #Else
    
        ReDim reportsToQuery(0): reportsToQuery(0) = ReturnReportType
        ReDim combinedFutures(0): combinedFutures(0) = IsWorkbookForFuturesAndOptions
        
        With Variable_Sheet
            Set cftcDateRange = .Range("Last_Updated_CFTC")
            If reportsToQuery(0) = "D" Then Set iceDateRange = .Range("Last_Updated_ICE")
        End With
        
    #End If
        
CollectDetails:

    If Debug_Mode = True Then
    
        If MsgBox("Test Weekly Data Retrieval ?", vbYesNo, "Choose what to debug") = vbYes Then
            debugWeeklyRetrieval = True
        ElseIf MsgBox("Test Multi-Week Historical Retrieval ?", vbYesNo, "Choose what to debug") = vbYes Then
            DBM_Historical_Retrieval = True
        End If
        ' Get the report type to test.
        #If DatabaseFile Then
        
            Do
                report = UCase(InputBox("Select 1 of L,D,T"))
            Loop While IsError(Application.Match(report, Array("L", "D", "T"), 0))
            
            reportsToQuery(0) = report
                        
            On Error Resume Next
            ' Get the combined status to test.
            Do
                queryingCombinedData = CBool(InputBox("Select 1 of the following:" & vbNewLine & vbNewLine & "  Futures Only: 0" & vbNewLine & "  Futures + Options: 1"))
            Loop While Err.Number <> 0
            
            combinedFutures(0) = queryingCombinedData
        
        #End If
        
        On Error GoTo 0
    End If

Retrieve_Latest_Data:

    On Error GoTo 0
    
    With TestTimers
        .description = "New Data Query [" & Time & "]"
        .StartTask totalTime
    End With
    
    For Each report In reportsToQuery 'Legacy data must be retrieved first so that price data only needs to be retrieved once
        
        For Each queryingCombinedData In combinedFutures 'True must be first so that price data can be retrieved for futures only data
            
            On Error GoTo 0
            
            Dim checkedViaExecutable As Boolean
            uploadDataToDatabaseOrWorksheet = False
            
            #If DatabaseFile Then
                
                On Error GoTo Try_Get_New_Data
                
                If Not executeableReturn Is Nothing Then
                
                    Dim innerValues As Collection, goNextLoop As Boolean
                    
                    Set innerValues = executeableReturn(report)(CStr(queryingCombinedData))
                    ' Less than or equal to one means that the exe completed successfully.
                    If innerValues("status") <= 1 Then
                        goNextLoop = True
                        checkedViaExecutable = True
                        
                        If innerValues("latest date") > cftcDateRange.value Or innerValues("status") = 1 Then

                            With GetStoredReportDetails(CStr(report))
                                If queryingCombinedData = .UsingCombined Then .PendingUpdateInDatabase = True
                            End With
                            
                            newDataSuccessfullyHandled = True
                            
                        End If
                        
                    Else
                        checkedViaExecutable = False
                    End If
                    
                    If CFTC_Incoming_Date < innerValues("latest date") Then CFTC_Incoming_Date = innerValues("latest date")
                    
                    Set innerValues = Nothing
                    
                    If goNextLoop And Not Debug_Mode Then
                        goNextLoop = False
                        GoTo Next_Combined_Value
                    End If
                    
                End If
                
            #End If
            
Try_Get_New_Data:
            
            On Error GoTo 0

            If Not Overwrite_All_Data Then useAPIData = True
            
            reportDataTimer = "( " & report & " ) Combined: (" & queryingCombinedData & ")"
            
            Set individualCotTask = TestTimers.ReturnTimedTask(reportDataTimer) 'Initializing this way allows you to create subtasks
            
            individualCotTask.Start
            
            If queryingCombinedData And report = legacy_initial Then
                Legacy_Combined_Data = True
            Else
                Legacy_Combined_Data = False
            End If
            
            #If DatabaseFile Then
                ' Retrieve the date the data was last queried for.
                With individualCotTask.SubTask(databaseDateQuery)
                    
                    .Start
                    databaseReturnedDate = TryGetLatestDate(Last_Update_CFTC, reportType:=CStr(report), getFuturesAndOptions:=CBool(queryingCombinedData), queryIceContracts:=False)
                    
                    If Not databaseReturnedDate Then
                        
                        If DataBase_Not_Found_CLCTN Is Nothing Then Set DataBase_Not_Found_CLCTN = New Collection
                    
                        DataBase_Not_Found_CLCTN.Add "Missing database for " & Evaluate("VLOOKUP(""" & report & """,Report_Abbreviation,2,FALSE)")
                    
                        'The Legacy_Combined data is the only one for which price data is queried.
                        If Legacy_Combined_Data Then exitSubroutine = True
                        
                     End If
                     
                    .EndTask
                    If Not databaseReturnedDate Then Exit For
                    
                End With
                
            #Else
                Last_Update_CFTC = cftcDateRange.Value2
            #End If
            
            individualCotTask.SubTask(dataRetrieval).Start
            
            If CFTC_Incoming_Date = 0 Or Last_Update_CFTC < CFTC_Incoming_Date Or Debug_Mode Then
                         
                On Error GoTo CFTC_Retrieval_Failed
                
                Set cftcMappedFieldInfo = Nothing
                
                CFTC_Data = HTTP_Weekly_Data(Last_Update_CFTC, Auto_Retrieval:=Scheduled_Retrieval, _
                                retrieveCombinedData:=CBool(queryingCombinedData), _
                                reportType:=CStr(report), useApi:=useAPIData, _
                                columnMap:=cftcMappedFieldInfo, _
                                DebugMD:=debugWeeklyRetrieval)
                
                cftcApiFailed = Not useAPIData
                
                If CFTC_Incoming_Date = 0 Then CFTC_Incoming_Date = CFTC_Data(UBound(CFTC_Data, 1), rawCftcDateColumn)
                
            Else
                'New data not available for Non Legacy Combined
                GoTo Stop_Timers_And_Update_If_Allowed
            End If
            
            If report = "D" Then
                
                On Error GoTo ICE_Retrieval_Failed
                
                Check_ICE = True
                
                #If DatabaseFile Then
                    Call TryGetLatestDate(Last_Update_ICE, reportType:=CStr(report), getFuturesAndOptions:=CBool(queryingCombinedData), queryIceContracts:=True)
                #Else
                    Last_Update_ICE = iceDateRange.Value2
                #End If
                
                If Not iceHasBeenQueried Then
                    iceHasBeenQueried = True
                    Set Weekly_ICE_CLCTN = Weekly_ICE(CDate(CFTC_Data(UBound(CFTC_Data, 1), rawCftcDateColumn)))
                End If
                
                iceKey = IIf(queryingCombinedData = True, "futures+options", "futures-only")
                
                With Weekly_ICE_CLCTN
                    ICE_Data = .Item(iceKey)
                    .Remove iceKey
                    If .count = 0 Or Not iceDateRange Is Nothing Then Set Weekly_ICE_CLCTN = Nothing
                End With
                
                ICE_Incoming_Date = ICE_Data(1, rawCftcDateColumn)
                
            Else
                Check_ICE = False
            End If
            
Finished_Querying_Weekly_Data:
            
            On Error GoTo 0
            
            If debugWeeklyRetrieval Then
            
                With individualCotTask.SubTask(dataRetrieval)
                
                    .Pause
                    
                    If MsgBox("Weekly Retrieval has completed. Would you like to continue?", vbYesNo) = vbNo Then
                        exitSubroutine = True
                        GoTo Stop_Timers_And_Update_If_Allowed
                    Else
                        .Continue
                    End If
                    
                End With
                
            End If
            
            If Not Debug_Mode And CFTC_Incoming_Date = Last_Update_CFTC And (Not Check_ICE Or (Check_ICE And ICE_Incoming_Date = Last_Update_ICE)) Then

                If Legacy_Combined_Data Then exitSubroutine = True
                
            ElseIf DBM_Historical_Retrieval Or ((cftcApiFailed And CFTC_Incoming_Date - Last_Update_CFTC > 7) Or (Check_ICE And ICE_Incoming_Date - Last_Update_ICE > 7)) Then
                
                If (cftcApiFailed And CFTC_Incoming_Date - Last_Update_CFTC > 7) Or DBM_Historical_Retrieval Then
                    
                    #If Mac Then
                        MsgBox "CFTC API currently unavailable. Please try again later."
                        exitSubroutine = True
                        GoTo Exit_Procedure
                    #Else
                        Download_CFTC = True
                    #End If
                Else
                    Download_CFTC = False
                End If
                
                If Check_ICE And (ICE_Incoming_Date - Last_Update_ICE > 7 Or DBM_Historical_Retrieval) Then
                    Download_ICE = True
                Else
                    Download_ICE = False
                End If
                
                On Error GoTo Exit_Procedure
                'Download missing data and overwrite the current array.
                Historical_Query = Missing_Data(getFuturesAndOptions:=CBool(queryingCombinedData), _
                    ICE_Data:=ICE_Data, CFTC_Data:=CFTC_Data, _
                    Download_ICE_Data:=Download_ICE, Download_CFTC_Data:=Download_CFTC, _
                    reportType:=CStr(report), _
                    CFTC_Last_Updated_Day:=Last_Update_CFTC, ICE_Last_Updated_Day:=Last_Update_ICE + 2, _
                    DebugMD:=DBM_Historical_Retrieval)
                 
                On Error GoTo 0
                
                If Not (Download_ICE And Download_CFTC) Then
                    
                    If Download_CFTC And (Check_ICE And ICE_Incoming_Date - Last_Update_ICE > 0) Then
                        'Determine if the most recently queried Ice Data needs to be added
                        Set Data_CLCTN = New Collection
                        
                        With Data_CLCTN
                            .Add Historical_Query
                            .Add ICE_Data
                        End With
                        
                        Historical_Query = CombineArraysInCollection(Data_CLCTN, Append_Type.Multiple_2d)
                        
                    ElseIf Download_ICE And CFTC_Incoming_Date - Last_Update_CFTC > 0 Then
                        'Determine if CFTC data needs to be added
                        Set Data_CLCTN = New Collection
                        
                        With Data_CLCTN
                            .Add Historical_Query
                            .Add CFTC_Data
                        End With
                        
                        Historical_Query = CombineArraysInCollection(Data_CLCTN, Append_Type.Multiple_2d)
                        
                    End If
                    
                End If
                uploadDataToDatabaseOrWorksheet = True
                
            ElseIf (CFTC_Incoming_Date - Last_Update_CFTC) > 0 Or (Check_ICE And ICE_Incoming_Date - Last_Update_ICE > 0) Or Debug_Mode = True Then  'If just a 1 week difference
                ' Concatenate data if needed and prep for databaseupload/ paste to worksheet.
                Set Data_CLCTN = New Collection
                
                With Data_CLCTN
                
                    If Check_ICE And (ICE_Incoming_Date - Last_Update_ICE > 0 Or Debug_Mode) Then .Add ICE_Data
                    If CFTC_Incoming_Date - Last_Update_CFTC > 0 Or Debug_Mode Then .Add CFTC_Data

                    If .count = 1 Then
                        ' No concatenation is needed so store array in variable.
                        Historical_Query = .Item(1)
                    ElseIf .count = 2 Then
                        ' Concatenate the arrays.
                        Historical_Query = CombineArraysInCollection(Data_CLCTN, Append_Type.Multiple_2d)
                    End If
                    
                End With
                
                uploadDataToDatabaseOrWorksheet = True
                
            End If

Stop_Timers_And_Update_If_Allowed:

            With individualCotTask
                
                .SubTask(dataRetrieval).EndTask
            
                If uploadDataToDatabaseOrWorksheet = True And Not exitSubroutine Then
                    ' Upload data to database/spreadsheet and retrieve prices.
                    With .SubTask(uploadTime)
                        .Start
                        
                        On Error GoTo Catch_Block_Query_Failed
                        Call Block_Query(Query:=Historical_Query, reportType:=CStr(report), _
                                        isDataFuturesAndOptions:=CBool(queryingCombinedData), availableContractInfo:=availableContractInfo, _
                                        debugOnly:=Debug_Mode, mappedFields:=cftcMappedFieldInfo, Overwrite_Worksheet:=Overwrite_All_Data)
                        .EndTask
                        newDataSuccessfullyHandled = True
                    End With
                    
                End If
                
                .EndTask
                
            End With
            
Next_Combined_Value:
            On Error GoTo 0
            If Debug_Mode Or exitSubroutine Then Exit For
        
        Next queryingCombinedData
        
Next_Report_Release_Type:
        
        On Error GoTo 0
        
        If Not checkedViaExecutable Then
            With individualCotTask
                If .isRunning Then .EndTask
            End With
        End If
        
        If Debug_Mode Or exitSubroutine Then Exit For
                        
    Next report
    
Exit_Procedure:
    
    On Error GoTo 0
    
    If newDataSuccessfullyHandled And Not exitSubroutine Then
        
        Data_Updated_Successfully = True
        '-------------------------------------------------------------------------------------------
        
        With cftcDateRange
            If CFTC_Incoming_Date > .Value2 Then
                'Update_Text CFTC_Incoming_Date   'Update Text Boxes "My_Date" on the HUB and Weekly worksheets.
                .Value2 = CFTC_Incoming_Date
            End If
        End With
        
        If Check_ICE And Not iceDateRange Is Nothing Then
            With iceDateRange
                If ICE_Incoming_Date > .Value2 Then
                    .Value2 = ICE_Incoming_Date
                End If
            End With
        End If
        '----------------------------------------------------------------------------------------
        If Not Debug_Mode Then
        
            If Not Scheduled_Retrieval Then HUB.Activate 'If ran manually then bring the User to the HUB
            Courtesy                                     'Change Status Bar_Message
            
            If newDataSuccessfullyHandled Then
            
                #If DatabaseFile Then
        
                    Select Case ThisWorkbook.ActiveSheet.name
                        Case LC.name: RefreshTableData "L"
                        Case DC.name: RefreshTableData "D"
                        Case TC.name: RefreshTableData "T"
                    End Select
                    
                    With TestTimers
                        report = "Query all databases for latest contracts."
                        .StartTask CStr(report)
                         Latest_Contracts
                        .EndTask CStr(report)
                    End With
                                    
                #End If
                
            End If
            
        End If
        
    ElseIf Not DataBase_Not_Found_CLCTN Is Nothing And Not Scheduled_Retrieval Then

        For Each report In DataBase_Not_Found_CLCTN
            MsgBox report
        Next report
        
    ElseIf Not (Scheduled_Retrieval Or Debug_Mode Or CFTC_Retrieval_Error) Then
    
        MsgBox "No new data could be retrieved." & _
        vbNewLine & _
        vbNewLine & _
            "The next release is scheduled for " & vbNewLine & vbTab & Format(CFTC_Release_Dates(False), "[$-x-sysdate]dddd, mmmm dd, yyyy") & " 3:30 PM Eastern time." & _
        vbNewLine & _
        vbNewLine & _
            "Enabling Test Mode will allow you to continue, but only new/missing rows will be added to the database. " & _
        vbNewLine & _
        vbNewLine & _
            "Otherwise, try again after new data has been released. Check the release schedule for more information.", , Title:="New data is unavailable."
    
        Application.StatusBar = vbNullString
        
    ElseIf CFTC_Retrieval_Error Then
    
        MsgBox "An error occured while attempting to retrieve CFTC data. Please check your internet connection"
        
    End If
    
    With TestTimers
        .EndTask totalTime
        Debug.Print .ToString
    End With
    
Finally:
    
    Re_Enable
    Running_Weekly_Retrieval = False
    
    Exit Sub

Deny_Debug_Mode:
    
    Debug_Mode = False
    Resume Retrieve_Latest_Data

ICE_Retrieval_Failed:
    
    Check_ICE = False
    Resume Finished_Querying_Weekly_Data

CFTC_Retrieval_Failed:
    
    CFTC_Retrieval_Error = True
    exitSubroutine = True
    
    Resume Stop_Timers_And_Update_If_Allowed
    'This will stop the data retrieval timer and the individualCotTask timer since uploadDataToDatabaseOrWorksheet will evaluate to False
    'ICE Data is dependent on new CFTC dates to query the correct URL
ExecutableFailed:
    
    Debug.Print Err.description
    Set executeableReturn = Nothing
    Resume CollectDetails

Catch_Block_Query_Failed:
    
    'If Not Scheduled_Retrieval Then
        
        MsgBox "Failed to upload data to database or worksheet." & vbNewLine & vbNewLine & _
                "Error code: " & Err.Number & vbNewLine & _
                "Description: " & Err.description & vbNewLine & _
                "Source: " & Err.Source & vbNewLine & vbNewLine & "Contact me at Mochim_UC@outlook.com"
    'End If
    
    Resume Next_Combined_Value
    
Catch_Database_Not_Found:
    If Legacy_Combined_Data Then exitSubroutine = True
    Resume Next_Report_Release_Type
End Sub

Private Sub Block_Query(ByRef Query, reportType As String, isDataFuturesAndOptions As Boolean, _
                        availableContractInfo As Collection, debugOnly As Boolean, _
                        mappedFields As Collection, Optional Overwrite_Worksheet As Boolean = False)
'===================================================================================================================
    'Purpose: Data within Query will be pruned for wanted columns and uploaded to either a database or worksheet.
    'Inputs: Query - Array that holds data to store.
    '        reportType - Report that is being uploaded.
    '        isDataFuturesAndOptions - True if data is futures + options; else, futures only.
    '        availableContractInfo - Collection of Contract instances. If on non-database file then this
    '                      contains only contracts within the file. IF database file then all contracts with a price symbol available.
    '        Overwrite_Worksheet - If True then all data on a worksheet will be replaced with the data matching its contract code.
    '        mappedFields - Collection of FieldInfo instances that represent each column in Query.
    'Outputs:
'===================================================================================================================
    Dim row As Long, iCount As Integer, databaseVersion As Boolean, uniqueContractCount As Integer
    
    Dim Block() As Variant, Contract_CLCTN As New Collection, priceColumn As Byte, contractCode As String
    
    Dim retrievePriceData As Boolean, placeDataOnWorksheetDebug As Boolean, contractCodeColumn As Byte

    Dim Progress_CHK As CheckBox, progressBarActive As Boolean
    
    Const IceContractCodeColumn As Byte = 4
    
    Set Progress_CHK = Weekly.Shapes("Progress_CHKBX").OLEFormat.Object
    
    #If Not DatabaseFile Then
    
        Dim columnFilter() As Variant, missingWeeksCount As Integer, Last_Calculated_Column As Integer, _
        current_Filters() As Variant, firstCalculatedColumn As Byte, WS_Data() As Variant, Table_Range As Range, Table_Data_CLCTN As New Collection
        
        Dim WantedColumnForAPI As New Collection, Z As Long, tempKey As String
        
        Const Time1 As Integer = 156, Time2 As Integer = 26, Time3 As Integer = 52

        Last_Calculated_Column = Variable_Sheet.Range("Last_Calculated_Column").Value2
        
        columnFilter = Filter_Market_Columns(True, False, False, reportType, Create_Filter:=True)
         
        ' +1 To account for filter returning a 1 based array
        ' +1 To get wanted value
        ' = +2
        priceColumn = UBound(filter(columnFilter, xlSkipColumn, False)) + 2
        firstCalculatedColumn = priceColumn + 2
        
    #Else
        priceColumn = UBound(Query, 2) + 1
        ReDim Preserve Query(LBound(Query, 1) To UBound(Query, 1), LBound(Query, 2) To priceColumn)  'Expand for calculations
        databaseVersion = True
    #End If
    
    Dim priceField As New FieldInfo: priceField.Constructor "price", priceColumn, "Price"
    mappedFields.Add priceField, "price"
    
    If (databaseVersion And isDataFuturesAndOptions And reportType = "L") Or Not databaseVersion Then
        
        contractCodeColumn = mappedFields("cftc_contract_market_code").ColumnIndex
        ReDim Block(1 To UBound(Query, 2))
        ' Parse array rows into collections keyed to their contract code.
        ' Array should be date sorted
        For row = LBound(Query, 1) To UBound(Query, 1)
        
            For iCount = LBound(Query, 2) To UBound(Query, 2)
                Block(iCount) = Query(row, iCount)
            Next iCount
            
            On Error GoTo Catch_MissingCollection
            Contract_CLCTN(Block(contractCodeColumn)).Add Block
            
        Next row
        
        Erase Query
        
        uniqueContractCount = Contract_CLCTN.count
    
        If uniqueContractCount = 0 Then
            MsgBox "An error occured. No unique contracts were retrieved." & reportType & "-C: " & isDataFuturesAndOptions
            Re_Enable
            End
        End If
        
        retrievePriceData = True
        placeDataOnWorksheetDebug = True
        
        On Error GoTo 0
        
        If debugOnly Then
            
            If MsgBox("Debug mode is active. Do you want to test price retrieval?", vbYesNo, "Test price retrieval?") = vbNo Then
                
                retrievePriceData = False
                
                #If DatabaseFile Then
                    GoTo Upload_Data
                #End If
                
            End If
            
            #If Not DatabaseFile Then
                If MsgBox("Test pasting to worksheet?", vbYesNo, "Test data paste?") = vbNo Then placeDataOnWorksheetDebug = False
            #End If
            
        End If
        
Block_Query_Main_Function:

        On Error GoTo 0

        If Progress_CHK.value = xlOn And Not databaseVersion Then
            ' Display Progress Bar control.
            ' Arguements are passed Byref and given values in the below Sub.
            With Progress_Bar
                .Show
                .InitializeValues CLng(uniqueContractCount)
            End With
            progressBarActive = True
        End If
        
        If Not Progress_CHK Is Nothing Then Set Progress_CHK = Nothing
        
        On Error GoTo 0
        
        For iCount = uniqueContractCount To 1 Step -1   'Loop list of wanted Contract Codes
            
            ' Removes the collection. A combined version orf collection elements will be added later.
            Block = CombineArraysInCollection(Contract_CLCTN(iCount), Append_Type.Multiple_1d)
            
            contractCode = Block(1, contractCodeColumn)
            Contract_CLCTN.Remove contractCode
            
            If HasKey(availableContractInfo, contractCode) Then
                
                #If Not DatabaseFile Then
                    Block = Filter_Market_Columns(False, True, False, reportType, False, Block, False, columnFilter)
                    ReDim Preserve Block(LBound(Block, 1) To UBound(Block, 1), LBound(Block, 2) To Last_Calculated_Column)  'Expand for calculations
                #End If
                
                If retrievePriceData Then Call TryGetPriceData(Block, priceColumn, availableContractInfo(contractCode), overwriteAllPrices:=False, datesAreInColumnOne:=Not databaseVersion) 'Gets Price_Info
            
            ElseIf Not databaseVersion Then
                GoTo NextAvailableContract
            End If
            
            #If Not DatabaseFile Then
            
                Set Table_Range = availableContractInfo(contractCode).TableSource.DataBodyRange
                
                missingWeeksCount = UBound(Block, 1)
                
                If Not Overwrite_Worksheet Then
                    '--Append New Data to bottom of already existing table data
                    WS_Data = Table_Range.Value2
                    
                    If UBound(WS_Data, 1) > 1 Then
                        If WS_Data(1, 1) > WS_Data(2, 1) Then WS_Data = Reverse_2D_Array(WS_Data)
                    End If
                    
                    With Table_Data_CLCTN
                        .Add Array(Old_Data, WS_Data), "Old"
                        .Add Array(Block_Data, Block), "Block"
                    End With
                    
                    Block = CombineArraysInCollection(Table_Data_CLCTN, Append_Type.Add_To_Old)
                
                    Set Table_Data_CLCTN = Nothing
                
                End If
                
                Select Case reportType
                    Case "L"
                        Block = Legacy_Multi_Calculations(Block, missingWeeksCount, firstCalculatedColumn, Time1, Time2)
                    Case "D"
                        Block = Disaggregated_Multi_Calculations(Block, missingWeeksCount, firstCalculatedColumn, Time1, Time2)
                    Case "T"
                        Block = TFF_Multi_Calculations(Block, missingWeeksCount, firstCalculatedColumn, Time1, Time2, Time3)
                End Select
                
                If placeDataOnWorksheetDebug Then
                
                    Call ChangeFilters(Table_Range.ListObject, current_Filters)
                                    
                    If Not Overwrite_Worksheet Then
                        Call Paste_To_Range(Data_Input:=Block, Sheet_Data:=WS_Data, Table_DataB_RNG:=Table_Range, Overwrite_Data:=Overwrite_Worksheet) 'Paste to bottom of table
                    Else
                        Call Paste_To_Range(Data_Input:=Block, Table_DataB_RNG:=Table_Range, Overwrite_Data:=Overwrite_Worksheet)
                    End If
                    
                    With Table_Range.ListObject.Sort
                        If .SortFields.count > 0 Then .Apply
                    End With
                        
                    Call RestoreFilters(Table_Range.ListObject, current_Filters)
                    
                End If
            #Else
                ' Adds array with price data retrieved to collection.
                Contract_CLCTN.Add Block, contractCode
            #End If
            
NextAvailableContract:
            
            If progressBarActive And Not databaseVersion Then
                If iCount = 1 Then
                    Unload Progress_Bar
                Else
                    Progress_Bar.IncrementBar
                End If
            End If

        Next iCount
        
        #If DatabaseFile Then
            ' Legacy Combined data needs to be rejoined together into a single array.
            Query = CombineArraysInCollection(Contract_CLCTN, Append_Type.Multiple_2d)
        #End If
        
    End If
    
Upload_Data:

    #If DatabaseFile Then
        On Error GoTo 0
        Call Update_Database(dataToUpload:=Query, uploadingFuturesAndOptions:=isDataFuturesAndOptions, _
        reportType:=reportType, debugOnly:=debugOnly, fieldInfoByEditedName:=mappedFields)
    
    #End If
    
    Exit Sub

Progress_Checkbox_Missing:

    'Set Progress_Control = Nothing
    Resume Block_Query_Main_Function
    
Catch_MissingCollection:

    Contract_CLCTN.Add New Collection, Block(contractCodeColumn)
    Resume
    
End Sub
Private Function Missing_Data(ByRef CFTC_Data As Variant, ByVal CFTC_Last_Updated_Day As Date, ByVal ICE_Last_Updated_Day As Date, ByRef ICE_Data As Variant, reportType As String, getFuturesAndOptions As Boolean, Download_ICE_Data As Boolean, Download_CFTC_Data As Boolean, Optional DebugMD As Boolean = False) As Variant 'Should change to function; Block will find the amount of missed time and download appropriate files
'===================================================================================================================
    'Purpose: Determines which files need to be downloaded for when multiple weeks of data have been missed
    'Inputs:
    'Outputs:
'===================================================================================================================
    
    Dim File_CLCTN As New Collection, MacB As Boolean, obj As Object, _
    Hyperlink_RNG As Range, New_Data As New Collection, iceUrl As Variant
    
    #If Mac Then
        MacB = True
    #End If
    
    If DebugMD Then
        If Not MacB Then If MsgBox("Do you want to test MAC OS data retrieval?", vbYesNo) = vbYes Then MacB = True
        
        If Download_CFTC_Data Then CFTC_Last_Updated_Day = DateAdd("yyyy", -2, CFTC_Last_Updated_Day)
        If Download_ICE_Data Then ICE_Last_Updated_Day = DateAdd("yyyy", -2, ICE_Last_Updated_Day)
    End If
    
    Application.DisplayAlerts = False
        
    If Download_ICE_Data Then
    
        Retrieve_Historical_Workbooks _
            Path_CLCTN:=File_CLCTN, _
            ICE_Contracts:=True, _
            CFTC_Contracts:=False, _
            Mac_User:=MacB, _
            reportType:=reportType, _
            downloadFuturesAndOptions:=getFuturesAndOptions, _
            ICE_Start_Date:=ICE_Last_Updated_Day, _
            ICE_End_Date:=ICE_Data(1, rawCftcDateColumn)
        
        'ICE_Query
        For Each iceUrl In File_CLCTN
            New_Data.Add ICE_Query(CStr(iceUrl), CDate(ICE_Last_Updated_Day))(IIf(getFuturesAndOptions, "futures+options", "futures-only"))
        Next
        'New_Data.Add Historical_Parse(File_CLCTN, retrieveCombinedData:=getFuturesAndOptions, reportType:=reportType, Yearly_C:=True, After_This_Date:=ICE_Last_Updated_Day, Kill_Previous_Workbook:=DebugMD)
    End If
    
    #If Not Mac Then
    
        If Download_CFTC_Data Then
            
            Set File_CLCTN = Nothing
            
            Retrieve_Historical_Workbooks _
                Path_CLCTN:=File_CLCTN, _
                ICE_Contracts:=False, _
                CFTC_Contracts:=True, _
                Mac_User:=MacB, _
                CFTC_Start_Date:=CFTC_Last_Updated_Day, _
                CFTC_End_Date:=CFTC_Data(1, rawCftcDateColumn), _
                reportType:=reportType, _
                downloadFuturesAndOptions:=getFuturesAndOptions
            
            New_Data.Add Historical_Parse(File_CLCTN, retrieveCombinedData:=getFuturesAndOptions, reportType:=reportType, Yearly_C:=True, After_This_Date:=CFTC_Last_Updated_Day, Kill_Previous_Workbook:=DebugMD)
            
        End If
        
    #End If
    
    Application.DisplayAlerts = True
    
    If New_Data.count = 1 Then
        Missing_Data = New_Data(1)
    ElseIf New_Data.count > 1 Then
        Missing_Data = CombineArraysInCollection(New_Data, Append_Type.Multiple_2d)
    End If

End Function
Public Function Weekly_ICE(Most_Recent_CFTC_Date As Date) As Collection

'Dim Path_CLCTN As New Collection

    Dim ICE_URL As String
    
    ICE_URL = Get_ICE_URL(Most_Recent_CFTC_Date)
    
    'If isMAC Then
    
        On Error GoTo Exit_Sub
        
        Set Weekly_ICE = ICE_Query(ICE_URL, Most_Recent_CFTC_Date)
        
    'Else
    
    '    On Error GoTo ICE_QueryT_Retrieval
    '
    '    With Path_CLCTN
    '
    '        .Add Environ("TEMP") & "\" & Date & "_Weekly_ICE.csv", "ICE"
    '
    '        If Dir(.Item(1)) = vbNullString Then Call Get_File(ICE_URL, .Item("ICE"))
    '
    '    End With
    '
    '    Weekly_ICE = Historical_Parse(Path_CLCTN, Weekly_ICE_Data:=True, After_This_Date:=LAst_Updated, reportType:=reportType, retrieveCombinedData:=Retrieve_Combined)
    
    'End If
    
    Exit Function
    
Exit_Sub:
    
End Function
Private Function Get_ICE_URL(Query_Date As Date) As String
'===================================================================================================================
    'Purpose: Creates a link to where to get ICE files based on Query_Date.
    'Inputs:  Query_Date - Date to create link for.
    'Outputs: URL sring that links to the representative Query_Date.
'===================================================================================================================

    Get_ICE_URL = "https://www.theice.com/publicdocs/cot_report/automated/COT_" & Format(Query_Date, "ddmmyyyy") & ".csv"
    'https://www.theice.com/publicdocs/cot_report/automated/COT_15112022.csv
End Function

Private Function ICE_Query(Weekly_ICE_URL As String, greaterThanDate As Date) As Collection

    Dim Data_Query As QueryTable, data As Variant, Data_Row() As Variant, URL As String, _
    Column_Filter() As Variant, Y As Byte, bb As Boolean, getFuturesAndOptions As Boolean, _
    Found_Data_Query As Boolean, Error_While_Refreshing As Boolean, Filtered_CLCTN As Collection
    
    Const connectionName As String = "ICE Data Refresh Connection", queryName = "ICE Data Refresh"
    
    With Application
    
        bb = .EnableEvents
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    For Each Data_Query In QueryT.QueryTables
    
        If Data_Query.name Like "*" & queryName & "*" Then
            Found_Data_Query = True
            Exit For
        End If
        
    Next Data_Query
    
    Column_Filter = Filter_Market_Columns(convert_skip_col_to_general:=True, reportType:="D", Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True, ICE:=True)
    
    If Not Found_Data_Query Then 'If QueryTable isn't found then create it
    
Recreate_Query:
    
        Set Data_Query = QueryT.QueryTables.Add(Connection:="TEXT;" & Weekly_ICE_URL, Destination:=QueryT.Range("A1"))
        
        With Data_Query
            
            .BackgroundQuery = False
            .SaveData = False
            .AdjustColumnWidth = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlOverwriteCells
            
            .TextFileColumnDataTypes = Column_Filter
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileCommaDelimiter = True
            
            .name = queryName
            
Try_NameConnection:
            .WorkbookConnection.RefreshWithRefreshAll = False
        End With
        
        On Error GoTo 0
        
        Erase Column_Filter
    
    Else
        ' Update Connection string
        On Error GoTo Catch_FailedConnectionUpdate
        With Data_Query
            .Connection = "TEXT;" & Weekly_ICE_URL
            .TextFileColumnDataTypes = Column_Filter
        End With
        
        On Error GoTo 0
        
    End If
    
    With Data_Query
        On Error GoTo Failed_To_Refresh 'Recreate Query and try again exactly 1 more time
        .Refresh False
        
        On Error GoTo Aggregation_Failed
        
        Set Filtered_CLCTN = New Collection
        
        With Filtered_CLCTN
            For Y = 1 To 2
                getFuturesAndOptions = Not getFuturesAndOptions
                .Add Historical_Excel_Aggregation(ThisWorkbook, getFuturesAndOptions, Date_Input:=CLng(greaterThanDate), ICE_Contracts:=True, QueryTable_To_Filter:=Data_Query), IIf(getFuturesAndOptions = True, "futures+options", "futures-only")
            Next Y
        End With
        
        With .ResultRange
            On Error Resume Next
            .Parent.ShowAllData
            .ClearContents 'Clear the Range
        End With
        
    End With
    
    With Application
        .DisplayAlerts = True
        .EnableEvents = bb
    End With
    
    Set ICE_Query = Filtered_CLCTN
    
    With Data_Query
        .WorkbookConnection.Delete
        .Delete
    End With
    
    Set Data_Query = Nothing
    
    Exit Function

Catch_FailedConnectionUpdate:
    
    If Err.Number = 1004 Then
        'Connection has been deleted.
        
    End If
    Err.Raise Err.Number, , Err.description
    
Catch_WorkbookConnectionNameTaken: 'Error handler is available when editing parameters for a new querytable and the connection name is already taken by a different query
    
    ThisWorkbook.Connections(connectionName).Delete
    
'    If Err.Number = 5 Then
'        'Invalid procedure call or arguement.
'    End If
    
    Resume Try_NameConnection
    
Failed_To_Refresh:
        
    With Data_Query
        .WorkbookConnection.Delete
        .Delete
    End With
    
    If Error_While_Refreshing = True Then
        On Error GoTo 0
        Err.Raise 5
    Else
        Error_While_Refreshing = True
        Resume Recreate_Query
    End If
    
Aggregation_Failed:

End Function

#If Not DatabaseFile Then

    Private Sub Convert_Workbook_Version()
    '===================================================================================================================
        'Purpose: Prompts the user for whether they wanat to download futures only or futures + options data.
        '         All data in the worksheet will then be replaced.
        'Inputs:
        'Outputs:
    '===================================================================================================================

        Dim User_Selection As Long, Retrieve_Futures_Only As Boolean
        
        User_Selection = CLng(InputBox("( 1 ) for Futures Only" & vbNewLine & vbNewLine & "{ 2 ) for Futures + Options."))
        
        If User_Selection = 1 Then
            Retrieve_Futures_Only = True
        ElseIf User_Selection = 2 Then
            Retrieve_Futures_Only = False
        Else
            Exit Sub
        End If
        
        With Variable_Sheet
            .Range("Combined_Workbook").Value2 = Not Retrieve_Futures_Only
            .Range("Last_Updated_CFTC").Value2 = 0
        
            If ReturnReportType = "D" Then .Range("Last_Updated_ICE").Value2 = 0
        
        End With
        
        Call New_Data_Query(Scheduled_Retrieval:=True, Overwrite_All_Data:=True)
        
        MsgBox "Conversion Complete"
        
    End Sub

    Public Sub New_CFTC_Data()
    
        Dim Current_Contracts As Collection, retrieveCombinedData As Boolean, File_Paths As New Collection, _
        wantedContractCode As String, invalidContractCode As Boolean, New_Data() As Variant, First_Calculated_Column As Byte
        
        Dim WS As Worksheet, Symbol_Row As Long, I As Byte, reportType As String
    
        Set Current_Contracts = GetAvailableContractInfo
        
        reportType = ReturnReportType
        
        retrieveCombinedData = IsWorkbookForFuturesAndOptions()
        
        First_Calculated_Column = 3 + WorksheetFunction.CountIf(Variable_Sheet.ListObjects(reportType & "_User_Selected_Columns").DataBodyRange.columns(2), True)
        
        Do
            wantedContractCode = InputBox("Enter a 6 digit CFTC contract code")
            
            If wantedContractCode = "" Then Exit Sub
            
            If HasKey(Current_Contracts, wantedContractCode) Then
                MsgBox "Contract Code is already present within the workbook"
                invalidContractCode = True
            Else
                invalidContractCode = False
            End If
            
        Loop While Len(wantedContractCode) <> 6 Or invalidContractCode
                
        On Error GoTo No_Data_Retrieved_From_API
        
        Dim columnMap As New Collection
        New_Data = CFTC_API_Method(reportType, retrieveCombinedData, DateSerial(1970, 1, 1), False, columnMap, wantedContractCode)
        
        New_Data = Filter_Market_Columns(False, True, False, reportType, True, New_Data, False)
        
        On Error GoTo 0
        
        wantedContractCode = New_Data(1, UBound(New_Data, 2))
        
        ReDim Preserve New_Data(LBound(New_Data, 1) To UBound(New_Data, 1), LBound(New_Data, 2) To UBound(New_Data, 2) + 1)
        
        With Range("Symbols_TBL")
        
            Symbol_Row = WorksheetFunction.Match(wantedContractCode, .columns(1), 0)
            
            If Symbol_Row <> 0 Then
                
                For I = 3 To 4
                    
                    If Not IsEmpty(.Cells(Symbol_Row, I)) Then
                        Call TryGetPriceData(New_Data, UBound(New_Data, 2), Array(.Cells(Symbol_Row, I), IIf(I = 3, True, False)), datesAreInColumnOne:=True, overwriteAllPrices:=True)
                        Exit For
                    End If
                Next I
        
            End If
        
        End With
        
        ReDim Preserve New_Data(LBound(New_Data, 1) To UBound(New_Data, 1), LBound(New_Data, 2) To Range("Last_Calculated_Column").Value2)
        
        Select Case reportType
            Case "L"
                New_Data = Legacy_Multi_Calculations(New_Data, UBound(New_Data, 1), First_Calculated_Column, 156, 26)
            Case "D"
                New_Data = Legacy_Multi_Calculations(New_Data, UBound(New_Data, 1), First_Calculated_Column, 156, 26)
            Case "T"
                New_Data = TFF_Multi_Calculations(New_Data, UBound(New_Data, 1), First_Calculated_Column, 156, 26, 52)
        End Select
        
        Application.ScreenUpdating = False
        
        Set WS = ThisWorkbook.Worksheets.Add
        
        With WS
        
            .columns(1).NumberFormat = "yyyy-mm-dd"
            .columns(First_Calculated_Column - 3).NumberFormat = "@"
            
            Call Paste_To_Range(Sheet_Data:=New_Data, Historical_Paste:=True, Target_Sheet:=WS)
            
            WS.ListObjects(1).name = "CFTC_" & wantedContractCode
            
        End With
        
        If Symbol_Row = 0 Then
        
            Range("Symbols_TBL").ListObject.ListRows.Add.Range.Value2 = Array(wantedContractCode, New_Data(UBound(New_Data, 1), 2), Empty, Empty)
            
            MsgBox "A new row has been added to the availbale symbols table. Please fill in the missing Symbol information if available."
        
        End If
        
        Re_Enable
    
    Exit Sub
No_Data_Retrieved_From_API:
    
    Re_Enable
    MsgBox "Data couldn't be retrieved from API"
    
    End Sub
#Else
    Function ConvertSymbolDataToJson() As String
    
        Dim stuff As New Collection, CD As ContractInfo, Item As Variant, quote As String
        
        quote = "\" & Chr(34)
        
        With stuff
            For Each Item In GetAvailableContractInfo
                Set CD = Item
                .Add (quote & CD.contractCode & quote & ":" & quote & CD.priceSymbol & quote)
            Next
        End With
        
       ConvertSymbolDataToJson = "{" & Join(ConvertCollectionToArray(stuff), ",") & "}"
        
    End Function
    Function ListDatabasePathsInJson() As String
    
        Dim Item As Variant, reportDetails As LoadedData, stuff As New Collection, quote As String
        
        quote = "\" & Chr(34)
        
        For Each Item In Array("L", "D", "T")
            Set reportDetails = GetStoredReportDetails(CStr(Item))
            stuff.Add quote & IIf(Item = "T", "TFF", reportDetails.FullReportName) & quote & ":" & quote & Replace(reportDetails.CurrentDatabasePath, "\", "\\") & quote
        Next Item
        ListDatabasePathsInJson = "{" & Join(ConvertCollectionToArray(stuff), ",") & "}"
        
    End Function
    Function RunCSharpExtractor() As Collection
    
        Dim commandArgs() As Variant, cmd As String
            
        ReDim commandArgs(2)
        commandArgs(0) = Range("CSharp_Exe").Value2
        
        If LenB(Dir(commandArgs(0))) = 0 Then Err.Raise 53, , "C# executable couldn't be found."
        
        commandArgs(1) = ListDatabasePathsInJson()
        commandArgs(2) = ConvertSymbolDataToJson()
        
        cmd = Join(QuotedForm(commandArgs), " ")
       
        Dim oShell As Object
        Set oShell = CreateObject("WScript.Shell")
    
        'run command
        Dim oExec As Object
        Dim oOutput As Object
        
        Application.StatusBar = "Querying new data with " & commandArgs(0)
        
        Set oExec = oShell.Exec(cmd)
        Set oOutput = oExec.StdOut
        
        Dim result As String
        
        result = oOutput.ReadAll
        oExec.Terminate
        Dim innerCollection As Collection, output As New Collection, I As Integer, programResponse() As String, report As String, reportInfo() As String, dataStart As Byte, Item As Variant, kvp() As String
        
        programResponse = Split(result, vbNewLine)
        
        For I = UBound(programResponse) - 1 To UBound(programResponse) - 6 Step -1
            
            On Error Resume Next
            
            report = Left(programResponse(I), 1)
            ' Ensure there is a collection for each Report Type.
            output.Add New Collection, report
            
            dataStart = InStr(1, programResponse(I), "{") + 1
            reportInfo = Split(Mid$(programResponse(I), dataStart, Len(programResponse(I)) - dataStart), ",")
            
            Set innerCollection = New Collection
            On Error GoTo 0
            'Add an inner collection keyed to whether or not the data is combined.
            output(report).Add innerCollection, IIf(InStrB(1, LCase$(programResponse(I)), "true") > 0, CStr(True), CStr(False))
            
            Dim elementName As String
            
            With innerCollection
                For Each Item In reportInfo
                    kvp = Split(Item, ":")
                    On Error GoTo DefaultStringAddition
                    
                    elementName = LCase$(Trim$(kvp(0)))
                    kvp(1) = Trim$(kvp(1))
                    
                    Select Case elementName
                        Case "combined"
                            .Add CBool(kvp(1)), elementName
                        Case "status"
                            .Add CByte(kvp(1)), elementName
                        Case "latest date"
                            .Add CDate(kvp(1)), elementName
                        Case Else
                            .Add kvp(1), elementName
                    End Select
                                    
                Next Item
            End With
        Next I
        
        Set RunCSharpExtractor = output
        Debug.Print result
        Application.StatusBar = vbNullString
        
        Exit Function
        
DefaultStringAddition:
        innerCollection.Add Trim$(kvp(1)), Trim$(kvp(0))
        Resume Next
    End Function

#End If


