Attribute VB_Name = "Data_Retrieval"
Public Const rawCftcDateColumn As Byte = 3
Public Data_Updated_Successfully As Boolean

Public Const ERROR_RETRIEVAL_FAILED As Long = vbObjectError + 513
Public Const ERROR_SOCRATA_SUCCESS_NO_DATA As Long = vbObjectError + 514
Private Const ERROR_BLOCK_QUERY_NO_UNIQUE_CONTRACTS As Long = vbObjectError + 520

Public Enum ReportStatusCode
    NoUpdateAvailable = 0
    Updated = 1
    Failure = 2
    attemptingRetrieval = 3
    AttemptingUpdate = 4
    NotInitialized = 5
    CheckingDataAvailability = 6
End Enum

Public Enum Append_Type
    Add_To_Old = 1
    Multiple_1d = 2
    Multiple_2d = 3
End Enum

Public Enum ArrayID
    Block_Data = 14
    Old_Data = 44
    Weekly_Data = 33
End Enum

Public Enum OpenInterestType
    FuturesAndOptions = -1
    FuturesOnly = 0
    OptionsOnly = 1
End Enum

Option Explicit
Sub New_Data_Query(Optional Scheduled_Retrieval As Boolean = False, Optional Overwrite_All_Data As Boolean = False, Optional IsWorbookOpenEvent As Boolean = False)
'===================================================================================================================
    'Purpose: Retrieves CFTC data that hasn't been stored either on a worksheet or database.
    'Inputs: Scheduled_Retrieval - If true then error messages will not be displayed.
    '        Overwrite_All_Data - Reserved for non-Database version. If true then all data on a worksheet will be replaced.
    'Outputs:
'===================================================================================================================

    Dim Last_Update_CFTC As Date, CFTC_Incoming_Date As Date, ICE_Incoming_Date As Date, Last_Update_ICE As Date

    Dim cftcDataA() As Variant, iceDataA() As Variant, Historical_Query() As Variant, reportsToQuery$()

    Dim report$, iReport As Byte, iOiType As Byte

    Dim cftcDateRange As Range, iceKey$
    'Booleans
    Dim debugWeeklyRetrieval As Boolean, DBM_Historical_Retrieval As Boolean, Debug_Mode As Boolean, _
    Download_CFTC As Boolean, Download_ICE As Boolean, Check_ICE As Boolean, queryFuturesAndOptions As Boolean, _
    uploadDataToDatabaseOrWorksheet As Boolean, isLegacyCombined As Boolean, _
    exitSubroutine As Boolean, CFTC_Retrieval_Error As Boolean, newDataSuccessfullyHandled As Boolean, processingReport As Boolean
    'Collections
    Dim cftcMappedFieldInfo As Collection, DataBase_Not_Found_CLCTN As Collection, _
    availableContractInfo As Collection, Data_CLCTN As Collection, Weekly_ICE_CLCTN As Collection
    
    Dim TestTimers As New TimedTask, individualCotTask As TimedTask

    Const dataRetrieval$ = "C.O.T data retrieval", uploadTime$ = "Upload Time", _
    databaseDateQuery$ = "Query database for latest date.", cRetrievalTask$ = "C# Retrieval", procedureName$ = "New_Data_Query"
     
    Const legacy_initial$ = "L"
    
    Dim useSocrataAPI As Boolean, socrataApiFailed As Boolean, executableSuccess As Boolean
        
    Call IncreasePerformance

    On Error GoTo Deny_Debug_Mode
        If Weekly.Shapes("Test_Toggle").OLEFormat.Object.value = xlOn Then Debug_Mode = True 'Determine if Debug status
    On Error GoTo Catch_General_Error
    
    #If DatabaseFile Then
        
        Dim openInterestTypesToQuery(1) As OpenInterestType, executeableReturn As Collection, _
        databaseReturnedDate As Boolean
        Const latestDateKey$ = "latest date"
        #If Mac Then
            GateMacAccessToWorkbook
        #End If
        'Legacy data must be retrieved first so that price data only needs to be retrieved once.
        reportsToQuery = Split("L,D,T", ",")
        
        openInterestTypesToQuery(0) = OpenInterestType.FuturesAndOptions
        openInterestTypesToQuery(1) = OpenInterestType.FuturesOnly
        
        Set cftcDateRange = Variable_Sheet.Range("Last_Updated_CFTC")
        
    #Else
        Dim openInterestTypesToQuery(0) As OpenInterestType, iceDateRange As Range
        
        ReDim reportsToQuery(0): reportsToQuery(0) = ReturnReportType
        openInterestTypesToQuery(0) = IsWorkbookForFuturesAndOptions
        
        With Variable_Sheet
            Set cftcDateRange = .Range("Last_Updated_CFTC")
            If reportsToQuery(0) = "D" Then Set iceDateRange = .Range("Last_Updated_ICE")
        End With
        
    #End If
            
    TestTimers.Start "New Data Query [" & Time & "]"
    
    If Debug_Mode = True Then
        
        With TestTimers
            .Pause
        
            If MsgBox("Test Weekly Data Retrieval ?", vbYesNo, "Choose what to debug") = vbYes Then
                debugWeeklyRetrieval = True
            ElseIf MsgBox("Test Multi-Week Historical Retrieval ?", vbYesNo, "Choose what to debug") = vbYes Then
                DBM_Historical_Retrieval = True
            End If
    
            #If DatabaseFile Then
            
                Do
                    report = UCase(InputBox("Select 1 of L,D,T"))
                Loop While IsError(Application.Match(report, reportsToQuery, 0))
                
                reportsToQuery(0) = report
                            
                On Error Resume Next
                ' Get the combined status to test.
                Do
                    queryFuturesAndOptions = CBool(InputBox("Select 1 of the following:" & vbNewLine & vbNewLine & "  Futures Only: 0" & vbNewLine & "  Futures + Options: 1"))
                Loop While Err.Number <> 0
                
                openInterestTypesToQuery(0) = queryFuturesAndOptions
            
            #End If
            
            On Error GoTo Catch_General_Error
            .Continue
        End With
    Else
        #If DatabaseFile Then
            With TestTimers.StartSubTask(cRetrievalTask)
                On Error GoTo Catch_ExecutableFailed
                Set executeableReturn = RunCSharpExtractor()
                On Error GoTo Catch_General_Error
                .EndTask
            End With
        #End If
    End If

Retrieve_Latest_Data:

    For iReport = LBound(reportsToQuery) To UBound(reportsToQuery)
        
        report = reportsToQuery(iReport)
        
        For iOiType = LBound(openInterestTypesToQuery) To UBound(openInterestTypesToQuery)
            
            processingReport = True
            
            On Error GoTo Catch_General_Error
            
            queryFuturesAndOptions = openInterestTypesToQuery(iOiType)
            uploadDataToDatabaseOrWorksheet = False
            isLegacyCombined = (queryFuturesAndOptions And report = legacy_initial)
            Check_ICE = False
            
            #If DatabaseFile Then
                                
                If Not Debug_Mode And iOiType = LBound(openInterestTypesToQuery) And iReport = LBound(reportsToQuery) And Not isLegacyCombined Then
                    MsgBox "Legacy Combined data needs to be retrieved first so that price data only has to be retrieved once."
                    GoTo Exit_Procedure
                End If
                
                On Error GoTo Try_Get_New_Data
                
                If Not executeableReturn Is Nothing Then
                
                    With executeableReturn(report)(CStr(queryFuturesAndOptions))
                        Select Case .Item("status")
                            Case ReportStatusCode.NoUpdateAvailable, ReportStatusCode.Updated
                                
                                executableSuccess = True
                                
                                If CFTC_Incoming_Date < .Item(latestDateKey) Then
                                    CFTC_Incoming_Date = .Item(latestDateKey)
                                End If
                                
                                If .Item(latestDateKey) > cftcDateRange.Value2 Or .Item("status") = ReportStatusCode.Updated Then
                                    With GetStoredReportDetails(report)
                                        Select Case .OpenInterestType.Value2
                                            Case queryFuturesAndOptions, OpenInterestType.OptionsOnly
                                                .PendingUpdateInDatabase.Value2 = True
                                        End Select
                                    End With
                                    newDataSuccessfullyHandled = True
                                End If
                            Case Else
                                executableSuccess = False
                        End Select
                    End With
                    
                    If executableSuccess And Not Debug_Mode Then
                        GoTo Next_Combined_Value
                    End If
                End If
            #End If
            
Try_Get_New_Data:
            
            On Error GoTo Catch_General_Error
            useSocrataAPI = Not Overwrite_All_Data
            
            Set individualCotTask = TestTimers.StartSubTask("( " & report & " ) Combined: (" & queryFuturesAndOptions & ")")
            
            #If DatabaseFile Then

                With individualCotTask.StartSubTask(databaseDateQuery)
                    databaseReturnedDate = TryGetLatestDate(Last_Update_CFTC, reportType:=report, versionToQuery:=openInterestTypesToQuery(iOiType), queryIceContracts:=False)
                        
                    If Not databaseReturnedDate Then
                        If DataBase_Not_Found_CLCTN Is Nothing Then Set DataBase_Not_Found_CLCTN = New Collection
                        DataBase_Not_Found_CLCTN.Add "Missing database for " & Evaluate("VLOOKUP(""" & report & """,Report_Abbreviation,2,FALSE)")
                        'The Legacy_Combined data is the only one for which price data is queried.
                        If isLegacyCombined Then exitSubroutine = True
                     End If
                     
                    .EndTask
                End With
                
                If Not databaseReturnedDate Then Exit For
                
            #Else
                Last_Update_CFTC = cftcDateRange.Value2
            #End If
            
            individualCotTask.StartSubTask dataRetrieval
            
            Dim cftcAlreadyUpdated As Boolean, iceAlreadyUpdated As Boolean
            
            If CFTC_Incoming_Date = TimeSerial(0, 0, 0) Or Last_Update_CFTC < CFTC_Incoming_Date Or Debug_Mode Then
                         
                On Error GoTo Catch_CFTCRetrievalFailed
                
                Set cftcMappedFieldInfo = Nothing
                
                cftcDataA = HTTP_Weekly_Data(Last_Update_CFTC, suppressMessages:=Scheduled_Retrieval, _
                                retrieveCombinedData:=queryFuturesAndOptions, _
                                reportType:=report, useApi:=useSocrataAPI, _
                                columnMap:=cftcMappedFieldInfo, _
                                testAllMethods:=debugWeeklyRetrieval)
                
                socrataApiFailed = Not useSocrataAPI
                CFTC_Incoming_Date = cftcDataA(UBound(cftcDataA, 1), rawCftcDateColumn)
                cftcAlreadyUpdated = (CFTC_Incoming_Date = Last_Update_CFTC)
                
            Else
                'New data not available for Non Legacy Combined
                GoTo Stop_Timers_And_Update_If_Allowed
            End If
            
Try_IceRetrieval:
            If report = "D" Then
                
                On Error GoTo ICE_Retrieval_Failed
                                               
                #If DatabaseFile Then
                    Call TryGetLatestDate(Last_Update_ICE, reportType:=report, versionToQuery:=openInterestTypesToQuery(iOiType), queryIceContracts:=True)
                #Else
                    Last_Update_ICE = iceDateRange.Value2
                #End If
                
                If Last_Update_ICE < CFTC_Incoming_Date Then
                    
                    Check_ICE = True
                    
                    If Weekly_ICE_CLCTN Is Nothing Then
                        Set Weekly_ICE_CLCTN = Weekly_ICE(CDate(cftcDataA(UBound(cftcDataA, 1), rawCftcDateColumn)))
                    End If
                    
                    iceKey = IIf(queryFuturesAndOptions = True, "futures+options", "futures-only")
                    
                    With Weekly_ICE_CLCTN
                        iceDataA = .Item(iceKey)
                        .Remove iceKey
                        ' if empty or not using DatabaseFile
                        If .count = 0 Or UBound(openInterestTypesToQuery) = 0 Then Set Weekly_ICE_CLCTN = Nothing
                    End With
                    
                    ICE_Incoming_Date = iceDataA(1, rawCftcDateColumn)
                    iceAlreadyUpdated = (ICE_Incoming_Date = Last_Update_ICE)
                
                End If

            End If
            
Finished_Querying_Weekly_Data:
            
            On Error GoTo Catch_General_Error
            
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
            
            If Not Debug_Mode And cftcAlreadyUpdated And (Not Check_ICE Or (Check_ICE And iceAlreadyUpdated)) Then
                If isLegacyCombined Then exitSubroutine = True
            ElseIf DBM_Historical_Retrieval Or ((socrataApiFailed And CFTC_Incoming_Date - Last_Update_CFTC > 7) Or (Check_ICE And ICE_Incoming_Date - Last_Update_ICE > 7)) Then
                
                If (socrataApiFailed And CFTC_Incoming_Date - Last_Update_CFTC > 7) Or DBM_Historical_Retrieval Then
                    
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
                
                Download_ICE = (Check_ICE And (ICE_Incoming_Date - Last_Update_ICE > 7 Or DBM_Historical_Retrieval))
                
                On Error GoTo Catch_General_Error

                Historical_Query = Missing_Data(getFuturesAndOptions:=queryFuturesAndOptions, _
                    maxDateICE:=ICE_Incoming_Date, maxDateCFTC:=CFTC_Incoming_Date, _
                    Download_ICE_Data:=Download_ICE, Download_CFTC_Data:=Download_CFTC, _
                    reportType:=report, _
                    CFTC_Last_Updated_Day:=Last_Update_CFTC, ICE_Last_Updated_Day:=Last_Update_ICE + 2, _
                    DebugMD:=DBM_Historical_Retrieval)
                                                 
                If Not (Download_ICE And Download_CFTC) Then
                    
                    If Download_CFTC And (Check_ICE And ICE_Incoming_Date - Last_Update_ICE > 0) Then
                        'Determine if the most recently queried Ice Data needs to be added
                        Set Data_CLCTN = New Collection
                        
                        With Data_CLCTN
                            .Add Historical_Query
                            .Add iceDataA
                        End With
                        Historical_Query = CombineArraysInCollection(Data_CLCTN, Append_Type.Multiple_2d)
                                                                        
                    ElseIf Download_ICE And CFTC_Incoming_Date - Last_Update_CFTC > 0 Then
                        'Determine if CFTC data needs to be added
                        Set Data_CLCTN = New Collection
                        
                        With Data_CLCTN
                            .Add Historical_Query
                            .Add cftcDataA
                        End With
                        
                        Historical_Query = CombineArraysInCollection(Data_CLCTN, Append_Type.Multiple_2d)
                    End If
                End If
                uploadDataToDatabaseOrWorksheet = True
                
            ElseIf (CFTC_Incoming_Date - Last_Update_CFTC) > 0 Or (Check_ICE And ICE_Incoming_Date - Last_Update_ICE > 0) Or Debug_Mode = True Then
                
                Set Data_CLCTN = New Collection
                With Data_CLCTN
                    If Check_ICE And (ICE_Incoming_Date - Last_Update_ICE > 0 Or Debug_Mode) Then .Add iceDataA
                    If CFTC_Incoming_Date - Last_Update_CFTC > 0 Or Debug_Mode Then .Add cftcDataA

                    If .count = 1 Then
                        Historical_Query = .Item(1)
                    ElseIf .count = 2 Then
                        Historical_Query = CombineArraysInCollection(Data_CLCTN, Append_Type.Multiple_2d)
                    End If
                End With
                uploadDataToDatabaseOrWorksheet = True
                
            End If
            
            Set Data_CLCTN = Nothing
            Erase cftcDataA
            If report = "D" Then Erase iceDataA
            
Stop_Timers_And_Update_If_Allowed:

            With individualCotTask
                
                .StopSubTask dataRetrieval
            
                If IsArrayAllocated(Historical_Query) And uploadDataToDatabaseOrWorksheet = True And Not exitSubroutine Then
                    ' Upload data to database/spreadsheet and retrieve prices.
                    With .StartSubTask(uploadTime)
                        
                        If availableContractInfo Is Nothing Then
                            On Error GoTo Catch_General_Error
                            Set availableContractInfo = GetAvailableContractInfo()
                            
                        End If
                        
                        On Error GoTo Catch_Block_Query_Failed
                        Call Block_Query(query:=Historical_Query, reportType:=report, _
                                        isDataFuturesAndOptions:=openInterestTypesToQuery(iOiType), availableContractInfo:=availableContractInfo, _
                                        debugOnly:=Debug_Mode, mappedFields:=cftcMappedFieldInfo, Overwrite_Worksheet:=Overwrite_All_Data)
                        On Error GoTo Catch_General_Error
                        Erase Historical_Query
                        
                        .EndTask
                    End With
                    
                    newDataSuccessfullyHandled = True
                    
                End If
                
                .EndTask
                
            End With
            
Next_Combined_Value:
            On Error GoTo Catch_General_Error
            processingReport = False
            If Debug_Mode Or exitSubroutine Then Exit For
        Next iOiType
        
Next_Report_Release_Type:
        On Error GoTo Catch_General_Error
        
        If Not executableSuccess Then
            With individualCotTask
                If .isRunning Then .EndTask
            End With
        End If
        
        If Debug_Mode Or exitSubroutine Then Exit For
    Next iReport
    
Exit_Procedure:
    
    On Error GoTo Catch_General_Error
    
    If newDataSuccessfullyHandled And Not exitSubroutine And Not CFTC_Retrieval_Error Then
        
        Data_Updated_Successfully = True
        '-------------------------------------------------------------------------------------------
        
        With cftcDateRange
            If CFTC_Incoming_Date > .Value2 Then
                'Update_Text CFTC_Incoming_Date   'Update Text Boxes "My_Date" on the HUB and Weekly worksheets.
                .Value2 = CFTC_Incoming_Date
            End If
        End With
        
        #If Not DatabaseFile Then
            If Check_ICE And Not iceDateRange Is Nothing Then
                With iceDateRange
                    If ICE_Incoming_Date > .Value2 Then
                        .Value2 = ICE_Incoming_Date
                    End If
                End With
            End If
        #End If
        '----------------------------------------------------------------------------------------
        If Not Debug_Mode Then
        
            If Not Scheduled_Retrieval Then HUB.Activate 'If ran manually then bring the User to the HUB
            Courtesy                                     'Change Status Bar_Message
            
            If newDataSuccessfullyHandled Then
            
                #If DatabaseFile Then
                                        
                    With ThisWorkbook
                        On Error Resume Next
                        RefreshTableData .Worksheets(.ActiveSheet.Name).WorksheetReportType
                        On Error GoTo Catch_General_Error
                    End With
                    
                    With TestTimers.StartSubTask("Query all databases for latest contracts.")
                         Latest_Contracts
                        .EndTask
                    End With
                                    
                #End If
                
            End If
            
        End If
        
    ElseIf Not DataBase_Not_Found_CLCTN Is Nothing And Not Scheduled_Retrieval Then
        
        With DataBase_Not_Found_CLCTN
            For iReport = 1 To .count
                MsgBox .Item(iReport)
            Next iReport
        End With
        
    ElseIf Not (Scheduled_Retrieval Or Debug_Mode) Then
    
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
        
    End If
    
Finally:
    With TestTimers
        .EndTask
        .DPrint
    End With
    
    Re_Enable
    Exit Sub

Deny_Debug_Mode:
    
    Debug_Mode = False
    Resume Next

ICE_Retrieval_Failed:
    
    Check_ICE = False
    Resume Finished_Querying_Weekly_Data

Catch_CFTCRetrievalFailed:

    Select Case Err.Number
        Case ERROR_SOCRATA_SUCCESS_NO_DATA
            'Retrieval didn't fail. Just no new data.
            CFTC_Incoming_Date = Last_Update_CFTC
            socrataApiFailed = False
            cftcAlreadyUpdated = True
            Resume Try_IceRetrieval
        Case ERROR_RETRIEVAL_FAILED
            MsgBox "Data retrieval methods have failed." & vbNewLine & vbNewLine & _
           "Check your internet connection. If this error persists please contact me at MoshiM_UC@outlook.com with your operating system and Excel version."
        Case Else
            DisplayErr Err, procedureName
    End Select
    
    CFTC_Retrieval_Error = True
    exitSubroutine = True
    Resume Stop_Timers_And_Update_If_Allowed
    
#If DatabaseFile Then
Catch_ExecutableFailed:
    TestTimers.SubTask(cRetrievalTask).EndTask
    If IsOnCreatorComputer Then DisplayErr Err, procedureName, "Failed to parse values from .exe response."
    Set executeableReturn = Nothing
    Resume Next
#End If

Catch_General_Error:
    DisplayErr Err, IIf(processingReport, report & "_" & ConvertOpenInterestTypeToName(CLng(queryFuturesAndOptions)), vbNullString)
    Resume Finally
    
Catch_Block_Query_Failed:
    Erase Historical_Query
    DisplayErr Err, procedureName, IIf(processingReport, report & "_" & ConvertOpenInterestTypeToName(CLng(queryFuturesAndOptions)), vbNullString)
    Resume Next_Combined_Value
    
Catch_Database_Not_Found:
    If isLegacyCombined Then exitSubroutine = True
    Resume Next_Report_Release_Type
    
End Sub
Private Sub Block_Query(ByRef query() As Variant, reportType$, isDataFuturesAndOptions As OpenInterestType, _
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
    Dim iRow As Long, iCount As Long, databaseVersion As Boolean, uniqueContractCount As Long
    
    Dim Block() As Variant, Contract_CLCTN As New Collection, priceColumn As Byte, contractCode$
    
    Dim retrievePriceData As Boolean, placeDataOnWorksheetDebug As Boolean, contractCodeColumn As Byte, yahooCookie$

    Dim Progress_CHK As CheckBox, progressBarActive As Boolean
    
    Const IceContractCodeColumn As Byte = 4
    
    On Error GoTo Catch_GeneralError
    Set Progress_CHK = Weekly.Shapes("Progress_CHKBX").OLEFormat.Object
    
    #If Not DatabaseFile Then
    
        Dim columnfilter() As Variant, missingWeeksCount As Long, Last_Calculated_Column As Long, _
        current_Filters() As Variant, firstCalculatedColumn As Byte, WS_Data() As Variant, Table_Range As Range, Table_Data_CLCTN As New Collection
        
        Dim WantedColumnForAPI As New Collection, Z As Long, tempKey$
        
        Const Time1 As Long = 156, Time2 As Long = 26, Time3 As Long = 52

        Last_Calculated_Column = Variable_Sheet.Range("Last_Calculated_Column").Value2
        
        columnfilter = Filter_Market_Columns(True, False, False, reportType, Create_Filter:=True)
         
        ' +1 To account for filter returning a 1 based array
        ' +1 To get wanted value
        ' = +2
        priceColumn = UBound(filter(columnfilter, xlSkipColumn, False)) + 2
        firstCalculatedColumn = priceColumn + 2
        
    #Else
        priceColumn = UBound(query, 2) + 1
        ReDim Preserve query(LBound(query, 1) To UBound(query, 1), LBound(query, 2) To priceColumn)  'Expand for calculations
        databaseVersion = True
    #End If
    
    Dim priceField As New FieldInfo: priceField.Constructor "price", priceColumn, "Price"
    mappedFields.Add priceField, "price"
    
    If (databaseVersion And isDataFuturesAndOptions = OpenInterestType.FuturesAndOptions And reportType = "L") Or Not databaseVersion Then
        
        contractCodeColumn = mappedFields("cftc_contract_market_code").ColumnIndex
        ReDim Block(1 To UBound(query, 2))
        ' Parse array rows into collections keyed to their contract code.
        ' Array should be date sorted
        For iRow = LBound(query, 1) To UBound(query, 1)
        
            For iCount = LBound(query, 2) To UBound(query, 2)
                Block(iCount) = query(iRow, iCount)
            Next iCount
            On Error GoTo Catch_MissingCollection
            Contract_CLCTN(Block(contractCodeColumn)).Add Block
            
        Next iRow
        On Error GoTo Catch_GeneralError
        Erase query
        
        uniqueContractCount = Contract_CLCTN.count
        
        If uniqueContractCount = 0 Then
            On Error GoTo 0
            Err.Raise ERROR_BLOCK_QUERY_NO_UNIQUE_CONTRACTS, "Block_Query", "No unique contracts available for " & reportType & "-[Futures_Options: " & CBool(isDataFuturesAndOptions) & "]"
        End If
        
        retrievePriceData = True
        placeDataOnWorksheetDebug = True
        
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

        For iCount = uniqueContractCount To 1 Step -1   'Loop list of wanted Contract Codes
            
            ' Removes the collection. A combined version orf collection elements will be added later.
            Block = CombineArraysInCollection(Contract_CLCTN(iCount), Append_Type.Multiple_1d)
            
            contractCode = Block(1, contractCodeColumn)
            ' Remove Collection.
            Contract_CLCTN.Remove contractCode
            
            If HasKey(availableContractInfo, contractCode) Then
                
                #If Not DatabaseFile Then
                    Block = Filter_Market_Columns(False, True, False, reportType, False, Block, False, columnfilter)
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
                
                Erase Block
                Erase WS_Data
                Set Table_Range = Nothing
                
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
            query = CombineArraysInCollection(Contract_CLCTN, Append_Type.Multiple_2d)
        #End If
        
    End If
    
Upload_Data:
    #If DatabaseFile And Not Mac Then
        Call Update_Database(dataToUpload:=query, versionToUpdate:=isDataFuturesAndOptions, _
        reportType:=reportType, debugOnly:=debugOnly, suppliedFieldInfoByEditedName:=mappedFields)
    #End If
    
    Exit Sub

Progress_Checkbox_Missing:

    'Set Progress_Control = Nothing
    Resume Block_Query_Main_Function
    
Catch_MissingCollection:

    Contract_CLCTN.Add New Collection, Block(contractCodeColumn)
    Resume
Catch_GeneralError:
    Call PropagateError(Err, "Block_Query", "Failed to upload data to database or worksheet.")
End Sub
Private Function Missing_Data(ByVal maxDateCFTC As Date, ByVal CFTC_Last_Updated_Day As Date, ByVal ICE_Last_Updated_Day As Date, ByVal maxDateICE As Date, reportType$, getFuturesAndOptions As Boolean, Download_ICE_Data As Boolean, Download_CFTC_Data As Boolean, Optional DebugMD As Boolean = False) As Variant()
'===================================================================================================================
    'Purpose: Determines which files need to be downloaded for when multiple weeks of data have been missed.
    'Inputs:
    'Outputs:
'===================================================================================================================
    
    Dim File_CLCTN As New Collection, MacB As Boolean, New_Data As New Collection, iceUrl As Variant
    
    On Error GoTo Propagate
    
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
            ICE_End_Date:=maxDateICE
        
        'ICE_Query
        With New_Data
            For Each iceUrl In File_CLCTN
                .Add ICE_Query(CStr(iceUrl), CDate(ICE_Last_Updated_Day))(IIf(getFuturesAndOptions, "futures+options", "futures-only"))
            Next
        End With
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
                CFTC_End_Date:=maxDateCFTC, _
                reportType:=reportType, _
                downloadFuturesAndOptions:=getFuturesAndOptions
            
            New_Data.Add Historical_Parse(File_CLCTN, retrieveCombinedData:=getFuturesAndOptions, reportType:=reportType, parsingMultipleWeeks:=True, After_This_Date:=CFTC_Last_Updated_Day, Kill_Previous_Workbook:=DebugMD)
            
        End If
        
    #End If
    
    Application.DisplayAlerts = True
    
    If New_Data.count = 1 Then
        Missing_Data = New_Data(1)
    ElseIf New_Data.count > 1 Then
        Missing_Data = CombineArraysInCollection(New_Data, Append_Type.Multiple_2d)
    End If
    
    Exit Function
Propagate:
    Call PropagateError(Err, "Missing_Data")
End Function
Public Function Weekly_ICE(Most_Recent_CFTC_Date As Date) As Collection

'Dim Path_CLCTN As New Collection

    Dim ICE_URL$
    
    ICE_URL = Get_ICE_URL(Most_Recent_CFTC_Date)
    
    On Error GoTo Exit_Sub
    Set Weekly_ICE = ICE_Query(ICE_URL, Most_Recent_CFTC_Date)
    
    Exit Function
    
Exit_Sub:
    PropagateError Err, "Weekly_ICE"
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

Private Function ICE_Query(Weekly_ICE_URL$, greaterThanDate As Date) As Collection

    Dim Data_Query As QueryTable, data As Variant, Data_Row() As Variant, url$, _
    Y As Byte, bb As Boolean, getFuturesAndOptions As Boolean, _
    Found_Data_Query As Boolean, Error_While_Refreshing As Boolean, Filtered_CLCTN As Collection
    
    Const connectionName$ = "ICE Data Refresh Connection", queryName = "ICE Data Refresh"
    
    With Application
        bb = .EnableEvents
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    For Each Data_Query In QueryT.QueryTables
    
        If Data_Query.Name Like "*" & queryName & "*" Then
            Found_Data_Query = True
            Exit For
        End If
        
    Next Data_Query
    
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
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileCommaDelimiter = True
            .Name = queryName
            
Try_NameConnection:
            .WorkbookConnection.RefreshWithRefreshAll = False
        End With
        
        On Error GoTo 0
    
    Else
        ' Update Connection string
        On Error GoTo Catch_FailedConnectionUpdate
        With Data_Query
            .Connection = "TEXT;" & Weekly_ICE_URL
        End With
        
        On Error GoTo 0
        
    End If
    
    With Data_Query
        
        On Error GoTo Failed_To_Refresh 'Recreate Query and try again exactly 1 more time
        .TextFileColumnDataTypes = Filter_Market_Columns(convert_skip_col_to_general:=True, reportType:="D", Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True, ICE:=True)
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
            .parent.ShowAllData
            .ClearContents
            Err.Clear
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
    PropagateError Err, "ICE_Query"
    
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
        PropagateError Err, "ICE_Query"
    Else
        Error_While_Refreshing = True
        Resume Recreate_Query
    End If
    
Aggregation_Failed:
    PropagateError Err, "ICE_Query"
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
        wantedContractCode$, invalidContractCode As Boolean, New_Data() As Variant, First_Calculated_Column As Byte
        
        Dim WS As Worksheet, Symbol_Row As Long, iColumn As Byte, reportType$
    
        Set Current_Contracts = GetAvailableContractInfo
        
        reportType = ReturnReportType
        
        retrieveCombinedData = IsWorkbookForFuturesAndOptions()
        
        First_Calculated_Column = 3 + WorksheetFunction.CountIf(GetAvailableFieldsTable(reportType).DataBodyRange.columns(2), True)
        
        Do
            wantedContractCode = InputBox("Enter a 6 digit CFTC contract code")
            
            If wantedContractCode = vbNullString Then
                Exit Sub
            ElseIf HasKey(Current_Contracts, wantedContractCode) Then
                MsgBox "Contract Code is already present within the workbook"
                invalidContractCode = True
            Else
                invalidContractCode = False
            End If
            
        Loop While Len(wantedContractCode) <> 6 Or invalidContractCode
                
        On Error GoTo No_Data_Retrieved_From_API
        
        Dim fieldInfoByName As Collection
        
        If TryGetCftcWithSocrataAPI(New_Data, reportType, retrieveCombinedData, False, fieldInfoByName, wantedContractCode) Then
        
            New_Data = Filter_Market_Columns(False, True, False, reportType, True, New_Data, False)
            
            On Error GoTo 0
            
            wantedContractCode = New_Data(1, UBound(New_Data, 2))
            
            ReDim Preserve New_Data(LBound(New_Data, 1) To UBound(New_Data, 1), LBound(New_Data, 2) To UBound(New_Data, 2) + 1)
            
            On Error GoTo Catch_SymbolMissing
            
            With Range("Symbols_TBL")
            
                Symbol_Row = WorksheetFunction.Match(wantedContractCode, .columns(1), 0)
                
                If Symbol_Row <> 0 Then
                    
                    For iColumn = 3 To 4
                        If Not IsEmpty(.Cells(Symbol_Row, iColumn)) Then
                            Call TryGetPriceData(New_Data, UBound(New_Data, 2), Array(.Cells(Symbol_Row, iColumn), IIf(iColumn = 3, True, False)), datesAreInColumnOne:=True, overwriteAllPrices:=True)
                            Exit For
                        End If
                    Next iColumn
            
                End If
            
            End With
            
Try_DoCalculations:
            On Error GoTo 0
            ReDim Preserve New_Data(LBound(New_Data, 1) To UBound(New_Data, 1), LBound(New_Data, 2) To Range("Last_Calculated_Column").Value2)
            
            Select Case reportType
                Case "L"
                    New_Data = Legacy_Multi_Calculations(New_Data, UBound(New_Data, 1), First_Calculated_Column, 156, 26)
                Case "D"
                    New_Data = Disaggregated_Multi_Calculations(New_Data, UBound(New_Data, 1), First_Calculated_Column, 156, 26)
                Case "T"
                    New_Data = TFF_Multi_Calculations(New_Data, UBound(New_Data, 1), First_Calculated_Column, 156, 26, 52)
            End Select
            
            Application.ScreenUpdating = False
            
            Set WS = ThisWorkbook.Worksheets.Add
            
            With WS
            
                .columns(1).NumberFormat = "yyyy-mm-dd"
                .columns(First_Calculated_Column - 3).NumberFormat = "@"
                
                Call Paste_To_Range(Sheet_Data:=New_Data, Historical_Paste:=True, Target_Sheet:=WS)
                
                .ListObjects(1).Name = "CFTC_" & wantedContractCode
                
            End With
            
            If Symbol_Row = 0 Then
                Range("Symbols_TBL").ListObject.ListRows.Add.Range.Value2 = Array(wantedContractCode, New_Data(UBound(New_Data, 1), 2), Empty, Empty)
                MsgBox "A new row has been added to the availbale symbols table. Please fill in the missing Symbol information if available."
            End If
            
        End If
Finally:
        Re_Enable
    
        Exit Sub
No_Data_Retrieved_From_API:
        MsgBox "Data couldn't be retrieved from API"
        Resume Finally
Catch_SymbolMissing:
        Resume Try_DoCalculations
        
    End Sub
#Else
    Function ConvertSymbolDataToJson$()
    
        Dim stuff As New Collection, CD As ContractInfo, Item As Variant, quote$
        
        quote = "\" & Chr(34)
        
        With stuff
            For Each Item In GetAvailableContractInfo
                Set CD = Item
                .Add (quote & CD.contractCode & quote & ":" & quote & CD.priceSymbol & quote)
            Next
        End With
        
       ConvertSymbolDataToJson = "{" & Join(ConvertCollectionToArray(stuff), ",") & "}"
        
    End Function
    Function ListDatabasePathsInJson$()
    
        Dim Item As Variant, reportDetails As LoadedData, stuff As New Collection, quote$
        
        quote = "\" & Chr(34)
        
        For Each Item In Array("L", "D", "T")
            Set reportDetails = GetStoredReportDetails(CStr(Item))
            stuff.Add quote & IIf(Item = "T", "TFF", reportDetails.FullReportName) & quote & ":" & quote & Replace(reportDetails.CurrentDatabasePath, "\", "\\") & quote
        Next Item
        ListDatabasePathsInJson = "{" & Join(ConvertCollectionToArray(stuff), ",") & "}"
        
    End Function
    Function RunCSharpExtractor() As Collection
    
        Dim commandArgs$(2), cmd$, result$
                
        commandArgs(0) = Range("CSharp_Exe").Value2
        
        If Not FileOrFolderExists(commandArgs(0)) Then Err.Raise 53, "RunCSharpExtractor", "C# executable couldn't be found."
        
        commandArgs(1) = ListDatabasePathsInJson()
        commandArgs(2) = ConvertSymbolDataToJson()
        
        cmd = Join(QuotedForm(commandArgs), " ")
        
        Application.StatusBar = "Querying new data with " & commandArgs(0)
        
        With CreateObject("WScript.Shell").exec(cmd)
            result = .StdOut.ReadAll
            .Terminate
        End With
        
        Dim innerCollection As Collection, output As New Collection, i As Long, programResponse$(), report$, reportInfo$(), dataStart As Byte, Item As Variant, kvp$()
        
        programResponse = Split(result, vbNewLine)
        
        For i = UBound(programResponse) - 1 To UBound(programResponse) - 6 Step -1
            
            On Error Resume Next
            
            report = Left$(programResponse(i), 1)
            ' Ensure there is a collection for each Report Type.
            output.Add New Collection, report
            
            dataStart = InStr(1, programResponse(i), "{") + 1
            reportInfo = Split(Mid$(programResponse(i), dataStart, Len(programResponse(i)) - dataStart), ",")
            
            Set innerCollection = New Collection
            On Error GoTo 0
            'Add an inner collection keyed to whether or not the data is combined.
            output(report).Add innerCollection, IIf(InStrB(1, LCase$(programResponse(i)), "true") > 0, CStr(True), CStr(False))
            
            Dim elementName$
            
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
        Next i
        
        Set RunCSharpExtractor = output
        Debug.Print result
        Application.StatusBar = vbNullString
        
        Exit Function
        
DefaultStringAddition:
        innerCollection.Add Trim$(kvp(1)), Trim$(kvp(0))
        Resume Next
    End Function

#End If


