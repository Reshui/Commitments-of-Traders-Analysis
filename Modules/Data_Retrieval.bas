Attribute VB_Name = "Data_Retrieval"

Public Data_Updated_Successfully As Boolean

Public Enum RetrievalErr
    RetrievalFailed = vbObjectError + 513
    SocrataSuccessNoNewData = vbObjectError + 514
    FinalProcessorNoContractsInInput = vbObjectError + 515
End Enum

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

Public Enum OpenInterestEnum
    '-1 =True; 0 =False; 1 NA
    FuturesAndOptions = -1
    FuturesOnly = 0
    OptionsOnly = 1
End Enum

Public Enum ReportEnum
    eLegacy = 1
    eDisaggregated = 2
    eTFF = 3
End Enum

Option Explicit
Sub New_Data_Query(Optional Scheduled_Retrieval As Boolean = False, Optional Overwrite_All_Data As Boolean = False, Optional IsWorbookOpenEvent As Boolean = False, Optional workbookEventProfiler As TimedTask)
'===================================================================================================================
    'Summary: Retrieves CFTC data that hasn't been stored either on a worksheet or database.
    'Inputs: Scheduled_Retrieval - If true then error messages will not be displayed.
    '        Overwrite_All_Data - Reserved for non-Database version. If true then all data on a worksheet will be replaced.
'===================================================================================================================
    Dim Last_Update_CFTC As Date, CFTC_Incoming_Date As Date, ICE_Incoming_Date As Date, Last_Update_ICE As Date

    Dim cftcDataA() As Variant, iceDataA() As Variant, retrievedData() As Variant, reportsToQuery$()

    Dim reportInitial$, iReport As Long, iOiType As Long

    Dim cftcDateRange As Range
    'Booleans
    Dim debugWeeklyRetrieval As Boolean, DBM_Historical_Retrieval As Boolean, Debug_Mode As Boolean, _
    Download_CFTC As Boolean, Download_ICE As Boolean, Check_ICE As Boolean, queryFuturesAndOptions As Boolean, _
    uploadDataToDatabaseOrWorksheet As Boolean, isLegacyCombined As Boolean, _
    exitSubroutine As Boolean, CFTC_Retrieval_Error As Boolean, newDataSuccessfullyHandled As Boolean, processingReport As Boolean
    'Collections
    Dim cftcMappedFieldInfo As Collection, DataBase_Not_Found_CLCTN As Collection, _
    availableContractInfo As Collection, Data_CLCTN As Collection, Weekly_ICE_CLCTN As Collection
    
    Dim taskProfiler As TimedTask, individualCotTask As TimedTask

    Const dataRetrieval$ = "C.O.T data retrieval", uploadTime$ = "Upload Time", _
    databaseDateQuery$ = "Query database for latest date.", cRetrievalTask$ = "C# COT Executable", ProcedureName$ = "New_Data_Query"

    Dim useSocrataAPI As Boolean, socrataApiFailed As Boolean, executableSuccess As Boolean, cftcAlreadyUpdated As Boolean, iceAlreadyUpdated As Boolean
    
    Const RawCftcDateColumn& = 3
    
    Call IncreasePerformance

    On Error GoTo Deny_Debug_Mode
    Debug_Mode = Weekly.Shapes("Test_Toggle").OLEFormat.Object.value = xlOn And Not Scheduled_Retrieval
    On Error GoTo Catch_General_Error

    #If DatabaseFile Then
        
        Dim openInterestTypesToQuery(1) As OpenInterestEnum, executeableReturn As Object, _
        recentDateQueryCompleted As Boolean
        Const latestDateKey$ = "Latest Date", statusKey$ = "Status"
        
        #If Mac Then
            GateMacAccessToWorkbook
        #End If

        'Legacy data must be retrieved first so that price data only needs to be retrieved once.
        reportsToQuery = Split("L,D,T", ",")

        openInterestTypesToQuery(0) = OpenInterestEnum.FuturesAndOptions
        openInterestTypesToQuery(1) = OpenInterestEnum.FuturesOnly

        Set cftcDateRange = Variable_Sheet.Range("Last_Updated_CFTC")
        
    #Else
        Dim openInterestTypesToQuery(0) As OpenInterestEnum, iceDateRange As Range
        
        ReDim reportsToQuery(0)
        
        reportsToQuery(0) = ReturnReportType
        openInterestTypesToQuery(0) = IsWorkbookForFuturesAndOptions
        
        With Variable_Sheet
            Set cftcDateRange = .Range("Last_Updated_CFTC")
            If reportsToQuery(0) = "D" Then Set iceDateRange = .Range("Last_Updated_ICE")
        End With
        
    #End If
    
    If workbookEventProfiler Is Nothing Then
        Set taskProfiler = New TimedTask
        taskProfiler.Start "New Data Query [" & Time & "]"
    Else
        Set taskProfiler = workbookEventProfiler.StartSubTask("New Data Query [" & Time & "]")
    End If
    
    If Debug_Mode = True Then
        
        With taskProfiler
            .Pause
        
            If MsgBox("Test Weekly Data Retrieval ?", vbYesNo, "Choose what to debug") = vbYes Then
                debugWeeklyRetrieval = True
            ElseIf MsgBox("Test Multi-Week Historical Retrieval ?", vbYesNo, "Choose what to debug") = vbYes Then
                DBM_Historical_Retrieval = True
            End If
    
            #If DatabaseFile Then
                Do
                    reportInitial = UCase(InputBox("Select 1 of L,D,T"))
                Loop While IsError(Application.Match(reportInitial, reportsToQuery, 0))
                
                reportsToQuery(0) = reportInitial
                            
                On Error Resume Next
                ' Get the combined status to test.
                Do
                    queryFuturesAndOptions = CBool(InputBox("Select 1 of the following:" & String$(2, vbNewLine) & "  Futures Only: 0" & vbNewLine & "  Futures + Options: 1"))
                Loop While Err.Number <> 0
                
                openInterestTypesToQuery(0) = queryFuturesAndOptions
            #End If
            
            On Error GoTo Catch_General_Error
            .Continue
        End With
    Else
        #If DatabaseFile Then
            If DoesUserPermit_SqlServer() And Evaluate("=UseExternalExecutable") Then
                With taskProfiler.StartSubTask(cRetrievalTask)
                    On Error GoTo Catch_ExecutableFailed
                    Set executeableReturn = RunCSharpExtractor()
                    On Error GoTo Catch_General_Error
                    .EndTask
                End With
            End If
        #Else
            If IsCreatorActiveUser() Then Exit Sub
        #End If
    End If

Retrieve_Latest_Data:
    
    Dim eReport As ReportEnum
    
    For iReport = LBound(reportsToQuery) To UBound(reportsToQuery)
        
        reportInitial = reportsToQuery(iReport)
        eReport = ConvertInitialToReportTypeEnum(reportInitial)
        ' This is to ensure that the collection is recreated for each enum.
        Set cftcMappedFieldInfo = Nothing
        
        For iOiType = LBound(openInterestTypesToQuery) To UBound(openInterestTypesToQuery)
            
            processingReport = True
            On Error GoTo Catch_General_Error
            
            queryFuturesAndOptions = openInterestTypesToQuery(iOiType)
            isLegacyCombined = (queryFuturesAndOptions And eReport = eLegacy)
            
            cftcAlreadyUpdated = False: iceAlreadyUpdated = False: uploadDataToDatabaseOrWorksheet = False: Check_ICE = False
            
            #If DatabaseFile Then
                If Not Debug_Mode And iOiType = LBound(openInterestTypesToQuery) And iReport = LBound(reportsToQuery) And Not isLegacyCombined Then
                    MsgBox "Legacy Combined data needs to be retrieved first so that price data only has to be retrieved once."
                    GoTo Exit_Procedure
                End If
                
                On Error GoTo Try_Get_New_Data
                ' If data was updated by C# executable then goto next loop.
                If Not executeableReturn Is Nothing Then
                    With executeableReturn.item(CStr(queryFuturesAndOptions)).item(reportInitial)
                        Select Case .item(statusKey)
                            Case ReportStatusCode.NoUpdateAvailable, ReportStatusCode.Updated
                                executableSuccess = True
                                If CFTC_Incoming_Date < .item(latestDateKey) Then CFTC_Incoming_Date = .item(latestDateKey)
                                
                                If .item(latestDateKey) > cftcDateRange.Value2 Or .item(statusKey) = ReportStatusCode.Updated Then
                                    With GetStoredReportDetails(eReport)
                                        Select Case .OpenInterestType.Value2
                                            Case queryFuturesAndOptions, OpenInterestEnum.OptionsOnly
                                                .PendingUpdateInDatabase.Value2 = True
                                        End Select
                                    End With
                                    newDataSuccessfullyHandled = True
                                End If
                                
                                If Not Debug_Mode Then GoTo Next_Combined_Value
                            Case Else
                                executableSuccess = False
                        End Select
                    End With
                End If
            #End If
Try_Get_New_Data:
            On Error GoTo Catch_General_Error
            useSocrataAPI = Not Overwrite_All_Data
            
            Set individualCotTask = taskProfiler.StartSubTask("[" & reportInitial & "] Combined: (" & queryFuturesAndOptions & ")")
            
            #If DatabaseFile Then
                With individualCotTask.StartSubTask(databaseDateQuery)
                    
                    recentDateQueryCompleted = TryGetLatestDate(Last_Update_CFTC, eReport:=eReport, versionToQuery:=openInterestTypesToQuery(iOiType), queryIceContracts:=False)
                    .EndTask
                    
                    If Not recentDateQueryCompleted Then
                    
                        If DataBase_Not_Found_CLCTN Is Nothing Then Set DataBase_Not_Found_CLCTN = New Collection
                        
                        DataBase_Not_Found_CLCTN.Add "Missing database for " & GetStoredReportDetails(eReport).FullReportName.Value2
                        'The Legacy_Combined data is the only one for which price data is queried.
                        If isLegacyCombined Then exitSubroutine = True
                        Exit For
                        
                     End If
                     
                End With
            #Else
                Last_Update_CFTC = cftcDateRange.Value2
            #End If
            
            individualCotTask.StartSubTask dataRetrieval
            
            If CFTC_Incoming_Date = TimeSerial(0, 0, 0) Or Last_Update_CFTC < CFTC_Incoming_Date Or Debug_Mode Then
                         
                On Error GoTo Catch_CFTCRetrievalFailed
                cftcDataA = HTTP_Weekly_Data(Last_Update_CFTC, suppressMessages:=Scheduled_Retrieval, _
                                retrieveCombinedData:=queryFuturesAndOptions, _
                                reportType:=eReport, useApi:=useSocrataAPI, _
                                columnMap:=cftcMappedFieldInfo, _
                                testAllMethods:=debugWeeklyRetrieval, _
                                DebugActive:=Debug_Mode)
                
                socrataApiFailed = Not useSocrataAPI
                CFTC_Incoming_Date = cftcDataA(UBound(cftcDataA, 1), RawCftcDateColumn)
                cftcAlreadyUpdated = (CFTC_Incoming_Date = Last_Update_CFTC)
            Else
                cftcAlreadyUpdated = True
            End If
            
Try_IceRetrieval:
            ' If Disaggregated report then get the most recent ICE data if cftc update available.
            If eReport = eDisaggregated Then
                
                On Error GoTo ICE_Retrieval_Failed
                                               
                #If DatabaseFile Then
                    If Not TryGetLatestDate(Last_Update_ICE, eReport:=eReport, versionToQuery:=openInterestTypesToQuery(iOiType), queryIceContracts:=True) Then
                        GoTo Finished_Querying_Weekly_Data
                    End If
                #Else
                    Last_Update_ICE = iceDateRange.Value2
                #End If
                
                If Last_Update_ICE < CFTC_Incoming_Date Or Debug_Mode Then
                    Check_ICE = True
                    
                    If Weekly_ICE_CLCTN Is Nothing Then
                        Set Weekly_ICE_CLCTN = Weekly_ICE(CFTC_Incoming_Date)
                    End If
                    
                    Dim iceKey$
                    
                    iceKey = ConvertOpenInterestTypeToName(openInterestTypesToQuery(iOiType))
                    
                    With Weekly_ICE_CLCTN
                        iceDataA = .item(iceKey)
                        .Remove iceKey
                        ' if empty or not using DatabaseFile
                        If .Count = 0 Or UBound(openInterestTypesToQuery) = 0 Then Set Weekly_ICE_CLCTN = Nothing
                    End With
                    
                    ICE_Incoming_Date = iceDataA(1, RawCftcDateColumn)
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
                    End If
                    .Continue
                End With
            End If
            
            If Debug_Mode Or Not cftcAlreadyUpdated Or (Check_ICE And Not iceAlreadyUpdated) Then
                    'If isLegacyCombined Then exitSubroutine = True
                If DBM_Historical_Retrieval Or ((socrataApiFailed And CFTC_Incoming_Date - Last_Update_CFTC > 7) Or (Check_ICE And ICE_Incoming_Date - Last_Update_ICE > 7)) Then
                    
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
    
                    retrievedData = Missing_Data(getFuturesAndOptions:=queryFuturesAndOptions, _
                        maxDateICE:=ICE_Incoming_Date, maxDateCFTC:=CFTC_Incoming_Date, _
                        Download_ICE_Data:=Download_ICE, Download_CFTC_Data:=Download_CFTC, _
                        eReport:=eReport, _
                        localCftcDate:=Last_Update_CFTC, localIceDate:=Last_Update_ICE + 2, _
                        DebugMD:=DBM_Historical_Retrieval)
                                                     
                    If Not (Download_ICE And Download_CFTC) Then
                        If Download_CFTC And (Check_ICE And ICE_Incoming_Date - Last_Update_ICE > 0) Then
                            'Determine if the most recently queried Ice Data needs to be added
                            Set Data_CLCTN = New Collection
                            With Data_CLCTN
                                .Add retrievedData
                                .Add iceDataA
                            End With
                            retrievedData = CombineArraysInCollection(Data_CLCTN, Append_Type.Multiple_2d)
                        ElseIf Download_ICE And CFTC_Incoming_Date - Last_Update_CFTC > 0 Then
                            'Determine if CFTC data needs to be added
                            Set Data_CLCTN = New Collection
                            With Data_CLCTN
                                .Add retrievedData
                                .Add cftcDataA
                            End With
                            retrievedData = CombineArraysInCollection(Data_CLCTN, Append_Type.Multiple_2d)
                        End If
                    End If
                    uploadDataToDatabaseOrWorksheet = True
                    
                ElseIf Not cftcAlreadyUpdated Or (Check_ICE And Not iceAlreadyUpdated) Or Debug_Mode = True Then
                    
                    Set Data_CLCTN = New Collection
                    With Data_CLCTN
                        If Check_ICE And (ICE_Incoming_Date - Last_Update_ICE > 0 Or Debug_Mode) And IsArrayAllocated(iceDataA) Then .Add iceDataA
                        If (CFTC_Incoming_Date - Last_Update_CFTC > 0 Or Debug_Mode) And IsArrayAllocated(cftcDataA) Then .Add cftcDataA
                        
                        If .Count > 0 Then
                            Select Case .Count
                                Case 1:
                                    retrievedData = .item(1)
                                Case 2:
                                    If UBound(.item(1), 2) <> UBound(.item(2), 2) Then
                                        Err.Raise vbObjectError + 599, Description:="ICE data and CFTC data don't have an equal number of fields."
                                    End If
                                    retrievedData = CombineArraysInCollection(Data_CLCTN, Append_Type.Multiple_2d)
                            End Select
                            uploadDataToDatabaseOrWorksheet = True
                        End If
                    End With
                    
                End If
            End If
            
            Set Data_CLCTN = Nothing
            Erase cftcDataA
            If eReport = eDisaggregated Then Erase iceDataA
            
Stop_Timers_And_Update_If_Allowed:
            With individualCotTask
                .StopSubTask dataRetrieval
                If IsArrayAllocated(retrievedData) And uploadDataToDatabaseOrWorksheet = True And Not exitSubroutine Then
                    ' Upload data to database/spreadsheet and retrieve prices.
                    With .StartSubTask(uploadTime)
                        If availableContractInfo Is Nothing Then
                            On Error GoTo Catch_General_Error
                            Set availableContractInfo = GetAvailableContractInfo(True)
                        End If
                        
                        If cftcMappedFieldInfo Is Nothing Then Set cftcMappedFieldInfo = GetExpectedLocalFieldInfo(eReport, False, False, False, False)
                        
                        On Error GoTo Catch_FinaProcessor_Failed
                        Call CommitmentsOfTradersFinalProcessor(cotData:=retrievedData, reportType:=eReport, _
                                        oiType:=openInterestTypesToQuery(iOiType), availableContractInfo:=availableContractInfo, _
                                        debugOnly:=Debug_Mode, mappedFields:=cftcMappedFieldInfo, Overwrite_Worksheet:=Overwrite_All_Data)
                        
                        On Error GoTo Catch_General_Error
                        Erase retrievedData
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
                If .IsRunning Then .EndTask
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
                    
                    With taskProfiler.StartSubTask("Query all databases for latest contracts.")
                         Latest_Contracts
                        .EndTask
                    End With
                    
                    On Error Resume Next
                    With taskProfiler.StartSubTask("Refresh data tables.")
                        Call RefreshAllDatabaseTables(False)
                        .EndTask
                    End With
                    On Error GoTo Catch_General_Error
                    
                #End If
            End If
        End If
        
    ElseIf Not DataBase_Not_Found_CLCTN Is Nothing And Not Scheduled_Retrieval Then
        
        With DataBase_Not_Found_CLCTN
            For iReport = 1 To .Count
                MsgBox .item(iReport)
            Next iReport
        End With
        
    ElseIf Not (Scheduled_Retrieval Or Debug_Mode) Then
        On Error GoTo NoDataMessage
        
        MsgBox "Data is already updated." & _
        vbNewLine & _
        vbNewLine & _
            "The next release is scheduled for " & vbNewLine & vbTab & Format$(CFTC_Release_Dates(False, False), "dddd, mmmm d, yyyy") & " 3:30 PM Eastern Time." & _
        vbNewLine & _
        vbNewLine & _
            "Enabling Test Mode will allow you to continue, but only new/missing rows will be added to the database. " & _
        vbNewLine & _
        vbNewLine & _
            "Otherwise, try again after new data has been released. Check the release schedule for more information.", , Title:="Data already updated."
        
        On Error GoTo 0
        Application.StatusBar = vbNullString
    End If
    
Finally:
    With taskProfiler
        .EndTask
        If workbookEventProfiler Is Nothing Then .DPrint
    End With
    
    Re_Enable
    Exit Sub
NoDataMessage:
    MsgBox "No new data available."
    Resume Next
    
Deny_Debug_Mode:
    Debug_Mode = False
    Resume Next

ICE_Retrieval_Failed:
    Check_ICE = False
    Resume Finished_Querying_Weekly_Data

Catch_CFTCRetrievalFailed:

    Select Case Err.Number
        Case RetrievalErr.SocrataSuccessNoNewData
            'Retrieval didn't fail. Just no new data.
            CFTC_Incoming_Date = Last_Update_CFTC
            socrataApiFailed = False
            cftcAlreadyUpdated = True
            Resume Try_IceRetrieval
            
        Case RetrievalErr.RetrievalFailed
            MsgBox "Data retrieval methods have failed." & String$(2, vbNewLine) & _
                   "Check your internet connection. If this error persists please contact me at MoshiM_UC@outlook.com with your operating system and Excel version."
           
        Case Else
            DisplayErr Err, ProcedureName
    End Select
    
    CFTC_Retrieval_Error = True
    exitSubroutine = True
    Resume Stop_Timers_And_Update_If_Allowed
    
#If DatabaseFile Then
Catch_ExecutableFailed:
    taskProfiler.StopSubTask cRetrievalTask
    If IsCreatorActiveUser Then DisplayErr Err, ProcedureName, "Failed to parse values from .exe response."
    Set executeableReturn = Nothing
    Resume Next
#End If

Catch_General_Error:
    DisplayErr Err, IIf(processingReport, reportInitial & "_" & ConvertOpenInterestTypeToName(CLng(queryFuturesAndOptions)), vbNullString)
    Resume Finally
    
Catch_FinaProcessor_Failed:
    Erase retrievedData
    DisplayErr Err, ProcedureName, IIf(processingReport, reportInitial & "_" & ConvertOpenInterestTypeToName(CLng(queryFuturesAndOptions)), vbNullString)
    Resume Next_Combined_Value
    
Catch_Database_Not_Found:
    If isLegacyCombined Then exitSubroutine = True
    Resume Next_Report_Release_Type
    
End Sub
Private Sub CommitmentsOfTradersFinalProcessor(ByRef cotData() As Variant, reportType As ReportEnum, oiType As OpenInterestEnum, _
                        availableContractInfo As Collection, debugOnly As Boolean, _
                        mappedFields As Collection, Optional Overwrite_Worksheet As Boolean = False)
'===================================================================================================================
    'Summary: Data within [cotData] will be pruned for wanted columns,have price data added if needed and uploaded to either a database or worksheet.
    'Inputs: [cotData] - Array that holds data to store.
    '        reportType - Report that is being uploaded.
    '        oiType - True if data is futures + options; else, futures only.
    '        availableContractInfo - Collection of Contract instances. If on non-database file then this
    '                      contains only contracts within the file. IF database file then all contracts with a price symbol available.
    '        Overwrite_Worksheet - If True then all data on a worksheet will be replaced with the data matching its contract code.
    '        mappedFields - Collection of FieldInfo instances that represent each column in [cotData].
'===================================================================================================================
    Dim iRow As Long, iCount As Long, isDatabaseVersion As Boolean, uniqueContractCount As Long
    
    Dim Block() As Variant, Contract_CLCTN As Collection, priceColumn As Long, contractCode$
    
    Dim retrievePriceData As Boolean, contractCodeColumn As Long

    Dim progressBarActive As Boolean, enableProgressBar As Boolean
    
    Const BLOCK_QUERY_ERR As Long = vbObjectError + 571, ProcedureName$ = "CommitmentsOfTradersFinalProcessor"
    
    On Error Resume Next
    enableProgressBar = Weekly.Shapes("Progress_CHKBX").OLEFormat.Object.value = xlOn
    
    On Error GoTo Propagate
    
    If mappedFields Is Nothing Then
        Err.Raise BLOCK_QUERY_ERR, Description:="Field map [mappedFields] hasn't been declared."
    ElseIf mappedFields.Count = 0 Then
        Err.Raise BLOCK_QUERY_ERR, Description:="Field map [mappedFields] is empty."
    End If
        
    #If Not DatabaseFile Then
    
        Dim columnfilter() As Variant, missingWeeksCount As Long, Last_Calculated_Column As Long, _
        current_Filters() As Variant, firstCalculatedColumn As Long, WS_Data() As Variant, _
        outputTable As ListObject, Table_Data_CLCTN As New Collection, placeDataOnWorksheetDebug As Boolean
        
        Const Time1 As Long = 156, Time2 As Long = 26, Time3 As Long = 52

        Last_Calculated_Column = Variable_Sheet.Range("Last_Calculated_Column").Value2
        
        columnfilter = Filter_Market_Columns(True, False, False, reportType, Create_Filter:=True)
         
        ' +1 To account for filter returning a 1 based array
        ' +1 To get wanted value
        ' = +2
        priceColumn = UBound(Filter(columnfilter, xlSkipColumn, False)) + 2
        firstCalculatedColumn = priceColumn + 2
        placeDataOnWorksheetDebug = True
    #Else
        On Error GoTo Catch_NotEnoughDimensions
        priceColumn = UBound(cotData, 2) + 1
        
        On Error GoTo Catch_ReDimension_Failure
        ReDim Preserve cotData(LBound(cotData, 1) To UBound(cotData, 1), LBound(cotData, 2) To priceColumn)
        
        isDatabaseVersion = True
        On Error GoTo Propagate
    #End If
    
    If Not HasKey(mappedFields, "price") Then
        mappedFields.Add CreateFieldInfoInstance("price", CInt(priceColumn), "Price", False, True, False), "price"
    End If
    
    If (isDatabaseVersion And oiType = OpenInterestEnum.FuturesAndOptions And reportType = eLegacy) Or Not isDatabaseVersion Then
        
        On Error GoTo Catch_Missing_Contract_Code_Field
        contractCodeColumn = mappedFields("cftc_contract_market_code").ColumnIndex
        
        On Error GoTo Propagate
        ReDim Block(LBound(cotData, 2) To UBound(cotData, 2))
        Set Contract_CLCTN = New Collection
        ' Parse array rows into collections keyed to their contract code.
        ' Array should be date sorted already.
        On Error GoTo Catch_MissingCollection
        
        For iRow = LBound(cotData, 1) To UBound(cotData, 1)
            For iCount = LBound(cotData, 2) To UBound(cotData, 2)
                Block(iCount) = cotData(iRow, iCount)
            Next iCount
            Contract_CLCTN(Block(contractCodeColumn)).Add Block
        Next iRow
        
        On Error GoTo Propagate
        
        uniqueContractCount = Contract_CLCTN.Count
        
        If uniqueContractCount = 0 Then
            Err.Raise RetrievalErr.FinalProcessorNoContractsInInput, Description:="No unique contracts available for " & ConvertReportTypeEnum(reportType) & "-[Futures_Options: " & CBool(oiType) & "]"
        End If
        
        retrievePriceData = True
        
        If debugOnly Then
            retrievePriceData = MsgBox("Debug mode is active. Do you want to test price retrieval?", vbYesNo, "Test price retrieval?") = vbYes
            
            If Not retrievePriceData Then
                #If DatabaseFile Then
                    GoTo Upload_Data
                #End If
            End If
            
            #If Not DatabaseFile Then
                placeDataOnWorksheetDebug = MsgBox("Test pasting to worksheet?", vbYesNo, "Test data paste?") = vbYes
            #End If
        End If
        
        If retrievePriceData Then Erase cotData
        
        If enableProgressBar And Not isDatabaseVersion Then
            ' Display Progress Bar control.
            ' Arguements are passed Byref and given values in the below Sub.
            With Progress_Bar
                .Show
                .InitializeValues CLng(uniqueContractCount)
            End With
            progressBarActive = True
        End If

        For iCount = uniqueContractCount To 1 Step -1
            
            ' Removes the collection. A combined version orf collection elements will be added later.
            Block = CombineArraysInCollection(Contract_CLCTN(iCount), Append_Type.Multiple_1d)
            contractCode = Block(1, contractCodeColumn)
            ' Remove Collection.
            Contract_CLCTN.Remove contractCode
            
            If HasKey(availableContractInfo, contractCode) Then
                #If Not DatabaseFile Then
                    Block = Filter_Market_Columns(False, True, False, reportType, False, Block, False, columnfilter)
                    'Expand for calculations
                    ReDim Preserve Block(LBound(Block, 1) To UBound(Block, 1), LBound(Block, 2) To Last_Calculated_Column)
                #End If
                
                If retrievePriceData And availableContractInfo(contractCode).HasSymbol Then
                    On Error Resume Next
                    Call TryGetPriceData(Block, priceColumn, availableContractInfo(contractCode), overwriteAllPrices:=False, datesAreInColumnOne:=Not isDatabaseVersion)
                    On Error GoTo Propagate
                End If
                
            ElseIf Not isDatabaseVersion Then
                GoTo NextAvailableContract
            End If
            
            #If Not DatabaseFile Then
            
                Set outputTable = availableContractInfo(contractCode).TableSource
                missingWeeksCount = UBound(Block, 1)
                
                If Not Overwrite_Worksheet Then
                    '--Append New Data to bottom of already existing table data
                    WS_Data = outputTable.DataBodyRange.Value2
                    
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
                    Case eLegacy
                        Block = Legacy_Multi_Calculations(Block, missingWeeksCount, firstCalculatedColumn, Time1, Time2)
                    Case eDisaggregated
                        Block = Disaggregated_Multi_Calculations(Block, missingWeeksCount, firstCalculatedColumn, Time1, Time2)
                    Case eTFF
                        Block = TFF_Multi_Calculations(Block, missingWeeksCount, firstCalculatedColumn, Time1, Time2, Time3)
                End Select
                
                If placeDataOnWorksheetDebug Then
                    Call ChangeFilters(outputTable, current_Filters)
                                    
                    If Not Overwrite_Worksheet Then
                        Call Paste_To_Range(Data_Input:=Block, Sheet_Data:=WS_Data, Table_DataB_RNG:=outputTable.DataBodyRange, Overwrite_Data:=Overwrite_Worksheet)
                    Else
                        Call Paste_To_Range(Data_Input:=Block, Table_DataB_RNG:=outputTable.DataBodyRange, Overwrite_Data:=Overwrite_Worksheet)
                    End If
                    
                    With outputTable.Sort
                        If .SortFields.Count > 0 Then .Apply
                    End With
                        
                    Call RestoreFilters(outputTable, current_Filters)
                End If
                
                Erase Block
                Erase WS_Data
                Set outputTable = Nothing
                
            #Else
                ' Adds array with price data retrieved to collection.
                Contract_CLCTN.Add Block, contractCode
            #End If
NextAvailableContract:
            If progressBarActive And Not isDatabaseVersion Then
                If iCount = 1 Then
                    Unload Progress_Bar
                Else
                    Progress_Bar.IncrementBar
                End If
            End If
        Next iCount
        
        #If DatabaseFile Then
            If Contract_CLCTN.Count > 0 Then
                cotData = CombineArraysInCollection(Contract_CLCTN, Append_Type.Multiple_2d)
            Else
                Err.Raise BLOCK_QUERY_ERR, Description:="Attempt to re-combine arrays after price retrieval failed. Contract collection is empty."
            End If
        #End If
        
    End If
    
Upload_Data:
    
    #If DatabaseFile And Not Mac Then
        On Error GoTo Propagate
        If IsArrayAllocated(cotData) And Not mappedFields Is Nothing Then
            Call Update_Database(dataToUpload:=cotData, versionToUpdate:=oiType, eReport:=reportType, debugOnly:=debugOnly, suppliedFieldInfoByEditedName:=mappedFields, enableProgressBar:=enableProgressBar)
        Else
            Err.Raise BLOCK_QUERY_ERR, Description:="Source array or field map is unallocated."
        End If
    #End If
    
    Exit Sub
Propagate:
    Call PropagateError(Err, ProcedureName)
Catch_ReDimension_Failure:
    Call PropagateError(Err, ProcedureName, "Failed to extend input data to fit price column.")
Catch_NotEnoughDimensions:
    Call PropagateError(Err, ProcedureName, "Source array doesn't have a 2nd dimension.")
Catch_Missing_Contract_Code_Field:
    
    Dim knownField As FieldInfo, names$(), fromSocrata As Boolean
    
    ReDim names(mappedFields.Count)
    iCount = LBound(names)
    
    For Each knownField In mappedFields
        With knownField
            names(iCount) = .EditedName
            If Not fromSocrata Then
                If .IsSocrataField Then fromSocrata = True
            End If
        End With
        iCount = iCount + 1
    Next knownField

    With Err
        .Description = "'cftc_contract_market_code' isn't a member of mappedFields." & vbNewLine & _
                       "From Socrata: " & fromSocrata & " | [mappedFields] count : " & mappedFields.Count & vbNewLine & .Description & vbNewLine & "List: " & Join(names, ", ")
    End With
    GoTo Propagate
Catch_MissingCollection:
    Select Case Err.Number
        Case 5 ' Key not found.
            Contract_CLCTN.Add New Collection, Block(contractCodeColumn)
            Resume
        Case 9 ' Subscript out of range.
            If IsArrayAllocated(Block) Then
                PropagateError Err, ProcedureName, "Variable contractCodeColumn(" & contractCodeColumn & ") represents an invalid index for the Block(" & LBound(Block) & " to " & UBound(Block) & ") array."
            Else
                PropagateError Err, ProcedureName, "Block array is unallocated."
            End If
        Case Else
            PropagateError Err, ProcedureName, "Error caught by Catch_MissingCollection handler."
    End Select
End Sub
Private Function Missing_Data(ByVal maxDateCFTC As Date, ByVal localCftcDate As Date, ByVal localIceDate As Date, _
                ByVal maxDateICE As Date, eReport As ReportEnum, getFuturesAndOptions As Boolean, _
                ByVal Download_ICE_Data As Boolean, ByVal Download_CFTC_Data As Boolean, _
                Optional DebugMD As Boolean = False) As Variant()
'===================================================================================================================
'   Summary: Determines which files need to be downloaded for when multiple weeks of data have been missed.
'   Inputs:
'       maxDateCFTC - Most recent date available from the CFTC
'       localCftcDate - Most recent CFTC date saved locally
'       localIceDate - Most recent ICE date saved locally
'       maxDateICE -
'       eReport - ReportEnum used to select data to download.
'       Download_CFTC_Data - True if you want to download CFTC data.
'       Download_ICE_Data - True if you want to download ICE data.
'       staticIceByUrl - Collection used to store yearly ICE data.
'===================================================================================================================
    
    Dim File_CLCTN As New Collection, MacB As Boolean, New_Data As New Collection, iceURL$, _
    i As Long, iceData() As Variant, oneYearData As Collection
    
    Static staticIceByUrl As Collection
    
    On Error GoTo Propagate
    
    #If Mac Then
        MacB = True
    #End If
    
    If DebugMD Then
        If Not MacB Then If MsgBox("Do you want to test MAC OS data retrieval?", vbYesNo) = vbYes Then MacB = True
                
        If eReport = eDisaggregated Then
            Download_ICE_Data = MsgBox("Download ICE?", vbYesNo) = vbYes
            If Download_ICE_Data Then localIceDate = DateAdd("yyyy", -2, localIceDate)
        Else
            Download_ICE_Data = False
        End If
        
        Download_CFTC_Data = MsgBox("Download CFTC?", vbYesNo) = vbYes
        If Download_CFTC_Data Then localCftcDate = DateAdd("yyyy", -2, localCftcDate)
    End If
    
    Application.DisplayAlerts = False
        
    If Download_ICE_Data And eReport = eDisaggregated Then
    
        Retrieve_Historical_Workbooks _
            Path_CLCTN:=File_CLCTN, _
            ICE_Contracts:=True, _
            CFTC_Contracts:=False, _
            Mac_User:=MacB, _
            eReport:=eReport, _
            downloadFuturesAndOptions:=getFuturesAndOptions, _
            ICE_Start_Date:=localIceDate, _
            ICE_End_Date:=maxDateICE
        
        If staticIceByUrl Is Nothing Then Set staticIceByUrl = New Collection
        
        For i = 1 To File_CLCTN.Count
            iceURL = File_CLCTN(i)
            
            If Not HasKey(staticIceByUrl, iceURL) Then
                Set oneYearData = ICE_Query(iceURL, localIceDate)
                staticIceByUrl.Add oneYearData, iceURL
            Else
                Set oneYearData = staticIceByUrl(iceURL)
            End If
            
            With oneYearData
                iceData = .item(ConvertOpenInterestTypeToName(CInt(getFuturesAndOptions)))
                If IsArrayAllocated(iceData) Then
                    New_Data.Add iceData
                End If
            End With
        Next i
      
    End If
    
    With New_Data
        #If Not Mac Then
            If Download_CFTC_Data Then
                Set File_CLCTN = Nothing
                
                Retrieve_Historical_Workbooks _
                    Path_CLCTN:=File_CLCTN, _
                    ICE_Contracts:=False, _
                    CFTC_Contracts:=True, _
                    Mac_User:=MacB, _
                    CFTC_Start_Date:=localCftcDate, _
                    CFTC_End_Date:=maxDateCFTC, _
                    eReport:=eReport, _
                    downloadFuturesAndOptions:=getFuturesAndOptions
                    
                .Add Historical_Parse(File_CLCTN, retrieveCombinedData:=getFuturesAndOptions, reportType:=eReport, After_This_Date:=localCftcDate, Kill_Previous_Workbook:=DebugMD)
            End If
        #End If
        
        Application.DisplayAlerts = True
        
        If .Count = 1 Then
            Missing_Data = New_Data(1)
        ElseIf .Count > 1 Then
            Missing_Data = CombineArraysInCollection(New_Data, Append_Type.Multiple_2d)
        End If
    End With
    
    Exit Function
Propagate:
    Application.DisplayAlerts = True
    Call PropagateError(Err, "Missing_Data")
End Function
Private Function Weekly_ICE(Most_Recent_CFTC_Date As Date) As Collection
'===================================================================================================================
'   Summary: Generates a weekly ice url and gets weekly data.
'   Inputs:
'      Most_Recent_CFTC_Date - Date used to generate a weekly ice url.
'   Returns:
'      A Collection of Weekly ICE data keyed to Open Interest Name.
'===================================================================================================================
    Dim weeklyIceUrl$
    On Error GoTo Propagate
    weeklyIceUrl = "https://www.theice.com/publicdocs/cot_report/automated/COT_" & Format$(Most_Recent_CFTC_Date, "ddmmyyyy") & ".csv"
    Set Weekly_ICE = ICE_Query(weeklyIceUrl, Most_Recent_CFTC_Date)
    Exit Function
Propagate:
    PropagateError Err, "Weekly_ICE"
End Function

Private Function ICE_Query(iceURL$, greaterThanDate As Date) As Collection
'===================================================================================================================
'   Summary: Queriees data from ICE using a querytable.
'   Inputs:
'      greaterThanDate -Used to filter any returned records by date.
'   Returns:
'      A Collection object containing Futures & Options and Futures Only data by key.
'===================================================================================================================
    Dim Data_Query As QueryTable, Y As Long, BB As Boolean, getFuturesAndOptions As Boolean, _
    Found_Data_Query As Boolean, Filtered_CLCTN As Collection
    
    Const connectionName$ = "ICE Data Refresh Connection", queryName = "ICE Data Refresh"
    
    With Application
        BB = .EnableEvents
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    On Error GoTo Finally
    
    For Each Data_Query In QueryT.QueryTables
        If Data_Query.name Like "*" & queryName & "*" Then
            Found_Data_Query = True
            Exit For
        End If
    Next Data_Query
    
    If Not Found_Data_Query Then 'If QueryTable isn't found then create it
        With QueryT
            Set Data_Query = .QueryTables.Add(Connection:="TEXT;" & iceURL, Destination:=.Range("A1"))
        End With
    End If
    
    With Data_Query
        .Connection = "TEXT;" & iceURL
        .BackgroundQuery = False
        .SaveData = False
        .AdjustColumnWidth = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileCommaDelimiter = True
        .name = queryName
        .WorkbookConnection.RefreshWithRefreshAll = False
        .TextFileColumnDataTypes = Filter_Market_Columns(convert_skip_col_to_general:=True, reportTypeEnum:=eDisaggregated, Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True, ICE:=True)
        .Refresh False
        
        .ResultRange.Replace " ", Empty, xlWhole
        
        Set Filtered_CLCTN = New Collection
        
        With Filtered_CLCTN
            For Y = 1 To 2
                getFuturesAndOptions = Not getFuturesAndOptions
                .Add Historical_Excel_Aggregation(ThisWorkbook, getFuturesAndOptions, Date_Input:=greaterThanDate, ICE_Contracts:=True, QueryTable_To_Filter:=Data_Query), ConvertOpenInterestTypeToName(CInt(getFuturesAndOptions))
            Next Y
        End With
        
        With .ResultRange
            On Error Resume Next
            .Parent.ShowAllData
            .ClearContents
            Err.Clear
        End With
    End With
    
    Set ICE_Query = Filtered_CLCTN
Finally:

    With Application
        .DisplayAlerts = True
        .EnableEvents = BB
    End With
    
    If Not Data_Query Is Nothing Then
        With Data_Query
            .WorkbookConnection.Delete
            .Delete
        End With
        Set Data_Query = Nothing
    End If
    
    If Err.Number <> 0 Then PropagateError Err, "ICE_Query"
    
End Function

#If Not DatabaseFile Then

    Private Sub Convert_Workbook_Version()
    '===================================================================================================================
        'Summary: Prompts the user for whether they wanat to download futures only or futures + options data.
        '         All data in the worksheet will then be replaced.
    '===================================================================================================================

        Dim User_Selection As Long, Retrieve_Futures_Only As Boolean
        
        User_Selection = CLng(InputBox("( 1 ) for Futures Only" & String$(2, vbNewLine) & "{ 2 ) for Futures + Options."))
        
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
            If ConvertInitialToReportTypeEnum(ReturnReportType) = eDisaggregated Then .Range("Last_Updated_ICE").Value2 = 0
        End With
        
        Call New_Data_Query(Scheduled_Retrieval:=True, Overwrite_All_Data:=True)
        
        MsgBox "Conversion Complete"
        
    End Sub

    Public Sub Add_New_CFTC_Contract()
    '===================================================================================================================
    'Summary: Queries the Socrata API for specific contract data by contract code and outputs it in a new worksheet.
    '===================================================================================================================
        Dim availableContracts As Collection, wantedContractCode$, invalidContractCode As Boolean, cftcData() As Variant, _
        First_Calculated_Column As Long
        
        Dim ws As Worksheet, eReport As ReportEnum, apiStatusCode As SocrataStatus
    
        Set availableContracts = GetAvailableContractInfo()
        
        eReport = ConvertInitialToReportTypeEnum(ReturnReportType())
        
        First_Calculated_Column = 3 + WorksheetFunction.CountIf(GetAvailableFieldsTable(eReport).DataBodyRange.columns(2), True)
        
        Do
            wantedContractCode = InputBox("Enter a 6 digit CFTC contract code")
            
            If wantedContractCode = vbNullString Then
                Exit Sub
            Else
                invalidContractCode = False
            End If
            
        Loop While Len(wantedContractCode) <> 6 Or invalidContractCode
                
        On Error GoTo No_Data_Retrieved_From_API
        
        Dim fieldInfoByName As Collection
        
        If TryGetCftcWithSocrataAPI(cftcData, eReport, IsWorkbookForFuturesAndOptions(), apiStatusCode, False, fieldInfoByName, wantedContractCode) Then
        
            cftcData = Filter_Market_Columns(False, True, False, eReport, True, cftcData, False)
            
            On Error GoTo 0
            
            ReDim Preserve cftcData(LBound(cftcData, 1) To UBound(cftcData, 1), LBound(cftcData, 2) To UBound(cftcData, 2) + 1)
            
            Dim newPriceSymbol$, CD As ContractInfo
                        
            If HasKey(availableContracts, wantedContractCode) Then
                Set CD = availableContracts(wantedContractCode)
            Else
                With Symbols.ListObjects("Symbols_TBL")
                    Set CD = New ContractInfo
                    On Error GoTo Catch_SymbolMissing
                    CD.InitializeContract wantedContractCode, newPriceSymbol, False
                End With
            End If
            
            With CD
                If .HasSymbol Then Call TryGetPriceData(cftcData, UBound(cftcData, 2), CD, datesAreInColumnOne:=True, overwriteAllPrices:=True)
            End With
Try_DoCalculations:
            On Error GoTo 0
            ReDim Preserve cftcData(LBound(cftcData, 1) To UBound(cftcData, 1), LBound(cftcData, 2) To Variable_Sheet.Range("Last_Calculated_Column").Value2)
            
            Select Case eReport
                Case eLegacy
                    cftcData = Legacy_Multi_Calculations(cftcData, UBound(cftcData, 1), First_Calculated_Column, 156, 26)
                Case eDisaggregated
                    cftcData = Disaggregated_Multi_Calculations(cftcData, UBound(cftcData, 1), First_Calculated_Column, 156, 26)
                Case eTFF
                    cftcData = TFF_Multi_Calculations(cftcData, UBound(cftcData, 1), First_Calculated_Column, 156, 26, 52)
            End Select
            
            Application.ScreenUpdating = False
                        
            If CD.TableSource Is Nothing Then
                Set ws = ThisWorkbook.Worksheets.Add
                With ws
                    .columns(1).NumberFormat = "yyyy-mm-dd"
                    .columns(First_Calculated_Column - 3).NumberFormat = "@"
                    Call Paste_To_Range(Sheet_Data:=cftcData, Historical_Paste:=True, Target_Sheet:=ws)
                    .ListObjects(1).name = "CFTC_" & wantedContractCode
                End With
                
                If Not CD.HasSymbol Then
                    Symbols.Range("Symbols_TBL").ListObject.ListRows.Add.Range.Value2 = Array(wantedContractCode, cftcData(UBound(cftcData, 1), 2), Empty, Empty)
                    MsgBox "A new row has been added to the availbale symbols table. Please fill in the missing Symbol information if available."
                End If
            
            Else
                Call Paste_To_Range(Sheet_Data:=cftcData, Historical_Paste:=True, Target_Sheet:=CD.SourceWorksheet, Overwrite_Data:=True)
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
    Private Function ConvertSymbolDataToJson$()
    '===================================================================================================================
    'Summary: Generates a JSON string with contract codes as a key and price symbols as a value.
    'Outputs: A serialized JSON object.
    '===================================================================================================================
        Dim jsonSerializer As New JsonParserB, dict As Object, CD As ContractInfo
        
        Set dict = GetDictionaryObject()
                
        For Each CD In GetAvailableContractInfo(True)
            With CD
                If .HasSymbol Then dict.item(.contractCode) = .PriceSymbol
            End With
        Next CD
        
       ConvertSymbolDataToJson = jsonSerializer.Serialize(dict, encloseStringsInEscapedQuotes:=True)
        
    End Function
    Private Function ListDatabasePathsInJson$()
    '===================================================================================================================
    'Summary: Generates a JSON string with report initials as a key and database paths for values.
    'Outputs: A serialized JSON object.
    '===================================================================================================================
        Dim dict As Object, availableReportEnums() As ReportEnum, i As Long, jsonSerializer As New JsonParserB
        
        Set dict = GetDictionaryObject()
        availableReportEnums = ReportEnumArray()
        
        For i = LBound(availableReportEnums) To UBound(availableReportEnums)
            With GetStoredReportDetails(availableReportEnums(i))
                dict.item(IIf(availableReportEnums(i) = eTFF, "TFF", .FullReportName.Value2)) = .CurrentDatabasePath.Value2
            End With
        Next i
        
        ListDatabasePathsInJson = jsonSerializer.Serialize(dict, encloseStringsInEscapedQuotes:=True)
        
    End Function
    Private Function RunCSharpExtractor() As Object
    '===================================================================================================================
    'Summary: Attempts to run a C# executable if available as an alternative to VBA data retrieval and database upload.
    'Inputs:
    '   timerClass - Used to time how long certain tasks take.
    'Returns:
    '   A deserialized JSON object.
    '===================================================================================================================
        Dim commandArgs$(3), shellCommand$, cmdOutput$, outputPath$
        
        commandArgs(0) = Variable_Sheet.Range("CSharp_Exe").Value2
        
        If FileOrFolderExists(commandArgs(0)) And LenB(commandArgs(0)) > 0 Then
        
            outputPath = Environ("Temp") & "\MoshiM_C_Output.txt"
            On Error GoTo Propagate
            
            commandArgs(0) = QuotedForm(commandArgs(0))
            commandArgs(1) = QuotedForm(ConvertSymbolDataToJson(), ensureAddition:=True)
            commandArgs(2) = CLng(True)
            commandArgs(3) = "> """ & outputPath & Chr(34)
            
            shellCommand = Join(commandArgs, " ")
            Erase commandArgs
            
            #If Not Mac Then
                Application.StatusBar = "Running " & commandArgs(0)
                Dim shellTerminal As New IWshRuntimeLibrary.WshShell, FSO As New Scripting.FileSystemObject
                                                
'                With shellTerminal.Exec("cmd.exe /c """ & shellCommand & """")
'                    cmdOutput = .StdOut.ReadAll
'                    .Terminate
'                End With
                With FSO
                    If .FileExists(outputPath) Then .DeleteFile outputPath
                     shellTerminal.Run "cmd.exe /c """ & shellCommand & """", WindowStyle:=7, waitonreturn:=True
                    If .FileExists(outputPath) Then
                        cmdOutput = .OpenTextFile(outputPath, 1).ReadAll
                        .DeleteFile outputPath
                    End If
                End With
                                                
                Set shellTerminal = Nothing: shellCommand = vbNullString: Set FSO = Nothing
                Application.StatusBar = "Executable has completed."
            #End If
                                                        
            If LenB(cmdOutput) > 0 Then
                Dim json$, jsonSerializer As New JsonParserB
                Debug.Print cmdOutput
                json = Split(Split(cmdOutput, "<json>")(1), "</json>")(0)
                Set RunCSharpExtractor = jsonSerializer.Deserialize(json, True, False, False)
            End If
            Application.StatusBar = vbNullString
        End If
        
        Exit Function
Propagate:
        Application.StatusBar = vbNullString
        PropagateError Err, "RunCSharpExtractor"
    End Function

#End If


