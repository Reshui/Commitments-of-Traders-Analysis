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

    Dim cftcDataA() As Variant, iceDataA() As Variant, Historical_Query() As Variant, reportsToQuery$()

    Dim reportInitial$, iReport As Byte, iOiType As Byte

    Dim cftcDateRange As Range, iceKey$
    'Booleans
    Dim debugWeeklyRetrieval As Boolean, DBM_Historical_Retrieval As Boolean, Debug_Mode As Boolean, _
    Download_CFTC As Boolean, Download_ICE As Boolean, Check_ICE As Boolean, queryFuturesAndOptions As Boolean, _
    uploadDataToDatabaseOrWorksheet As Boolean, isLegacyCombined As Boolean, _
    exitSubroutine As Boolean, CFTC_Retrieval_Error As Boolean, newDataSuccessfullyHandled As Boolean, processingReport As Boolean
    'Collections
    Dim cftcMappedFieldInfo As Collection, DataBase_Not_Found_CLCTN As Collection, _
    availableContractInfo As Collection, Data_CLCTN As Collection, Weekly_ICE_CLCTN As Collection, storedYearlyIceCLCTN As Collection
    
    Dim taskProfiler As TimedTask, individualCotTask As TimedTask

    Const dataRetrieval$ = "C.O.T data retrieval", uploadTime$ = "Upload Time", _
    databaseDateQuery$ = "Query database for latest date.", cRetrievalTask$ = "C# Retrieval", ProcedureName$ = "New_Data_Query"
     
    Dim legacyReportEnum As ReportEnum: legacyReportEnum = eLegacy
    
    Dim useSocrataAPI As Boolean, socrataApiFailed As Boolean, executableSuccess As Boolean, cftcAlreadyUpdated As Boolean, iceAlreadyUpdated As Boolean
    
    Call IncreasePerformance

    On Error GoTo Deny_Debug_Mode
        If Weekly.Shapes("Test_Toggle").OLEFormat.Object.value = xlOn Then Debug_Mode = True 'Determine if Debug status
    On Error GoTo Catch_General_Error

    #If DatabaseFile Then
        
        Dim openInterestTypesToQuery(1) As OpenInterestType, executeableReturn As Object, _
        databaseReturnedDate As Boolean
        Const latestDateKey$ = "Latest Date", statusKey$ = "Status"
        
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
                    queryFuturesAndOptions = CBool(InputBox("Select 1 of the following:" & vbNewLine & vbNewLine & "  Futures Only: 0" & vbNewLine & "  Futures + Options: 1"))
                Loop While Err.Number <> 0
                
                openInterestTypesToQuery(0) = queryFuturesAndOptions
            
            #End If
            
            On Error GoTo Catch_General_Error
            .Continue
        End With
    Else
        #If DatabaseFile Then
            With taskProfiler.StartSubTask(cRetrievalTask)
                On Error GoTo Catch_ExecutableFailed
                Set executeableReturn = RunCSharpExtractor(.ReturnReference())
                On Error GoTo Catch_General_Error
                .EndTask
            End With
        #Else
            If IsOnCreatorComputer Then Exit Sub
        #End If
    End If

Retrieve_Latest_Data:
    
    Dim eReport As ReportEnum
    
    For iReport = LBound(reportsToQuery) To UBound(reportsToQuery)
        
        reportInitial = reportsToQuery(iReport)
        eReport = ConvertInitialToReportTypeEnum(reportInitial)
        
        Set cftcMappedFieldInfo = Nothing
        
        For iOiType = LBound(openInterestTypesToQuery) To UBound(openInterestTypesToQuery)
            
            processingReport = True
            On Error GoTo Catch_General_Error
            
            queryFuturesAndOptions = openInterestTypesToQuery(iOiType)
            isLegacyCombined = (queryFuturesAndOptions And eReport = legacyReportEnum)
            
            cftcAlreadyUpdated = False: iceAlreadyUpdated = False: uploadDataToDatabaseOrWorksheet = False: Check_ICE = False
            
            #If DatabaseFile Then
                If Not Debug_Mode And iOiType = LBound(openInterestTypesToQuery) And iReport = LBound(reportsToQuery) And Not isLegacyCombined Then
                    MsgBox "Legacy Combined data needs to be retrieved first so that price data only has to be retrieved once."
                    GoTo Exit_Procedure
                End If
                
                On Error GoTo Try_Get_New_Data
                ' If data was updated by C# executable then goto next loop.
                If Not executeableReturn Is Nothing Then
                    With executeableReturn.Item(CStr(queryFuturesAndOptions)).Item(reportInitial)
                        
                        Select Case .Item(statusKey)
                            Case ReportStatusCode.NoUpdateAvailable, ReportStatusCode.Updated
                                executableSuccess = True
                                If CFTC_Incoming_Date < .Item(latestDateKey) Then CFTC_Incoming_Date = .Item(latestDateKey)
                                
                                If .Item(latestDateKey) > cftcDateRange.Value2 Or .Item(statusKey) = ReportStatusCode.Updated Then
                                    With GetStoredReportDetails(eReport)
                                        Select Case .OpenInterestType.Value2
                                            Case queryFuturesAndOptions, OpenInterestType.OptionsOnly
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
            
            Set individualCotTask = taskProfiler.StartSubTask("( " & reportInitial & " ) Combined: (" & queryFuturesAndOptions & ")")
            
            #If DatabaseFile Then
                With individualCotTask.StartSubTask(databaseDateQuery)
                    databaseReturnedDate = TryGetLatestDate(Last_Update_CFTC, reportType:=eReport, versionToQuery:=openInterestTypesToQuery(iOiType), queryIceContracts:=False)
                    .EndTask
                    
                    If Not databaseReturnedDate Then
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
                CFTC_Incoming_Date = cftcDataA(UBound(cftcDataA, 1), rawCftcDateColumn)
                cftcAlreadyUpdated = (CFTC_Incoming_Date = Last_Update_CFTC)
                
            Else
                'New data not available for Non Legacy Combined
                GoTo Stop_Timers_And_Update_If_Allowed
            End If
            
Try_IceRetrieval:
            If eReport = eDisaggregated Then
                
                On Error GoTo ICE_Retrieval_Failed
                                               
                #If DatabaseFile Then
                    Call TryGetLatestDate(Last_Update_ICE, reportType:=eReport, versionToQuery:=openInterestTypesToQuery(iOiType), queryIceContracts:=True)
                #Else
                    Last_Update_ICE = iceDateRange.Value2
                #End If
                
                If Last_Update_ICE < CFTC_Incoming_Date Or Debug_Mode Then
                    Check_ICE = True
                    
                    If Weekly_ICE_CLCTN Is Nothing Then
                        Set Weekly_ICE_CLCTN = Weekly_ICE(CFTC_Incoming_Date)
                    End If
                    
                    iceKey = ConvertOpenInterestTypeToName(openInterestTypesToQuery(iOiType))
                    
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
    
                    Historical_Query = Missing_Data(getFuturesAndOptions:=queryFuturesAndOptions, _
                        maxDateICE:=ICE_Incoming_Date, maxDateCFTC:=CFTC_Incoming_Date, _
                        Download_ICE_Data:=Download_ICE, Download_CFTC_Data:=Download_CFTC, _
                        eReport:=eReport, _
                        localCftcDate:=Last_Update_CFTC, localIceDate:=Last_Update_ICE + 2, _
                        DebugMD:=DBM_Historical_Retrieval, yearlyIceCLCTN:=storedYearlyIceCLCTN)
                                                     
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
                    
                ElseIf Not cftcAlreadyUpdated Or (Check_ICE And Not iceAlreadyUpdated) Or Debug_Mode = True Then
                    Set Data_CLCTN = New Collection
                    With Data_CLCTN
                        If Check_ICE And (ICE_Incoming_Date - Last_Update_ICE > 0 Or Debug_Mode) And IsArrayAllocated(iceDataA) Then .Add iceDataA
                        If (CFTC_Incoming_Date - Last_Update_CFTC > 0 Or Debug_Mode) And IsArrayAllocated(cftcDataA) Then .Add cftcDataA
                        
                        If .count > 0 Then
                            Select Case .count
                                Case 1:
                                    Historical_Query = .Item(1)
                                Case 2:
                                    If UBound(.Item(1), 2) <> UBound(.Item(2), 2) Then
                                        Err.Raise vbObjectError + 599, Description:="ICE data and CFTC data don't have an equal number of fields."
                                    End If
                                    Historical_Query = CombineArraysInCollection(Data_CLCTN, Append_Type.Multiple_2d)
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
            
                If IsArrayAllocated(Historical_Query) And uploadDataToDatabaseOrWorksheet = True And Not exitSubroutine Then
                    ' Upload data to database/spreadsheet and retrieve prices.
                    With .StartSubTask(uploadTime)
                        
                        If availableContractInfo Is Nothing Then
                            On Error GoTo Catch_General_Error
                            Set availableContractInfo = GetAvailableContractInfo()
                        End If
                        
                        If cftcMappedFieldInfo Is Nothing Then Set cftcMappedFieldInfo = GetExpectedLocalFieldInfo(eReport, False, False, False, False)
                        
                        On Error GoTo Catch_Block_Query_Failed
                        Call Block_Query(cotData:=Historical_Query, reportType:=eReport, _
                                        oiType:=openInterestTypesToQuery(iOiType), availableContractInfo:=availableContractInfo, _
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
                    
                    With taskProfiler.StartSubTask("Query all databases for latest contracts.")
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
            DisplayErr Err, ProcedureName
    End Select
    
    CFTC_Retrieval_Error = True
    exitSubroutine = True
    Resume Stop_Timers_And_Update_If_Allowed
    
#If DatabaseFile Then
Catch_ExecutableFailed:
    taskProfiler.StopSubTask cRetrievalTask
    If IsOnCreatorComputer Then DisplayErr Err, ProcedureName, "Failed to parse values from .exe response."
    Set executeableReturn = Nothing
    Resume Next
#End If

Catch_General_Error:
    DisplayErr Err, IIf(processingReport, reportInitial & "_" & ConvertOpenInterestTypeToName(CLng(queryFuturesAndOptions)), vbNullString)
    Resume Finally
    
Catch_Block_Query_Failed:
    Erase Historical_Query
    DisplayErr Err, ProcedureName, IIf(processingReport, reportInitial & "_" & ConvertOpenInterestTypeToName(CLng(queryFuturesAndOptions)), vbNullString)
    Resume Next_Combined_Value
    
Catch_Database_Not_Found:
    If isLegacyCombined Then exitSubroutine = True
    Resume Next_Report_Release_Type
    
End Sub
Private Sub Block_Query(ByRef cotData() As Variant, reportType As ReportEnum, oiType As OpenInterestType, _
                        availableContractInfo As Collection, debugOnly As Boolean, _
                        mappedFields As Collection, Optional Overwrite_Worksheet As Boolean = False)
'===================================================================================================================
    'Summary: Data within Query will be pruned for wanted columns and uploaded to either a database or worksheet.
    'Inputs: Query - Array that holds data to store.
    '        reportType - Report that is being uploaded.
    '        oiType - True if data is futures + options; else, futures only.
    '        availableContractInfo - Collection of Contract instances. If on non-database file then this
    '                      contains only contracts within the file. IF database file then all contracts with a price symbol available.
    '        Overwrite_Worksheet - If True then all data on a worksheet will be replaced with the data matching its contract code.
    '        mappedFields - Collection of FieldInfo instances that represent each column in Query.
'===================================================================================================================
    Dim iRow As Long, iCount As Long, databaseVersion As Boolean, uniqueContractCount As Long
    
    Dim Block() As Variant, Contract_CLCTN As New Collection, priceColumn As Byte, contractCode$
    
    Dim retrievePriceData As Boolean, placeDataOnWorksheetDebug As Boolean, contractCodeColumn As Byte

    Dim progressBarActive As Boolean, enableProgressBar As Boolean
    
    Const BLOCK_QUERY_ERR As Long = vbObjectError + 571, ProcedureName$ = "Block_Query"
    
    On Error Resume Next
    enableProgressBar = Weekly.Shapes("Progress_CHKBX").OLEFormat.Object.value = xlOn
    
    On Error GoTo Propagate
    
    If mappedFields Is Nothing Then
        Err.Raise BLOCK_QUERY_ERR, Description:="Field map [mappedFields] hasn't been declared."
    ElseIf mappedFields.count = 0 Then
        Err.Raise BLOCK_QUERY_ERR, Description:="Field map [mappedFields] is empty."
    End If
        
    #If Not DatabaseFile Then
    
        Dim columnfilter() As Variant, missingWeeksCount As Long, Last_Calculated_Column As Long, _
        current_Filters() As Variant, firstCalculatedColumn As Byte, WS_Data() As Variant, outputTable As ListObject, Table_Data_CLCTN As New Collection
        
        Const Time1 As Long = 156, Time2 As Long = 26, Time3 As Long = 52

        Last_Calculated_Column = Variable_Sheet.Range("Last_Calculated_Column").Value2
        
        columnfilter = Filter_Market_Columns(True, False, False, reportType, Create_Filter:=True)
         
        ' +1 To account for filter returning a 1 based array
        ' +1 To get wanted value
        ' = +2
        priceColumn = UBound(Filter(columnfilter, xlSkipColumn, False)) + 2
        firstCalculatedColumn = priceColumn + 2
        
    #Else
    
        On Error GoTo Catch_NotEnoughDimensions
        priceColumn = UBound(cotData, 2) + 1
        
        On Error GoTo Catch_ReDimension_Failure
        ReDim Preserve cotData(LBound(cotData, 1) To UBound(cotData, 1), LBound(cotData, 2) To priceColumn)
        
        databaseVersion = True
        On Error GoTo Propagate
        
    #End If
    
    If Not HasKey(mappedFields, "price") Then
        mappedFields.Add CreateFieldInfoInstance("price", CInt(priceColumn), "Price", False, True, False), "price"
    End If
    
    If (databaseVersion And oiType = OpenInterestType.FuturesAndOptions And reportType = eLegacy) Or Not databaseVersion Then
        
        On Error GoTo Catch_Missing_Contract_Code_Field
        contractCodeColumn = mappedFields("cftc_contract_market_code").ColumnIndex
        
        ReDim Block(LBound(cotData, 2) To UBound(cotData, 2))
        ' Parse array rows into collections keyed to their contract code.
        ' Array should be date sorted
        On Error GoTo Catch_MissingCollection
        For iRow = LBound(cotData, 1) To UBound(cotData, 1)
            For iCount = LBound(cotData, 2) To UBound(cotData, 2)
                Block(iCount) = cotData(iRow, iCount)
            Next iCount
            Contract_CLCTN(Block(contractCodeColumn)).Add Block
        Next iRow
        
        On Error GoTo Propagate
        Erase cotData
        
        uniqueContractCount = Contract_CLCTN.count
        
        If uniqueContractCount = 0 Then
            Err.Raise ERROR_BLOCK_QUERY_NO_UNIQUE_CONTRACTS, Description:="No unique contracts available for " & ConvertReportTypeEnum(reportType) & "-[Futures_Options: " & CBool(oiType) & "]"
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

        If enableProgressBar And Not databaseVersion Then
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
                
                If retrievePriceData And availableContractInfo(contractCode).HasSymbol Then Call TryGetPriceData(Block, priceColumn, availableContractInfo(contractCode), overwriteAllPrices:=False, datesAreInColumnOne:=Not databaseVersion)
            
            ElseIf Not databaseVersion Then
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
                        If .SortFields.count > 0 Then .Apply
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
            If progressBarActive And Not databaseVersion Then
                If iCount = 1 Then
                    Unload Progress_Bar
                Else
                    Progress_Bar.IncrementBar
                End If
            End If
        Next iCount
        
        #If DatabaseFile Then
            If Contract_CLCTN.count > 0 Then
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
            Call Update_Database(dataToUpload:=cotData, versionToUpdate:=oiType, eReport:=reportType, debugOnly:=debugOnly, suppliedFieldInfoByEditedName:=mappedFields)
        Else
            Err.Raise BLOCK_QUERY_ERR, Description:="Source array or field map is unallocated."
        End If
    #End If
    
    Exit Sub
Propagate:
    'Stop: Resume
    Call PropagateError(Err, ProcedureName)
Catch_ReDimension_Failure:
    Call PropagateError(Err, ProcedureName, "Failed to extend input data to fit price column.")
Catch_NotEnoughDimensions:
    Call PropagateError(Err, ProcedureName, "Source array doesn't have a 2nd dimension.")
Catch_ContractCodeIndexError:
    Call PropagateError(Err, ProcedureName, "Contract code for current row couldn't be retrieved for Block(contractCodeColumn).")
Catch_Missing_Contract_Code_Field:
    
    Dim knownField As FieldInfo, names$(), fromSocrata As Boolean
    
    ReDim names(mappedFields.count)
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
                       "From Socrata: " & fromSocrata & " | [mappedFields] count : " & mappedFields.count & vbNewLine & .Description & vbNewLine & "List: " & Join(names, ", ")
    End With
    GoTo Propagate
Catch_MissingCollection:
    On Error GoTo Catch_ContractCodeIndexError
    Contract_CLCTN.Add New Collection, Block(contractCodeColumn)
    On Error GoTo Catch_MissingCollection
    Resume
End Sub
Private Function Missing_Data(ByVal maxDateCFTC As Date, ByVal localCftcDate As Date, ByVal localIceDate As Date, _
                ByVal maxDateICE As Date, eReport As ReportEnum, getFuturesAndOptions As Boolean, _
                ByVal Download_ICE_Data As Boolean, ByVal Download_CFTC_Data As Boolean, _
                Optional DebugMD As Boolean = False, Optional yearlyIceCLCTN As Collection) As Variant()
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
'       yearlyIceCLCTN - Collection used to store yearly ICE data.
'===================================================================================================================
    
    Dim File_CLCTN As New Collection, MacB As Boolean, New_Data As New Collection, iceURL$, _
    i As Byte, iceData() As Variant, oneYearData As Collection
    
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
        
        If yearlyIceCLCTN Is Nothing Then Set yearlyIceCLCTN = New Collection
        
        For i = 1 To File_CLCTN.count
            iceURL = File_CLCTN(i)
            
            If Not HasKey(yearlyIceCLCTN, iceURL) Then
                Set oneYearData = ICE_Query(iceURL, localIceDate)
                yearlyIceCLCTN.Add oneYearData, iceURL
            Else
                Set oneYearData = yearlyIceCLCTN(iceURL)
            End If
            
            With oneYearData
                iceData = .Item(ConvertOpenInterestTypeToName(CInt(getFuturesAndOptions)))
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
                    
                .Add Historical_Parse(File_CLCTN, retrieveCombinedData:=getFuturesAndOptions, reportType:=eReport, parsingMultipleWeeks:=True, After_This_Date:=localCftcDate, Kill_Previous_Workbook:=DebugMD)
            End If
        #End If
        
        Application.DisplayAlerts = True
        
        If .count = 1 Then
            Missing_Data = New_Data(1)
        ElseIf .count > 1 Then
            Missing_Data = CombineArraysInCollection(New_Data, Append_Type.Multiple_2d)
        End If
    End With
    
    Exit Function
Propagate:
    Application.DisplayAlerts = True
    Call PropagateError(Err, "Missing_Data")
End Function
Private Function Weekly_ICE(Most_Recent_CFTC_Date As Date) As Collection
    
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
    'Summary: Queriees data from ICE using a querytable.
    'Inputs:
    '   greaterThanDate -Used to filter any returned records by date.
    'Returns:
    '   A Collection object containing Futures & Options and Futures Only data by key.
'===================================================================================================================
    Dim Data_Query As QueryTable, Y As Byte, BB As Boolean, getFuturesAndOptions As Boolean, _
    Found_Data_Query As Boolean, Filtered_CLCTN As Collection
    
    Const connectionName$ = "ICE Data Refresh Connection", queryName = "ICE Data Refresh"
    
    With Application
        BB = .EnableEvents
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    On Error GoTo Finally
    
    For Each Data_Query In QueryT.QueryTables
        If Data_Query.Name Like "*" & queryName & "*" Then
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
        .Name = queryName
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
            .parent.ShowAllData
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
            If ConvertInitialToReportTypeEnum(ReturnReportType) = eDisaggregated Then .Range("Last_Updated_ICE").Value2 = 0
        End With
        
        Call New_Data_Query(Scheduled_Retrieval:=True, Overwrite_All_Data:=True)
        
        MsgBox "Conversion Complete"
        
    End Sub

    Public Sub New_CFTC_Data()
    
        Dim Current_Contracts As Collection, retrieveCombinedData As Boolean, File_Paths As New Collection, _
        wantedContractCode$, invalidContractCode As Boolean, New_Data() As Variant, First_Calculated_Column As Byte
        
        Dim WS As Worksheet, Symbol_Row As Long, iColumn As Byte, reportType As ReportEnum
    
        Set Current_Contracts = GetAvailableContractInfo
        
        reportType = ConvertInitialToReportTypeEnum(ReturnReportType)
        
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
            
            With Symbols.Range("Symbols_TBL")
            
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
            ReDim Preserve New_Data(LBound(New_Data, 1) To UBound(New_Data, 1), LBound(New_Data, 2) To Variable_Sheet.Range("Last_Calculated_Column").Value2)
            
            Select Case reportType
                Case eLegacy
                    New_Data = Legacy_Multi_Calculations(New_Data, UBound(New_Data, 1), First_Calculated_Column, 156, 26)
                Case eDisaggregated
                    New_Data = Disaggregated_Multi_Calculations(New_Data, UBound(New_Data, 1), First_Calculated_Column, 156, 26)
                Case eTFF
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
                Symbols.Range("Symbols_TBL").ListObject.ListRows.Add.Range.Value2 = Array(wantedContractCode, New_Data(UBound(New_Data, 1), 2), Empty, Empty)
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
    Private Function ConvertSymbolDataToJson$()
    '===================================================================================================================
        'Summary: Generates a JSON string with contract codes as a key and price symbols as a value.
        'Outputs: A serialized JSON object.
    '===================================================================================================================
        Dim JsonSerializer As New JsonParserB, dict As Object, CD As ContractInfo
        
        Set dict = CreateObject("Scripting.Dictionary")
                
        For Each CD In GetAvailableContractInfo
            With CD
                If .HasSymbol Then dict.Item(.contractCode) = .priceSymbol
            End With
        Next CD
        
       ConvertSymbolDataToJson = JsonSerializer.Serialize(dict, encloseStringsInEscapedQuotes:=True)
        
    End Function
    Private Function ListDatabasePathsInJson$()
    
        Dim dict As Object, availableReportEnums() As ReportEnum, i As Byte, JsonSerializer As New JsonParserB
        
        Set dict = CreateObject("Scripting.Dictionary")
        availableReportEnums = ReportEnumArray
        
        For i = LBound(availableReportEnums) To UBound(availableReportEnums)
            With GetStoredReportDetails(availableReportEnums(i))
                dict.Item(IIf(availableReportEnums(i) = eTFF, "TFF", .FullReportName.Value2)) = .CurrentDatabasePath.Value2
            End With
        Next i
        
        ListDatabasePathsInJson = JsonSerializer.Serialize(dict, encloseStringsInEscapedQuotes:=True)
        
    End Function
    Public Function RunCSharpExtractor(timerClass As TimedTask) As Object
    '===================================================================================================================
        'Summary: Attempts to run a C# executable if available as an alternative to VBA data retrieval and database upload.
        'Inputs:
        '   timerClass - Used to time how long certain tasks take.
        'Returns:
        '   A JSON object is returned.
    '===================================================================================================================
        Dim commandArgs$(3), shellCommand$, result$, outputPath$
                
        commandArgs(0) = Variable_Sheet.Range("CSharp_Exe").Value2
        
        On Error GoTo Propagate
        
        If FileOrFolderExists(commandArgs(0)) Then
        
            With timerClass.StartSubTask("Generate CMD")
                outputPath = Environ("Temp") & "\MoshiM_C_Output.txt"
                commandArgs(0) = QuotedForm(commandArgs(0))
                commandArgs(1) = QuotedForm(ListDatabasePathsInJson(), ensureAddition:=True)
                commandArgs(2) = QuotedForm(ConvertSymbolDataToJson(), ensureAddition:=True)
                commandArgs(3) = "> """ & outputPath & Chr(34)
                shellCommand = Join(commandArgs, " ")
                .EndTask
            End With
            
            Application.StatusBar = "Querying new data with " & commandArgs(0)
            Erase commandArgs
            
            With timerClass.StartSubTask("Executable runtime.")
                CreateObject("WScript.Shell").Run "cmd.exe /c """ & shellCommand & """", 0, True
                .EndTask
            End With
            
            shellCommand = vbNullString
            
            With timerClass.StartSubTask("Retrieve output info.")
                result = CreateObject("Scripting.FileSystemObject").OpenTextFile(outputPath, 1).ReadAll
                .EndTask
            End With
            
            Kill outputPath
            
            Dim json$, jp As New JsonParserB
            json = Split(Split(result, "<json>")(1), "</json>")(0)
            
            With timerClass.StartSubTask("Deserialize JSON")
                Set RunCSharpExtractor = jp.Deserialize(json, True, False, False)
                .EndTask
            End With
            
            Debug.Print result
            Application.StatusBar = vbNullString
        Else
            Err.Raise 53, Description:="C# executable couldn't be found."
        End If
        
        Exit Function
Propagate:
        PropagateError Err, "RunCSharpExtractor"
    End Function

#End If


