Attribute VB_Name = "Data_Retrieval"
Public Const rawCftcDateColumn As Byte = 3
Public Retrieval_Halted_For_User_Interaction As Boolean
Public Data_Updated_Successfully As Boolean
Public Running_Weekly_Retrieval As Boolean

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
'======================================================================================================
'Retrieves the latest data and outputs it to the worksheet if available
'======================================================================================================

    Dim Last_Update_CFTC As Long, CFTC_Incoming_Date As Long, ICE_Incoming_Date As Long, Last_Update_ICE As Long

    Dim Debug_Mode As Boolean, CFTC_Data() As Variant, report As Variant, ICE_Data() As Variant, Historical_Query() As Variant

    Dim DBM_Weekly_Retrieval As Boolean, DBM_Historical_Retrieval As Boolean, queryingCombinedData As Variant

    Dim newDataFound As Boolean, cftcDateRange As Range, iceDateRange As Range  ', All_Available_Contracts() As Variant
    
    Dim Download_CFTC As Boolean, Download_ICE As Boolean, Check_ICE As Boolean, uploadDataToDatabaseOrWorksheet As Boolean, Weekly_ICE_CLCTN As Collection, iceKey As String
    
    Dim DataBase_Not_Found_CLCTN As Collection, Symbol_Info As Collection, Data_CLCTN As Collection, Legacy_Combined_Data As Boolean, Weekly_Ice_Queried As Boolean
    
    Dim TestTimers As New TimerC, reportDataTimer As String, individualCotTask As TimedTask, databaseMissingOccured As Boolean, exitSubroutine As Boolean, CFTC_Retrieval_Error As Boolean
    
    Const dataRetrieval As String = "C.O.T data retrieval", uploadTime As String = "Upload Time", _
    totalTime As String = "Total runtime", databaseDateQuery As String = "Query database for latest date"
     
    Const contractCodeColumn As Byte = 4, legacy_initial As String = "L", apiCodeColumn As Byte = 6
        
    Dim reportsToQuery() As String, combinedFutures() As Boolean, useAPIData As Boolean
    
    #If DatabaseFile Then
    
        Dim databaseExistsOnLocalMachine As Boolean
        
        #If Mac Then
            GateMacAccessToWorkbook
        #End If
        
        reportsToQuery = Split("L,D,T", ",")
        ReDim combinedFutures(1): combinedFutures(0) = True
        Set cftcDateRange = Variable_Sheet.Range("Most_Recently_Queried_Date")
    #Else
    
        ReDim reportsToQuery(0): reportsToQuery(0) = ReturnReportType
        ReDim combinedFutures(0): combinedFutures(0) = combined_workbook
        
        With Variable_Sheet
            Set cftcDateRange = .Range("Last_Updated_CFTC")
            If reportsToQuery(0) = "D" Then Set iceDateRange = .Range("Last_Updated_ICE")
        End With
        
    #End If
    'This function will either return all contracts for which price data is available or all contracts in the workbook.
    Set Symbol_Info = ContractDetails
    
    Running_Weekly_Retrieval = True
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    On Error GoTo Deny_Debug_Mode
    
    If Weekly.Shapes("Test_Toggle").OLEFormat.Object.value = 1 Then Debug_Mode = True 'Determine if Debug status
    
    If Debug_Mode = True Then
    
        If MsgBox("Test Weekly Data Retrieval ?", vbYesNo, "Choose what to debug") = vbYes Then
            DBM_Weekly_Retrieval = True
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
                     Last_Update_CFTC = Latest_Date(reportType:=CStr(report), combined_wb_bool:=CBool(queryingCombinedData), ICE_Query:=False, databaseExists:=databaseExistsOnLocalMachine)
                    .EndTask
                End With
                ' databaseExistsOnLocalMachine is given a value when assigning a value to Last_Update_CFTC.
                If Not databaseExistsOnLocalMachine Then
                    
                    If DataBase_Not_Found_CLCTN Is Nothing Then Set DataBase_Not_Found_CLCTN = New Collection
                    
                    databaseMissingOccured = True
                    
                    DataBase_Not_Found_CLCTN.Add "Missing database for " & Evaluate("VLOOKUP(""" & report & """,Report_Abbreviation,2,FALSE)")
                    
                    'The Legacy_Combined data is the only one for which price data is queried.
    
                    If Legacy_Combined_Data Then exitSubroutine = True
                    
                    Exit For

                End If
                
            #Else
                Last_Update_CFTC = cftcDateRange.Value2
            #End If
            
            individualCotTask.SubTask(dataRetrieval).Start
            
            If CFTC_Incoming_Date = 0 Or Last_Update_CFTC < CFTC_Incoming_Date Or Debug_Mode Then
                         
                On Error GoTo CFTC_Retrieval_Failed
                                                
                CFTC_Data = HTTP_Weekly_Data(Last_Update_CFTC, Auto_Retrieval:=Scheduled_Retrieval, _
                                retrieveCombinedData:=CBool(queryingCombinedData), _
                                reportType:=CStr(report), useApi:=useAPIData, _
                                DebugMD:=DBM_Weekly_Retrieval)
                
                If CFTC_Incoming_Date = 0 Then CFTC_Incoming_Date = CFTC_Data(UBound(CFTC_Data, 1), rawCftcDateColumn)
                
            Else
                'New data not available for Non Legacy Combined
                GoTo Stop_Timers_And_Update_If_Allowed
            End If

            If report = "D" Then
                
                On Error GoTo ICE_Retrieval_Failed
                
                Check_ICE = True
                
                #If DatabaseFile Then
                    Last_Update_ICE = Latest_Date(reportType:=CStr(report), combined_wb_bool:=CBool(queryingCombinedData), ICE_Query:=True, databaseExists:=databaseExistsOnLocalMachine)
                #Else
                    Last_Update_ICE = iceDateRange.Value2
                #End If
                
                If Not Weekly_Ice_Queried Then
                    Weekly_Ice_Queried = True
                    Set Weekly_ICE_CLCTN = Weekly_ICE(CBool(queryingCombinedData), CDate(CFTC_Data(UBound(CFTC_Data, 1), rawCftcDateColumn)))
                End If
                
                iceKey = IIf(queryingCombinedData = True, "futures + options", "futures-only")
                
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
            
            If DBM_Weekly_Retrieval Then
            
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
                
            ElseIf Not useAPIData And (CFTC_Incoming_Date - Last_Update_CFTC > 7 Or (Check_ICE And ICE_Incoming_Date - Last_Update_ICE > 7) Or DBM_Historical_Retrieval) Then
                
                If CFTC_Incoming_Date - Last_Update_CFTC > 7 Or DBM_Historical_Retrieval Then
                    Download_CFTC = True
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
                Historical_Query = Missing_Data(retrieve_combined_data:=CBool(queryingCombinedData), _
                    ICE_Data:=ICE_Data, CFTC_Data:=CFTC_Data, _
                    Download_ICE_Data:=Download_ICE, Download_CFTC_Data:=Download_CFTC, _
                    reportType:=CStr(report), _
                    CFTC_Last_Updated_Day:=Last_Update_CFTC, ICE_Last_Updated_Day:=Last_Update_ICE, _
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
                        
                        Historical_Query = Multi_Week_Addition(Data_CLCTN, Append_Type.Multiple_2d)
                        
                    ElseIf Download_ICE And CFTC_Incoming_Date - Last_Update_CFTC > 0 Then
                        'Determine if CFTC data needs to be added
                        Set Data_CLCTN = New Collection
                        
                        With Data_CLCTN
                            .Add Historical_Query
                            .Add CFTC_Data
                        End With
                        
                        Historical_Query = Multi_Week_Addition(Data_CLCTN, Append_Type.Multiple_2d)
                        
                    End If
                    
                End If

                newDataFound = True
                uploadDataToDatabaseOrWorksheet = True
                
            ElseIf (CFTC_Incoming_Date - Last_Update_CFTC) > 0 Or (Check_ICE And ICE_Incoming_Date - Last_Update_ICE > 0) Or Debug_Mode = True Then  'If just a 1 week difference
                ' Concatenate data if needed and prep for databaseupload/ paste to worksheet.
                Set Data_CLCTN = New Collection
                
                With Data_CLCTN
                
                    If Check_ICE And (ICE_Incoming_Date - Last_Update_ICE > 0 Or Debug_Mode) Then
                        ' If new ICE data has been added and is wanted.
                        .Add ICE_Data
                    End If
                    
                    If CFTC_Incoming_Date - Last_Update_CFTC > 0 Or Debug_Mode Then
                        ' If new CFTC data has been added.
                        .Add CFTC_Data
                    End If
                    
                    If .count = 1 Then
                        ' No concatenation is needed so store array in variable.
                        Historical_Query = .Item(1)
                    ElseIf .count = 2 Then
                        ' Concatenate the arrays.
                        Historical_Query = Multi_Week_Addition(Data_CLCTN, Append_Type.Multiple_2d)
                    Else
                        
                        With individualCotTask
                            .SubTask(dataRetrieval).EndTask
                            .EndTask
                        End With
                        
                        Exit For
                        
                    End If
                    
                End With
                
                newDataFound = True
                uploadDataToDatabaseOrWorksheet = True
                
            End If

Stop_Timers_And_Update_If_Allowed:

            With individualCotTask
                
                .SubTask(dataRetrieval).EndTask
            
                If uploadDataToDatabaseOrWorksheet = True And Not exitSubroutine Then
                    ' Upload data to database/spreadsheet and retrieve prices.
                    With .SubTask(uploadTime)
                        .Start
                         Call Block_Query(Query:=Historical_Query, Code_Column:=contractCodeColumn, reportType:=CStr(report), _
                                        processing_combined_data:=CBool(queryingCombinedData), Symbol_Info:=Symbol_Info, _
                                        debugOnly:=Debug_Mode, dataRetrievedFromAPI:=useAPIData, Overwrite_Worksheet:=Overwrite_All_Data)
                        .EndTask
                    End With
                    
                    uploadDataToDatabaseOrWorksheet = False
                    
                End If
                
                .EndTask
                
            End With
            
Next_Combined_Value:
            If Debug_Mode Or exitSubroutine Then Exit For
        
        Next queryingCombinedData
        
Next_Report_Release_Type:

        With individualCotTask
            If .isRunning Then .EndTask
        End With
        
        If Debug_Mode Or exitSubroutine Then Exit For
                        
    Next report
    
Exit_Procedure:
    
    On Error GoTo 0
    
    Re_Enable
    
    If newDataFound And Not exitSubroutine Then
        
        Data_Updated_Successfully = True
        '-------------------------------------------------------------------------------------------
        
        With cftcDateRange
            If CFTC_Incoming_Date > .value Then
                Update_Text CFTC_Incoming_Date   'Update Text Boxes "My_Date" on the HUB and Weekly worksheets.
                .value = CFTC_Incoming_Date
            End If
        End With
        
        If Check_ICE And Not iceDateRange Is Nothing Then
            With iceDateRange
                If ICE_Incoming_Date > .value Then
                    .value = ICE_Incoming_Date
                End If
            End With
        End If
        '----------------------------------------------------------------------------------------
        If Not Debug_Mode Then
        
            If Not Scheduled_Retrieval Then HUB.Activate 'If ran manually then bring the User to the HUB
            Courtesy                                     'Change Status Bar_Message
            
            If newDataFound Then
            
                #If DatabaseFile Then
        
                    Select Case ThisWorkbook.ActiveSheet.name
                        Case LC.name: Manage_Table_Visual "L", LC
                        Case DC.name: Manage_Table_Visual "D", DC
                        Case TC.name: Manage_Table_Visual "T", TC
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
        
    ElseIf databaseMissingOccured And Not Scheduled_Retrieval Then

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
    
    Running_Weekly_Retrieval = False
    
    With TestTimers
        .EndTask totalTime
        Debug.Print .ToString
    End With
    
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
    
End Sub

Private Sub Block_Query(ByRef Query, Code_Column As Byte, reportType As String, _
                    processing_combined_data As Boolean, Symbol_Info As Collection, debugOnly As Boolean, _
                    dataRetrievedFromAPI As Boolean, Optional Overwrite_Worksheet As Boolean = False)

'======================================================================================================
'Subroutine takes a properly formatted array and outputs all contracts to where they need to go.
'Overwrite_Worksheet =True when all data on a worksheet needs to be replaced
'======================================================================================================

    Dim X As Long, C As Integer, nonDatabaseVersion As Boolean, uniqueContractCount As Integer
    
    Dim Block() As Variant, Contract_CLCTN As New Collection, priceColumn As Byte, contractCode As String
    
    Dim retrievePriceData As Boolean, placeDataOnWorksheetDebug As Boolean
    '=======================================================================================================================

    Dim Progress_CHK As CheckBox, progressBarActive As Boolean

    Set Progress_CHK = Weekly.Shapes("Progress_CHKBX").OLEFormat.Object
    
    #If Not DatabaseFile Then
    
        Dim columnFilter() As Variant, missingWeeksCount As Integer, Last_Calculated_Column As Integer, _
        current_Filters() As Variant, DL_Reference As Byte, WS_Data() As Variant, Table_Range As Range, Table_Data_CLCTN As New Collection
        
        
        Dim WantedColumnForAPI As New Collection, Z As Long, tempKey As String
        
        Const Time1 As Integer = 156, Time2 As Integer = 26, Time3 As Integer = 52

        Last_Calculated_Column = Range("Last_Calculated_Column").Value2
        
        columnFilter = Filter_Market_Columns(True, False, False, reportType, Create_Filter:=True)
         
        ' +1 To account for filter returning a 1 based array
        ' +1 To get wanted value
        ' = +2
        priceColumn = UBound(filter(columnFilter, xlSkipColumn, False)) + 2

        DL_Reference = priceColumn + 2
        
        nonDatabaseVersion = True
        
    #Else
        priceColumn = UBound(Query, 2) + 1
        ReDim Preserve Query(LBound(Query, 1) To UBound(Query, 1), LBound(Query, 2) To priceColumn)  'Expand for calculations
    #End If
    
    If (Not nonDatabaseVersion And processing_combined_data And reportType = "L") Or nonDatabaseVersion Then
        
        ReDim Block(1 To UBound(Query, 2))
        ' Parse array rows into collections keyed to their contract code.
        ' Array should be date sorted
        For X = LBound(Query, 1) To UBound(Query, 1)
        
            For C = LBound(Query, 2) To UBound(Query, 2)
                Block(C) = Query(X, C)
            Next C
            
            On Error GoTo Missing_Collection
            Contract_CLCTN(Block(Code_Column)).Add Block
            
        Next X
        
        Erase Query
        
        uniqueContractCount = Contract_CLCTN.count
    
        If uniqueContractCount = 0 Then
            MsgBox "An error occured. No unique contracts were retrieved." & reportType & "-C: " & processing_combined_data
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

        If Progress_CHK.value = 1 And nonDatabaseVersion Then
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
        
        For C = uniqueContractCount To 1 Step -1   'Loop list of wanted Contract Codes
            
            On Error GoTo Contract_Code_Is_Missing
            
            Block = Multi_Week_Addition(Contract_CLCTN(C), Append_Type.Multiple_1d)
            
            contractCode = Block(1, Code_Column)
            ' Removes the collection, a combined version orf collection elements will be added later.
            Contract_CLCTN.Remove contractCode
            
            If HasKey(Symbol_Info, contractCode) Then
            
                #If Not DatabaseFile Then
                    Block = Filter_Market_Columns(False, True, False, reportType, False, Block, False, columnFilter)
                    ReDim Preserve Block(LBound(Block, 1) To UBound(Block, 1), LBound(Block, 2) To Last_Calculated_Column)  'Expand for calculations
                #End If
                
                If retrievePriceData Then Call Retrieve_Tuesdays_CLose(Block, priceColumn, Symbol_Info(contractCode), overwrite_all_prices:=False, dates_in_column_1:=nonDatabaseVersion) 'Gets Price_Info
            
            ElseIf nonDatabaseVersion Then
                GoTo NextAvailableContract
            End If
            
            #If Not DatabaseFile Then
            
                Set Table_Range = Symbol_Info(contractCode).TableSource.DataBodyRange
                
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
                    
                    Block = Multi_Week_Addition(Table_Data_CLCTN, Append_Type.Add_To_Old)
                
                    Set Table_Data_CLCTN = Nothing
                
                End If
                
                Select Case reportType
                    Case "L"
                        Block = Legacy_Multi_Calculations(Block, missingWeeksCount, DL_Reference, Time1, Time2)
                    Case "D"
                        Block = Disaggregated_Multi_Calculations(Block, missingWeeksCount, DL_Reference, Time1, Time2)
                    Case "T"
                        Block = TFF_Multi_Calculations(Block, missingWeeksCount, DL_Reference, Time1, Time2, Time3)
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
            
            If progressBarActive And nonDatabaseVersion Then
                If C = 1 Then
                    Unload Progress_Bar
                Else
                    Progress_Bar.IncrementBar
                End If
            End If

        Next C
        
        #If DatabaseFile Then
            ' Legacy Combined data needs to be rejoined together into a single array.
            Query = Multi_Week_Addition(Contract_CLCTN, Append_Type.Multiple_2d)
        #End If
        
    End If
    
Upload_Data:

    #If DatabaseFile Then
        
        On Error GoTo Database_Update_Error
        
        Call Update_DataBase(dataToUpload:=Query, uploadingCombinedData:=processing_combined_data, _
        reportType:=reportType, debugOnly:=debugOnly)
    
    #End If
    
    Exit Sub

Progress_Checkbox_Missing:

    'Set Progress_Control = Nothing
    Resume Block_Query_Main_Function
     
Contract_Code_Is_Missing:
    Resume NextAvailableContract
    
Database_Update_Error:

    MsgBox "Unhandled Error in database update step."
    
Missing_Collection:

    Contract_CLCTN.Add New Collection, Block(Code_Column)
    Resume
    
End Sub
Private Sub Update_Text(ByVal new_date As Long)
'======================================================================================================
'Updates shapes on the Weekly and Hub worksheet so that they show the last CFTC update date
'======================================================================================================

    Dim AN1 As String, TWS() As Variant
    
    AN1 = "[v] " & MonthName(Month(new_date)) & " " & Day(new_date) & ", " & Year(new_date)
    
    On Error Resume Next
    
    TWS = Array(Weekly, HUB)
    
    For new_date = LBound(TWS) To UBound(TWS)
        TWS(new_date).Shapes("My_Date").TextFrame.Characters.Text = AN1
    Next new_date
    
    End Sub
    Private Function Missing_Data(ByRef CFTC_Data As Variant, ByVal CFTC_Last_Updated_Day As Long, ByVal ICE_Last_Updated_Day As Long, ByRef ICE_Data As Variant, reportType As String, retrieve_combined_data As Boolean, Download_ICE_Data As Boolean, Download_CFTC_Data As Boolean, Optional DebugMD As Boolean = False) As Variant 'Should change to function; Block will find the amount of missed time and download appropriate files
    '======================================================================================================
    'Determines which files need to be downloaded for when multiple weeks of data have been missed
    '======================================================================================================
    Dim File_CLCTN As New Collection, MacB As Boolean, obj As Object, _
    Hyperlink_RNG As Range, T As Byte, New_Data As New Collection
    
    #If Mac Then
        MacB = True
    #End If
    
    If DebugMD Then
        If Not MacB Then If MsgBox("Do you want to test MAC OS data retrieval?", vbYesNo) = vbYes Then MacB = True
        
        If Download_CFTC_Data Then CFTC_Last_Updated_Day = DateAdd("yyyy", -2, CFTC_Last_Updated_Day)
        If Download_ICE_Data Then ICE_Last_Updated_Day = DateAdd("yyyy", -2, ICE_Last_Updated_Day)
    End If
    
    Application.DisplayAlerts = False
    
    If Download_ICE_Data And Download_CFTC_Data And CFTC_Last_Updated_Day = ICE_Last_Updated_Day Then
    
            Retrieve_Historical_Workbooks _
                Path_CLCTN:=File_CLCTN, _
                ICE_Contracts:=True, _
                CFTC_Contracts:=True, _
                Mac_User:=MacB, _
                reportType:=reportType, _
                combined_data_version:=retrieve_combined_data, _
                ICE_Start_Date:=CDate(ICE_Last_Updated_Day), ICE_End_Date:=CDate(ICE_Data(1, rawCftcDateColumn)), _
                CFTC_Start_Date:=CDate(CFTC_Last_Updated_Day), CFTC_End_Date:=CDate(CFTC_Data(1, rawCftcDateColumn))
            
           New_Data.Add Historical_Parse(File_CLCTN, retrieveCombinedData:=retrieve_combined_data, reportType:=reportType, Yearly_C:=True, After_This_Date:=ICE_Last_Updated_Day, Kill_Previous_Workbook:=DebugMD)
    
    Else
    
        If Download_ICE_Data Then
        
            Retrieve_Historical_Workbooks _
                Path_CLCTN:=File_CLCTN, _
                ICE_Contracts:=True, _
                CFTC_Contracts:=False, _
                Mac_User:=MacB, _
                reportType:=reportType, _
                combined_data_version:=retrieve_combined_data, _
                ICE_Start_Date:=CDate(ICE_Last_Updated_Day), _
                ICE_End_Date:=CDate(ICE_Data(1, rawCftcDateColumn))
            
            New_Data.Add Historical_Parse(File_CLCTN, retrieveCombinedData:=retrieve_combined_data, reportType:=reportType, Yearly_C:=True, After_This_Date:=ICE_Last_Updated_Day, Kill_Previous_Workbook:=DebugMD)
        
        End If
    
        If Download_CFTC_Data Then
            
            Set File_CLCTN = Nothing
            
            Retrieve_Historical_Workbooks _
                Path_CLCTN:=File_CLCTN, _
                ICE_Contracts:=False, _
                CFTC_Contracts:=True, _
                Mac_User:=MacB, _
                CFTC_Start_Date:=CDate(CFTC_Last_Updated_Day), _
                CFTC_End_Date:=CDate(CFTC_Data(1, rawCftcDateColumn)), _
                reportType:=reportType, _
                combined_data_version:=retrieve_combined_data
            
            New_Data.Add Historical_Parse(File_CLCTN, retrieveCombinedData:=retrieve_combined_data, reportType:=reportType, Yearly_C:=True, After_This_Date:=CFTC_Last_Updated_Day, Kill_Previous_Workbook:=DebugMD)
            
        End If
    
    End If
    
    Application.DisplayAlerts = True
    
    If New_Data.count = 1 Then
        Missing_Data = New_Data(1)
    ElseIf New_Data.count > 1 Then
        Missing_Data = Multi_Week_Addition(New_Data, Append_Type.Multiple_2d)
    End If

End Function
Public Function Weekly_ICE(Retrieve_Combined As Boolean, Most_Recent_CFTC_Date As Date) As Collection

'Dim Path_CLCTN As New Collection

    Dim ICE_URL As String
    
    ICE_URL = Get_ICE_URL(Most_Recent_CFTC_Date)
    
    'If isMAC Then
    
        On Error GoTo Exit_Sub
        
        Set Weekly_ICE = ICE_Query(Retrieve_Combined, ICE_URL)
        
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
    Get_ICE_URL = "https://www.theice.com/publicdocs/cot_report/automated/COT_" & Format(Query_Date, "ddmmyyyy") & ".csv"
    
    
    
    'https://www.theice.com/publicdocs/cot_report/automated/COT_15112022.csv
End Function

Private Function ICE_Query(ByVal Retrieve_Combined As Boolean, Weekly_ICE_URL As String) As Collection

    Dim Data_Query As QueryTable, data As Variant, Data_Row() As Variant, URL As String, _
    Column_Filter() As Variant, Y As Byte, bb As Boolean, _
    Found_Data_Query As Boolean, Error_While_Refreshing As Boolean, Filtered_CLCTN As Collection
    
    
    With Application
    
        bb = .EnableEvents
        
        .EnableEvents = False
        .DisplayAlerts = False
        
    End With
    
    For Each Data_Query In QueryT.QueryTables
    
        If Data_Query.name Like "*ICE Data Refresh*" Then
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
            
            .name = "ICE Data Refresh"
            
            On Error GoTo Delete_Connection
            
Name_Connection:
    
            With .WorkbookConnection
                .RefreshWithRefreshAll = False
                .name = "ICE Data Refresh"
            End With
            
        End With
        
        On Error GoTo 0
        
        Erase Column_Filter
    
    Else
        ' Update Connection string
        With Data_Query
            .Connection = "TEXT;" & Weekly_ICE_URL
            .TextFileColumnDataTypes = Column_Filter
        End With
        
    End If
    
    On Error GoTo Failed_To_Refresh 'Recreate Query and try again exactly 1 more time
    
    With Data_Query
    
        .Refresh False
        
        On Error GoTo Aggregation_Failed
        
        Set Filtered_CLCTN = New Collection
        
        With Filtered_CLCTN
            For Y = 1 To 2
                Retrieve_Combined = Not Retrieve_Combined
                .Add Historical_Excel_Aggregation(ThisWorkbook, Retrieve_Combined, ICE_Contracts:=True, QueryTable_To_Filter:=Data_Query), IIf(Retrieve_Combined = True, "futures + options", "futures-only")
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
    
    Exit Function

Delete_Connection: 'Error handler is available when editing parameters for a new querytable and the connection name is already taken by a different query

    ThisWorkbook.Connections("ICE Data Refresh").Delete
        
    On Error GoTo 0
    
    Resume Name_Connection
    
Failed_To_Refresh:
        
    Data_Query.Delete
    
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
            .Range("Combined_Workbook").value = Not Retrieve_Futures_Only
            .Range("Last_Updated_CFTC").value = 0
        
            If ReturnReportType = "D" Then .Range("Last_Updated_ICE").value = 0
        
        End With
        
        Call New_Data_Query(Scheduled_Retrieval:=True, Overwrite_All_Data:=True)
        
        MsgBox "Conversion Complete"
        
    End Sub

    Public Sub New_CFTC_Data()
    
        Dim Current_Contracts As Collection, retrieveCombinedData As Boolean, File_Paths As New Collection, _
        wantedContractCode As String, invalidContractCode As Boolean, New_Data() As Variant, First_Calculated_Column As Byte
        
        Dim WS As Worksheet, Symbol_Row As Long, I As Byte, reportType As String
    
        Set Current_Contracts = ContractDetails
        
        reportType = ReturnReportType
        
        retrieveCombinedData = Range("Combined_Workbook").value
        
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
        
        New_Data = CFTC_API_Method(reportType, retrieveCombinedData, DateSerial(1970, 1, 1), False, wantedContractCode)
        
        New_Data = Filter_Market_Columns(False, True, False, reportType, True, New_Data, False)
        
        On Error GoTo 0
        
        wantedContractCode = New_Data(1, UBound(New_Data, 2))
        
        ReDim Preserve New_Data(LBound(New_Data, 1) To UBound(New_Data, 1), LBound(New_Data, 2) To UBound(New_Data, 2) + 1)
        
        With Range("Symbols_TBL")
        
            Symbol_Row = WorksheetFunction.Match(wantedContractCode, .columns(1), 0)
            
            If Symbol_Row <> 0 Then
                
                For I = 3 To 4
                    
                    If Not IsEmpty(.Cells(Symbol_Row, I)) Then
                        Call Retrieve_Tuesdays_CLose(New_Data, UBound(New_Data, 2), Array(.Cells(Symbol_Row, I), IIf(I = 3, True, False)), dates_in_column_1:=True, overwrite_all_prices:=True)
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
        
            Range("Symbols_TBL").ListObject.ListRows.Add.Range.value = Array(wantedContractCode, New_Data(UBound(New_Data, 1), 2), Empty, Empty)
            
            MsgBox "A new row has been added to the availbale symbols table. Please fill in the missing Symbol information if available."
        
        End If
        
        Re_Enable
    
    Exit Sub
No_Data_Retrieved_From_API:
    
    Re_Enable
    MsgBox "Data couldn't be retrieved from API"
    
    End Sub

#End If
