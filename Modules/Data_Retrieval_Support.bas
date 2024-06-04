Attribute VB_Name = "Data_Retrieval_Support"

Option Explicit

Public Sub Retrieve_Historical_Workbooks(ByRef Path_CLCTN As Collection, ByVal ICE_Contracts As Boolean, ByVal CFTC_Contracts As Boolean, _
                                               ByVal Mac_User As Boolean, _
                                               ByVal reportType As String, _
                                               ByVal downloadFuturesAndOptions As Boolean, _
                                            Optional ByVal CFTC_Start_Date As Date, _
                                            Optional ByVal CFTC_End_Date As Date, _
                                            Optional ByVal ICE_Start_Date As Date, _
                                            Optional ByVal ICE_End_Date As Date, _
                                            Optional ByVal Historical_Archive_Download As Boolean = False)
'===================================================================================================================
    'Purpose: Downloads CFTC .zip files.
    'Inputs: Path_CLTCN - Collection to store file paths to extracted CoT data.
    '        ICE_Contracts - True if ICE data should be downloaded.
    '        CFTC_Contracts - True if CFTC data should be downloaded.
    '        Mac_User - True if script is being run on a MAC.
    '        reportType - Type of report to download.
    '        downloadFuturesAndOptions - True if futures + options should be retrieved else futures only.
    '        CFTC_Start_Date - Min cftc date.
    '        CFTC_End_Date - Max cftc date.
    '        Historical_Archive_Download - If true then download all data available.
    'Outputs:
'===================================================================================================================
    Dim fileNameWithinZip As String, Path_Separator As String, AnnualOF_FilePath As String, Destination_Folder As String, zipFileNameAndPath As String, _
    fullFileName As String, multiYearFileExtractedFromZip As String, Partial_Url As String, URL As String, multiYearZipFileFullName As String, combinedOrFutures As String, Multi_Year_URL As String
    
    Dim Queried_Date As Long, Download_Year As Long, Final_Year As Long, multiYearName As String
    
    Const TXT As String = ".txt", ZIP As String = ".zip", CSV As String = ".csv", ID_String As String = "B.A.T"
    
    Const mainFolderName As String = "COT_Historical_MoshiM"
    
    On Error GoTo Failed_To_Download
    
    #If Not Mac Then
        
        Path_Separator = Application.PathSeparator
        
        Destination_Folder = Environ("TEMP") & Path_Separator & mainFolderName & Path_Separator & reportType & Path_Separator & IIf(downloadFuturesAndOptions = True, "Combined", "Futures Only")
        
        If Not FileOrFolderExists(Destination_Folder) Then
            
            '/c =execute the command and then exit
            
            Shell ("cmd /c mkdir """ & Destination_Folder & """")
            
            Do Until FileOrFolderExists(Destination_Folder)
                
            Loop
        End If
        
    #Else
        '/Users/rondebruin/Library/Containers/com.microsoft.Excel/Data

'        This setion is for if files are downloaded and stored on client computer.
'        As of May 2024 MAc users only need this sub for getting urls to ice data.
'        Path_Separator = "/"
'        Destination_Folder = BasicMacAvailablePathMac & Path_Separator & mainFolderName & Path_Separator & IIf(downloadFuturesAndOptions = True, "Combined", "Futures Only") 'Keep variable as an empty string.User will decide path
'        If Not FileOrFolderExists(Destination_Folder) Then
'            Call CreateRootDirectories(Destination_Folder)
'        End If
        
    #End If
    
    With Path_CLCTN
    
        #If Not Mac Then
        
            If CFTC_Contracts Then
            
                If Not downloadFuturesAndOptions Then  'IF Futures Only Workbook
                
                    combinedOrFutures = "_Futures_Only"
                    
                    Select Case reportType
                    
                        Case "L"
                        
                            fileNameWithinZip = "annual" & TXT
                            
                            Partial_Url = "https://www.cftc.gov/files/dea/history/deacot"
                            Multi_Year_URL = "https://www.cftc.gov/files/dea/history/deacot1986_2016" & ZIP
                            multiYearName = "FUT86_16"
                            
                        Case "D"
                        
                            fileNameWithinZip = "f_year" & TXT
                            Partial_Url = "https://www.cftc.gov/files/dea/history/fut_disagg_txt_"
                            Multi_Year_URL = "https://www.cftc.gov/files/dea/history/fut_disagg_txt_hist_2006_2016" & ZIP
                            multiYearName = "F_DisAgg06_16"
                            
                        Case "T"
                        
                            fileNameWithinZip = "FinFutYY" & TXT
                            Partial_Url = "https://www.cftc.gov/files/dea/history/fut_fin_txt_"
                            Multi_Year_URL = "https://www.cftc.gov/files/dea/history/fin_fut_txt_2006_2016" & ZIP
                            multiYearName = "F_TFF_2006_2016"
                            
                    End Select
                
                Else 'Combined Contracts
                
                    combinedOrFutures = "_Combined"
                    
                    Select Case reportType
                    
                        Case "L"
                        
                            fileNameWithinZip = "annualof.txt"
                            
                            Partial_Url = "https://www.cftc.gov/files/dea/history/deahistfo" 'TXT URL
                            Multi_Year_URL = "https://www.cftc.gov/files/dea/history/deahistfo_1995_2016" & ZIP
                            multiYearName = "Com95_16"
                            
                        Case "D"
                        
                            fileNameWithinZip = "c_year" & TXT
                            Partial_Url = "https://www.cftc.gov/files/dea/history/com_disagg_txt_"
                            'https://www.cftc.gov/files/dea/history/com_disagg_txt_hist_2006_2016.zip
                            Multi_Year_URL = "https://www.cftc.gov/files/dea/history/com_disagg_txt_hist_2006_2016" & ZIP
                            
                            multiYearName = "C_DisAgg06_16"
                            
                        Case "T"
                        
                            fileNameWithinZip = "FinComYY" & TXT
                            'https://www.cftc.gov/files/dea/history/com_fin_txt_2014.zip
                            Partial_Url = "https://www.cftc.gov/files/dea/history/com_fin_txt_"
                            Multi_Year_URL = "https://www.cftc.gov/files/dea/history/fin_com_txt_2006_2016" & ZIP
                            multiYearName = "C_TFF_2006_2016"
                            
                    End Select
                
                End If
                
                If Year(CFTC_Start_Date) <= 2016 Then 'All report types have a compiled file for data before 2016
                    Historical_Archive_Download = True
                    CFTC_Start_Date = DateSerial(2017, 1, 1) 'So we can start dates in 2017 instead
                End If
                
                multiYearZipFileFullName = Destination_Folder & Path_Separator & reportType & "_COT_MultiYear_Archive" & combinedOrFutures & ZIP
                
                AnnualOF_FilePath = Destination_Folder & Path_Separator & fileNameWithinZip
        
                Download_Year = Year(CFTC_Start_Date)
                
                Final_Year = Year(CFTC_End_Date)
                
                Queried_Date = CFTC_End_Date
                
                '-1 is for if historical archive download needs to be executed
                For Download_Year = Download_Year - 1 To Final_Year
                        
                    If Not Historical_Archive_Download Then 'if not doing a download where multi year files are needed ie 2006-2016
                    
                        If Download_Year = Year(CFTC_Start_Date) - 1 Then
                            GoTo Skip_Download_Loop 'if on first loop
                        Else
                            URL = Partial_Url & Download_Year & ZIP 'Declare URL of Zip file
                        End If
                        
                    ElseIf Historical_Archive_Download Then
                        URL = Multi_Year_URL
                    End If

                    If Historical_Archive_Download Then
                        fullFileName = Destination_Folder & Path_Separator & reportType & "_" & multiYearName & combinedOrFutures & TXT
                    ElseIf Final_Year = Download_Year Then
                        fullFileName = Destination_Folder & Path_Separator & reportType & "_Weekly_" & CLng(Queried_Date) & "_" & Download_Year & combinedOrFutures & TXT
                    Else
                        fullFileName = Destination_Folder & Path_Separator & reportType & "_" & Download_Year & combinedOrFutures & TXT
                    End If
                    
                    If Not FileOrFolderExists(fullFileName) Then   'If wanted workbook doesn't exist
                        
                        If Historical_Archive_Download Then
                            zipFileNameAndPath = multiYearZipFileFullName
                        Else
                            zipFileNameAndPath = Replace$(fullFileName, TXT, ZIP)
                        End If
                        
                        If Not FileOrFolderExists(zipFileNameAndPath) Then
                            #If Mac Then
                                Call DownloadFile(URL, zipFileNameAndPath)
                            #Else
                                Call Get_File(URL, zipFileNameAndPath)
                            #End If
                        End If

                        If Not Historical_Archive_Download Then
                        
                            If FileOrFolderExists(AnnualOF_FilePath) Then Kill AnnualOF_FilePath    'If file within Zip folder exists within file directory then kill it
                        
                            #If Mac Then
                                Call UnzipFile(zipFileNameAndPath, Destination_Folder, fileNameWithinZip)
                            #Else
                                Call entUnZip1File(zipFileNameAndPath, Destination_Folder, fileNameWithinZip) 'Unzip specified file
                            #End If
                            
                            Name AnnualOF_FilePath As fullFileName
                            
                        ElseIf Historical_Archive_Download Then
                        
                            multiYearFileExtractedFromZip = Destination_Folder & Path_Separator & multiYearName & TXT
                            
                            If FileOrFolderExists(multiYearFileExtractedFromZip) Then Kill multiYearFileExtractedFromZip
    
                            #If Mac Then
                                Call UnzipFile(zipFileNameAndPath, Destination_Folder, multiYearName & TXT)
                            #Else
                                Call entUnZip1File(zipFileNameAndPath, Destination_Folder, multiYearName & TXT) 'Unzip specified file
                            #End If
                            
                            Name multiYearFileExtractedFromZip As fullFileName
                            
                        End If
                            
                    End If
                    
                    .Add fullFileName, fullFileName
        
Skip_Download_Loop:
                    Historical_Archive_Download = False
        
                Next Download_Year
                
            End If
        
        #End If
        
        If ICE_Contracts Then
            
            If Year(ICE_Start_Date) < 2011 Then
                ICE_Start_Date = DateSerial(2011, 1, 1)
            End If
            
            Download_Year = Year(ICE_Start_Date)
            Final_Year = Year(ICE_End_Date)
            
            Queried_Date = ICE_End_Date
            
            For Download_Year = Download_Year To Final_Year
            
                URL = "https://www.theice.com/publicdocs/futures/COTHist" & Download_Year & ".csv"
                
                Select Case Download_Year
                    Case Final_Year
                        fullFileName = Destination_Folder & Path_Separator & "ICE_Weekly_" & Queried_Date & "_" & Download_Year & ".csv"
                    Case Else
                        fullFileName = Destination_Folder & Path_Separator & "ICE_" & Download_Year & ".csv"
                End Select
                
                .Add URL, URL
    
            Next Download_Year
            
        End If
        
    End With
    
    Exit Sub
    
Failed_To_Download:
    Call PropagateError(Err, "Retrieve_Historical_Workbooks()", "Check your internet connection. If error persists, contact me at MoshiM_UC@outlook.com .")
End Sub
Public Function IsWorkbookOutdated(Optional workbookPath As String) As Boolean

'===================================================================================================================
    'Purpose: Tests if a given file has been updated with the most recent data available..
    'Inputs: workbookPath - File path  of file to test.
    'Outputs: True if data doesn't need updating; else, false.
'===================================================================================================================
    Dim Last_Release As Date

    On Error GoTo Default_True
    
    Last_Release = CFTC_Release_Dates(True) 'Returns Local date and time for the most recent release
    
    If LenB(workbookPath) > 0 And CDbl(Last_Release) <> 0 Then
        IsWorkbookOutdated = (FileDateTime(workbookPath) < Last_Release)
    Else
       IsWorkbookOutdated = True
    End If
    
    Exit Function
    
Default_True:
    IsWorkbookOutdated = True
    
End Function

Public Function HTTP_Weekly_Data(previousUpdateDate As Date, reportType As String, retrieveCombinedData As Boolean, ByRef useApi As Boolean, ByRef columnMap As Collection, Optional suppressMessages As Boolean = False, _
                                Optional testAllMethods As Boolean = False) As Variant
'===================================================================================================================
    'Purpose: Uses multiple methods of data retrieval from the CFTC.
    'Inputs: previousUpdateDate - Date converted to long for which data was last updated to.
    '        reportType - One of L,D,T to represent what type of report to retrieve.
    '        retrieveCombinedData - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    '        useApi - If true then the function will attempt to retrieve data via API.
    '        suppressMessages - true if error messages should be repressed.
    '        columnMap - An empty collection that wil store FieldInfo instances for each column found within the output.
    'Outputs: An array of weekly data if ap method fails; else, all data since last_update.
'===================================================================================================================
    Dim PowerQuery_Available As Boolean, Power_Query_Failed As Boolean, _
    Text_Method_Failed As Boolean, Query_Table_Method_Failed As Boolean, _
    MAC_OS As Boolean, dataRetrieved As Boolean, successCount As Byte, tempData() As Variant, attemptCount As Byte
    
    Dim retrievalTimer As TimedTask, savedState As Boolean
        
    Const PowerQTask As String = "Power Query Retrieval", _
    QueryTask As String = "QueryTable Retrieval", HTTPTask As String = "HTTP Retrieval", _
    ApiTask = "Socrata API", ProcedureName = "HTTP_Weekly_Data()"
    
    #If Mac Then
        MAC_OS = True
        PowerQuery_Available = False 'Use standalone QueryTable rather than QueryTable wrapped in listobject
    #Else
        On Error GoTo Default_No_Power_Query
        
        If Val(Application.Version) < 16 Then 'IF excel version is prior to Excel 2016 then
            PowerQuery_Available = IsPowerQueryAvailable 'Check if Power Query is available
        Else
            PowerQuery_Available = True
        End If
    #End If
    
Retrieval_Process:

    If testAllMethods Then
        Set retrievalTimer = New TimedTask
        retrievalTimer.Start "Time Retrieval Methods."
    End If
    
    savedState = ThisWorkbook.Saved
    
    If useApi Then
        
        If testAllMethods Then
            If MsgBox("Test Socrata API Method", vbYesNo) <> vbYes Then GoTo QueryTable_Method
            attemptCount = attemptCount + 1
            retrievalTimer.StartSubTask ApiTask
        End If

        On Error GoTo Catch_SocrataRetrievalFailed

        If TryGetCftcWithSocrata(tempData, reportType, retrieveCombinedData, testAllMethods, columnMap, mostRecentStoredDate:=previousUpdateDate) Then
            
            On Error GoTo 0
            If IsArrayAllocated(tempData) Then
                HTTP_Weekly_Data = tempData
                Erase tempData
            Else
                Err.Raise Data_Retrieval.ERROR_SOCRATA_SUCCESS_NO_DATA, ProcedureName, "No new data could be retrieved from Socrata's API."
            End If
            
            dataRetrieved = True
        End If
        
        If testAllMethods Then
            retrievalTimer.StopSubTask ApiTask
            successCount = successCount + 1
        End If
        
    End If
    
QueryTable_Method:

    If dataRetrieved = False Or testAllMethods Then
        
        If testAllMethods Then
            If MsgBox("Test Querytable Method", vbYesNo) <> vbYes Then GoTo PowerQuery_Method
            attemptCount = attemptCount + 1
            retrievalTimer.StartSubTask QueryTask
        End If
        
        On Error GoTo QueryTable_Failed
            
        HTTP_Weekly_Data = CFTC_Data_QueryTable_Method(reportType:=reportType, retrieveCombinedData:=retrieveCombinedData)
        
        If testAllMethods Then
            retrievalTimer.StopSubTask QueryTask
            successCount = successCount + 1
        End If
        
        dataRetrieved = True
        
    End If
    
PowerQuery_Method:

    If Not MAC_OS Then
    
        If (Not dataRetrieved And PowerQuery_Available) Or testAllMethods Then
        
            If testAllMethods Then
                If MsgBox("Test PowerQuery Method", vbYesNo) <> vbYes Then GoTo TXT_Method
                attemptCount = attemptCount + 1
                retrievalTimer.StartSubTask PowerQTask
            End If
            
            On Error GoTo PowerQuery_Failed
                
            HTTP_Weekly_Data = CFTC_Data_PowerQuery_Method(reportType, retrieveCombinedData)
                
            If testAllMethods Then
                retrievalTimer.StopSubTask PowerQTask
                successCount = successCount + 1
            End If
        
            dataRetrieved = True
        
        End If
        
TXT_Method:
    
        If testAllMethods Or Not dataRetrieved Then     'TXT file Method
            
            If testAllMethods Then
                If MsgBox("Test Txt Method", vbYesNo) <> vbYes Then GoTo Finally
                attemptCount = attemptCount + 1
                retrievalTimer.StartSubTask HTTPTask
            End If
            
            On Error GoTo TXT_Failed
                
            HTTP_Weekly_Data = CFTC_Data_Text_Method(previousUpdateDate, reportType:=reportType, retrieveCombinedData:=retrieveCombinedData)
                
            If testAllMethods Then
                retrievalTimer.StopSubTask HTTPTask
                successCount = successCount + 1
            End If
        
            dataRetrieved = True
            
        End If
    
    End If
                                                                                                                      
Finally:
    
    On Error GoTo 0
    
    ThisWorkbook.Saved = savedState
    
    If testAllMethods Then retrievalTimer.DPrint
    
    If dataRetrieved And columnMap Is Nothing And Not useApi Then
        Set columnMap = GetExpectedLocalFieldInfo(reportType, False, False, False, False)
    End If
    
    If Not dataRetrieved Or (testAllMethods And successCount <> attemptCount) Then
        Err.Raise Data_Retrieval.ERROR_RETRIEVAL_FAILED, "HTTP_Weekly_Data", "All retrieval methods have failed."
    End If
    
    Exit Function

PowerQuery_Failed:

    If testAllMethods Then
        DisplayErrorIfAvailable Err, ProcedureName
        retrievalTimer.StopSubTask PowerQTask
    End If
    
    Resume TXT_Method
    
TXT_Failed:

    If testAllMethods Then
        DisplayErrorIfAvailable Err, ProcedureName
        retrievalTimer.StopSubTask HTTPTask
    End If
    
    Resume Finally
    
QueryTable_Failed:

    If testAllMethods Then
        DisplayErrorIfAvailable Err, ProcedureName
        retrievalTimer.StopSubTask QueryTask
    End If
    
    If Not MAC_OS Then
        Resume PowerQuery_Method
    Else
        Resume Finally
    End If
    
Default_No_Power_Query:

    PowerQuery_Available = False
    Resume Retrieval_Process

Catch_SocrataRetrievalFailed:

    If testAllMethods Then
        DisplayErrorIfAvailable Err, ProcedureName
        retrievalTimer.StopSubTask ApiTask
    End If
    
    useApi = False
    Resume QueryTable_Method

End Function

Public Function TryGetCftcWithSocrata(outputA() As Variant, reportType As String, getFuturesAndOptions As Boolean, _
        debugModeActive As Boolean, ByRef fieldInfoByEditedName As Collection, _
        Optional contractCode As String = vbNullString, Optional ByVal mostRecentStoredDate As Date) As Boolean
'===================================================================================================================
    'Purpose: Retrieve data from the CFTC's Public Reporting Environment.
    'Inputs:
    '        outputA - Array that will store retrieved data if successfull.
    '        mostRecentStoredDate - Date which data was last updated to.
    '        reportType - One of L,D,T to represent what type of report to retrieve.
    '        getFuturesAndOptionsData - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    '        contractCode - If supplied with a value than only data that with this contract code will be retrieved.
    '        fieldInfoByEditedName - Empty Collection that will store information for wanted fields.
    'Output: True if data was successfully retrieved.
'===================================================================================================================

    Dim tempDataCLCTN As Collection, apiCode As String, reportKey As String, _
    apiURL As String, dataFilters As String, queryReturnLimit As Long, dataQuery As QueryTable, _
    socrataData() As Variant, iCount As Long, apiColumnNames() As Variant, _
    wantedFieldsA() As Variant, imperfectOperator As String
    
    Const apiBaseURL As String = "https://publicreporting.cftc.gov/resource/"
    
    On Error GoTo Catch_GeneralError
    
    If mostRecentStoredDate = TimeSerial(0, 0, 0) Then
        mostRecentStoredDate = DateSerial(1970, 1, 1)
    End If
    
    If mostRecentStoredDate = CDate(0) Then mostRecentStoredDate = DateSerial(1970, 1, 1)
    
    If LenB(contractCode) > 0 Then contractCode = " AND cftc_contract_market_code='" & contractCode & "'"
    
    ' General purpose array that will work for all Report types. Unneeded values will be discarded.
    Dim columnTypes(1 To 200) As XlColumnDataType, codeColumn As Byte, dateColumn As Byte
    
    dateColumn = 3: codeColumn = 6
    ' The query table needs to import contract codes as text.
    For iCount = LBound(columnTypes) To UBound(columnTypes)
        
        Select Case iCount
            Case 1, 4, 5, 8, 10
                columnTypes(iCount) = xlSkipColumn
            Case dateColumn, codeColumn
                columnTypes(iCount) = xlTextFormat
            Case Else
                columnTypes(iCount) = xlGeneralFormat
        End Select
        
    Next iCount
    ' Creates a collection of api codes keyed to their report type.
    Set tempDataCLCTN = New Collection
    With tempDataCLCTN
        For iCount = 0 To 2
            reportKey = Array("L", "D", "T")(iCount)
            If getFuturesAndOptions Then
                .Add Array("jun7-fc8e", "kh3c-gbw2", "yw9f-hn96")(iCount), reportKey
            Else
                .Add Array("6dca-aqww", "72hh-3qpy", "gpe5-46if")(iCount), reportKey
            End If
        Next iCount
    End With
    
    apiCode = tempDataCLCTN(reportType)
    Set tempDataCLCTN = Nothing
    
    queryReturnLimit = IIf(debugModeActive, 400, 40000)
    imperfectOperator = IIf(debugModeActive, ">=", ">")
    
    dataFilters = "?$where=report_date_as_yyyy_mm_dd" & imperfectOperator & "'" & Format(mostRecentStoredDate, "yyyy-mm-dd") & "T00:00:00.000'" & _
                        contractCode & "&$order=report_date_as_yyyy_mm_dd" & "&$limit=" & queryReturnLimit
                        
    apiURL = apiBaseURL & apiCode & ".csv" & dataFilters
    
    Set dataQuery = QueryT.QueryTables.Add(Connection:="TEXT;" & apiURL, Destination:=QueryT.Range("A1"))
        
    With dataQuery
        
        .BackgroundQuery = False
        .SaveData = False
        .AdjustColumnWidth = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileCommaDelimiter = True
        .TextFileColumnDataTypes = columnTypes
        Erase columnTypes
Name_Connection:
        
        .WorkbookConnection.RefreshWithRefreshAll = False
        ' Loop until the API doesn't return anything.
        iCount = 0
        Do
            iCount = iCount + 1
            
            If iCount > 1 Then
                .Connection = "TEXT;" & apiURL & "&$offset=" & queryReturnLimit * (iCount - 1)
                .TextFileCommaDelimiter = True
            End If
            
            On Error GoTo Catch_RetrievalFailed
            
            Application.StatusBar = "Retrieveing set number [ " & iCount & " ] for Report : " & reportType & " Combined data: " & getFuturesAndOptions
            .Refresh False
            Application.StatusBar = vbNullString
            On Error GoTo Catch_GeneralError
            
            With .ResultRange
            
                .Replace ".", Empty, xlWhole
                ' >1 since column names will always be returned.
                If .Rows.count > 1 Then
                
                    If iCount = 1 Then apiColumnNames = Application.Transpose(Application.Transpose(.Rows(1).Value2))

                    If tempDataCLCTN Is Nothing Then Set tempDataCLCTN = New Collection
                    
                    tempDataCLCTN.Add .Range(.Cells(2, 1), .Cells(.Rows.count, .columns.count)).Value2
                    
                End If
                
            End With
            ' If debugModeActive Then Exit Do
        Loop While .ResultRange.Rows.count = queryReturnLimit + 1 And debugModeActive = False
        
        .ResultRange.ClearContents
        .WorkbookConnection.Delete
        .Delete
        Set dataQuery = Nothing
    End With
    
    On Error GoTo Catch_GeneralError
    
    If Not tempDataCLCTN Is Nothing Then
    
        Select Case tempDataCLCTN.count
            Case 1
                socrataData = tempDataCLCTN(1)
            Case Is > 1
                socrataData = CombineArraysInCollection(tempDataCLCTN, Append_Type.Multiple_2d)
        End Select
        
        Set tempDataCLCTN = Nothing
        
        If IsArrayAllocated(socrataData) Then
            
            Dim basicField As FieldInfo, columnInOutput As Byte, columnInApiData As Byte
                        
            wantedFieldsA = Application.Transpose(GetAvailableFieldsTable(reportType).DataBodyRange.columns(1).Value2)
            
            Set fieldInfoByEditedName = CreateFieldInfoMap(apiColumnNames, wantedFieldsA, externalHeadersFromSocrataAPI:=True)
            
            Erase apiColumnNames: Erase wantedFieldsA

            With fieldInfoByEditedName
                ReDim outputA(1 To UBound(socrataData, 1), 1 To .count)
                codeColumn = .Item("cftc_contract_market_code").ColumnIndex
                dateColumn = .Item("report_date_as_yyyy_mm_dd").ColumnIndex
                'cftcRegionCodeColumn = .Item("cftc_region_code").ColumnIndex
            End With
            
            For Each basicField In fieldInfoByEditedName 'GetExpectedLocalFieldInfo(reportType, False, False, False, False)
                
                columnInOutput = columnInOutput + 1
                'On Error GoTo Catch_WantedColumnNotFound
                
                With basicField 'fieldInfoByEditedName(basicField.EditedName)
                    On Error GoTo Catch_GeneralError
                    
                    If Not .IsMissing Then
                        columnInApiData = .ColumnIndex
                        For iCount = LBound(socrataData, 1) To UBound(socrataData, 1)
                            Select Case columnInApiData
                                Case codeColumn
                                    ' Ensure that it was imported as a string.
                                    If Not VarType(socrataData(iCount, columnInApiData)) = 8 Then
                                        outputA(iCount, columnInOutput) = CStr(socrataData(iCount, columnInApiData))
                                    Else
                                        outputA(iCount, columnInOutput) = socrataData(iCount, columnInApiData)
                                    End If
                                Case dateColumn
                                    outputA(iCount, columnInOutput) = CDate(Left$(socrataData(iCount, columnInApiData), 10))
                                Case Else
                                    outputA(iCount, columnInOutput) = socrataData(iCount, columnInApiData)
                            End Select
                        Next iCount
                    End If
                    ' The field reflects column within the api data. Adjust it to match column in outputA.
                    If columnInApiData <> columnInOutput Then .AdjustColumnIndex columnInOutput
                End With
Next_Field:
            Next basicField
                    
        End If
        
    End If
    
    TryGetCftcWithSocrata = True
    
    Exit Function
    
Catch_RetrievalFailed:
    
    If Not dataQuery Is Nothing Then
        With dataQuery
            .WorkbookConnection.Delete
            .Delete
        End With
    End If
    
    Application.StatusBar = vbNullString
    Call PropagateError(Err, "TryGetCftcWithSocrata()", "An error occurred while attempting to connect to the Socrata API for [ " & reportType & " ] getFuturesAndOptions=" & getFuturesAndOptions & ".")
    
Catch_GeneralError:
    Erase outputA
    Call PropagateError(Err, "TryGetCftcWithSocrata()")

Catch_WantedColumnNotFound:
    Resume Next_Field
End Function

Public Function CFTC_Data_PowerQuery_Method(reportType As String, retrieveCombinedData As Boolean) As Variant()
'===================================================================================================================
    'Purpose: Retrieves the latest Weekly data with Power Query.
    'Inputs: reportType - One of L,D,T to represent what type of report to retrieve.
    '        retrieveCombinedData - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    'Outputs: An array of the most recent weekly CFTC data.
    'Notes: Use only on Windows.
'===================================================================================================================
    Dim URL As String, Formula_AR() As String, quotation As String, Y As Byte, table_name As String
    
    On Error GoTo Failure
    
    table_name = Evaluate("=VLOOKUP(""" & reportType & """,Report_Abbreviation,2,FALSE)")
    
    quotation = Chr(34)
    
    URL = "https://www.cftc.gov/dea/newcot/"
    
    Y = Application.Match(reportType, Array("L", "D", "T"), 0) - 1
    
    If Not retrieveCombinedData Then 'Futures Only
        URL = URL & Array("deafut.txt", "f_disagg.txt", "FinFutWk.txt")(Y)
    Else
        URL = URL & Array("deacom.txt", "c_disagg.txt", "FinComWk.txt")(Y)
    End If
    
    With ThisWorkbook
        'Change Query URL
        Formula_AR = Split(.Queries(table_name).Formula, quotation, 3) 'Split with quotation mark
        Formula_AR(1) = URL
        .Queries(table_name).Formula = Join(Formula_AR, quotation)
    End With

    With Weekly.ListObjects(table_name)
        .QueryTable.Refresh False                               'Refresh Weekly Data Table
        CFTC_Data_PowerQuery_Method = .DataBodyRange.Value2     'Store contents of table in an array
    End With
    
    Exit Function
Failure:
    PropagateError Err, "CFTC_Data_PowerQuery_Method()"
End Function
Public Function CFTC_Data_Text_Method(Last_Update As Date, reportType As String, retrieveCombinedData As Boolean) As Variant()
'===================================================================================================================
    'Purpose: Retrieves the latest Weekly using HTTP methods found on the Windows version of Excel.
    'Inputs: reportType - One of L,D,T to represent what type of report to retrieve.
    '        retrieveCombinedData - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    '        Last_Update - Date that data was last retrieved for.
    'Outputs: An array of the most recent weekly CFTC data.
    'Notes: Use only on Windows.
'===================================================================================================================
    Dim File_Path As New Collection, URL As String, Y As Byte
    
    On Error GoTo Failure
    URL = "https://www.cftc.gov/dea/newcot/"
    
    Y = Application.Match(reportType, Array("L", "D", "T"), 0) - 1
    
    If Not retrieveCombinedData Then 'Futures Only
        URL = URL & Array("deafut.txt", "f_disagg.txt", "FinFutWk.txt")(Y)
    Else
        URL = URL & Array("deacom.txt", "c_disagg.txt", "FinComWk.txt")(Y)
    End If
    
    With File_Path
    
        .Add Environ("TEMP") & "\" & Date & "_" & reportType & "_Weekly.txt", "Weekly Text File" 'Add file path of file to be downloaded
    
        Call Get_File(URL, .Item("Weekly Text File")) 'Download the file to the above path
        
    End With
    
    CFTC_Data_Text_Method = Historical_Parse(File_Path, retrieveCombinedData:=retrieveCombinedData, CFTC_TXT:=True, reportType:=reportType, After_This_Date:=Last_Update) 'return array
    Exit Function
Failure:
    PropagateError Err, "CFTC_Data_Text_Method()"
End Function
Public Function CFTC_Data_QueryTable_Method(reportType As String, retrieveCombinedData As Boolean) As Variant()
'===================================================================================================================
    'Purpose: Retrieves the latest Weekly data with Power Query.
    'Inputs: reportType - One of L,D,T to represent what type of report to retrieve.
    '        retrieveCombinedData - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    'Outputs: An array of the most recent weekly CFTC data.
    'Notes: Use only on Windows.
'===================================================================================================================
    Dim Data_Query As QueryTable, data() As Variant, URL As String, _
     Y As Byte, reEnableEventsOnExit As Boolean, _
    Found_Data_Query As Boolean, Error_While_Refreshing As Boolean
    
    Dim Workbook_Type As String
    
    With Application
    
        reEnableEventsOnExit = .EnableEvents
        
        .EnableEvents = False
        .DisplayAlerts = False
        
    End With
    
    Workbook_Type = IIf(retrieveCombinedData, "Combined", "Futures_Only")
    
    For Each Data_Query In QueryT.QueryTables
        If InStrB(1, Data_Query.Name, reportType & "_CFTC_Data_Weekly_" & Workbook_Type) > 0 Then
            Found_Data_Query = True
            Exit For
        End If
    Next Data_Query
    
    If Not Found_Data_Query Then 'If QueryTable isn't found then create it
    
Recreate_Query:
        
        URL = "https://www.cftc.gov/dea/newcot/"
        
        Y = Application.Match(reportType, Array("L", "D", "T"), 0) - 1
        
        If Not retrieveCombinedData Then 'Futures Only
            URL = URL & Array("deafut.txt", "f_disagg.txt", "FinFutWk.txt")(Y)
        Else
            URL = URL & Array("deacom.txt", "c_disagg.txt", "FinComWk.txt")(Y)
        End If
        
        Set Data_Query = QueryT.QueryTables.Add(Connection:="TEXT;" & URL, Destination:=QueryT.Range("A1"))
        
        With Data_Query
            
            .BackgroundQuery = False
            .SaveData = False
            .AdjustColumnWidth = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlOverwriteCells
            
            .TextFileColumnDataTypes = Filter_Market_Columns(convert_skip_col_to_general:=True, reportType:=reportType, Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True)
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileCommaDelimiter = True
            
            .Name = reportType & "_CFTC_Data_Weekly_" & Workbook_Type
            
            On Error GoTo Delete_Connection
            
Name_Connection:
    
            With .WorkbookConnection
                .RefreshWithRefreshAll = False
                .Name = reportType & "_Weekly CFTC Data: " & Workbook_Type
            End With
            
        End With
        
        On Error GoTo 0
    
    End If
    
    On Error GoTo Failed_To_Refresh 'Recreate Query and try again exactly 1 more time
    
    With Data_Query
    
        .Refresh False
        
        With .ResultRange
            .Replace ".", Null, xlWhole
            CFTC_Data_QueryTable_Method = .value 'Store Data in an Array
            .ClearContents 'Clear the Range
        End With
        
    End With
    
    With Application
        .DisplayAlerts = True
        .EnableEvents = reEnableEventsOnExit
    End With
    
    Exit Function

Delete_Connection: 'Error handler is available when editing parameters for a new querytable and the connection name is already taken by a different query

    ThisWorkbook.Connections("Weekly CFTC Data: " & Workbook_Type).Delete
        
    On Error GoTo 0
    
    Resume Name_Connection
    
Failed_To_Refresh:
        
    If Not Data_Query Is Nothing Then
        With Data_Query
            .WorkbookConnection.Delete
            .Delete
        End With
    End If
    
    If Error_While_Refreshing = True Then
        PropagateError Err, "CFTC_Data_QueryTable_Method()"
    Else
        Error_While_Refreshing = True
        Resume Recreate_Query
    End If
    
End Function
Public Function Historical_Parse(ByVal File_CLCTN As Collection, reportType As String, retrieveCombinedData As Boolean, _
                                  Optional ByRef contractCode As String = vbNullString, _
                                  Optional After_This_Date As Date = 0, _
                                  Optional Kill_Previous_Workbook As Boolean = False, _
                                  Optional parsingMultipleWeeks As Boolean, _
                                  Optional Weekly_ICE_Data As Boolean, _
                                  Optional CFTC_TXT As Boolean, _
                                  Optional Parse_All_Data As Boolean) As Variant()
'===================================================================================================================
    'Purpose: Retrieves data from Excel Workbooks.
    'Inputs: reportType - One of L,D,T to represent what type of report to retrieve.
    '        retrieveCombinedData - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    '        File_CLCTN - Collection of file paths.
    '        contract_code - If given a value, then Excel workbooks will be filtered for a specific contract code.
    '        After_This_Date - Data after this date will be retrieved.
    '        Kill_Previous_Workbook - If a previous workbook exists then delete it.
    '        parsingMultipleWeeks - Not ALL data may have been downloaded. Maybe only specific years.
    '        Specified_Contract - True if a single contract is wanted
    '        Weekly_ICE_Data -
    '        CFTC_TXT -
    '        Parse_All_Data -
    'Outputs:
    'Notes:
'===================================================================================================================
    Dim Contract_WB As Workbook, Contract_WB_Path As String, ICE_Data As Boolean
    
    Dim OS_BasedPathSeparator As String, filterForSpecificContract As Boolean
    
    On Error GoTo Historical_Parse_General_Error_Handle
    
    filterForSpecificContract = LenB(contractCode) > 0
    
    #If Mac Then
        OS_BasedPathSeparator = "/"
    #Else
        OS_BasedPathSeparator = "\"
    #End If
    
    Application.ScreenUpdating = False

    Select Case True
        'Parse all data is for when all data is being downloaded
        Case parsingMultipleWeeks, filterForSpecificContract, Parse_All_Data
    
            Contract_WB_Path = Left$(File_CLCTN(1), InStrRev(File_CLCTN(1), OS_BasedPathSeparator))
        
            If parsingMultipleWeeks Then
            
                Contract_WB_Path = Contract_WB_Path & reportType & "_COT_Yearly_Contracts_" & IIf(retrieveCombinedData, "Combined", "Futures_Only") & ".xlsb"
            
            ElseIf filterForSpecificContract Or Parse_All_Data Then  'If using the new contract macro
            
                Contract_WB_Path = Contract_WB_Path & reportType & "_COT_Historical_Archive_" & IIf(retrieveCombinedData, "Combined", "Futures_Only") & ".xlsb"
                
            End If
        
            If Not FileOrFolderExists(Contract_WB_Path) Then
                ' Compile text files into a single document.
                Set Contract_WB = Historical_TXT_Compilation(File_CLCTN, reportType:=reportType, Saved_Workbook_Path:=Contract_WB_Path, OnMAC:=False, parsingFuturesAndOptions:=retrieveCombinedData)
            ElseIf IsWorkbookOutdated(Contract_WB_Path) Or parsingMultipleWeeks Or Kill_Previous_Workbook = True Then
                On Error Resume Next
                Kill Contract_WB_Path
                On Error GoTo 0
                Set Contract_WB = Historical_TXT_Compilation(File_CLCTN, reportType:=reportType, Saved_Workbook_Path:=Contract_WB_Path, OnMAC:=False, parsingFuturesAndOptions:=retrieveCombinedData)
            Else
                Set Contract_WB = Workbooks.Open(Contract_WB_Path)
                Contract_WB.Windows(1).Visible = False
            End If
            
            Historical_Parse = Historical_Excel_Aggregation(Contract_WB, getFuturesAndOptions:=retrieveCombinedData, contractCodeToFilterFor:=contractCode, Date_Input:=After_This_Date, ICE_Contracts:=False)
            
            Contract_WB.Close False 'Close without saving
            
        Case Weekly_ICE_Data
            
            Set Contract_WB = Workbooks.Open(File_CLCTN.Item("ICE"))
            
            With Contract_WB
            
                .Windows(1).Visible = False
                Historical_Parse = Historical_Excel_Aggregation(Contract_WB, getFuturesAndOptions:=retrieveCombinedData, Date_Input:=After_This_Date, ICE_Contracts:=True)
                .Close False
                
                If retrieveCombinedData = False Then Kill File_CLCTN.Item("ICE")
                
            End With
            
        Case CFTC_TXT 'Result=2D Array stored in Collection2D Array(s) stored in Collection from .txt file(s)
            Historical_Parse = Weekly_Text_File(File_CLCTN, reportType:=reportType, retrieveCombinedData:=retrieveCombinedData)
            
    End Select
    
    Application.StatusBar = vbNullString
    
    Exit Function

Historical_Parse_General_Error_Handle:
    Call PropagateError(Err, "Historical_Parse()")
End Function
Public Function Historical_TXT_Compilation(File_Collection As Collection, _
Saved_Workbook_Path As String, OnMAC As Boolean, reportType As String, parsingFuturesAndOptions As Boolean) As Workbook
    
    Dim File_TXT As Variant, fileNumber As Long, Data_STR As String, File_Path() As String, newWorkbook As Workbook
    
    Dim InfoF() As Variant, columnFormatTypesA() As Variant, D As Long, ICE_Filter As Boolean, ICE_Count As Byte, OS_BasedPathSeparator As String
    
    Dim File_Name As String, CFTC_Count As Byte, file_text As String, outputFileNumber As Long, outputFileName As String 'g ', DD As Double
    
    Const COMMA As String = ","
    
    On Error GoTo Query_Table_Method_For_TXT_Retrieval
        
    If OnMAC Then
        OS_BasedPathSeparator = "/"
    Else
        OS_BasedPathSeparator = "\"
    End If
    
    outputFileNumber = FreeFile
    outputFileName = Left$(File_Collection(1), InStrRev(File_Collection(1), OS_BasedPathSeparator)) & "Historic.txt"
    
    If FileOrFolderExists(outputFileName) Then Kill outputFileName
    
    Open outputFileName For Append As #outputFileNumber 'Write contents of string to text File
    
    fileNumber = FreeFile
    'Open each file in the collection and write their contents to a string.
    For Each File_TXT In File_Collection
    
        Application.StatusBar = "Parsing " & File_TXT
        DoEvents
        
        Open File_TXT For Input As fileNumber
            
            File_Name = Right$(File_TXT, Len(File_TXT) - InStrRev(File_TXT, OS_BasedPathSeparator))
            
            If File_Name Like "*ICE*" Then
                D = 0
                ICE_Count = ICE_Count + 1
                Do Until EOF(fileNumber)
                    D = D + 1
                    Line Input #fileNumber, Data_STR
                    
                    If (D > 1 And ICE_Count > 1) Or ICE_Count = 1 Then
                        'Only allow printing of headers if on first file
                        Print #outputFileNumber, Data_STR
                    End If
                Loop
            Else
                CFTC_Count = CFTC_Count + 1
                D = 0
                Do Until EOF(fileNumber)
                    D = D + 1
                    Line Input #fileNumber, Data_STR
                    
                    If (D > 1 And CFTC_Count > 1) Or CFTC_Count = 1 Then
                        'Only allow printing of headers if on first file
                        Print #outputFileNumber, Data_STR
                    End If
                Loop
            End If
            
        Close fileNumber
        
        'If LCase$(File_TXT) Like "*weekly*" Then Kill File_TXT
        
    Next File_TXT

    Close #outputFileNumber
    
    Application.StatusBar = "TXT file compilation was successful. Creating Workbook."
    DoEvents
    
    columnFormatTypesA = Filter_Market_Columns(convert_skip_col_to_general:=True, reportType:=reportType, Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True, ICE:=False)

    ReDim InfoF(1 To UBound(columnFormatTypesA, 1))
    
    For D = 1 To UBound(columnFormatTypesA, 1) 'Fill in column numbers for use when supplying column filters to OpenTxt
        InfoF(D) = Array(D, columnFormatTypesA(D))
    Next D
    
    Erase columnFormatTypesA
    On Error GoTo Query_Table_Method_For_TXT_Retrieval
    
    #If Mac Then
        D = xlMacintosh
    #Else
        D = xlWindows
    #End If
    With Workbooks
    
        .OpenText fileName:=outputFileName, origin:=D, startRow:=1, DataType:=xlDelimited, _
                                    TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, COMMA:=True, _
                                    FieldInfo:=InfoF, DecimalSeparator:=".", ThousandsSeparator:=",", TrailingMinusNumbers:=False, _
                                    Local:=False
        Set newWorkbook = Workbooks(.count)

    End With
    
   With newWorkbook
        .Windows(1).Visible = False
        On Error Resume Next
        If Not OnMAC Then
            newWorkbook.SaveAs Saved_Workbook_Path, FileFormat:=xlExcel12
        End If
        On Error GoTo 0
    End With
    
    Set Historical_TXT_Compilation = newWorkbook
    Exit Function
Query_Table_Method_For_TXT_Retrieval:
    
    On Error GoTo Parent_Handler

    InfoF = Query_Text_Files(File_Collection, combined_wb:=parsingFuturesAndOptions, reportType:=reportType)
    
    Application.StatusBar = "Data compilation was successful. Creating Workbook."
    DoEvents
    
    Set newWorkbook = Workbooks.Add
    
    With newWorkbook
    
        .Windows(1).Visible = False
        
        With .Worksheets(1)
            .DisplayPageBreaks = False
            .columns("C:C").NumberFormat = "@" 'Format as text
            .Range("A1").Resize(UBound(InfoF, 1), UBound(InfoF, 2)).Value2 = InfoF
        End With
        
    End With
    Set Historical_TXT_Compilation = newWorkbook
    Exit Function
    
Parent_Handler:
    Call PropagateError(Err, "Historical_TXT_Compilation", "An error occurred while compiling text files.")
End Function
Public Function Historical_Excel_Aggregation(Contract_WB As Workbook, _
                                        getFuturesAndOptions As Boolean, _
                                        Optional contractCodeToFilterFor As String = vbNullString, _
                                        Optional Date_Input As Date = 0, _
                                        Optional ICE_Contracts As Boolean = False, _
                                        Optional Weekly_CFTC_TXT As Boolean = False, Optional QueryTable_To_Filter As QueryTable) As Variant()
'===================================================================================================================
    'Purpose: Filters and sorts data on a worksheet.
    'Inputs: Contract_WB - Workbook that contains workbook.
    '        contractCodeToFilterFor - If given a value then data will be filtered for this contract code.
    '        combined_workbook - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    '        Date_Input - If not 0 then all data > than this will be filtered for.
    '        filterForSpecificContract - True if specified contract is wanted.
    '        Weekly_CFTC_TXT - True if file data is from the cftc website. Note the url available text file.
    '        QueryTable_To_Filter - Data may be within a query table.
    'Outputs: An array.
    'Notes:
'===================================================================================================================
    Dim VAR_DTA() As Variant, Comparison_Operator As String, iRow As Long
    
    Dim Combined_CLMN As Byte, Disaggregated_Filter_STR As String 'Used if filtering ICE Contracts for Futures and Options
    
    Dim Filtering_QueryTable As Boolean, Source_RNG As Range, filterForSpecificContract As Boolean
    
    Const yymmdd_column As Byte = 2
    Const Contract_Code_CLMN As Byte = 4 'Column that holds Contract identifiers
    Const ICE_Contract_Code_CLMN As Byte = 7
    Const Date_Field As Byte = 3
    filterForSpecificContract = LenB(contractCodeToFilterFor) > 0
    On Error GoTo Finally
    
    Filtering_QueryTable = (Not QueryTable_To_Filter Is Nothing)
    
    If Not Filtering_QueryTable Then
        Application.StatusBar = "Filtering Data."
        DoEvents
        Set Source_RNG = Contract_WB.Worksheets(1).UsedRange
    Else
        Set Source_RNG = QueryTable_To_Filter.ResultRange
    End If
    
    If Source_RNG.Cells.count = 1 Then 'If worksheet is empty then display message
        GoTo Scripts_Failed_To_Collect_Data
    End If

    On Error GoTo Finally
    
    If ICE_Contracts Or Weekly_CFTC_TXT Then 'Weekly_CFTC_TXT should be unique to CFTC Weekly Text Files at the time of writing
        Comparison_Operator = ">="
    Else
        Comparison_Operator = ">"
    End If
    
    If ICE_Contracts Then
        Disaggregated_Filter_STR = IIf(getFuturesAndOptions, "*Combined*", "*FutOnly*")
        'Find column to be sorted based on the column header.
        On Error GoTo Catch_CombinedColumn_Not_Found
        Combined_CLMN = Application.Match("FutOnly_or_Combined", Source_RNG.Rows(1).Value2, 0)
        Comparison_Operator = Comparison_Operator & Format(IIf(Date_Input = 0, DateSerial(2000, 1, 1), Date_Input), "YYMMDD")
    Else
        Comparison_Operator = Comparison_Operator & CLng(Date_Input)
    End If
    
    On Error GoTo Finally
    
Check_If_Code_Exists:

    With Source_RNG
    
        On Error Resume Next
        .Parent.ShowAllData
        On Error GoTo Finally
        'Sort date column in ascending order.
        .Sort key1:=.Cells(2, IIf(ICE_Contracts = True, yymmdd_column, Date_Field)), ORder1:=xlAscending, header:=IIf(Weekly_CFTC_TXT, xlNo, xlYes), MatchCase:=False
        ' Filter for wanted dates.
        .AutoFilter Field:=IIf(ICE_Contracts = True, yymmdd_column, Date_Field), Criteria1:=Comparison_Operator, Operator:=xlFilterValues
        
        If ICE_Contracts Then
            ' Sort by Combined Contracts or Futures Only.
            .Sort key1:=.Cells(2, Combined_CLMN), ORder1:=xlAscending, header:=xlYes, MatchCase:=False
            'Filter for "Combined" if condition met.
            .AutoFilter Field:=Combined_CLMN, Criteria1:=Disaggregated_Filter_STR, Operator:=xlFilterValues, VisibleDropDown:=False
        End If

        If filterForSpecificContract Then
            .AutoFilter Field:=Contract_Code_CLMN, Criteria1:=UCase(contractCodeToFilterFor), Operator:=xlFilterValues, VisibleDropDown:=False
            On Error GoTo Catch_ContractCode_Not_Found
        Else
            On Error GoTo Catch_NoVisibleData
        End If
        
        With .SpecialCells(xlCellTypeVisible)
            On Error GoTo Finally
            If .Cells.count > Source_RNG.Rows(1).Cells.count Then
            
                If Weekly_CFTC_TXT Then
                    VAR_DTA = .value
                Else
                    If .Areas.count = 1 Then
                        ' Data excluding headers.
                        VAR_DTA = .Offset(1).Resize(.Rows.count - 1).value
                    Else
                        VAR_DTA = .Areas(2).value
                    End If
                End If
                
                If ICE_Contracts Then
                
                    For iRow = LBound(VAR_DTA, 1) To UBound(VAR_DTA, 1)
                        
                        If IsEmpty(VAR_DTA(iRow, Contract_Code_CLMN)) Then
                            ' Convert Dates from YYMMDD
                            VAR_DTA(iRow, Date_Field) = DateSerial(Left(VAR_DTA(iRow, yymmdd_column), 2) + 2000, Mid(VAR_DTA(iRow, yymmdd_column), 3, 2), Right(VAR_DTA(iRow, yymmdd_column), 2))
                            ' Map contract codes to different column
                            VAR_DTA(iRow, Contract_Code_CLMN) = VAR_DTA(iRow, ICE_Contract_Code_CLMN)
                            VAR_DTA(iRow, ICE_Contract_Code_CLMN) = Empty
                        End If
                        
                    Next iRow
                    
                End If
            
                Historical_Excel_Aggregation = VAR_DTA
                
            ElseIf filterForSpecificContract Then
                GoTo Catch_ContractCode_Not_Found
            End If
            
        End With 'End .SpecialCells(xlCellTypeVisible)
        
    End With
    
    If Not Filtering_QueryTable Then
        Application.StatusBar = vbNullString
        DoEvents
    End If

Finally:
    
    If Err.Number <> 0 Then
        If Not Contract_WB Is ThisWorkbook Then
            With Contract_WB
                .Close False
                Kill .fullName
            End With
            Application.StatusBar = vbNullString
        End If
        PropagateError Err, "Historical_Excel_Aggregation()"
    End If
    
    Exit Function
    
Catch_ContractCode_Not_Found: 'Used when user has input an invalid contract code

    If MsgBox("The Selected Contract Code [" & contractCodeToFilterFor & "] wasn't found" & vbNewLine & "Would you like to try again with a different Contract Code?", vbYesNo, "Please choose") _
                = vbYes Then
        contractCodeToFilterFor = UCase(Application.InputBox("Please supply the Contract Code of the desired contract"))
        GoTo Check_If_Code_Exists
    Else
        Application.StatusBar = vbNullString:
        If Not Contract_WB Is ThisWorkbook Then
            Contract_WB.Close False
        End If
        
        Call Re_Enable
        End
    End If
Catch_NoVisibleData:
    Err.description = "Attempt to retrieve data from compiled worksheet failed. No visible data after filtering."
    GoTo Finally
Scripts_Failed_To_Collect_Data:
    Err.description = "No data found on worksheet."
    GoTo Finally
Catch_CombinedColumn_Not_Found:
    Err.description = "Could not locate Combined column in Disaggregated report."
    GoTo Finally
End Function
Public Function Weekly_Text_File(File_Path As Collection, reportType As String, retrieveCombinedData As Boolean) As Variant()

    Dim File_IO As Variant, D As Byte, FilterC() As Variant, InfoF() As Variant, Contract_WB As Workbook
    
    FilterC = Filter_Market_Columns(convert_skip_col_to_general:=True, Return_Filter_Columns:=True, reportType:=reportType, Return_Filtered_Array:=False, Create_Filter:=True)
    
    ReDim InfoF(1 To UBound(FilterC, 1))
    
    For D = 1 To UBound(FilterC, 1)
        InfoF(D) = Array(D, FilterC(D))
    Next D
    
    Erase FilterC
    
    #If Mac Then
        D = xlMacintosh
    #Else
        D = xlWindows
    #End If
    
    For Each File_IO In File_Path
    
        On Error GoTo Error_While_Opening_Text_File
    
        With Workbooks
            .OpenText fileName:=File_IO, origin:=D, startRow:=1, DataType:=xlDelimited, _
                                TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, COMMA:=True, _
                                FieldInfo:=InfoF, DecimalSeparator:=".", ThousandsSeparator:=",", TrailingMinusNumbers:=False, _
                                Local:=False
                           
            Set Contract_WB = Workbooks(.count)
        End With
        
        With Contract_WB
            .Windows(1).Visible = False
             Weekly_Text_File = .Worksheets(1).UsedRange.value
            .Close False
        End With
        
        Kill File_IO
    
Next_File:
    
    Next File_IO
    
    Exit Function

Error_While_Opening_Text_File:
    PropagateError Err, "Weekly_Text_File()", "Error while attempting to open a Weekly based Text File."
    
End Function
Public Sub Filter_Market_Arrays(ByRef Contract_CLCTN As Collection, reportType As String, Optional ICE_Market As Boolean = False)
    
    Dim TempB As Variant, FilterC() As Variant, T As Long, Array_Count As Long, Unknown_Filter As Boolean
          
    With Contract_CLCTN
    
        Array_Count = .count
        
        If Array_Count > 1 Then
            FilterC = Filter_Market_Columns(convert_skip_col_to_general:=True, reportType:=reportType, Return_Filter_Columns:=True, Return_Filtered_Array:=False, ICE:=ICE_Market) '1 Based Positionl array filter
            Unknown_Filter = False
        Else
            Unknown_Filter = True
        End If
        
        For T = .count To 1 Step -1
            
            TempB = .Item(T)
            
            .Remove T
            
            TempB = Filter_Market_Columns(convert_skip_col_to_general:=True, Return_Filter_Columns:=False, _
                                            Return_Filtered_Array:=True, _
                                            inputA:=TempB, _
                                            ICE:=ICE_Market, _
                                            Column_Status:=FilterC, _
                                            Create_Filter:=Unknown_Filter, reportType:=reportType)
                                            
            If T = .count + 1 Then 'If last item in Collection then Simply re-add
                .Add TempB
            Else
                .Add TempB, Before:=T
            End If
            
        Next T
    
    End With

End Sub
Public Function Filter_Market_Columns(Return_Filter_Columns As Boolean, _
                                       Return_Filtered_Array As Boolean, _
                                       convert_skip_col_to_general As Boolean, _
                                       reportType As String, _
                                       Optional Create_Filter As Boolean = True, _
                                       Optional ByVal inputA As Variant, _
                                       Optional ICE As Boolean = False, _
                                       Optional ByVal Column_Status As Variant) As Variant
'======================================================================================================
'Generates an array referencing RAW data columns to determine if they should be kept or not
'If and array is given an return_filtered_array=True then the array will be filtered column wise based on the previous array
'======================================================================================================

    Dim ZZ As Long, output() As Variant, V As Byte, Y As Byte, columnOffset As Byte, columnsRemaining As Byte, _
    contractIdField As Byte, alternateCftcCodeColumn As Byte, _
    columnInOutput As Byte, finalColumnIndex As Byte, nameField As Byte, filterLength As Byte
    
    Dim CFTC_Wanted_Columns() As Variant, dateField As Byte, skip_value As Byte, twoDimensionalArray As Boolean
    
    CFTC_Wanted_Columns = GetAvailableFieldsTable(reportType).DataBodyRange.columns(2).Value2
    
    If ICE Then
        dateField = 2
        contractIdField = 7
    Else
        dateField = 3
        contractIdField = 4
        nameField = 1
    End If
        
    Select Case reportType
        Case "L":
            alternateCftcCodeColumn = 127
        Case "D":
            alternateCftcCodeColumn = 187
        Case "T":
            alternateCftcCodeColumn = 83
    End Select
    
    If convert_skip_col_to_general Then
        skip_value = xlGeneralFormat
    Else
        skip_value = xlSkipColumn
    End If
    
    If IsArray(inputA) Or IsMissing(inputA) Then
        filterLength = UBound(CFTC_Wanted_Columns, 1)
    Else
        filterLength = inputA.count
    End If
    
    If Create_Filter = True And IsMissing(Column_Status) Then 'IF column Status is empty or if it is empty
        
        ReDim Column_Status(1 To filterLength)

        For V = LBound(Column_Status) To UBound(Column_Status)
                
            ' Allows entry into block regardless of if ICE or CFTC is needed for dates or contract code
            On Error GoTo Catch_OutsideBounds
            
            If (CFTC_Wanted_Columns(V, 1) = True Or V = dateField Or V = contractIdField) Then
            
                Select Case V
                
                    Case dateField 'column 2 or 3 depending on if ICE or not
                        Column_Status(V) = IIf(ICE, xlGeneralFormat, xlYMDFormat) 'xlMDYFormat
                    Case nameField, contractIdField
                        Column_Status(V) = xlTextFormat
                    Case 2, 3, 4, 7 'These numbers may overlap with dates column or contract field
                                    'The previous cases will prevent it from executing unnecessarily depending on if ICE or not
                        Column_Status(V) = skip_value
                    Case Else
                        Column_Status(V) = xlGeneralFormat
                End Select
            Else
                If V = alternateCftcCodeColumn And convert_skip_col_to_general Then
                    Column_Status(V) = xlTextFormat
                Else
                    Column_Status(V) = skip_value
                End If
            End If
Assign_Next_FilterColumn:
        Next V
        
    End If
    On Error GoTo 0
    
    If Return_Filter_Columns = True Then
        Filter_Market_Columns = Column_Status
    ElseIf Return_Filtered_Array = True Then
        
         'Don't worry about text files..they are filtered in the same sub that they are opened in
         'FYI dateField would be 2 if doing TXT files...2 is used for ICE contracts because of exchange inconsistency
        On Error Resume Next
        
        If IsArray(inputA) Then
            Y = 0
            Do 'Determine the total number of dimensions
                Y = Y + 1
                V = LBound(inputA, Y)
            Loop Until Err.Number <> 0
            On Error GoTo 0
            If Y - 1 = 2 Then twoDimensionalArray = True
        ElseIf TypeName(inputA) = "Collection" Then
            twoDimensionalArray = False
        End If
        
        If twoDimensionalArray Then
            ReDim output(1 To UBound(inputA, 1), 1 To UBound(filter(Column_Status, xlSkipColumn, False)) + 1)
            finalColumnIndex = UBound(output, 2)
        Else
            ReDim output(1 To UBound(filter(Column_Status, xlSkipColumn, False)) + 1)
            finalColumnIndex = UBound(output, 1)
        End If
        
        Y = 0
        
        For V = LBound(Column_Status) To UBound(Column_Status)
            
            If Column_Status(V) <> xlSkipColumn Then
                
                Select Case V
                
                    Case nameField
                        columnInOutput = 2
                    Case dateField
                        columnInOutput = 1
                    Case contractIdField
                        columnInOutput = finalColumnIndex
                    Case Else
                        'Find the next value that excludes the above cases
                        Do
                            Y = Y + 1
                        Loop Until 2 < Y And Y < finalColumnIndex
                        
                        columnInOutput = Y
                        
                End Select
                
                If twoDimensionalArray Then
                    For ZZ = LBound(output, 1) To UBound(output, 1)
                        output(ZZ, columnInOutput) = inputA(ZZ, V)
                    Next ZZ
                Else
                    If IsObject(inputA(V)) Then
                        Set output(columnInOutput) = inputA(V)
                    Else
                        output(columnInOutput) = inputA(V)
                    End If
                End If
                
            End If
            
        Next V
        
        Filter_Market_Columns = output
        
    End If
    
    Exit Function
    
Catch_OutsideBounds:
    If Not IsArray(inputA) And Err.Number = 9 Then
        Column_Status(V) = xlGeneralFormat
        Resume Assign_Next_FilterColumn
    Else
        Err.Raise Err.Number, "Filter_Market_Columns", Err.description
    End If
    
End Function
Public Function Query_Text_Files(ByVal TXT_File_Paths As Collection, reportType As String, combined_wb As Boolean) As Variant
'===================================================================================================================
    'Purpose: Queries text files in TXT_File_Paths and adds their contents(array) to a collection
    'Inputs: reportType - One of L,D,T to represent what type of report to retrieve.
    '        combined_wb  - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    'Outputs: An array of the most recent weekly CFTC data.
    'Notes: Use only on Windows.
'===================================================================================================================
    Dim QT As QueryTable, file As Variant, Found_QT As Boolean, Field_Info() As Variant, Output_Arrays As New Collection, _
    Field_Info_ICE() As Variant
     
    Dim headerCount As Byte
    
     On Error GoTo Propagate
     
    For Each QT In QueryT.QueryTables 'Search for the following query if it exists
        If InStrB(1, QT.Name, "TXT Import") > 0 Then
            Found_QT = True
            Exit For
        End If
    Next QT
    
    Field_Info = Filter_Market_Columns(convert_skip_col_to_general:=True, reportType:=reportType, Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True) '^^ CFTC Column filter
    
    If reportType = "D" Then 'ICE Data column filter
        Field_Info_ICE = Filter_Market_Columns(convert_skip_col_to_general:=True, reportType:=reportType, Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True, ICE:=True)
    End If
    
    For Each file In TXT_File_Paths
        
        Application.StatusBar = "Querying: " & file
        DoEvents
        
        If Not Found_QT Then
            
            Set QT = QueryT.QueryTables.Add(Connection:="TEXT;" & file, Destination:=QueryT.Cells(1, 1))
            
            With QT
                .Name = "TXT Import"
                .BackgroundQuery = False
                .SaveData = False
            End With
            
            Found_QT = True 'So that this statement isn't executed again
            
        End If
        
        With QT
            
            .Connection = "TEXT;" & file
            .TextFileCommaDelimiter = True
            .TextFileConsecutiveDelimiter = False
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            
            If file Like "*.csv" And reportType = "D" Then 'ICE Workbooks
                .TextFileColumnDataTypes = Field_Info_ICE
            Else
                .TextFileColumnDataTypes = Field_Info
            End If
            
            .RefreshStyle = xlOverwriteCells
            .AdjustColumnWidth = False
            .Destination = QueryT.Cells(1, 1)
            
            .Refresh False
            
            headerCount = headerCount + 1
            
            With .ResultRange
                If headerCount = 1 Then
                    Output_Arrays.Add .Value2
                Else
                    Output_Arrays.Add .Offset(1).Resize(.Rows.count - 1).Value2
                End If
                .ClearContents
            End With
        
        End With
    
    Next file
    
    If Output_Arrays.count > 1 Then
        Query_Text_Files = CombineArraysInCollection(Output_Arrays, Append_Type.Multiple_2d)
    Else
        Query_Text_Files = Output_Arrays(1)
    End If
    
    QT.Delete
    
    Exit Function
    
Propagate:
    If Not QT Is Nothing Then
        QT.Delete
    End If
    
    PropagateError Err, "Query_Text_Files()"
    
End Function
Public Function TryGetPriceData(ByRef inputData As Variant, ByVal inputDataPriceColumn As Byte, contractDataOBJ As ContractInfo, _
    overwriteAllPrices As Boolean, datesAreInColumnOne As Boolean) As Boolean

'===================================================================================================================
    'Purpose: Retrieves price data.
    'Inputs: inputData -
    '        inputDataPriceColumn - Column within inputData to store prices in.
    '        contractDataOBJ - Contract instance that contains symbol info and where to get prices from.
    '        overwriteAllPrices - Clears price column in inputData.
    '        datesAreInColumnOne -  If true then dates are assumed to be in column 1 else in column 3.
    'Outputs:
'===================================================================================================================

    Dim Y As Long, Start_Date As Date, End_Date As Date, URL As String, _
    Year_1970 As Date, x As Long, Yahoo_Finance_Parse As Boolean, Stooq_Parse As Boolean
    
    Dim priceData() As String, Initial_Split_CHR As String, D_OHLC_AV() As String, foundData As Boolean
    
    Dim closePriceColumn As Byte, Secondary_Split_STR As String, Response_STR As String, QT_Connection_Type As String
    
    Dim End_Date_STR As String, Start_Date_STR As String, Query_Name As String, priceSymbol As String, reverseSortOrder As Boolean
    
    Dim QT As QueryTable, QueryTable_Found As Boolean, Using_QueryTable As Boolean, Query_Data() As Variant, dateColumn As Byte
    
    Const unmodified_COT_DateColumn As Byte = 3
    
    With contractDataOBJ
        If Not .HasSymbol Then Exit Function
        priceSymbol = .priceSymbol
        Yahoo_Finance_Parse = .UseYahooPrices
        Stooq_Parse = Not Yahoo_Finance_Parse
    End With
    
    If Not datesAreInColumnOne Then
        dateColumn = unmodified_COT_DateColumn
    Else
        dateColumn = 1
    End If
    
    If inputData(1, dateColumn) > inputData(UBound(inputData, 1), dateColumn) Then
    'For this sub to work correctly It must be ordered from oldest to newest
        reverseSortOrder = True
        inputData = Reverse_2D_Array(inputData)
    End If
    
    On Error GoTo Exit_Price_Parse
    
    Start_Date = inputData(1, dateColumn)
    End_Date = inputData(UBound(inputData, 1), dateColumn)
    
    If Yahoo_Finance_Parse Then 'CSV File
        
        Query_Name = "Yahoo Finance Query"
        Year_1970 = DateSerial(1970, 1, 1) 'Yahoo bases there URLs on the date converted to UNIX time
        End_Date = DateAdd("d", 1, End_Date) '1 more day than is in range to encapsulate that day
        Start_Date_STR = DateDiff("s", Year_1970, Start_Date) 'Convert to UNIX time
        End_Date_STR = DateDiff("s", Year_1970, End_Date) 'An extra day is added to encompass the End Day
        
        URL = "https://query1.finance.yahoo.com/v7/finance/download/" & priceSymbol & _
                "?period1=" & Start_Date_STR & _
                "&period2=" & End_Date_STR & _
                "&interval=1d&events=history&includeAdjustedClose=true"
          
        QT_Connection_Type = "TEXT;"
        
    Else 'CSV file from STOOQ
        
        Query_Name = "Stooq Query"
        
        End_Date_STR = Format(End_Date, "yyyymmdd")
        Start_Date_STR = Format(Start_Date, "yyyymmdd")
        URL = "https://stooq.com/q/d/l/?s=" & priceSymbol & "&d1=" & Start_Date_STR & "&d2=" & End_Date_STR & "&i=d"
        QT_Connection_Type = "URL;"
    
    End If
    
    End_Date_STR = vbNullString
    Start_Date_STR = vbNullString
        
    #If Mac Then
    
        On Error GoTo Exit_Price_Parse
        'On Error GoTo 0
        Using_QueryTable = True
    
        For Each QT In QueryT.QueryTables           'Determine if QueryTable Exists
            
            If InStrB(1, QT.Name, Query_Name) > 0 Then 'Instr method used in case Excel appends a number to the name
                QueryTable_Found = True
                Exit For
            End If
            
        Next QT
        
        If Not QueryTable_Found Then Set QT = QueryT.QueryTables.Add(QT_Connection_Type & URL, QueryT.Cells(1, 1))
        
        With QT
        
            If Not QueryTable_Found Then
            
                .BackgroundQuery = False
                .Name = Query_Name
                ' If an error occurs then delete the already existing connection and then try again.
                On Error GoTo Workbook_Connection_Name_Already_Exists
                    .WorkbookConnection.Name = Replace$(Query_Name, "Query", "Prices")
                On Error GoTo Exit_Price_Parse
                
            Else
                .Connection = QT_Connection_Type & URL
            End If
            
            .RefreshOnFileOpen = False
            .RefreshStyle = xlOverwriteCells
            .SaveData = False
            
            On Error GoTo Remove_QT_And_Connection 'Delete both the Querytable and the connection and exit the sub
    
             .Refresh False
            
            On Error GoTo Exit_Price_Parse
            
            With .ResultRange
                ' .value returns an array of comma separated values in a single column.
                If Yahoo_Finance_Parse Or Stooq_Parse Then Query_Data = .Value2
                .ClearContents
            End With
            
        End With
        
        Set QT = Nothing
        Query_Name = vbNullString
        QT_Connection_Type = vbNullString
        
    #Else
    
        On Error GoTo Exit_Price_Parse
        
        'Dim HTTP2 As New MSXML2.XMLHTTP60
        
        Dim HTTP2 As Object
        
        Set HTTP2 = CreateObject("Msxml2.ServerXMLHTTP")
    
        With HTTP2
            .Open "GET", URL, False
            .send
            Response_STR = .responseText
        End With
        
        Set HTTP2 = Nothing
        
    #End If
    
    URL = vbNullString
    
    On Error GoTo Exit_Price_Parse
      
    If Yahoo_Finance_Parse Or Stooq_Parse Then 'Parsing CSV Files
        
        If Not Using_QueryTable Then
            
            If InStrB(1, Response_STR, 404) = 1 Or LenB(Response_STR) = 0 Then Exit Function 'Something likely wrong with the URl
            
            If Yahoo_Finance_Parse Then
                Initial_Split_CHR = Mid$(Response_STR, InStr(1, Response_STR, "Volume") + Len("volume"), 1) 'Finding Splitting_Charachter
            ElseIf Stooq_Parse Then
                Initial_Split_CHR = vbNewLine
            End If
            
            priceData = Split(Response_STR, Initial_Split_CHR)
               
        Else
        
            ReDim priceData(0 To UBound(Query_Data, 1) - 1) 'redim to fit all rows of the query array
             
            For x = 0 To UBound(Query_Data, 1) - 1 'Add everything  to array
                priceData(x) = Query_Data(x + 1, 1)
            Next x
            
            Erase Query_Data
            
        End If
        
        If overwriteAllPrices Then
            'Data Table has been selected to have all price data overwritten
            For Y = LBound(inputData, 1) To UBound(inputData, 1)
                inputData(Y, inputDataPriceColumn) = Empty
            Next Y
        End If
        
        Secondary_Split_STR = Chr(44)
        x = LBound(priceData) + 1 'Skip headers
        
        closePriceColumn = 4 'Base 0 location of close prices within the queried array
        
    End If
    
    If LenB(Response_STR) > 0 Then Response_STR = vbNullString
    If LenB(Initial_Split_CHR) > 0 Then Initial_Split_CHR = vbNullString
    
    Y = 1
    
    Start_Date = CDate(Left$(priceData(x), InStr(1, priceData(x), Secondary_Split_STR) - 1))
    
    Do Until inputData(Y, dateColumn) >= Start_Date
    'Align the data based on the date
    
        If Y + 1 <= UBound(inputData, 1) Then
            Y = Y + 1
        Else
            If reverseSortOrder Then inputData = Reverse_2D_Array(inputData)
            Exit Function
        End If
        
    Loop
     
    For Y = Y To UBound(inputData, 1)
    
        On Error GoTo Error_While_Splitting
        
        Do Until Start_Date >= inputData(Y, dateColumn)
        'Loop until price dates meet or exceed wanted date
        '>= used in case there isnt  a price for the requested date
Increment_X:
    
            x = x + 1
            
            If x > UBound(priceData) Then
                Exit For
            Else
                Start_Date = CDate(Left$(priceData(x), InStr(1, priceData(x), Secondary_Split_STR) - 1))
            End If
            
        Loop
    
        On Error Resume Next
        
        If Start_Date = inputData(Y, dateColumn) Then
        
            D_OHLC_AV = Split(priceData(x), Secondary_Split_STR)
                    
            If Not IsNumeric(D_OHLC_AV(closePriceColumn)) Then 'find first value that came before that isn't empty
                inputData(Y, inputDataPriceColumn) = Empty
            ElseIf CDbl(D_OHLC_AV(closePriceColumn)) = 0 Then
                inputData(Y, inputDataPriceColumn) = Empty
            Else
                inputData(Y, inputDataPriceColumn) = CDbl(D_OHLC_AV(closePriceColumn))
                If Not foundData Then foundData = True
            End If
            
            Erase D_OHLC_AV
                
        End If
        
Ending_INcrement_X:
    Next Y
    
    TryGetPriceData = foundData
    
Exit_Price_Parse:
    
    Erase priceData
    If reverseSortOrder Then inputData = Reverse_2D_Array(inputData)
        
    Exit Function

Remove_QT_And_Connection:
    
    QT.Delete
    Exit Function
    
Workbook_Connection_Name_Already_Exists:

    ThisWorkbook.Connections(Replace(Query_Name, "Query", "Prices")).Delete
    
    QT.WorkbookConnection.Name = Replace(Query_Name, "Query", "Prices")
    Resume Next

Error_While_Splitting:

    If Err.Number = 13 Then 'type mismatch error from using cdate on a non-date string
        Resume Increment_X
    Else
        Exit Function
    End If
    
End Function

Public Sub Paste_To_Range(Optional Table_DataB_RNG As Range, Optional Data_Input As Variant, _
        Optional Sheet_Data As Variant, Optional Historical_Paste As Boolean = False, _
        Optional Target_Sheet As Worksheet, _
        Optional Overwrite_Data As Boolean = False)
'===================================================================================================================
    'Purpose: Places data at the bottom of a specified table.
    'Inputs: Table_DataB_RNG -
    '        Data_Input - Data to place in table when Historical_Paste is False.
    '        Sheet_Data - Data that is already present within a table or data to place if Historical_Paste is True..
    '        Historical_Paste - True if a table needs to be created and not normal weekly data additions.
    '        Target_Sheet - Worksheet that data will be placed on.
    '        Overwrite_Data - True if you want to clear any already present rows. ONly applicable if Historical_Paste is True
    'Outputs:
'===================================================================================================================
    Dim Model_Table As ListObject, Invalid_STR() As String, I As Long, _
    Invalid_Found() As Variant, newRowNumber As Long, rowNumber As Long, ColumnNumber As Long
    
    If Not Historical_Paste Then 'If Weekly/Block data addition
        
        If Not Overwrite_Data Then
            'Search in reverse order for dates that are too old to add to sheet.
            'Compare the Max date in data to upload and alrady on the sheet to determine how much if any of the data should be placed on the sheet.
            
            I = LBound(Data_Input, 1)

            Do While Data_Input(I, 1) <= Sheet_Data(UBound(Sheet_Data, 1), 1)
                I = I + 1
                If I > UBound(Data_Input, 1) Then Exit Do
            Loop

            If I > UBound(Data_Input, 1) Then
                Exit Sub
            ElseIf I <> LBound(Data_Input, 1) Then
            
                ReDim Invalid_Found(1 To UBound(Data_Input, 1) - I, 1 To UBound(Data_Input, 2))
                'Fill array with wanted data.
                For rowNumber = I To UBound(Data_Input, 1)
                
                    newRowNumber = newRowNumber + 1
                    
                    For ColumnNumber = 1 To UBound(Data_Input, 2)
                        Invalid_Found(newRowNumber, ColumnNumber) = Data_Input(rowNumber, ColumnNumber)
                    Next ColumnNumber
                    
                Next rowNumber
                
                Data_Input = Invalid_Found
            End If
        Else
            Table_DataB_RNG.ClearContents
            'Table_DataB_RNG.ListObject.AutoFilter.ShowAllData
        End If
        
        On Error GoTo No_Table
        
        With Table_DataB_RNG
            
            .Worksheet.DisplayPageBreaks = False
            .Cells(IIf(Overwrite_Data = False, .Rows.count + 1, 1), 1).Resize(UBound(Data_Input, 1), UBound(Data_Input, 2)).Value2 = Data_Input 'bottom row +1,1st column
            'Overwritten range depends on Overwrite Data Boolean, If true then overwrite all data on the worksheet
    
            With .ListObject
            
                If Not Overwrite_Data Then
                    'If just appending data.
                    If .DataBodyRange.Rows.count <> UBound(Data_Input, 1) + UBound(Sheet_Data, 1) Then
                        .Resize .Range.Resize(UBound(Data_Input, 1) + UBound(Sheet_Data, 1) + 1, .DataBodyRange.columns.count)
                        'resize to fit all data +1 to accomodate for headers
                    End If
                
                ElseIf .DataBodyRange.Rows.count <> UBound(Data_Input, 1) Then
                    .Resize .Range.Resize(UBound(Data_Input, 1) + 1, .DataBodyRange.columns.count)
                End If
                
            End With
            
        End With 'pastes the bottom row of the array if bottom date is greater than previous
        
    ElseIf Historical_Paste = True Then 'pastes to active sheet and retrieves headers from sheet

        If Overwrite_Data Then
            MsgBox "Within the Paste_To_Range sub OVerwrite_Data cannot be true if Historical_Paste is true."
            Exit Sub
        End If

        On Error GoTo PROC_ERR_Paste
    
        Set Model_Table = GetAvailableContractInfo(1).TableSource
            
        With Model_Table
            .DataBodyRange.Copy 'copy and paste formatting
            Target_Sheet.Range(.HeaderRowRange.Address).Value2 = .HeaderRowRange.Value2
        End With
        
        With Target_Sheet
        
            .Range("A2").Resize(UBound(Sheet_Data, 1), UBound(Sheet_Data, 2)).Value2 = Sheet_Data
            
            With .ListObjects.Add(xlSrcRange, .UsedRange, , xlYes)
                .DataBodyRange.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            End With
            
            .Hyperlinks.Add Anchor:=.Cells(1, 1), Address:=vbNullString, SubAddress:="'" & HUB.Name & "'!A1", TextToDisplay:=.Cells(1, 1).Value2
            
            On Error GoTo Re_Name '{Finding Valid Worksheet Name}
            .Name = Split(Sheet_Data(UBound(Sheet_Data, 1), 2), " -")(0)
        
        End With
        
        Application.StatusBar = "Data has been added to sheet. Calculating Formulas"
            
    End If
    
    On Error GoTo 0
    
    Exit Sub
        
Re_Name:
   MsgBox " If you were attempting to add a new contract then the Worksheet name could not be changed automatically."
    Resume Next
PROC_ERR_Paste:
    MsgBox "Error: (" & Err.Number & ") " & Err.description, vbCritical
    Resume Next
No_Table:
    MsgBox "If you are seeing this then either a table could not be found in cell A1 or your version " & _
    "of Excel does not support the listbody object. Further data will not be updated. Contact me via email."
    Call Re_Enable: End
End Sub

Public Function CreateFieldInfoMap(externalHeaders() As Variant, localDatabaseHeaders() As Variant, externalHeadersFromSocrataAPI As Boolean) As Collection
'==========================================================================================================
' Creates a Collection of FieldInfo insances for fields that are found within both externalHeaders and localDatabaseHeaders.
' Variables:
'   externalHeaders: 1D array of column names associated with each field from apiData
'   databaseFieldsByEditedName: Columns from a localy saved database.
'==========================================================================================================

    Dim iCount As Byte, externalHeaderIndexByEditedName As New Collection, Item As Variant, databaseFieldsByEditedName As New Collection, FI As FieldInfo

    On Error GoTo Abandon_Processes
    ' Column names from the api source are often spelled incorrectly or aren't standardized in their naming.
    With externalHeaderIndexByEditedName
        For iCount = LBound(externalHeaders) To UBound(externalHeaders)
            If externalHeadersFromSocrataAPI Then
                externalHeaders(iCount) = Replace(externalHeaders(iCount), "spead", "spread")
                externalHeaders(iCount) = Replace(externalHeaders(iCount), "postions", "positions")
                externalHeaders(iCount) = Replace(externalHeaders(iCount), "open_interest", "oi")
                externalHeaders(iCount) = Replace(externalHeaders(iCount), "__", "_")
                .Add iCount, externalHeaders(iCount)
            Else
                .Add iCount, EditDatabaseNames(CStr(externalHeaders(iCount)))
            End If
        Next iCount
    End With
    
    Dim FieldInfoMap As New Collection, endings() As String, editedName As String, mainLoopCount As Byte
    
    With databaseFieldsByEditedName
        For iCount = LBound(localDatabaseHeaders) To UBound(localDatabaseHeaders)
            editedName = EditDatabaseNames(CStr(localDatabaseHeaders(iCount)))
            .Add Array(editedName, localDatabaseHeaders(iCount)), editedName
        Next iCount
    End With
        
    Dim endingsIterator As Byte, endingStrippedName As String, digitIncrement As Byte, _
    foundMainEditedName As Boolean, secondaryIndex As Byte, newKey As String
    
    ' This array is ordered in the manner that they appear within the api columns.
    endings = Split("_all,_old,_other", ",")
    
    ' Loop through databaseFieldsByEditedName and determine if the edited name exists within externalHeaderIndexByEditedName.
    ' Regardless of if it does, create a FieldInfo instance and add to FieldInfoMap.
    For Each Item In databaseFieldsByEditedName
                       
        editedName = Item(0)
        mainLoopCount = mainLoopCount + 1
        
        If HasKey(FieldInfoMap, editedName) Then
            ' FieldInfo instance has already been added. Ensure its order within the collection.
            foundMainEditedName = True
            With FieldInfoMap
                Set FI = .Item(editedName)
                .Remove editedName
                .Add FI, FI.editedName, After:=databaseFieldsByEditedName(mainLoopCount - 1)(0)
            End With
        ElseIf HasKey(externalHeaderIndexByEditedName, editedName) Then
            ' Exact match between column name sources.
            Set FI = New FieldInfo
            FI.Constructor editedName, externalHeaderIndexByEditedName(editedName), CStr(Item(1))
            
            With FieldInfoMap
                If .count = 0 Then
                    .Add FI, editedName
                ElseIf .count > 0 Then
                    .Add FI, editedName, After:=databaseFieldsByEditedName(mainLoopCount - 1)(0)
                End If
            End With
            
            externalHeaderIndexByEditedName.Remove editedName
            foundMainEditedName = True
        Else
            
            For endingsIterator = LBound(endings) To UBound(endings)
                ' Checking if the name ends with the pattern.
                If editedName Like "*" + endings(endingsIterator) Then
                    
                    endingStrippedName = Replace$(editedName, endings(endingsIterator), vbNullString)
                    digitIncrement = 0
                    
                    For secondaryIndex = endingsIterator To UBound(endings)
                        
                        Dim apiFieldName As String, placementKnown As Boolean
                        
                        newKey = vbNullString
                        placementKnown = False
                        
                        If secondaryIndex = endingsIterator And HasKey(externalHeaderIndexByEditedName, endingStrippedName) Then
                            newKey = editedName
                            apiFieldName = endingStrippedName
                            placementKnown = True
                            foundMainEditedName = True
                        ElseIf secondaryIndex > endingsIterator Then
                            
                            digitIncrement = digitIncrement + 1
                            apiFieldName = endingStrippedName & "_" & digitIncrement
                            
                            If HasKey(externalHeaderIndexByEditedName, apiFieldName) Then
                                newKey = endingStrippedName + endings(secondaryIndex)
                            End If
                            
                        End If
                        
                        If LenB(newKey) > 0 Then
                            Set FI = New FieldInfo
                            FI.Constructor newKey, externalHeaderIndexByEditedName(apiFieldName), CStr(databaseFieldsByEditedName(newKey)(1))
                            
                            If placementKnown Then
                                FieldInfoMap.Add FI, newKey, After:=databaseFieldsByEditedName(mainLoopCount - 1)(0)
                            Else
                                FieldInfoMap.Add FI, newKey
                            End If
                                                        
                            ' Removal is just for viewing how many and which api columns weren't found.
                            externalHeaderIndexByEditedName.Remove apiFieldName
                        End If
                        
                    Next secondaryIndex
                    
                End If
            
            Next endingsIterator
                                                
        End If
        ' This conditional adds a FieldINfo instance with the IsMissing property set to true.
        If Not foundMainEditedName Then
            Set FI = New FieldInfo
            FI.Constructor editedName, 0, CStr(Item(1)), True
            'Place after previous field by name.
            FieldInfoMap.Add FI, editedName, After:=databaseFieldsByEditedName(mainLoopCount - 1)(0)
        End If
        foundMainEditedName = False
        
    Next Item
        
    Set CreateFieldInfoMap = FieldInfoMap
     
    Exit Function
    
Abandon_Processes:
    PropagateError Err, "CreateFieldInfoMap()"
End Function



