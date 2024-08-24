Attribute VB_Name = "Public_Functions"
#If Not Mac Then
    #If VBA7 Then
        Public Declare PtrSafe Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (tzInfo As TimeZoneInformation, localTime As SystemTime, returnedTimeUTC As SystemTime) As LongPtr
        Public Declare PtrSafe Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (tzInfo As TimeZoneInformation, localTime As SystemTime, returnedLocalTime As SystemTime) As LongPtr
        Public Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (tzInfo As TimeZoneInformation) As LongPtr
        Public Declare PtrSafe Function GetSystemTime Lib "kernel32" (returnedUTC As SystemTime) As LongPtr
    #Else
        Public Declare Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (tzInfo As TimeZoneInformation, localTime As SystemTime, returnedTimeUTC As SystemTime) As Long
        Public Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (tzInfo As TimeZoneInformation, localTime As SystemTime, returnedLocalTimeUTC As SystemTime) As Long
        Public Declare Function GetTimeZoneInformation Lib "kernel32" (tzInfo As TimeZoneInformation) As Long
        Public Declare Function GetSystemTime Lib "kernel32" (returnedUTC As SystemTime) As Long
    #End If
#End If

Public Type SystemTime
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type TimeZoneInformation
    Bias As Long 'The bias is the difference, in minutes, between Coordinated Universal Time (UTC) and local time.
    standardName(31) As Integer
    StandardDate As SystemTime
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SystemTime
    DaylightBias As Long
    'TimeZoneKeyName(127) As Integer
    'DynamicDaylightTimeDisabled As Boolean
End Type
Option Explicit
#If Not DatabaseFile Then

    Public Function ReturnCftcTable(WS As Worksheet) As ListObject
    '===================================================================================================================
        'Summary: Checks to see if a cftc table exists on a worksheet(WS) based on table names.
        'Inputs: WS - worksheet to check.
        'Returns: CFTC Listobject or Nothing.
        'Notes - ONly applicable to Non - Database version of file.
    '===================================================================================================================
        Dim Item As ListObject, tableName$
        
        For Each Item In WS.ListObjects
            tableName = Item.Name
            If tableName Like "CFTC_*" Or tableName Like "ICE_*" Then
                Set ReturnCftcTable = Item
                Exit Function
            End If
        Next Item
        
    End Function
    
    Public Function IsWorkbookForFuturesAndOptions() As Boolean
        On Error GoTo Propogate
        IsWorkbookForFuturesAndOptions = Variable_Sheet.Range("Combined_Workbook").Value2
        Exit Function
Propogate:
        PropagateError Err, "IsWorkbookForFuturesAndOptions", "'Combined_Workbook' range could not be found."
    End Function
    
    Public Function ReturnReportType$()
        On Error GoTo Propogate
        ReturnReportType = Variable_Sheet.Range("Report_Type").Value2
        Exit Function
Propogate:
        PropagateError Err, "ReturnReportType", "'Report_Type' range could not be found."
    End Function
    
#End If
    
Sub SendEmailFromOutlook(Body$, Subject$, toEmails$, ccEmails$, bccEmails$)
    Dim outApp As Object
    Dim outMail As Object
    On Error GoTo No_Outlook
    
    Set outApp = CreateObject("Outlook.Application")
    Set outMail = outApp.CreateItem(0)
 
    With outMail
        .To = toEmails
        .cc = ccEmails
        .BCC = bccEmails
        .Subject = Subject
        .HTMLBody = Body
        .send 'Send the email
    End With
 
    Set outMail = Nothing
    Set outApp = Nothing
    
    Exit Sub
    
No_Outlook:
    MsgBox "Microsoft Outlook object could not be loaded."
End Sub

Function SplitOutsideOfQuotes(ByVal inputA$, Delimiter$, Optional N_Delimiter$ = "*") As String()

    Dim x As Long, SA$()

    If InStrB(1, inputA, Chr(34)) = 0 Then 'if there are no quotation marks then split with the supplied delimiter
        SplitOutsideOfQuotes = Split(inputA, Delimiter)
    Else
        SA = Split(inputA, Chr(34))
        
        For x = LBound(SA) To UBound(SA) Step 2
            If InStrB(SA(x), Delimiter) <> 0 Then
               SA(x) = Replace$(SA(x), Delimiter, N_Delimiter)
            End If
        Next x
        SplitOutsideOfQuotes = Split(Join(SA, vbNullString), N_Delimiter)
    End If

End Function
Public Sub Courtesy()

    With Application

        If Not IsOnCreatorComputer Then
            .StatusBar = vbTab & vbTab & vbTab & "Brought to you by MoshiM. Consider donating to support the continued development of this project."
        Else
            .StatusBar = vbNullString
        End If
    
    End With

End Sub
Function GetUtcTime() As Date
'===================================================================================================================
    'Summary: Gets the current UTC time.
    'Returns: The current UTC datetime.
'===================================================================================================================
    #If Mac Then
        GetUtcTime = MacScript("set UTC to (current date) - (time to GMT)")
    #Else
        With CreateObject("WbemScripting.SWbemDateTime")
            .SetVarDate Now
           GetUtcTime = .GetVarDate(False)
        End With
    #End If
End Function

Function IsPowerQueryAvailable() As Boolean 'Determine if Power Query is available..use for when less than EXCEL 2016
    
    Dim bAvailable As Boolean
    On Error Resume Next
    bAvailable = Application.COMAddIns("Microsoft.Mashup.Client.Excel").Connect
    On Error GoTo 0
    IsPowerQueryAvailable = bAvailable
    'Debug.Print bAvailable
    
End Function
Sub Donators(Query_W As Worksheet, Target_T As Shape)
    '_______________________________________________________________
    'Take text from online text file and apply to shape
    Dim url$, QT As QueryTable, Disclaimer As Shape, My_Info$, DL$
    
    #If DatabaseFile Then
        If Variable_Sheet.Range("Github_Version").Value2 = True Then Exit Sub
    #End If
    
    On Error GoTo EXIT_DN_List

    DL = vbNewLine & vbNewLine

    My_Info = "Contact Email:   MoshiM_UC@outlook.com" & DL & _
                            "Skills:  C#, Python, Excel VBA, SQL, Data Analysis and Web Scraping." & DL & _
                            "Feel free to contact me for both personal and work related jobs."

    url = Replace$("https://www.dropbox.com/s/g75ij0agki217ow/CT%20Donators.txt?dl=0", _
            "www.dropbox.com", "dl.dropboxusercontent.com") 'URL leads to external text file
        
    Target_T.TextFrame.Characters.Text = vbNullString 'Clear text from shape

    Set QT = Query_W.QueryTables.Add("TEXT;" & url, Query_W.Range("A1")) 'Assign object to Variable

    With QT
        .BackgroundQuery = False
        .SaveData = False
        .AdjustColumnWidth = False
        .RefreshStyle = xlOverwriteCells
        .Refresh False
        
        With .ResultRange
            Target_T.TextFrame.Characters.Text = .Cells(1, 1) & vbNewLine & .Cells(2, 1) & DL & My_Info
            .ClearContents
        End With
    End With

Remove_QueryTable:

    With Target_T
        Set Disclaimer = .parent.Shapes("Disclaimer")
        .TextFrame.AutoSize = True
        .TextFrame.AutoSize = False
        .Width = Disclaimer.Width
        .Left = Disclaimer.Left
        .Top = Disclaimer.Top + Disclaimer.Height + 7
        .Visible = True
    End With

Finally:
    
    If Not QT Is Nothing Then
        With QT
            .WorkbookConnection.Delete
            .Delete
        End With
    End If

    Exit Sub

EXIT_DN_List:
    Target_T.TextFrame.Characters.Text = My_Info
    Resume Remove_QueryTable
    
End Sub
Public Function CFTC_Release_Dates(Find_Latest_Release As Boolean, convertToLocalTime As Boolean) As Date
'===================================================================================================================
    'Summary: Finds a wanted release date.
    'Inputs: Find_Latest_Release - If true then find the most recent release; else find next release.
    'Returns: Wanted Date.
'===================================================================================================================

    Dim wantedDateTime As Date, iRow As Byte, iColumn As Byte, variableDate As Date, releaseSchedule() As Variant, _
    Time_Zones() As Variant, easternTime As Date, Local_Time As Date, releaseYear As Long, releaseDay As Byte, releaseMonth As Byte, iCount As Byte
    
    Dim easternHourlyOffset As Long, currentEasternTime As Date
    
    On Error GoTo Propogate
    
    With Variable_Sheet
        Time_Zones = .ListObjects("Time_Zones").DataBodyRange.columns(2).Value2 'This Query is refrshed on Workbook Open
        releaseSchedule = .ListObjects("Release_Schedule").DataBodyRange.Value2 'Array of Release Dates
    End With
    
    easternTime = Time_Zones(1, 1)
    Local_Time = Time_Zones(2, 1)

    If easternTime = TimeSerial(0, 0, 0) Or Local_Time = TimeSerial(0, 0, 0) Then Err.Raise 13, , "Local or Eastern Time couldn't be determined."
    
    easternHourlyOffset = DateDiff("h", easternTime, Local_Time, vbSunday, vbFirstJan1)
    currentEasternTime = DateAdd("h", -easternHourlyOffset, Now)
    
    For iRow = LBound(releaseSchedule, 1) To UBound(releaseSchedule, 1)
    
        If IsNumeric(releaseSchedule(iRow, 1)) Then 'Checking in first column for Year
            releaseYear = CInt(releaseSchedule(iRow, 1))
            iCount = 0
            'While still within bounds of array, determine how many months are available for each year and initiate month enum to 1 - wanted.
            Do While iCount < UBound(releaseSchedule, 1) - iRow And iCount <= 12
                If LenB(releaseSchedule(iRow + 1 + iCount, 1)) = 0 Then Exit Do
                iCount = iCount + 1
            Loop
            releaseMonth = 12 - (iCount - 1)
        ElseIf LenB(releaseSchedule(iRow, 1)) <> 0 And releaseMonth > 0 Then
            For iColumn = 2 To UBound(releaseSchedule, 2)
                If LenB(releaseSchedule(iRow, iColumn)) <> 0 Then
                    
                    releaseDay = CByte(Replace$(releaseSchedule(iRow, iColumn), "*", vbNullString))
                    'Datetime 15:30 Eastern Time
                    variableDate = DateSerial(releaseYear, releaseMonth, releaseDay) + TimeSerial(15, 30, 0)
                    
                    'If looking for the previous release datetime.
                    If Find_Latest_Release Then
                        If variableDate > currentEasternTime Then
                            Exit For
                        Else
                            wantedDateTime = variableDate
                        End If
                    Else
                        'If finding the next release datetime.
                        If variableDate > currentEasternTime Then
                            wantedDateTime = variableDate
                            Exit For
                        End If
                    End If
                End If
            Next iColumn
            releaseMonth = releaseMonth + 1
            If variableDate > currentEasternTime Then Exit For
        End If
        
    Next iRow
    
    If wantedDateTime = TimeSerial(0, 0, 0) Then wantedDateTime = variableDate
    
    CFTC_Release_Dates = IIf(convertToLocalTime, DateAdd("h", easternHourlyOffset, wantedDateTime), wantedDateTime)
    
    Exit Function
Assign_local_time_to_now:
    Local_Time = Now
    Resume Next
Propogate:
    PropagateError Err, "CFTC_Release_Dates"
End Function

Public Function IsOnCreatorComputer() As Boolean
'===================================================================================================================
    'Summary: Determines if user is on creator computer.
    'Returns: Boolean representation of whether or not file is being run on creator computer.
'===================================================================================================================
    
#If Not Mac Then
    
    Dim creatorProperties As Object, storedSerialNumber As Long, isOnUserComputer As Boolean
    
    Const Function_Value_Key$ = "Creator_Computer_?"
    
    On Error GoTo Collection_Lacks_Key
    IsOnCreatorComputer = ThisWorkbook.Event_Storage(Function_Value_Key)
    Exit Function
    
Load_Password_File:

    On Error GoTo Exit_UUID
        
    Set creatorProperties = GetCreatorPasswordsAndCodes()
    
    If Not creatorProperties Is Nothing Then
        storedSerialNumber = creatorProperties("DRIVE_SERIAL_NUMBER")
        isOnUserComputer = (DoesDriveSerialNumberMatch(storedSerialNumber) And Environ$("COMPUTERNAME") = "CAMPBELL-PC")
    End If
    ThisWorkbook.Event_Storage.Add isOnUserComputer, Function_Value_Key
    
    IsOnCreatorComputer = isOnUserComputer
    Exit Function

Exit_UUID:

    ThisWorkbook.Event_Storage.Add False, Function_Value_Key
    IsOnCreatorComputer = False
    Exit Function

Collection_Lacks_Key:
    Resume Load_Password_File
#End If
End Function
Public Function DoesDriveSerialNumberMatch(My_Serial As Long) As Boolean

    Dim FS As Object, D As Drive, x As Long, tt$

    On Error GoTo No_Scripting

    With ThisWorkbook
        If .SerialN = 0 Then
            Set FS = CreateObject("Scripting.FileSystemObject")
            Set D = FS.GetDrive(FS.GetDriveName(FS.GetAbsolutePathName(Empty))) 'drvpath
            .SerialN = D.SerialNumber
        End If
        If .SerialN = My_Serial Then DoesDriveSerialNumberMatch = True
    End With

'Select Case d.DriveType
'Case 0: t = "Unknown"
'Case 1: t = "Removable"
'Case 2: t = "Fixed"
'Case 3: t = "Network"
'Case 4: t = "CD-ROM"
'Case 5: t = "RAM Disk"
'End Select

No_Scripting:

End Function
Function MAC_Identifier(MAC_Address_Input$) As Boolean
 
'Declaring the necessary variables.
    Dim strComputer$
    Dim objWMIService   As Object
    Dim colItems        As Object
    Dim objItem         As Object

    'Set the computer.
    strComputer = "."
 
    'The root\cimv2 namespace is used to access the Win32_NetworkAdapterConfiguration class.
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
 
    'A select query is used to get a collection of network adapters that have the property IPEnabled equal to true.
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
 
    'Loop through all the collection of adapters and return the MAC address of the first adapter that has a non-empty IP.
    For Each objItem In colItems
        If Not IsNull(objItem.IPAddress) Then
            If MAC_Address_Input = objItem.MACAddress Then MAC_Identifier = True
            Exit For
        End If
    Next

    Set objWMIService = Nothing
    Set colItems = Nothing
    Set objItem = Nothing
'
End Function

Public Function GetAvailableContractInfo() As Collection
'===================================================================================================================
    'Summary: Creates a collection of all available Contract instances.
    'Returns: A collection of contract instances.
'===================================================================================================================
    
    On Error GoTo Propagate
    #If DatabaseFile Then
        Set GetAvailableContractInfo = GetContractInfo_DbVersion
    #Else
    
        Const codeColumn As Byte = 1, yahooColumn As Byte = 3, nameColumn As Byte = 2
        Dim priceSymbolsTable As Range, currentSymbol$, Yahoo_Finance_Ticker As Boolean, WS As Worksheet, lo As ListObject
        
        Dim Code$, contractData As ContractInfo, rowIndex As Long, This_C As New Collection
        
        Set priceSymbolsTable = Symbols.ListObjects("Symbols_TBL").DataBodyRange

        For Each WS In ThisWorkbook.Worksheets
            For Each lo In WS.ListObjects
                With lo
                    If .Name Like "CFTC_*" Or .Name Like "ICE_*" Then
                        Code = Right$(.Name, Len(.Name) - InStr(1, .Name, "_", vbBinaryCompare))
                        
                        On Error Resume Next
                            currentSymbol = WorksheetFunction.VLookup(Code, priceSymbolsTable, 3, False)
                        On Error GoTo 0
                        
                        Yahoo_Finance_Ticker = LenB(currentSymbol) <> 0
                        
                        Set contractData = New ContractInfo
                        contractData.InitializeContract Code, currentSymbol, Yahoo_Finance_Ticker, lo
                        
                        This_C.Add contractData, Code
                        
                        currentSymbol = vbNullString
                    End If
                End With
            Next lo
        Next WS
        Set GetAvailableContractInfo = This_C
    #End If
    
    Exit Function
Propagate:
    PropagateError Err, "GetAvailableContractInfo"
End Function

Public Function IsLoadedUserform(User_Form_Name$) As Boolean
'===================================================================================================================
'Summary: Determines if a UserForm is loaded based on its name.
'Inputs: User_Form_Name - Name to check for.
'Returns: Stores the current time on the Variable_Sheet along with the local time on the running environment.
'===================================================================================================================

    Dim frm As Object

    For Each frm In VBA.UserForms
        If frm.Name = User_Form_Name Then
            IsLoadedUserform = True
            Exit Function
        End If
    Next frm

End Function

Public Function ConvertCollectionToArray(data As Collection) As Variant()
'=================================================================================================
'Summary: Parses eash item within a collection and returns an array of those items.
'=================================================================================================
    Dim outputA() As Variant, iCount As Long
    
    With data
        ReDim outputA(1 To .count)
        For iCount = 1 To .count
            If Not IsObject(.Item(iCount)) Then
                outputA(iCount) = .Item(iCount)
            Else
                Set outputA(iCount) = .Item(iCount)
            End If
        Next iCount
    End With
    
    ConvertCollectionToArray = outputA

End Function
Public Function ConvertArrayToCollection(ByVal data As Variant, useValuesAsKey As Boolean) As Collection
    
    Dim output As New Collection, G As Long, nDimensions As Byte, columnCount As Variant
    
    On Error Resume Next
    nDimensions = 1
    Do
        nDimensions = nDimensions + 1
        columnCount = UBound(data, nDimensions)
    Loop While Err.Number = 0
    
    If nDimensions - 1 = 2 Then
        data = WorksheetFunction.Transpose(data)
    End If
    
    With output
        For G = LBound(data) To UBound(data)
            If useValuesAsKey Then
                .Add data(G), CStr(data(G))
            Else
                .Add data(G)
            End If
        Next G
    End With
    
    Set ConvertArrayToCollection = output
    
End Function
Public Function StandardizedDatabaseFieldNames(DatabaseName$) As String
'=========================================================================================
' Standardizes [DatabaseName] for interface with socrata api.
'=========================================================================================
    Dim lcaseVersion$
    
    lcaseVersion = LCase$(DatabaseName)
    
    If InStrB(1, lcaseVersion, "yyyy") <> 0 Then
        lcaseVersion = "report_date_as_yyyy_mm_dd"
    Else
        If InStrB(lcaseVersion, " ") <> 0 Then lcaseVersion = Replace$(lcaseVersion, " ", "_")
        If InStrB(lcaseVersion, Chr(34)) <> 0 Then lcaseVersion = Replace$(lcaseVersion, Chr(34), vbNullString)
        If InStrB(lcaseVersion, "%") <> 0 Then lcaseVersion = Replace$(lcaseVersion, "%", "pct")
        If InStrB(lcaseVersion, "=") <> 0 Then lcaseVersion = Replace$(lcaseVersion, "=", "_")
        If InStrB(lcaseVersion, "(") <> 0 Then lcaseVersion = Replace$(lcaseVersion, "(", "_")
        If InStrB(lcaseVersion, ")") <> 0 Then lcaseVersion = Replace$(lcaseVersion, ")", vbNullString)
        If InStrB(lcaseVersion, "-") <> 0 Then lcaseVersion = Replace$(lcaseVersion, "-", "_")
        If InStrB(lcaseVersion, "commercial") <> 0 Then lcaseVersion = Replace$(lcaseVersion, "commercial", "comm")
        If InStrB(lcaseVersion, "reportable") <> 0 Then lcaseVersion = Replace$(lcaseVersion, "reportable", "rept")
        If InStrB(lcaseVersion, "total") <> 0 Then lcaseVersion = Replace$(lcaseVersion, "total", "tot")
        If InStrB(lcaseVersion, "concentration") <> 0 Then lcaseVersion = Replace$(lcaseVersion, "concentration", "conc")
        If InStrB(lcaseVersion, "spreading") <> 0 Then lcaseVersion = Replace$(lcaseVersion, "spreading", "spread")
        If InStrB(lcaseVersion, "_lt_") <> 0 Then lcaseVersion = Replace$(lcaseVersion, "_lt_", "_le_")
        If InStrB(lcaseVersion, "_in_initials") <> 0 Then lcaseVersion = Replace$(lcaseVersion, "_in_initials", vbNullString)
        
        Do While InStrB(1, lcaseVersion, "__") <> 0
            lcaseVersion = Replace$(lcaseVersion, "__", "_")
        Loop
        
        If InStrB(lcaseVersion, "open_interest_oi") <> 0 Then lcaseVersion = Replace$(lcaseVersion, "open_interest_oi", "oi")
        If InStrB(lcaseVersion, "open_interest") <> 0 Then lcaseVersion = Replace$(lcaseVersion, "open_interest", "oi")
        
    End If
        
    StandardizedDatabaseFieldNames = lcaseVersion

End Function

Public Function GETNUMBER(inputValue$, Optional index As Byte = 1) As Long
    
    Dim outputNumber$, i As Byte, _
    currentCharacter$, addToOutput As Boolean, decimalCount As Byte, threeCharacters$
    
    Dim numbersCollection As New Collection, finishedNumber As Boolean, lengthOfText As Byte
    
    lengthOfText = Len(inputValue)
    
    For i = 1 To lengthOfText
    
        currentCharacter = Mid$(inputValue, i, 1)
        
        If currentCharacter Like "#" Then
            addToOutput = True
            finishedNumber = False
        ElseIf i > 1 Then
            threeCharacters = Mid$(inputValue, i - 1, 3)
            If threeCharacters Like "#.#" Or threeCharacters Like "#,#" Then
                addToOutput = True
                decimalCount = decimalCount + 1
            End If
        Else
            finishedNumber = True
        End If
        
        If addToOutput Then
            outputNumber = outputNumber & currentCharacter
            addToOutput = False
        End If
        
        If (i = lengthOfText Or finishedNumber) And LenB(outputNumber) <> 0 Then
            finishedNumber = False
            numbersCollection.Add outputNumber
            outputNumber = vbNullString
        End If
        
    Next i
    
    GETNUMBER = numbersCollection(index) * 1
        
End Function
Public Function GetExpectedLocalFieldInfo(reportType As ReportEnum, filterUnwantedFields As Boolean, reArrangeToReflectSheet As Boolean, includePrice As Boolean, Optional adjustIndexes As Boolean = True) As Collection
'======================================================================================================================
'   Summary: Generates FieldInfo instances for field names stored on Variable Sheet.
'   Paramaters:
'       reportType - One of LDT to select which report type you want the fields for.
'       filterUnwantedFields - True if you want to remove unwanted fields from output collection.
'       reArrangeToReflectSheet - True to re-arrange order of collection fields to match known output on a worksheet
'       includePrice - True to add a Price column to the output collection.
'       adjustIndexes -
'======================================================================================================================
    Dim iCount As Integer, localCopyOfColumnNames() As Variant, columnMap As New Collection, FI As FieldInfo
    
    On Error GoTo Propagate
    localCopyOfColumnNames = GetAvailableFieldsTable(reportType).DataBodyRange.Value2
    
    With columnMap
    
        For iCount = 1 To UBound(localCopyOfColumnNames, 1)
            If Not filterUnwantedFields Or localCopyOfColumnNames(iCount, 2) = True Then
                Set FI = CreateFieldInfoInstance(StandardizedDatabaseFieldNames(CStr(localCopyOfColumnNames(iCount, 1))), iCount, CStr(localCopyOfColumnNames(iCount, 1)), False, CBool(localCopyOfColumnNames(iCount, 2)))
                .Add FI, FI.EditedName
            End If
        Next iCount
        
        If reArrangeToReflectSheet Or filterUnwantedFields Then
        
            If reArrangeToReflectSheet And filterUnwantedFields Then
                Set FI = .Item("cftc_contract_market_code")
                .Remove FI.EditedName
                .Add FI, FI.EditedName
    
                Set FI = .Item("report_date_as_yyyy_mm_dd")
                .Remove FI.EditedName
                .Add FI, FI.EditedName, Before:=1
            End If
            
            If adjustIndexes Then
                iCount = 0
                For Each FI In columnMap
                    iCount = iCount + 1
                    FI.ColumnIndex = iCount
                Next FI
            End If
            
        End If
        
        If includePrice Then
            Set FI = CreateFieldInfoInstance("price", 1 + IIf(filterUnwantedFields, .count, UBound(localCopyOfColumnNames, 1)), "Price", False)
            .Add FI, FI.EditedName
        End If
    
    End With
    
    Set GetExpectedLocalFieldInfo = columnMap
    Exit Function
Propagate:
    PropagateError Err, "GetExpectedLocalFieldInfo"
End Function
Public Function GetAvailableFieldsTable(reportType As ReportEnum) As ListObject
'======================================================================================================================
'   Summary: Returns a reference to a table on Variable_Sheet that contains a list of fields for a given [reportType]
'   Paramaters:
'       [reportType] - Report you want the fields for.
'======================================================================================================================
    On Error GoTo Catch_Table_Not_Found
    
    Dim tableName$: tableName = ConvertReportTypeEnum(reportType) & "_User_Selected_Columns"
    Set GetAvailableFieldsTable = Variable_Sheet.ListObjects(tableName)
    
    Exit Function
Catch_Table_Not_Found:
    PropagateError Err, "GetAvailableFieldsTable", tableName & " listobject could not be found on the " & Variable_Sheet.Name & " worksheet."
End Function
Public Function GetSocrataApiEndpoint(reportType As ReportEnum, oiType As OpenInterestType) As String
'===================================================================================================================
'Summary: Retrieves an API endpoint based on given parameters.
'Inputs:
'   reportType - A ReportEnum used to target an endpoint
'   oiType - An OpenInterestType enum used to target an end point.
'Returns: A socrata API endpoint string.
'===================================================================================================================
    Dim iRow As Byte

    For iRow = 0 To 2
        If Array(eLegacy, eDisaggregated, eTFF)(iRow) = reportType Then
            If oiType = FuturesAndOptions Then
                GetSocrataApiEndpoint = Array("jun7-fc8e", "kh3c-gbw2", "yw9f-hn96")(iRow)
            Else
                GetSocrataApiEndpoint = Array("6dca-aqww", "72hh-3qpy", "gpe5-46if")(iRow)
            End If
            Exit Function
        End If
    Next iRow

End Function
Public Function CFTC_CommodityGroupings() As Collection
'============================================================================================================
'Queries the Socrata API to get commodity groups and subgroups for each contract.
'============================================================================================================
    Dim outputCLCTN As New Collection, apiCode$, reportKey$, _
    apiUrl$, dataFilters$, apiData$(), iRow As Long, appProperties As Collection, httpResult$(), response$
    
    Const apiBaseURL$ = "https://publicreporting.cftc.gov/resource/"
    Const queryReturnLimit As Long = 5000
    
    Set appProperties = DisableApplicationProperties(True, False, True)
    
    On Error GoTo Finally
    
    apiCode = GetSocrataApiEndpoint(eLegacy, FuturesAndOptions)

    dataFilters = ".csv?$select=cftc_contract_market_code,commodity_group_name,commodity_subgroup_name" & _
                    "&$where=report_date_as_yyyy_mm_dd=" & Format$(Variable_Sheet.Range("Last_Updated_CFTC").Value2, "'yyyy-mm-dd'")
                    
    apiUrl = apiBaseURL & apiCode & dataFilters
    
    If TryGetRequest(apiUrl, response) Then
    
        httpResult = Split(Replace$(response, Chr(34), vbNullString), vbLf)
        Set outputCLCTN = New Collection
        
        With outputCLCTN
            For iRow = LBound(httpResult) + 1 To UBound(httpResult)
                If LenB(httpResult(iRow)) <> 0 Then
                    apiData = Split(httpResult(iRow), ",")
                    .Add Array(apiData(0), apiData(1), apiData(2)), apiData(0)
                End If
            Next iRow
        End With
        
    End If
    
    Set CFTC_CommodityGroupings = outputCLCTN
    
Finally:
    With HoldError(Err)
        EnableApplicationProperties appProperties
        If .ErrorAvailable Then Call PropagateError(.HeldError, "CFTC_CommodityGroupings")
    End With
End Function
'Public Function GetFuturesTradingCode(baseCode$, wantedMonth As Byte, wantedYear As Long, forYaahoo As Boolean) As String
'
'    Dim contractmonth$, exchangeCode$
'
'    Err.Raise 1
''        January – F
''        February -G
''        March -h
''        April -J
''        May -K
''        June -M
''        July -n
''        August -Q
''        September -u
''        October -V
''        November -X
''        December -Z
'    contractmonth = Split("F,G,H,J,K,M,N,Q,U,V,X,Z", ",")(wantedMonth - 1)
'    GetFuturesTradingCode = baseCode & contractmonth & Format(wantedYear, "yy") & "." & exchangeCode
'
'End Function

Public Function GetCreatorPasswordsAndCodes() As Object
'======================================================================================================================
'   Summary: Returns a deserialized json object of creator passwords.
'======================================================================================================================
    Dim storageFilePath$, availableProperties As Object, json$, jp As New JsonParserB
    
    On Error GoTo Catch_Error
    ' > Creates an error if OneDrive isn't installed
    storageFilePath = Environ$("OneDriveConsumer") & "\COT_Related_Creator.json"
        
    If FileOrFolderExists(storageFilePath) Then
        json = CreateObject("Scripting.FileSystemObject").OpenTextFile(storageFilePath, 1).ReadAll
        Set availableProperties = jp.Deserialize(json)
    End If
Finally:
    Set GetCreatorPasswordsAndCodes = availableProperties
    Exit Function
Catch_Error:
    Resume Finally
End Function
Public Function ConvertOpenInterestTypeToName(oiType As OpenInterestType) As String
'======================================================================================================================
'   Summary: Returns the string representation of [oiType].
'   Paramaters:
'       [oiType] - OpenInterestType you want the textual representation of.
'======================================================================================================================
    Select Case oiType
        Case OpenInterestType.OptionsOnly
            ConvertOpenInterestTypeToName = "Options Only"
        Case OpenInterestType.FuturesOnly
            ConvertOpenInterestTypeToName = "Futures Only"
        Case FuturesAndOptions
            ConvertOpenInterestTypeToName = "Futures & Options"
    End Select
End Function
Public Function ConvertReportTypeEnum(reportEnumToConvert As ReportEnum) As String
    
    Select Case reportEnumToConvert
        Case ReportEnum.eLegacy: ConvertReportTypeEnum = "L"
        Case ReportEnum.eDisaggregated: ConvertReportTypeEnum = "D"
        Case ReportEnum.eTFF: ConvertReportTypeEnum = "T"
    End Select
    
End Function

Public Function ConvertInitialToReportTypeEnum(reportInitialToConvert As String) As String
    Select Case reportInitialToConvert
        Case "L"
            ConvertInitialToReportTypeEnum = ReportEnum.eLegacy
        Case "D"
            ConvertInitialToReportTypeEnum = ReportEnum.eDisaggregated
        Case "T"
           ConvertInitialToReportTypeEnum = ReportEnum.eTFF
        Case Else
            Err.Raise vbObjectError + 700, "ConvertInitialToReportTypeEnum", reportInitialToConvert & " < is Invalid."
    End Select
End Function
Public Function ReportEnumArray() As ReportEnum()
    Dim outputA(0 To 2) As ReportEnum
    outputA(0) = eLegacy
    outputA(1) = eDisaggregated
    outputA(2) = eTFF
    ReportEnumArray = outputA
End Function
Public Function CreateFieldInfoInstance(EditedName$, ColumnIndex As Integer, mappedName$, Optional IsMissing As Boolean = False, _
Optional isWanted As Boolean = False, Optional fromSocrata As Boolean = False) As FieldInfo
    
    Dim FI As New FieldInfo
    
    On Error GoTo FailedToConstruct
    FI.Constructor EditedName, ColumnIndex, mappedName, IsMissing, isWanted, fromSocrata
    Set CreateFieldInfoInstance = FI
    
    Exit Function
    
FailedToConstruct:
    PropagateError Err, "CreateFieldInfoInstance", "Failed to construct FieldInfo instance."
End Function
Public Function ConstructDynamicCheckBox(chx As msforms.CheckBox, Optional onColor&, Optional offColor&) As DynamicCheckBox
    
    Dim dymChx As New DynamicCheckBox
    
    dymChx.Constructor chx, onColor, offColor
    Set ConstructDynamicCheckBox = dymChx
    
End Function
Public Function ConvertSystemTimeToDate(timeToConvert As SystemTime) As Date
    With timeToConvert
        ConvertSystemTimeToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function
Public Function ConvertDateToSystemTime(timeToConvert As Date) As SystemTime
    With ConvertDateToSystemTime
        .wYear = Year(timeToConvert)
        .wMonth = Month(timeToConvert)
        .wDay = Day(timeToConvert)
        .wHour = Hour(timeToConvert)
        .wMinute = Minute(timeToConvert)
        .wSecond = Second(timeToConvert)
    End With
End Function
Public Function ConvertLocalDatetimeToUTC(convertedDate As Date) As Date
    #If Mac Then
        'ConvertLocalDatetimeToUTC = utc_ConvertDate(convertedDate, True)
    #Else
        Dim tz As TimeZoneInformation, utcTime As SystemTime
        Call GetTimeZoneInformation(tz)
        Call TzSpecificLocalTimeToSystemTime(tz, ConvertDateToSystemTime(convertedDate), utcTime)
        ConvertLocalDatetimeToUTC = ConvertSystemTimeToDate(utcTime)
    #End If
End Function
