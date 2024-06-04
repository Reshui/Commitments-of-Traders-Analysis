Attribute VB_Name = "Public_Functions"
Option Explicit
Sub SendEmailFromOutlook(Body As String, Subject As String, toEmails As String, ccEmails As String, bccEmails As String)
    Dim outApp As Object
    Dim outMail As Object
    On Error GoTo No_Outlook
    
    Set outApp = CreateObject("Outlook.Application")
    Set outMail = outApp.CreateItem(0)
 
    With outMail
        .To = toEmails
        .CC = ccEmails
        .BCC = bccEmails
        .Subject = Subject
        .HTMLBody = Body
        .send 'Send the email
    End With
 
    Set outMail = Nothing
    Set outApp = Nothing
    
    Exit Sub
    
No_Outlook:
    MsgBox "Microsoft Outlook isn't installed."
End Sub

Function Quote_Delimiter_Array(ByVal inputA As String, Delimiter As String, Optional N_Delimiter As String = "*")

    Dim x As Long, SA() As String

    If InStrB(1, inputA, Chr(34)) = 0 Then 'if there are no quotation marks then split with the supplied delimiter
        
        Quote_Delimiter_Array = Split(inputA, Delimiter)
        Exit Function

    Else
        
        SA = Split(inputA, Chr(34))
        
        For x = LBound(SA) To UBound(SA) Step 2
            SA(x) = Replace(SA(x), Delimiter, N_Delimiter)
        Next x
        
        Quote_Delimiter_Array = Split(Join(SA), N_Delimiter)
        
    End If

End Function
Public Sub Courtesy()

    With Application

        If Not UUID Then
            .StatusBar = "Brought to you by MoshiM. Please consider donating to support the continued development of this project."
        Else
            .StatusBar = vbNullString
        End If
    
    End With

End Sub
Public Sub OpenExcelWorkbook(File_Name_And_Path As String, Optional keepVisible As Boolean = False) 'Open specific file

    Dim WBOpen As Workbook

    Set WBOpen = Workbooks.Open(File_Name_And_Path)      'Opens the Excel file/csv
    WBOpen.Windows(1).Visible = keepVisible            'Files will not be visible

End Sub
Public Function LastModified(ByRef filePath As String) As Date
   
    Dim KK As Object
    
    LastModified = vbNull

    If LenB(Trim$(filePath)) = 0 Then Exit Function

    Set KK = CreateObject("Scripting.FileSystemObject")
    
    With KK 'CreateObject("Scripting.FileSystemObject")
    
        If .FileExists(filePath) Then LastModified = .GetFile(filePath).DateLastModified
        
    End With

End Function
Function UTC() As Date
'===================================================================================================================
    'Purpose: Gets the current UTC time.
    'Inputs:
    'Outputs: The current UTC datetime.
    'Note:
'===================================================================================================================

    #If Mac Then
        UTC = MacScript("set UTC to (current date) - (time to GMT)")
    #Else

        With CreateObject("WbemScripting.SWbemDateTime")
            .SetVarDate Now
            UTC = .GetVarDate(False)
        End With

    #End If
    
'Debug.Print UTC

End Function
Function HasKey(col As Collection, key As String) As Boolean
'===================================================================================================================
    'Purpose: Determines if a given collection has a specific key.
    'Inputs: col - Collection to check.
    '        Key - key to check col for.
    'Outputs: True or false.
'===================================================================================================================

    Dim V As Boolean
    
    On Error GoTo Exit_Function
    V = IsObject(col.Item(key))
    HasKey = Not IsEmpty(V)

Exit_Function:
    'The key doesn't exist.
End Function

Function IsPowerQueryAvailable() As Boolean 'Determine if Power Query is available...use for when less than EXCEL 2016
    
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
    Dim URL As String, QT As QueryTable, Disclaimer As Shape, My_Info As String, DL As String
    
    #If DatabaseFile Then
        If Variable_Sheet.Range("Github_Version").Value2 = True Then Exit Sub
    #End If
    
    On Error GoTo EXIT_DN_List

    DL = vbNewLine & vbNewLine

    My_Info = "Contact Email:   MoshiM_UC@outlook.com" & DL & _
                            "Skills:  C#, Python, Excel VBA, SQL, Data Analysis and Web Scraping." & DL & _
                            "Feel free to contact me for both personal and work related jobs."

    URL = Replace("https://www.dropbox.com/s/g75ij0agki217ow/CT%20Donators.txt?dl=0", _
            "www.dropbox.com", "dl.dropboxusercontent.com") 'URL leads to external text file
        
    Target_T.TextFrame.Characters.Text = vbNullString 'Clear text from shape

    Set QT = Query_W.QueryTables.Add("TEXT;" & URL, Query_W.Range("A1")) 'Assign object to Variable

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

        Set Disclaimer = .Parent.Shapes("Disclaimer")
        
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
Public Function CFTC_Release_Dates(Find_Latest_Release As Boolean) As Date
'===================================================================================================================
    'Purpose: Finds a wanted release date.
    'Inputs: Find_Latest_Release - If true then find the most recent release; else find next release.
    'Outputs: Wanted Date.
'===================================================================================================================

    Dim Data_Release As Date, x As Byte, Y As Byte, INTE_D As Date, rs As Variant, _
    Time_Zones As Variant, EST As Date, Local_Time As Date, YearN As Long, DayN As Byte
    
    Dim EST_To_Local_Difference As Long, EST_Current_Time As Date
    
    With Variable_Sheet
        Time_Zones = .ListObjects("Time_Zones").DataBodyRange.Value2 'This Query is refrshed on Workbook Open
                rs = .ListObjects("Release_Schedule").DataBodyRange.Value2 'Array of Release Dates
    End With
    
    With WorksheetFunction
    
        EST = .VLookup("EST Time", Time_Zones, 2, False)
        On Error GoTo assign_local_time_to_now
        
        Local_Time = .VLookup("Local Time", Time_Zones, 2, False)
        On Error GoTo 0
        
    End With
    
    EST_To_Local_Difference = DateDiff("h", EST, Local_Time, vbSunday, vbFirstJan1)
    
    EST_Current_Time = DateAdd("h", -EST_To_Local_Difference, Now)
    
    For x = 1 To UBound(rs, 1)
    
        If IsNumeric(rs(x, 1)) Then 'Checking in first column for Year
            YearN = CInt(rs(x, 1))
        Else
        
            For Y = 2 To UBound(rs, 2) 'Start from 2nd Column
            
                If LenB(rs(x, Y)) > 0 Then 'Get the Release time in GMT
                    
                    DayN = CByte(Replace(rs(x, Y), "*", vbNullString))
                    
                    'INTE_D = DateSerial(YearN, rs(X, 1), DayN) + TimeSerial(15, 30, 0) 'Date and time 15:30 EST
                    
                    INTE_D = DateValue(rs(x, 1) & " " & DayN & ", " & YearN) _
                             + TimeSerial(15, 30, 0) 'Date and time 15:30 EST
                    
                    If Not Find_Latest_Release Then 'If finding the next release
                    
                        If INTE_D > EST_Current_Time Then
                            Data_Release = INTE_D
                            Exit For
                        End If
                        
                    Else                'If looking for the previous release date and time
                    
                        If INTE_D > EST_Current_Time Then
                            Exit For
                        Else
                            Data_Release = INTE_D
                        End If
                        
                    End If
                    
                End If
                
            Next Y
            
            If INTE_D > EST_Current_Time Then Exit For
            
        End If
        
    Next x
    
    If Data_Release = TimeSerial(0, 0, 0) Then Data_Release = INTE_D
    
    CFTC_Release_Dates = DateAdd("h", EST_To_Local_Difference, Data_Release) 'Latest Release Date in Local Time
    
    Exit Function

assign_local_time_to_now:

    Local_Time = Now
    Resume Next
    
End Function

Public Function UUID() As Boolean
'===================================================================================================================
    'Purpose: Determines if user is on creator computer.
    'Inputs:
    'Outputs: Boolean representation of whether or not file is being run on creator computer.
'===================================================================================================================

    Dim creatorProperties As Collection, storedSerialNumber As Long, isOnUserComputer As Boolean
    
    Const Function_Value_Key As String = "Creator_Computer_?"
    
    #If Mac Then
        Exit Function
    #End If
    
    On Error GoTo Collection_Lacks_Key
    UUID = ThisWorkbook.Event_Storage(Function_Value_Key)
    Exit Function
    
Load_Password_File:

    On Error GoTo Exit_UUID
        
    Set creatorProperties = GetCreatorPasswordsAndCodes()
    
    With ThisWorkbook
        If Not creatorProperties Is Nothing Then
            storedSerialNumber = CLng(creatorProperties("DRIVE_SERIAL_NUMBER"))
            isOnUserComputer = (DoesDriveSerialNumberMatch(storedSerialNumber) And Environ("COMPUTERNAME") = "CAMPBELL-PC")
        End If
        .Event_Storage.Add isOnUserComputer, Function_Value_Key
    End With
    
    UUID = isOnUserComputer
    Exit Function

Exit_UUID:

    ThisWorkbook.Event_Storage.Add False, Function_Value_Key
    UUID = False
    Exit Function

Collection_Lacks_Key:
    Resume Load_Password_File
End Function
Public Function DoesDriveSerialNumberMatch(My_Serial As Long) As Boolean

    Dim FS As Object, D As Drive, x As Long, TT As String

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
Function MAC_Identifier(MAC_Address_Input As String) As Boolean
 
'Declaring the necessary variables.
    Dim strComputer     As String
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
    'Purpose: Creates a collection of all available Contract instances.
    'Outputs: A collection of contract instances.
'===================================================================================================================
    Dim This_C As Collection
    
    On Error GoTo Propagate
    #If DatabaseFile Then
        Set This_C = GetContractInfo_DbVersion
    #Else
    
        Const codeColumn As Byte = 1, yahooColumn As Byte = 3, nameColumn As Byte = 2
        Dim priceSymbolsTable As Range, currentSymbol As String, Yahoo_Finance_Ticker As Boolean
        
        Dim Code As String, contractData As ContractInfo, rowIndex As Long
        
        Set priceSymbolsTable = Symbols.ListObjects("Symbols_TBL").DataBodyRange

        Dim WS As Worksheet, lo As ListObject
            
        Set This_C = New Collection
        
        For Each WS In ThisWorkbook.Worksheets
        
            For Each lo In WS.ListObjects
            
                With lo

                    If .Name Like "CFTC_*" Or .Name Like "ICE_*" Then
                    
                        Code = Right$(.Name, Len(.Name) - InStr(1, .Name, "_", vbBinaryCompare))
                        
                        On Error Resume Next
                            currentSymbol = WorksheetFunction.VLookup(Code, priceSymbolsTable, 3, False)
                        On Error GoTo 0
                        
                        Yahoo_Finance_Ticker = LenB(currentSymbol) > 0
                        
                        Set contractData = New ContractInfo
                        contractData.InitializeContract Code, currentSymbol, Yahoo_Finance_Ticker, lo
                        
                        This_C.Add contractData, Code
                        
                        currentSymbol = vbNullString
                    
                    End If
                    
                End With
                
            Next lo
            
        Next WS

    #End If

    Set GetAvailableContractInfo = This_C
    Exit Function
Propagate:
    Call PropagateError(Err, "GetAvailableContractInfo")
End Function

Public Function IsLoadedUserform(User_Form_Name As String) As Boolean
'===================================================================================================================
    'Purpose: Determines if a UserForm is loaded based on its name.
    'Inputs: User_Form_Name - Name to check for.
    'Outputs: Stores the current time on the Variable_Sheet along with the local time on the running environment.
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
    
    Dim output As New Collection, g As Long, nDimensions As Byte, columnCount As Variant
    
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
        For g = LBound(data) To UBound(data)
            If useValuesAsKey Then
                .Add data(g), CStr(data(g))
            Else
                .Add data(g)
            End If
        Next g
    End With
    
    Set ConvertArrayToCollection = output
    
End Function
Public Function EditDatabaseNames(DatabaseName As String) As String

    Dim lcaseVersion As String
    
    lcaseVersion = LCase$(DatabaseName)
    
    If InStrB(1, lcaseVersion, "yyyy") > 0 Then
        lcaseVersion = "report_date_as_yyyy_mm_dd"
    Else
        lcaseVersion = Replace$(lcaseVersion, " ", "_")
        lcaseVersion = Replace$(lcaseVersion, Chr(34), vbNullString)
        lcaseVersion = Replace$(lcaseVersion, "%", "pct")
        lcaseVersion = Replace$(lcaseVersion, "=", "_")
        lcaseVersion = Replace$(lcaseVersion, "(", "_")
        lcaseVersion = Replace$(lcaseVersion, ")", vbNullString)
        lcaseVersion = Replace$(lcaseVersion, "-", "_")
        lcaseVersion = Replace$(lcaseVersion, "commercial", "comm")
        lcaseVersion = Replace$(lcaseVersion, "reportable", "rept")
        lcaseVersion = Replace$(lcaseVersion, "total", "tot")
        lcaseVersion = Replace$(lcaseVersion, "concentration", "conc")
        lcaseVersion = Replace$(lcaseVersion, "spreading", "spread")
        lcaseVersion = Replace$(lcaseVersion, "_lt_", "_le_")
        
        Do While InStrB(1, lcaseVersion, "__") > 0
            lcaseVersion = Replace$(lcaseVersion, "__", "_")
        Loop
        
        lcaseVersion = Replace$(lcaseVersion, "open_interest_oi", "oi")
        lcaseVersion = Replace$(lcaseVersion, "open_interest", "oi")
        
    End If
        
    EditDatabaseNames = lcaseVersion

End Function

Public Function GETNUMBER(inputValue As String, Optional index As Byte = 1) As Long
    
    Dim outputNumber As String, I As Byte, _
    currentCharacter As String, addToOutput As Boolean, decimalCount As Byte, threeCharacters As String
    
    Dim numbersCollection As New Collection, finishedNumber As Boolean, lengthOfText As Byte
    
    lengthOfText = Len(inputValue)
    
    For I = 1 To lengthOfText
    
        currentCharacter = Mid$(inputValue, I, 1)
        
        If currentCharacter Like "#" Then
            addToOutput = True
            finishedNumber = False
        ElseIf I > 1 Then
            threeCharacters = Mid$(inputValue, I - 1, 3)
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
        
        If (I = lengthOfText Or finishedNumber) And LenB(outputNumber) > 0 Then
            finishedNumber = False
            numbersCollection.Add outputNumber
            outputNumber = vbNullString
        End If
        
    Next I
    
    GETNUMBER = numbersCollection(index) * 1
        
End Function
Public Function GetExpectedLocalFieldInfo(reportType As String, filterUnwantedFields As Boolean, reArrangeToReflectSheet As Boolean, includePrice As Boolean, Optional adjustIndexes As Boolean = True) As Collection
'=============================================================================================
'   Summary: Generates FieldInfo instances for field names stored on Variable Sheet.
'=============================================================================================
    Dim T As Byte, localCopyOfColumnNames() As Variant, columnMap As New Collection, FI As FieldInfo
    
    localCopyOfColumnNames = GetAvailableFieldsTable(reportType).DataBodyRange.Value2
    
    With columnMap
    
        For T = 1 To UBound(localCopyOfColumnNames, 1)
            If Not filterUnwantedFields Or localCopyOfColumnNames(T, 2) = True Then
                Set FI = New FieldInfo
                FI.Constructor EditDatabaseNames(CStr(localCopyOfColumnNames(T, 1))), T, CStr(localCopyOfColumnNames(T, 1)), False
                .Add FI, FI.editedName
            End If
        Next T
        
        If reArrangeToReflectSheet Then
        
            If filterUnwantedFields Then
                Set FI = .Item("cftc_contract_market_code")
                .Remove FI.editedName
                .Add FI, FI.editedName
    
                Set FI = .Item("report_date_as_yyyy_mm_dd")
                .Remove FI.editedName
                .Add FI, FI.editedName, 1
            End If
            
            If adjustIndexes Then
                For T = 1 To .count
                    .Item(T).AdjustColumnIndex T
                Next T
            End If
            
        End If
        
        If includePrice Then
            Set FI = New FieldInfo
            FI.Constructor "price", 1 + IIf(filterUnwantedFields, .count, UBound(localCopyOfColumnNames, 1)), "Price", False
            .Add FI, "price"
        End If
    
    End With
    
    Set GetExpectedLocalFieldInfo = columnMap
    
End Function
Public Function GetAvailableFieldsTable(reportType As String) As ListObject
    Set GetAvailableFieldsTable = Variable_Sheet.ListObjects(reportType & "_User_Selected_Columns")
End Function
Public Function CFTC_CommodityGroupings() As Collection
'============================================================================================================
'Queries the Socrata API to get commodity groups and subgroups for each contract.
'============================================================================================================
    Dim CC As New Collection, apiCode As String, reportKey As String, _
    apiURL As String, dataFilters As String, dataQuery As QueryTable, _
    retrievedData() As Variant, apiData() As Variant, iRow As Long, appProperties As Collection
    
    Const apiBaseURL As String = "https://publicreporting.cftc.gov/resource/"
    Const reportType As String = "L", queryReturnLimit As Long = 5000, getFuturesAndOptions As Boolean = True
    
    Set appProperties = DisableApplicationProperties(True, False, True)
    
    Dim columnTypes(0 To 2) As Variant
    
    ' The query table needs to import contract codes as text.
    For iRow = LBound(columnTypes) To UBound(columnTypes)
        columnTypes(iRow) = xlTextFormat
    Next iRow
    
    ' Creates a collection of api codes keyed to their report type.
    For iRow = 0 To 2
    
        reportKey = Array("L", "D", "T")(iRow)
        
        If getFuturesAndOptions Then
            CC.Add Array("jun7-fc8e", "kh3c-gbw2", "yw9f-hn96")(iRow), reportKey
        Else
            CC.Add Array("6dca-aqww", "72hh-3qpy", "gpe5-46if")(iRow), reportKey
        End If
        
    Next iRow
    
    apiCode = CC(reportType)
            
    dataFilters = "?$select=cftc_contract_market_code,commodity_group_name,commodity_subgroup_name" & _
                    "&$group=cftc_contract_market_code,commodity_group_name,commodity_subgroup_name" & _
                    "&$limit=" & queryReturnLimit
    
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
        .WorkbookConnection.RefreshWithRefreshAll = False

        On Error GoTo Finally
        .Refresh False
        
        With .ResultRange
            ' >1 since column names will always be returned.
            If .Rows.count > 1 Then
                apiData = .Range(.Cells(2, 1), .Cells(.Rows.count, .columns.count)).Value2
            End If
            .ClearContents
        End With
        
    End With
    
    Set CC = New Collection
    On Error GoTo Catch_Duplicate
    With CC
        For iRow = LBound(apiData, 1) To UBound(apiData, 1)
            .Add Array(apiData(iRow, 1), apiData(iRow, 2), apiData(iRow, 3)), apiData(iRow, 1)
Next_Grouping:
        Next
    End With
    
    On Error GoTo 0
    
    Set CFTC_CommodityGroupings = CC

Finally:

    EnableApplicationProperties appProperties
    
    If Not dataQuery Is Nothing Then
        With dataQuery
            .WorkbookConnection.Delete
            .Delete
        End With
    End If
    
    With Err
        If .Number <> 0 Then Call PropagateError(Err, "CFTC_CommodityGroupings()")
    End With
    
    Exit Function
    
Catch_Duplicate:

    Resume Next_Grouping
    
QueryTable_Already_Exists:

    With QueryT.QueryTables(reportType & "_CFTC_API_Weekly Combined:" & getFuturesAndOptions) '
        .WorkbookConnection.Delete
        .Delete
    End With
    
    Resume
    
End Function
Public Function GetFuturesTradingCode(baseCode As String, wantedMonth As Byte, wantedYear As Long, forYaahoo As Boolean) As String
    
    Dim contractmonth As String, exchangeCode As String
    
    Err.Raise 1
'        January – F
'        February -G
'        March -h
'        April -J
'        May -K
'        June -M
'        July -n
'        August -Q
'        September -u
'        October -V
'        November -X
'        December -Z
    contractmonth = Split("F,G,H,J,K,M,N,Q,U,V,X,Z", ",")(wantedMonth - 1)
    GetFuturesTradingCode = baseCode & contractmonth & Format(wantedYear, "yy") & "." & exchangeCode
        
End Function

Public Function GetCreatorPasswordsAndCodes() As Collection
    
    Dim storageFilePath As String, x As Long, fileContents As String, availableProperties As Collection, keyAndValue() As String
    
    On Error GoTo Catch_Error
    ' > Creates an error if OneDrive isn't installed
    storageFilePath = Environ("OneDriveConsumer") & "\COT_Related_Creator.txt"
        
    If FileOrFolderExists(storageFilePath) Then
        Set availableProperties = New Collection
        x = FreeFile
        'Open Stored text file and retrieve string for comparison.
        With availableProperties
            Open storageFilePath For Input As #x
                Do While Not EOF(x)
                    Input #x, fileContents
                    keyAndValue = Split(fileContents, ":", 2)
                    .Add keyAndValue(1), keyAndValue(0)
                Loop
            Close #x
        End With

    End If
Finally:
    Set GetCreatorPasswordsAndCodes = availableProperties
    Exit Function
Catch_Error:
    Resume Finally
End Function

#If Not DatabaseFile Then

    Public Function ReturnCftcTable(WS As Worksheet) As ListObject
    '===================================================================================================================
        'Purpose: Checks to see if a cftc table exists on a worksheet(WS) based on table names.
        'Inputs: WS - worksheet to check.
        'Outputs: CFTC Listobject or Nothing.
        'Notes - ONly applicable to Non - Database version of file.
    '===================================================================================================================
        Dim Item As ListObject, tableName As String
        
        For Each Item In WS.ListObjects
            tableName = Item.Name
            If tableName Like "CFTC_*" Or tableName Like "ICE_*" Then
                Set ReturnCftcTable = Item
                Exit Function
            End If
        Next Item
        
    End Function
    
    Public Function IsWorkbookForFuturesAndOptions() As Boolean
        IsWorkbookForFuturesAndOptions = Variable_Sheet.Range("Combined_Workbook").Value2
    End Function
    
    Public Function ReturnReportType() As String
        ReturnReportType = Variable_Sheet.Range("Report_Type").Value2
    End Function
    
#End If
