Attribute VB_Name = "Public_Functions"

Public Sub Run_This(WB As Workbook, ScriptN As String)

    Application.Run "'" & WB.Name & "'!" & ScriptN

End Sub
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
Sub Re_Enable()
Attribute Re_Enable.VB_Description = "Resets application variables that may interfere with Workbook display or calculation."
Attribute Re_Enable.VB_ProcData.VB_Invoke_Func = " \n14"

    With Application
    
        If .Calculation <> xlCalculationAutomatic Then .Calculation = xlCalculationAutomatic
        If .ScreenUpdating = False Then .ScreenUpdating = True
        If .DisplayStatusBar = False Then .DisplayStatusBar = True
        If .EnableEvents = False Then .EnableEvents = True
        
    End With

End Sub

Sub Remove_Worksheet_Formatting()
Attribute Remove_Worksheet_Formatting.VB_Description = "Removes all worksheet formatting from the currently active worksheet."
Attribute Remove_Worksheet_Formatting.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_Conditional_Formats__On_Worksheet Macro
'
' Keyboard Shortcut: Ctrl+Shift+X
    Cells.FormatConditions.Delete

End Sub
Sub ZoomToRange(ByVal ZoomThisRange As Range, ByVal PreserveRows As Boolean, WB As Workbook)

Application.ScreenUpdating = False

Dim Wind As Window

Set Wind = ActiveWindow

Application.GoTo ZoomThisRange.Cells(1, 1), True

With ZoomThisRange
    If PreserveRows = True Then
        .Resize(.Rows.Count, 1).Select
    Else
        .Resize(1, .Columns.Count).Select
    End If
End With

With Wind
    .Zoom = True
    .VisibleRange.Cells(1, 1).Select
End With

    If Not WB.Last_Used_Sheet Is Nothing And UUID Then 'accounting for if the variable has not been declared for normal use
        'do nothing
    Else
        Application.ScreenUpdating = True
    End If

End Sub
Function Quote_Delimiter_Array(ByVal InputA As String, Delimiter As String, Optional N_Delimiter As String = "*")

Dim X As Long, SA() As String

If InStr(1, InputA, Chr(34)) = 0 Then 'if there are no quotation marks then split with the supplied delimiter
    
    Quote_Delimiter_Array = Split(InputA, Delimiter)
    Exit Function

Else
    
    SA = Split(InputA, Chr(34))
    
    For X = LBound(SA) To UBound(SA) Step 2
        SA(X) = Replace(SA(X), Delimiter, N_Delimiter)
    Next X
    
    Quote_Delimiter_Array = Split(Join(SA), N_Delimiter)
    
End If

End Function

Public Function Change_Delimiter_Not_Between_Quotes(ByRef Current_String As Variant, ByVal Delimiter As String, Optional ByVal Changed_Delimiter As String = ">Ý") As Variant
    
'returns a 0 based array
    
Dim String_Array() As String, X As Long, Right_CHR As String

    If InStr(1, Current_String, Chr(34)) = 0 Then 'if there are no quotation marks then split with the supplied delimiter
        
        Change_Delimiter_Not_Between_Quotes = Split(Current_String, Delimiter)
        Exit Function

    End If
    
    Right_CHR = Right(Changed_Delimiter, 1) 'RightMost character in at least 2 character string that will be used as a replacement delimiter

    'Replace ALL quotation marks with the ChangedDelimiter[Quotation mark] EX: " --> $+
    Current_String = Replace(Current_String, Chr(34), Changed_Delimiter)

    String_Array = Split(Current_String, Left(Changed_Delimiter, 1))
    '1st character of Changed_Delimiter will be used to delimit a new array
    'element [0] will be an empty string if the first value in the delmited string begins with a Quotation mark.
    
    For X = LBound(String_Array) To UBound(String_Array) 'loop all elements of the array

        If Left(String_Array(X), 1) = Right_CHR And Not Left(String_Array(X), 2) = Right_CHR & Delimiter Then
            'If the string contains a valid comma
            'Checked by if [the First character is the 2nd Character in the Changed Delimiter] and the 2nd character isn't the delimiter
            'Then offset the string by 1 character to remove the 2nd portion of the changed Delimiter
            String_Array(X) = Right(String_Array(X), Len(String_Array(X)) - 1)
        
        Else
        
            If Left(String_Array(X), 1) = Right_CHR Then 'If 1st character = 2nd portion of the Changed Delimiter
                                                         'Then offset string by 1 and then repalce all [Delimiter]
                String_Array(X) = Replace(Right(String_Array(X), Len(String_Array(X)) - 1), Delimiter, Changed_Delimiter)
            
            Else 'Just replace
                
                String_Array(X) = Replace(String_Array(X), Delimiter, Changed_Delimiter)
            
            End If
            
        End If
        
    Next X
    'Join the Array elements back together {Do not add another delimiter] and split with the changed Delimiter
    Change_Delimiter_Not_Between_Quotes = Split(Join(String_Array), Changed_Delimiter)
    
    Erase String_Array
End Function
Public Function entUnZip1File(ByVal strZipFilename As Variant, ByVal strDstDir As Variant, ByVal strFilename As Variant) 'Opens zip file
                                                'path of file     path of Folder containing file              name of specified file within .zip file
    Const glngcCopyHereDisplayProgressBox = 256
'
  Dim intOptions, objShell, objSource, objTarget As Object
'
' Create the required Shell objects
  Set objShell = CreateObject("Shell.Application")
'
' Create a reference to the files and folders in the ZIP file
  Set objSource = objShell.Namespace(strZipFilename).items.Item(strFilename)
'
' Create a reference to the target folder
  Set objTarget = objShell.Namespace(strDstDir)
'
  intOptions = glngcCopyHereDisplayProgressBox
'
' UnZIP the files
  objTarget.CopyHere objSource, intOptions
'
' Release the objects
  Set objSource = Nothing
  Set objTarget = Nothing
  Set objShell = Nothing
'
entUnZip1File = 1
'
End Function
Public Function Quicksort(ByRef vArray As Variant, arrLbound As Long, arrUbound As Long)
'Sorts a one-dimensional VBA array from smallest to largest
'using a very fast quicksort algorithm variant.
Dim pivotVal As Variant
Dim vSwap    As Variant
Dim Temporary_Low   As Long
Dim Temporary_High    As Long
 
Temporary_Low = arrLbound

Temporary_High = arrUbound

pivotVal = vArray((arrLbound + arrUbound) \ 2) 'The element in the middle of the array
 
While (Temporary_Low <= Temporary_High) 'divide

   While (vArray(Temporary_Low) < pivotVal And Temporary_Low < arrUbound)
      
      Temporary_Low = Temporary_Low + 1
      
   Wend
  
   While (pivotVal < vArray(Temporary_High) And Temporary_High > arrLbound)
      Temporary_High = Temporary_High - 1
   Wend
 
   If (Temporary_Low <= Temporary_High) Then
      vSwap = vArray(Temporary_Low)
      vArray(Temporary_Low) = vArray(Temporary_High)
      vArray(Temporary_High) = vSwap
      Temporary_Low = Temporary_Low + 1
      Temporary_High = Temporary_High - 1
   End If
   
Wend
 
  If (arrLbound < Temporary_High) Then Quicksort vArray, arrLbound, Temporary_High 'conquer
  If (Temporary_Low < arrUbound) Then Quicksort vArray, Temporary_Low, arrUbound 'conquer
  
End Function
Public Function Get_File(file As String, SaveFilePathAndName As String)

Dim oStrm As Object, WinHttpReq As Object ', Extension As String, File_Name As String
    
Set WinHttpReq = CreateObject("Msxml2.ServerXMLHTTP")

    WinHttpReq.Open "GET", file, False
    WinHttpReq.send
    file = WinHttpReq.responseBody
    
    If WinHttpReq.Status = 200 Then
        
        'Application.StatusBar = "Retrieving file from: " & file
        
        Set oStrm = CreateObject("ADODB.Stream")
        With oStrm
            .Open
            .Type = 1
            .write WinHttpReq.responseBody
            .SaveToFile SaveFilePathAndName, 2 ' 1 = no overwrite, 2 = overwrite
            .Close
        End With
        
        'Application.StatusBar = vbNullString
        
    End If
    
'AppleScript:
'set u to "http://download.finance.yahoo.com/d/quotes.csv?s=AAPL&f=sl1d1t1c1ohgv&e=.csv"
'do shell script "curl -L -s " & File & " > ~/desktop/quotes.csv"

End Function
Public Function Courtesy()

With Application

    If Not UUID Then
        .StatusBar = "Brought to you by MoshiM. Please consider donating to support the continued development of this project."
    Else
        .StatusBar = vbNullString
    End If
 
End With

End Function
Public Function Open_File(File_Name_And_Path) 'Open specific file

Dim WBOpen As Workbook

    Set WBOpen = Workbooks.Open(File_Name_And_Path)      'Opens the Excel file/csv

    WBOpen.Windows(1).Visible = False            'Files will not be visible

End Function
Public Function LastModified(ByRef filePath As String) As Date
    'Set a default value
    Dim KK As Object
    
    LastModified = vbNull

    If Len(Trim$(filePath)) = 0 Then Exit Function

    Set KK = CreateObject("Scripting.FileSystemObject")
    
    With KK 'CreateObject("Scripting.FileSystemObject")
    
        If .FileExists(filePath) Then LastModified = .GetFile(filePath).DateLastModified
        
    End With

End Function
Function UTC() As Date

#If Mac Then

    UTC = MacScript("set UTC to (current date) - (time to GMT)")

#Else

    Dim dt As Object
    
    Set dt = CreateObject("WbemScripting.SWbemDateTime")
        
    With dt
        .SetVarDate Now
        UTC = .GetVarDate(False)
    End With

#End If
    
'Debug.Print UTC

End Function
Function HasKey(col As Collection, Key As String) As Boolean
    Dim V As Variant
    
    On Error GoTo Exit_Function
    
    V = IsObject(col.Item(Key))
    HasKey = Not IsEmpty(V)

Exit_Function:
    'If Err.Number <> 0 Then Err.Clear
End Function
Sub Increment_Progress(Label As MSForms.Label, New_Width As Double, Loop_Percentage As Double)
   
    Label.Width = New_Width
    
    Progress_Bar.Caption = "Completed : " & Round(Loop_Percentage * 100, 0) & " %"
    
    DoEvents
    
End Sub

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
Dim URL As String, QT As QueryTable, Disclaimer As Shape

If Range("Github_Version") = True Then Exit Sub

On Error GoTo EXIT_DN_List

Const DL As String = vbNewLine & vbNewLine

Const My_Info As String = "Contact Email:   MoshiM_UC@outlook.com" & DL & _
                          "Skills:  Python, Excel VBA, SQL, Data Analysis and Web Scraping." & DL & _
                          "Feel free to contact me for both personal and work related jobs."

URL = Replace("https://www.dropbox.com/s/g75ij0agki217ow/CT%20Donators.txt?dl=0", _
        "www.dropbox.com", "dl.dropboxusercontent.com") 'URL leads to external text file
      
With Target_T
    Target_T.TextFrame.Characters.Text = vbNullString 'Clear text from shape
End With

Set QT = Query_W.QueryTables.Add("TEXT;" & URL, Query_W.Range("A1")) 'Assign object to Variable

With QT

    .BackgroundQuery = False
    .SaveData = False
    .AdjustColumnWidth = False
    .RefreshStyle = xlOverwriteCells
    .WorkbookConnection.Name = "Donation_Information"
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

If Not QT Is Nothing Then
    With QT
        .WorkbookConnection.Delete
        .Delete
    End With
End If

Exit Sub

EXIT_DN_List:
    
    On Error Resume Next
    
    Target_T.TextFrame.Characters.Text = My_Info
    
    Resume Remove_QueryTable
    
End Sub
Public Function CFTC_Release_Dates(Find_Latest_Release As Boolean) As Date

Dim Data_Release As Date, X As Long, Y As Long, INTE_D As Date, rs As Variant, _
Time_Zones As Variant, EST As Date, Local_Time As Date, YearN As Long, DayN As Long

Dim EST_Local_Difference As Long, EST_Current_Time As Date

With Variable_Sheet
    Time_Zones = .ListObjects("Time_Zones").DataBodyRange.Value2 'This Query is refrshed on Workbook Open
            rs = .ListObjects("Release_Schedule").DataBodyRange.Value2 'Array of Release Dates
End With

With WorksheetFunction
    EST = .VLookup("EST Time", Time_Zones, 2, False)
    On Error GoTo assign_local_time_to_now
    Local_Time = .VLookup("Local Time", Time_Zones, 2, False)
End With

If Hour(EST) = Hour(Local_Time) Then
    
    EST_Local_Difference = 0
    
ElseIf Local_Time > EST Then 'if further East then EST

    EST_Local_Difference = DateDiff("h", EST, Local_Time, vbSunday, vbFirstJan1)
    
ElseIf EST > Local_Time Then 'IF further West than EST

    EST_Local_Difference = DateDiff("h", Local_Time, EST, vbSunday, vbFirstJan1)

End If

EST_Current_Time = EST + TimeSerial(DateDiff("h", Local_Time, Now, vbSunday, vbFirstJan1), 0, 0)

For X = 1 To UBound(rs, 1)

    If IsNumeric(rs(X, 1)) Then 'Checking in first column for Years
        
        YearN = CLng(rs(X, 1))
    
    Else
    
        For Y = 2 To UBound(rs, 2) 'Start from 2nd Column
        
            If rs(X, Y) <> vbNullString Then 'Get the Release time in GMT
                
                DayN = CLng(Replace(rs(X, Y), "*", vbNullString))
                
                INTE_D = DateValue(rs(X, 1) & " " & DayN & ", " & YearN) _
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
    
Next X

If Data_Release = TimeSerial(0, 0, 0) Then Data_Release = INTE_D

CFTC_Release_Dates = Data_Release - TimeSerial(EST_Local_Difference, 0, 0) 'Latest Release Date in Local Time

Exit Function

assign_local_time_to_now:

    Local_Time = Now
    Resume Next
End Function

Public Function UUID() As Boolean

Dim Text_S As String, CMD_Output As String, X As Byte, cmd As String, _
MY_ID As String, Storage_File As String, PWD_A() As String, My_Serial_N As Long, MY_MAC_Address As String

Const Function_Value_Key As String = "Creator_Computer_?"

#If Mac Then
    Exit Function
#End If

On Error GoTo Collection_Lacks_Key

    UUID = ThisWorkbook.Event_Storage(Function_Value_Key)
    
    Exit Function

Load_Password_File: On Error GoTo Exit_UUID 'return False

Storage_File = Environ("ONEDRIVE") & "\C.O.T Password.txt" ' > Creates an error if OneDrive isn't installed
    
If Dir(Storage_File) <> vbNullString Then 'If stored password file exists

    With ThisWorkbook

        X = FreeFile
        
        Open Storage_File For Binary As #X 'Open Stored text file and retrieve string for comparison
        
            MY_ID = Space$(LOF(X))
            Get #X, , MY_ID
            
        Close #X
        
        PWD_A = Split(MY_ID, ",")
        
        My_Serial_N = CLng(PWD_A(3)) '4th
        
        MY_MAC_Address = PWD_A(4) '5th
        
        If GetSerialN(My_Serial_N) And Environ("COMPUTERNAME") = "CAMPBELL-PC" Then

            ThisWorkbook.Event_Storage.Add True, Function_Value_Key
        Else
            ThisWorkbook.Event_Storage.Add False, Function_Value_Key
            
        End If
        
    End With

Else
    
    ThisWorkbook.Event_Storage.Add False, Function_Value_Key
    
End If

UUID = ThisWorkbook.Event_Storage(Function_Value_Key)

Exit Function

Exit_UUID:

    ThisWorkbook.Event_Storage.Add False, Function_Value_Key
    UUID = False

    Exit Function

Collection_Lacks_Key:
    Resume Load_Password_File
'Debug.Print "Function UUID completed in " & Timer - TT & " seconds"

End Function
Public Function GetSerialN(My_Serial As Long) As Boolean

Dim FS As Object, D As Drive, X As Long, TT As String

On Error GoTo No_Scripting

With ThisWorkbook

    If .SerialN = 0 Then

        Set FS = CreateObject("Scripting.FileSystemObject")
        
        Set D = FS.GetDrive(FS.GetDriveName(FS.GetAbsolutePathName(Empty))) 'drvpath

        .SerialN = D.SerialNumber

    End If

    If .SerialN = My_Serial Then GetSerialN = True

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

Public Sub ChangeFilters(w As ListObject, ByRef filterArray)

With w.AutoFilter

    With .Filters
        ReDim filterArray(1 To .Count, 1 To 3)
        On Error GoTo Show_Data
        For f = 1 To .Count
            With .Item(f)
                If .On Then
                    filterArray(f, 1) = .Criteria1
                    If .Operator Then
                        filterArray(f, 2) = .Operator
                        filterArray(f, 3) = .Criteria2
                    End If
                End If
            End With
        Next
    End With
Show_Data:
    .ShowAllData
    
End With

End Sub
Public Sub RestoreFilters(w As ListObject, ByVal filterArray)

Dim col As Long

With w.DataBodyRange

    For col = 1 To UBound(filterArray, 1)
    
        If Not IsEmpty(filterArray(col, 1)) Then
            If filterArray(col, 2) Then
                .AutoFilter Field:=col, _
                    Criteria1:=filterArray(col, 1), _
                        Operator:=filterArray(col, 2), _
                    Criteria2:=filterArray(col, 3)
            Else
                .AutoFilter Field:=col, _
                    Criteria1:=filterArray(col, 1)
            End If
            
        End If
        
    Next

End With

End Sub
Public Function Get_Price_Symbols() As Collection
'======================================================================================================
'Generates an array of contract information within the workbook
'Array rows are (contract code, worksheet index, worksheet name, table object, current symbol)
'Columns 1-3 will be output to the Variable worksheet
'
'======================================================================================================

Dim i As Long, This_C As New Collection, Code As String

Dim SymbolA() As Variant, Current_Symbol As String, Yahoo_Finance_Ticker As Boolean ', Contract_Code_Column As Long

SymbolA = Symbols.ListObjects("Symbols_TBL").DataBodyRange.value

For i = LBound(SymbolA, 1) To UBound(SymbolA, 1)

    If Not (IsError(SymbolA(i, 1)) Or IsEmpty(SymbolA(i, 1))) Then
        
        If Not IsEmpty(SymbolA(i, 3)) Then 'Yahoo Finance
        
            Current_Symbol = SymbolA(i, 3)
            Yahoo_Finance_Ticker = True
            
        ElseIf Not IsEmpty(SymbolA(i, 4)) Then 'Stooq
        
            Current_Symbol = SymbolA(i, 4)
            Yahoo_Finance_Ticker = False
            
        End If
        
        If Current_Symbol <> vbNullString Then
        
            Code = SymbolA(i, 1)
            
            This_C.Add Array(Current_Symbol, Yahoo_Finance_Ticker), Code
               
            Code = vbNullString
            Current_Symbol = vbNullString
        
        End If
        
    End If
    
Next i

Set Get_Price_Symbols = This_C

'Debug.Print Timer - Start_Time
    
End Function
Public Function IsLoadedUserform(User_Form_Name As String) As Boolean

Dim frm As Object

For Each frm In VBA.UserForms
    If frm.Name = User_Form_Name Then
        IsLoadedUserform = True
        Exit Function
    End If
Next frm

End Function
Public Function Reverse_2D_Array(ByVal Data As Variant, Optional ByRef selected_columns As Variant)

    Dim X As Long, Y As Long, Temp(1 To 2) As Variant, Projected_Row As Long
    
    Dim LB2 As Long, UB2 As Long, Z As Long

    If IsMissing(selected_columns) Then
        LB2 = LBound(Data, 2)
        UB2 = UBound(Data, 2)
    Else
        LB2 = LBound(selected_columns)
        UB2 = UBound(selected_columns)
    End If
    
    For X = LBound(Data, 1) To UBound(Data, 1)
        
        Projected_Row = UBound(Data, 1) - (X - LBound(Data, 1))
        
        If Projected_Row <= X Then Exit For
        
        For Y = LB2 To UB2
            
            If IsMissing(selected_columns) Then
                Z = Y
            Else
                Z = selected_columns(Y)
            End If
            
            Temp(1) = Data(X, Z)
            Temp(2) = Data(Projected_Row, Z)
            
            Data(X, Z) = Temp(2)
            Data(Projected_Row, Z) = Temp(1)
            
        Next Y

    Next X

    Reverse_2D_Array = Data

End Function
Public Function COT_ABR_Match(COT_Type_Abbrev As String) As Range

    Dim Column_1 As Range
    
    Set Column_1 = Variable_Sheet.ListObjects("Report_Abbreviation").DataBodyRange.Columns(1)
    
    Set COT_ABR_Match = Column_1.Find(COT_Type_Abbrev, , xlValues, xlWhole)
    
End Function
