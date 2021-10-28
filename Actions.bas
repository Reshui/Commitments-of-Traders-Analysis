Attribute VB_Name = "Actions"
Private HUB_Date_Color As Long
Private Workbook_Is_Outdated As Boolean
Public Close_Workbook_Macro As Boolean
Public Cancel_WB_Close As Boolean
Public Enable_Creator_Mode As Boolean

Private Sub Run_These_Key_Binds()

Dim Key_Bind() As String, Procedure() As String, X As Long, WBN As String ', Saved_State As Boolean

Key_Bind = Split("^b,^+b,^s,^+c,^w", ",")

Procedure = Split("ToTheHub,To_Charts,Custom_Save,Navigation_Userform,Close_Workbook", ",")

WBN = "'" & ThisWorkbook.Name & "'!"

With Application

    For X = LBound(Key_Bind) To UBound(Key_Bind)
        .OnKey Key_Bind(X), WBN & Procedure(X)
    Next X
    
End With
 
End Sub
Private Sub Remove_Key_Binds()

Dim Key_Bind() As String, X As Long

Key_Bind = Split("^b,^+b,^s,^+c,^w", ",")

With Application

    For X = LBound(Key_Bind) To UBound(Key_Bind)
        .OnKey Key_Bind(X)
    Next X
    
End With

End Sub
Public Sub Remove_Images()

'this sub is only run if on creator computer

Dim Variant_OBJ() As Variant, Wall_Path As String, X As Long, T As Long, _
Their_HUB As String, Their_Weekly As String, WallP As New Collection, _
OPI As Shape, Shape_Group As GroupShapes, AnT As Shape, Variant_OBJ_Names As String, CWB As Boolean

If Not UUID Then
    Exit Sub
Else
    Enable_Creator_Mode = False
End If

Wall_Path = Environ("USERPROFILE") & "\Desktop\Wallpapers\"
    
CWB = Combined_Workbook(Variable_Sheet)

Variant_OBJ = Array(HUB, Variable_Sheet, Weekly)
    
For X = LBound(Variant_OBJ) To UBound(Variant_OBJ)
  
    For Each OPI In Variant_OBJ(X).Shapes
    
        With OPI
            
            Select Case LCase(.Name)
            
                Case "macro_check"
                
                    If Not ThisWorkbook.Last_Used_Sheet Is Nothing And Not .Visible Then .Visible = True
                    .ZOrder (msoBringToFront)
                    
                    'If saving then make this shape visible
                Case "object_group"
                    
                    Set Shape_Group = .GroupItems
                    
                    For Each AnT In Shape_Group
                    
                        With AnT
                        
                            Select Case .Name
                                Case "DN_List", "Diagnostic"
                                    .Visible = False
                                Case Else
                                    .Visible = True
                            End Select
                        
                        End With
                        
                    Next AnT
                
                Case "temp", "unbound", "holding"
                     
                Case "wallpaper_items"
                
                    .Visible = False
                    
                Case Else
                
                    If Not Variant_OBJ(X) Is Weekly Then .Visible = False
                    
            End Select
            
        End With
        
    Next OPI

Next X

Variant_OBJ = Array(QueryT, MAC_SH)

For X = LBound(Variant_OBJ) To UBound(Variant_OBJ)

    Variant_OBJ(X).Visible = xlSheetVeryHidden

Next X

End Sub
Public Sub Creator_Version()
Attribute Creator_Version.VB_Description = "Locked macro."
Attribute Creator_Version.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Variant_OBJ() As Variant, Wall_Path As String, X As Long, T As Long, OPI As Shape, _
MY_HUB As String, My_Weekly As String, My_Variables As String, WallP As New Collection

If Not UUID Then
    Exit Sub
Else
    Enable_Creator_Mode = True
End If
    
Wall_Path = Environ("USERPROFILE") & "\Desktop\Wallpapers\"

Variant_OBJ = Array(HUB, Variable_Sheet, Weekly)
 
 For X = LBound(Variant_OBJ) To UBound(Variant_OBJ)
 
    For Each OPI In Variant_OBJ(X).Shapes
    
        With OPI
    
            Select Case LCase(.Name)
                
                Case "macro_check"
                
                Case "macros", "patch", "wallpaper_items"
                
                    .Visible = True
                    
                Case "object_group"
                    
                    For T = 1 To .GroupItems.Count 'make everything but Email visible
                        
                        With .GroupItems(T)
                        
                            Select Case .Name
                            
                                Case "Feedback", "DN_List", "Disclaimer", "DropBox Folder", "Diagnostic"
                                
                                    .Visible = False
                                    
                                Case Else
                                
                                    .Visible = True
                                    
                            End Select
                            
                        End With
                        
                    Next T
                
                Case Else
                
                    If Not Variant_OBJ(X) Is Weekly Then .Visible = False
                    
            End Select
        
        End With
                
    Next OPI
    
Next X

With Variable_Sheet 'load wallpaper strings into array and make worksheet visible
    Variant_OBJ = .ListObjects("Wallpaper_Selection").DataBodyRange.Value2
    .Visible = xlSheetVisible
End With

With WorksheetFunction 'load array strings to variables
    MY_HUB = Wall_Path & .VLookup("My_HUB", Variant_OBJ, 2, 0)
    My_Weekly = Wall_Path & .VLookup("My_Weekly", Variant_OBJ, 2, 0)
    My_Variables = Wall_Path & .VLookup("This_Sheet", Variant_OBJ, 2, 0)
End With

With WallP
    .Add Array(Weekly, My_Weekly)
    .Add Array(HUB, MY_HUB)
    .Add Array(Variable_Sheet, My_Variables)
End With

For X = 1 To WallP.Count
    
    Variant_OBJ = WallP(X)
    
    If Dir(Variant_OBJ(1)) <> vbNullString Then
        Variant_OBJ(0).SetBackgroundPicture Filename:=Variant_OBJ(1)
    Else
        MsgBox "Wallpaper not found for " & Variant_OBJ(0).Name
    End If
    
Next X

End Sub
Public Sub Worksheet_Protection_Toggle(Optional This_Sheet As Worksheet, Optional Allow_Color_Change As Boolean = True, Optional Manual_Trigger As Boolean = True)

Dim Sheet_Pass As String, FileNumber As Long, Path As String, INTC As Interior, SHP As Shape, _
HUB_Color_Change As Boolean, Creator As Boolean, PWD As String

Const Alternate_Password As String = "F84?59D87~$[]\=<ApPle>###43"

Creator = UUID
    
If This_Sheet Is Nothing Then Set This_Sheet = HUB

If Not Creator And This_Sheet Is HUB Then Exit Sub
    
With ThisWorkbook
    
    If Manual_Trigger And Creator And .ActiveSheet Is HUB And ActiveWorkbook Is ThisWorkbook Then
        
        HUB.Shapes("My_Date").TopLeftCell.Offset(1, 0).Select 'In case I need design mode
    
    End If
    
    If .Password_M = vbNullString And Creator Then
    
        On Error GoTo Password_File_Not_Found
        
        Path = Environ("ONEDRIVE") & "\C.O.T Password.txt" 'path of file containing password
        
        FileNumber = FreeFile
        Open Path For Input As #FileNumber                 'open text file and load delmited string to variable
            .Password_M = Input(LOF(FileNumber), #FileNumber)
            .Password_M = Split(.Password_M, Chr(44))(0)   'password will be first item in the string
        Close #FileNumber
        
    End If
    
    If This_Sheet Is HUB And Creator Then
        PWD = .Password_M
    ElseIf Not This_Sheet Is HUB Then
        PWD = Alternate_Password
    End If

End With
    
On Error GoTo Worksheet_Protection_Change_Error

With This_Sheet

    If This_Sheet Is HUB And Creator And Allow_Color_Change = True Then
        Set SHP = .Shapes("My_Date")
        Set INTC = .Range("A1").Interior
        HUB_Color_Change = True
    End If
    
    If .ProtectContents = True Then 'If the worksheet is protected
    
        .Unprotect PWD
        
        If HUB_Color_Change = True Then 'If not saving
        
            With SHP.Fill.ForeColor
                HUB_Date_Color = .RGB   'Store color in Variable
                INTC.Color = .RGB       'Store current color on worksheet
                .RGB = RGB(88, 143, 33) 'Change to Green Color
            End With
                
        End If
            
    ElseIf .ProtectContents = False Then 'If worksheet is not protected
    
        If HUB_Color_Change = True Then 'if not saving
         
            With INTC
            
                If HUB_Date_Color <> .Color And .Color <> RGB(255, 255, 255) Then
                
                    SHP.Fill.ForeColor.RGB = .Color  'Apply stored color in range back to Shape
                
                ElseIf HUB_Date_Color <> RGB(255, 255, 255) Then
                    
                    SHP.Fill.ForeColor.RGB = HUB_Date_Color 'use color stored in variable
                
                End If
                
                .ColorIndex = 0 'No fill on worksheet storage location
                
            End With
            
         End If
             
        .Protect Password:=PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                UserInterfaceOnly:=False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
                AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows:=True, _
                AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, AllowDeletingRows:=True, _
                AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
             
    End If
    
End With

Exit Sub

Worksheet_Protection_Change_Error:

    MsgBox "The protection status of worksheet: " & This_Sheet.Name & " couldn't be changed."
    Exit Sub

Password_File_Not_Found:

    MsgBox "Password File not found in OneDrive folder. Worksheets were not locked. If Password has been forgotten check DropBox or Onedrive."

End Sub
Public Sub Close_Workbook() 'CTRL+W
Attribute Close_Workbook.VB_Description = "CTRL+W"
Attribute Close_Workbook.VB_ProcData.VB_Invoke_Func = "w\n14"

Dim Stored_DTA_UPD_Time As Date, Stored_WB_UPD_Time As Date, SV As Variant, WBN As String, Saved_State As Boolean

Dim Want_To_Save As Long, MSG As String, Verified_Workbook_Close As Boolean

Close_Workbook_Macro = True 'Used in Before Close Event

If ThisWorkbook.Saved = False Then 'If there are unsaved changes

    MSG = "Do you want to save changes for this workbook?"

    Want_To_Save = MsgBox(MSG, vbYesNoCancel)
    
    Select Case Want_To_Save
    
        Case vbYes, vbNo 'If user clicks yes or no
            
            Verified_Workbook_Close = True
                 
            If Want_To_Save = vbYes Then
                
                Call Before_Save(Enable_Events_Toggle:=False) 'False so that after save isn't ran when saving workbook
                            
                With Application
                
                    .DisplayAlerts = False
                    
                    ThisWorkbook.Save                   'Before/After save events will not be executed
                    
                    .DisplayAlerts = True
                    
                    .EnableEvents = True
                    
               End With
                
            End If
            
        Case Else 'If user has pressed cancel
            
            Cancel_WB_Close = True
            Close_Workbook_Macro = False
            Exit Sub
            
    End Select

ElseIf ThisWorkbook.Saved = True Then
    
    Verified_Workbook_Close = True
    
End If

If Verified_Workbook_Close Then 'True as long as cancel or X button aren't clicked
    
    Remove_Key_Binds
    
    With ThisWorkbook
        Re_Enable 'Re-enable events
        WBN = "'" & .Name & "'!"
    End With
    
    SV = Variable_Sheet.ListObjects("Saved_Variables").DataBodyRange.Value2
        
    With WorksheetFunction
        Stored_DTA_UPD_Time = .VLookup("Scheduled Data Update Time", SV, 2, False)
        Stored_WB_UPD_Time = .VLookup("Scheduled Workbook Update Time", SV, 2, False)
    End With
    
    Erase SV
    
    With Application
        
        On Error Resume Next 'Unschedule OnTime events
        
        .OnTime Stored_WB_UPD_Time, WBN & "Update_Check", Schedule:=False
        .OnTime Stored_DTA_UPD_Time, WBN & "Schedule_Data_Update", Schedule:=False
        .StatusBar = vbNullString
        
        On Error GoTo 0
        
    End With

    With ThisWorkbook
    
        .Saved = True 'Workbook is already saved if user wanted it to be
        
        If Not .WB_EVNT_Close_Workbook Then .Close  'if closing workbook with CTRL+W instead of button click
        
    End With

End If

Close_Workbook_Macro = False

End Sub
Private Sub Update_Check()

Dim Stored_WB_UPD_RNG As Range, Schedule As Date, Error_STR As String

With Variable_Sheet.ListObjects("Saved_Variables").DataBodyRange.Columns(1)
    
    Set Stored_WB_UPD_RNG = .Cells(WorksheetFunction.Match("Scheduled Workbook Update Time", .Value2, 0), 2)

End With

Schedule = Now + TimeSerial(0, 10, 0)                'Schedule this procedure to run every 10 minutes

Stored_WB_UPD_RNG = Schedule                         'Save value to range

Application.OnTime Schedule, "'" & ThisWorkbook.Name & "'!Update_Check"   'Schedule Check

#If Mac Then
    On Error GoTo Date_Check_Error
    MAC_Update_Check     'Refresh and check in background with QueryTable
#Else
    On Error Resume Next
    
    Windows_Update_Check 'Refresh with a HTTP Request
    
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo Date_Check_Error
        MAC_Update_Check
    End If
    
#End If
    
Exit Sub

Date_Check_Error:
 
    Error_STR = "An error occurred while checking for Feature/Macro Updates or you aren't connected to the internet." & vbNewLine & _
    "IF you have a stable internet connection and continue to see this error then contact me at MoshiM_UC@outlook.com"

    ThisWorkbook.Event_Storage("Event_Error").Add Error_STR, "Event_Error_Update_Check"
    
Resume Error_Cleared

Error_Cleared:

    On Error Resume Next
    
    Application.OnTime Schedule, "'" & ThisWorkbook.Name & "'!Update_Check", Schedule:=False

End Sub
Private Sub Windows_Update_Check()

Dim WinHttpReq As Object, Workbook_Version As Double, URL As String, HTML As Object, _
X As Byte, File_Type As String

With Variable_Sheet.ListObjects("Saved_Variables").DataBodyRange

    Workbook_Version = WorksheetFunction.VLookup("Update", .Value2, 2, False)
    
End With

File_Type = Data_Retrieval.TypeF

X = Application.Match(File_Type, Array("L", "D", "T"), 0) - 1

If Not Combined_Workbook(Variable_Sheet) Then X = X + 3

URL = Replace("https://www.dropbox.com/s/78l4v2gp99ggp1g/Date_Check.txt?dl=0", _
              "www.dropbox.com", "dl.dropboxusercontent.com")
    
Set HTML = CreateObject("htmlFile")

Set WinHttpReq = CreateObject("MSXML2.XMLHTTP")
    
With WinHttpReq

    .Open "GET", URL, False 'File is a URL/web page: False means that it has to make the connection before moving on
    .send         'File is the URL of the file or webpage

    HTML.Body.innerHTML = .responseText

End With

If Round(Workbook_Version, 10) < Round(Split(HTML.Body.FirstChild.Data, ",")(X), 10) Then
    Workbook_Is_Outdated = True
    Update_File.Show 'Array elements in order [L,D,T]
End If
 
End Sub
Sub MAC_Update_Check(Optional QT As QueryTable)

Dim URL As String, X As Long, Workbook_Version As Double, File_Type As String, _
Query_L As QueryTable, Query_EVNT As New ClassQTE

If QT Is Nothing Then 'IF before refreshing the Quuery

    URL = Replace("https://www.dropbox.com/s/78l4v2gp99ggp1g/Date_Check.txt?dl=0", _
            "www.dropbox.com", "dl.dropboxusercontent.com")
    
    For Each Query_L In QueryT.QueryTables
        If InStr(1, Query_L.Name, "MAC_Creator_Update_Check") > 0 Then
            Set QT = Query_L
        End If
    Next Query_L
    
    If QT Is Nothing Then 'create Query_Table if it doesn't exist
    
        Set QT = QueryT.QueryTables.Add("TEXT;" & URL, Destination:=QueryT.Range("A1"))
    
        With QT
            .RefreshStyle = xlOverwriteCells
            .BackgroundQuery = True
            .Name = "MAC_Creator_Update_Check"
            .WorkbookConnection.Name = "Creator Update Checks {MAC}"
            .AdjustColumnWidth = False
            .SaveData = False
            .TextFileCommaDelimiter = True
        End With
    
    End If
    
    Query_EVNT.HookUpQueryTable QT, "MAC_Update_Check", ThisWorkbook, Variable_Sheet, True, Weekly
                                '0                1                   2            3           4        5
    QT.Refresh False 'refresh in background

Else

    With Variable_Sheet.ListObjects("Saved_Variables").DataBodyRange
    
        Workbook_Version = WorksheetFunction.VLookup("Update", .Value2, 2, False)
        
    End With
    
    File_Type = Data_Retrieval.TypeF
    
    X = Application.Match(File_Type, Array("L", "D", "T"), 0) 'Need 1 based so no -1
    
    If Not Combined_Workbook(Variable_Sheet) Then X = X + 3
    
    With QT.ResultRange 'Ran after Query has finished Refreshing
        
        If Round(Workbook_Version, 10) < Round(.Cells(1, X).Value2, 10) Then
            Workbook_Is_Outdated = True
            Update_File.Show
            
        End If
        
        .ClearContents
        
    End With

End If

End Sub
Private Sub Unlock_Project()

Dim g As Long, Path As String, RR As Range, WshShell As Object, PWD As String, Saved_State As Boolean
        
#If Mac Then

    Exit Sub

#Else
    
    If Not UUID Then Exit Sub
    
    Const PWD_Target As Long = 5
    
    With Variable_Sheet.ListObjects("Saved_Variables").DataBodyRange.Columns(1)
    
        Set RR = .Cells(WorksheetFunction.Match("Unlock_Project_Toggle", .Value2, 0), 2)
        
    End With

    If RR.Value2 = False Then
        
        Saved_State = ThisWorkbook.Saved
        
        Set WshShell = CreateObject("WScript.Shell")
        
        With ThisWorkbook
            
            g = FreeFile
            
            Path = Environ("ONEDRIVE") & "\C.O.T Password.txt"
            
            If Dir(Path) = vbNullString Then Exit Sub
            
            Open Path For Input As #g
                PWD = Input(LOF(g), #g)
                PWD = Split(PWD, Chr(44))(PWD_Target)
            Close #g
                
        End With
        'application.VBE.CommandBars.FindControls(
        With WshShell
        
            .SendKeys "%l", True                       'ALT L   Developer Tab
            
            .SendKeys "c", True                        'C       View Code for worksheet
            
            .SendKeys PWD, True    'Supply  Password
            
            .SendKeys "{ENTER}", True                  'Submit
            
            ThisWorkbook.VBProject.VBE.MainWindow.Visible = False 'Close Editor
            
        End With
        
        RR.Value2 = True
        
        ThisWorkbook.Saved = Saved_State
        
    Else
    
        ThisWorkbook.VBProject.VBE.MainWindow.Visible = True
        
    End If
    
#End If

End Sub
Sub Schedule_Data_Update(Optional Workbook_Open_EVNT As Boolean = False)
'Private Sub Schedule_Data_Update()

Dim INTE_D As Date, Schedule_Time As Date, Saved_State As Boolean, Stored_Update_Time As Date, _
    Last_Updated_ICE As Long
    
Dim File_Type As String, WBN As String, Last_Updated_CFTC As Long

Dim Schedule_Data_Triggered As Range, UserForm_OB As Object, Stored_DTA_UPD_RNG As Range, _
        Automatic_Checkbox As CheckBox, Data_Collection_Time As Double
        
Dim Allow_Schedule As Boolean, Proceed_To_Save_State As Boolean, CFTC_Updated As Boolean, ICE_Updated As Boolean, _
CFTC_Data() As Variant, ICE_Data() As Variant, MacUser As Boolean

#If Mac Then
    MacUser = True
#End If

If Workbook_Open_EVNT = True Then
    On Error GoTo Ask_For_Auto_Scheduling_Permissions
Else
    On Error GoTo Default_Disable_Scheduling
End If

Set Automatic_Checkbox = Weekly.Shapes("Auto-U-CHKBX").OLEFormat.Object

If Automatic_Checkbox.Value = -4146 Then 'If user doesn't want to auto-schedule and retrieve
    ThisWorkbook.Saved = True
    Exit Sub
End If

Check_If_Workbook_Is_Outdated:

If Actions.Workbook_Is_Outdated = True Then 'Cancel auto data refresh check & Turn off Auto-Update CheckBox
    
    On Error Resume Next
    
    Automatic_Checkbox.Value = -4146
    
    MsgBox "Automatic workbook data updates and scheduling have been terminated due to a workbook update availble on DropBox." & vbNewLine & _
           vbNewLine & _
           vbNewLine & _
           "It is highly encouraged that you download the latest version of this workbook."
           
    Exit Sub
    
End If

Retrieve_Info_Ranges: On Error GoTo 0

File_Type = Data_Retrieval.TypeF 'File type for Workbook [L,D,T]

With Variable_Sheet
    
    With .ListObjects("Saved_Variables").DataBodyRange
    
        Data = .Columns(1).Value2
        
        Set Schedule_Data_Triggered = .Cells(WorksheetFunction.Match("Triggered Data Schedule", Data, 0), 2)
        Set Stored_DTA_UPD_RNG = .Cells(WorksheetFunction.Match("Scheduled Data Update Time", Data, 0), 2)
        
        Data = .Value2
        
    End With
     
End With

On Error Resume Next

With WorksheetFunction

    Last_Updated_CFTC = .VLookup("Last_Updated_CFTC", Data, 2, False) 'Last updated DATE for COT data
    
    Allow_Schedule = .VLookup("Release Schedule Queried", Data, 2, False) 'Determined by macros in Query Table module
    
    If File_Type = "D" Then Last_Updated_ICE = .VLookup("Last_Updated_ICE", Data, 2, False)
    
    Erase Data 'No longer need saved_variables
    
End With

Schedule_Data_Triggered.Value2 = True 'record that this SUB has been called
'^^ used to determine
'if it is needed to fire Auto Refresh Macro on Check box click from weekly sheet

Stored_Update_Time = Stored_DTA_UPD_RNG.Value2 'Recorded time for the next CFTC Update

With ThisWorkbook

    WBN = "'" & .Name & "'!"
    
    If Not Workbook_Open_EVNT Then Saved_State = .Saved
    
    Data_Collection_Time = Timer
                                'Returns an array of the most Recent weekly CFTC COT data
    CFTC_Data = Application.Run(WBN & "HTTP_Weekly_Data", LAst_Updated, True, False)
    
    Data_Collection_Time = Round(Timer - Data_Collection_Time, 2)
    
    If (Not Not CFTC_Data) <> 0 Then  'If the returned array isn't empty.
            
        If CFTC_Data(1, 1) > Last_Updated_CFTC Then
            'If new data is available and the corresponding checkbox is on, then run the sorting macro
            Debug.Print "[" & File_Type & "]-Data Retrieval: " & Data_Collection_Time & " seconds."
            
            Application.Run WBN & "Sort_" & File_Type, True, CFTC_Data
            
            If Data_Retrieval.Retrieval_Halted_For_User_Interaction = True Then GoTo Scheduling_Next_Update 'If mac user and multiple weeks need to be downloaded
            
            CFTC_Updated = True
            
            Erase Data
            
        End If
    
    End If
    
    If File_Type = "D" And Not CFTC_Updated Then
        
        ICE_Data = Application.Run(WBN & "Weekly_ICE", Last_Updated_ICE, MacUser)
        
        If (Not Not ICE_Data) <> 0 Then ' if array isn't empty
            
            If ICE_Data(1, 1) > Last_Updated_ICE Then
            
                Application.Run WBN & "Sort_" & File_Type, True, CFTC_Data, ICE_Data
                
                If Data_Retrieval.Retrieval_Halted_For_User_Interaction = True Then GoTo Scheduling_Next_Update
            
                ICE_Updated = True
                
            End If
            
        End If
        
    End If

Scheduling_Next_Update:

    On Error Resume Next 'Unschedule any stored dates...possibly redundant but dosn't hurt
    
    Application.OnTime Stored_Update_Time, Procedure:=WBN & "Schedule_Data_Update", Schedule:=False
    
    If Allow_Schedule = True Then
    
        With ThisWorkbook.Event_Storage
            
            .Add True, "Currently Scheduling" 'Error Handle resume next
        
            If Err.Number <> 0 Then
                
                .Remove "Currently Scheduling"
                
                .Add True, "Currently Scheduling"
            
            End If
            
        End With
        
        Schedule_Time = CFTC_Release_Dates(False) 'Date and Time to schedule next data update check in Local Time.
            
        If Now > Schedule_Time Then 'if current local time exceeds local time of the next stored release then update the release schedule
    
            Application.Run WBN & "Time_Zones_Refresh" 'Updates Release Schedule Tab;e

            Schedule_Time = CFTC_Release_Dates(False) 'Returns Local Time for next update
            
        End If
            
        If Schedule_Time <> TimeSerial(0, 0, 0) And Schedule_Time > Now Then 'if succesful date & time found
            
            Select Case File_Type
                Case "L": LAst_Updated = 1
                Case "D": LAst_Updated = 2
                Case "T": LAst_Updated = 3
            End Select
            
            If Not Combined_Workbook(Variable_Sheet) Then LAst_Updated = LAst_Updated + 3
            
            Schedule_Time = Schedule_Time + TimeSerial(0, 0, LAst_Updated)
            
            Application.OnTime EarliestTime:=Schedule_Time, _
                                Procedure:=WBN & "Schedule_Data_Update", _
                                Schedule:=True
                              
            Stored_DTA_UPD_RNG.Value = Schedule_Time 'Save Date and Time to Range
            
        End If
        
        ThisWorkbook.Event_Storage.Remove ("Currently Scheduling")
        
    End If
    
    If ICE_Updated Or CFTC_Updated Then
        .Saved = False
    ElseIf Workbook_Open_EVNT = True Then
        .Saved = True   'Data wasn't updated so no changes to save state needed
    Else
        .Saved = Saved_State  'Save state is what it was when this procedure was originally called
    End If

End With
   
Exit Sub

Default_Disable_Scheduling:

    ThisWorkbook.Saved = True

    Exit Sub
    
Ask_For_Auto_Scheduling_Permissions:

    If MsgBox("Auto-Scheduling and Retrieval checkbox couldn't be located." & vbNewLine & vbNewLine & _
           "Would you like to auto-schedule the retrieval of data.", vbYesNo) = vbYes Then
        
        Resume Check_If_Workbook_Is_Outdated
    Else
        ThisWorkbook.Saved = True
        Exit Sub
    End If
    
CheckBox_Failed:

    Resume Retrieve_Info_Ranges
    
End Sub
Private Sub Update_Date_Text_File(IsCreator As Boolean)

Dim Path As String, Update As Double, FileNumber As Byte, FileN As String, Update_Range As Range, _
STR_AR() As String, X As Byte, File_Type As String

If Not ThisWorkbook.Last_Used_Sheet Is Nothing And IsCreator Then 'Only to be ran while saving by me

    With Variable_Sheet.Range("Saved_Variables").Columns(1)
    
        Set Update_Range = .Cells(WorksheetFunction.Match("Update", .Value2, 0), 2)
        
        File_Type = Data_Retrieval.TypeF
        
    End With
    
    FileN = "Date_Check.txt"
    
    X = Application.Match(File_Type, Array("L", "D", "T"), 0) - 1 '-1 adjusts for split function used below
    
    If Not Combined_Workbook(Variable_Sheet) Then X = X + 3 'offset by 3 if Futures Only Workbook
    
    Path = Environ("ONEDRIVE") & "\COT Workbooks\" & FileN
    
    FileNumber = FreeFile
    
    Update = Now
    
    Open Path For Input As #FileNumber 'open text file and store contents in a string
    
        FileN = Input(LOF(FileNumber), #FileNumber)
        
    Close #FileNumber

    STR_AR() = Split(FileN, ",") 'Split the string with a comma
    
    STR_AR(X) = Update           'Overwrite array element with the current time converted to double

    Open Path For Output As #FileNumber 'Join array elements together and write back to text file
        
        Print #FileNumber, Join(STR_AR, ",")
    
    Close #FileNumber

    Update_Range.Value2 = Update 'Update saved Last_Saved_Time and date within workbook
 
End If
    
End Sub
Public Sub Save_Workbooks()
Attribute Save_Workbooks.VB_Description = "Rus the Custom Save macro in all workbooks if available."
Attribute Save_Workbooks.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Valid_Workbooks As New Collection, Z As Long, Saved_STR As String, Active_WB As Workbook

Saved_STR = "The following workbooks were saved >" & vbNewLine & vbNewLine

Set Active_WB = ActiveWorkbook

On Error GoTo Save_Error

For Z = 1 To Workbooks.Count

    With Workbooks(Z)
    
        If Not .Name Like "PERSONAL.*" Then
        
            Application.EnableEvents = False
            
            Application.Run "'" & .Name & "'!Custom_Save"
            
            Valid_Workbooks.Add Workbooks(Z)
            
        End If
        
    End With
    
Resume_Workbook_Loop:

Next Z

Active_WB.Activate

Application.EnableEvents = True

On Error GoTo 0

For Z = 1 To Valid_Workbooks.Count
    
    If TypeOf Valid_Workbooks(Z) Is Workbook Then
    
        With Valid_Workbooks(Z)
            If .Saved = False Then .Saved = True
            Saved_STR = Saved_STR & vbTab & .Name & vbNewLine
        End With
        
    End If
    
Next Z

MsgBox Saved_STR
Exit Sub

Save_Error:

    'Err.Clear
    MsgBox ("Unable to save " & Workbooks(Z).FullName)
    
    Resume Resume_Workbook_Loop

End Sub
Sub Custom_SaveAS()

    Save_Workbook Save_To_DropBox:=True

End Sub
Sub Custom_Save()
Attribute Custom_Save.VB_Description = "Save workboook without warnings [CTRL+S]"
Attribute Custom_Save.VB_ProcData.VB_Invoke_Func = "s\n14"
    Save_Workbook Save_To_DropBox:=False
End Sub
Private Sub Save_Workbook(Optional Save_To_DropBox As Boolean = False)

Re_Enable

Dim Item  As Variant, Pic_Status As Boolean

With Application

    .DisplayAlerts = False

    Pic_Status = Enable_Creator_Mode
    
    Call Before_Save(Enable_Events_Toggle:=False)  'Enable events is turned off here

    With ThisWorkbook
    
        .RemovePersonalInformation = False

        If Not Save_To_DropBox Then
            .Save
        Else
        
            If UUID Then
            
                On Error Resume Next
                
                With Variable_Sheet
                
                    For Each Item In Array("Holding", "Temp", "Unbound")
                        .Shapes(Item).Delete
                    Next
                    
                End With
                
                On Error GoTo 0
            
            End If
            
            .SaveAs Application.GetSaveAsFilename
            
        End If
        
    End With
    
    Enable_Creator_Mode = Pic_Status
    
    Call After_Save 'Enable events is turned on here

    .DisplayAlerts = True

End With

End Sub
Private Sub Before_Save(Enable_Events_Toggle As Boolean)

Dim WBN As String, Saved_Variables_RNG As Range, Saved_Variables() As Variant, _
X As Long, Creator As Boolean
    
Const Save_Timer_Key As String = "Save Events Timer", Saved_Variables_Key As String = "Saved_Variables"

With Application
    .ScreenUpdating = False: .EnableEvents = False
End With

With ThisWorkbook
    WBN = "'" & .Name & "'!"
    Set .Last_Used_Sheet = .ActiveSheet
End With

With HUB

    .Shapes("Macro_Check").Visible = True
    .Disable_ActiveX_Events = True
    .Shapes("Diagnostic").Visible = False
    
    If Not ThisWorkbook.ActiveSheet Is HUB Then .Activate
    
End With

Creator = UUID

If Creator Then

    If HUB.ProtectContents = False Then Call Worksheet_Protection_Toggle(HUB, True, False)
    
    With ThisWorkbook.Event_Storage
    
        On Error Resume Next 'Remove timer in case of code debuging and it wasn't removed
        .Remove Save_Timer_Key
        On Error GoTo 0
        
        .Add Timer, Save_Timer_Key 'Add back to Collection
        
    End With
    
    With Application
    
        .Run WBN & "Remove_Images" 'Make Workbook SFW
        
        .Run WBN & "Update_Date_Text_File", Creator 'Update last saved date and time in text file..Text File will be uploaded to DropBox
        
    End With
    
    On Error Resume Next
    
    Application.VBE.CommandBars.FindControl(Type:=msoControlButton, ID:=578).Execute 'Compile the project
    
    On Error GoTo 0
    
 End If

Set Saved_Variables_RNG = Variable_Sheet.ListObjects("Saved_Variables").DataBodyRange

Saved_Variables = Saved_Variables_RNG.Value2 'Write range to array

With ThisWorkbook.Event_Storage

    On Error Resume Next
    
        .Remove Saved_Variables_Key
        
        .Add Array(Saved_Variables, Saved_Variables_RNG), Saved_Variables_Key
         
    On Error GoTo 0
    
End With

For X = LBound(Saved_Variables, 1) To UBound(Saved_Variables, 1) 'Change certain values to false based on name

    Select Case Saved_Variables(X, 1)
    
        Case "Release Schedule Queried", "Triggered Data Schedule", "Unlock_Project_Toggle", "On_Creator_Computer"
            
             Saved_Variables(X, 2) = False         'Set range to boolean location
            
    End Select
    
Next X

Saved_Variables_RNG.Columns(2).Value2 = Application.Index(Saved_Variables, 0, 2) 'overwrite the saved variable range in column 2

Application.EnableEvents = Enable_Events_Toggle 'Turn back on to allow After_Save if not running custom_save macro
    
End Sub
Private Sub After_Save()

Dim Misc As Variant, WBN As String, Creator As Boolean ', Remove_Item As Boolean

Const Save_Timer_Key As String = "Save Events Timer", Saved_Variables_Key As String = "Saved_Variables"

WBN = "'" & ThisWorkbook.Name & "'!"

Application.EnableEvents = False

With HUB
    .Shapes("Macro_Check").Visible = False 'Turns this textbox back off if macros are enabled
    .Disable_ActiveX_Events = False
End With

With ThisWorkbook.Event_Storage
    
    Misc = .Item(Saved_Variables_Key)
           .Remove Saved_Variables_Key
           
    Misc(1).Value2 = Misc(0) 'Overwrite range with Pre-Save Values
    
    Creator = UUID
    
    Erase Misc
    
    If Creator Then
    
        If Enable_Creator_Mode = True Then Application.Run WBN & "Creator_Version"
        
        Debug.Print "[" & Data_Retrieval.TypeF & "] <" & Save_Timer & ">  " & Round(Timer - .Item(Save_Timer_Key), 2) & "s {" & Now & "}"
        
        .Remove Save_Timer_Key

    End If
    
End With

If Not Creator Then

    With HUB

        .Shapes("Diagnostic").Visible = True
        .Shapes("DN_List").Visible = True
    
    End With

End If

With ThisWorkbook
        .Last_Used_Sheet.Activate
    Set .Last_Used_Sheet = Nothing
        .Saved = True
End With

With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

End Sub

