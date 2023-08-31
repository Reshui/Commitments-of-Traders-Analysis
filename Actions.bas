Attribute VB_Name = "Actions"
Private HUB_Date_Color As Long
Private Workbook_Is_Outdated As Boolean
Public Close_Workbook_Macro As Boolean
Public Cancel_WB_Close As Boolean
Private Const Saved_Variables_Key As String = "Saved_Variables"
Private Const Save_Timer_Key As String = "Save Events Timer"
Private useCreatorWallpapers As Boolean

'Ary = Application.Index(Range("A1:G1000").Value, Evaluate("row(1:200)"), Array(4, 7, 1))
Private Sub Run_These_Key_Binds()

    Dim Key_Bind() As String, Procedure() As String, X As Byte, WBN As String ', Saved_State As Boolean
    
    'Saved_State = ThisWorkbook.Saved
    
    Key_Bind = Split("^b,^s,^w", ",")
    
    Procedure = Split("ToTheHub,Custom_Save,Close_Workbook", ",")
    
    WBN = "'" & ThisWorkbook.name & "'!"
    
    With Application
    
        For X = LBound(Key_Bind) To UBound(Key_Bind)
            .OnKey Key_Bind(X), WBN & Procedure(X)
        Next X
        
    End With

 'ThisWorkbook.Saved = Saved_State
 
End Sub
Private Sub Remove_Key_Binds()

Dim Key_Bind() As String, X As Byte

'Saved_State = ThisWorkbook.Saved

    Key_Bind = Split("^b,^s,^w", ",")
    
    With Application
    
        For X = LBound(Key_Bind) To UBound(Key_Bind)
            .OnKey Key_Bind(X)
        Next X
        
    End With

'ThisWorkbook.Saved = Saved_State

End Sub
Public Sub Remove_Images(Optional executeAsPartOfSaveEvent As Boolean = False)
Attribute Remove_Images.VB_Description = "Locked Macro"
Attribute Remove_Images.VB_ProcData.VB_Invoke_Func = " \n14"
'======================================================================================================
'Hides/Shows certain worksheets or images
'======================================================================================================

    Dim Variant_OBJ() As Variant, Wall_Path As String, X As Byte, _
    Their_HUB As String, Their_Weekly As String, WallP As New Collection, _
    obj As Variant, Shape_Group As GroupShapes, AnT As Shape, Variant_OBJ_Names As String ', wallpaperChangeTimer As TimedTask
    
    If UUID Then
        
        If Not executeAsPartOfSaveEvent Then useCreatorWallpapers = False
        
        'Set wallpaperChangeTimer = New TimedTask:wallpaperChangeTimer.Start "Change to Non-Creator wallpapers."
    
        Wall_Path = Environ("USERPROFILE") & "\Desktop\Wallpapers\"
            
        Variant_OBJ = Array(HUB, Variable_Sheet, Weekly)
            
        For X = LBound(Variant_OBJ) To UBound(Variant_OBJ)
          
            For Each obj In Variant_OBJ(X).Shapes
            
                With obj
                    
                    Select Case LCase(.name)
                    
                        Case "macro_check"
                        
                            If Not ThisWorkbook.Last_Used_Sheet Is Nothing And Not .Visible Then .Visible = True
                            .ZOrder (msoBringToFront)
                            
                            'If saving then make this shape visible
                        Case "object_group"
                            
                            Set Shape_Group = .GroupItems
                            
                            For Each AnT In Shape_Group
                            
                                With AnT
                                
                                    Select Case .name
                                        Case "DN_List", "Diagnostic", "Donate"
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
                
            Next obj
        
        Next X
        
        With Variable_Sheet
        
            Variant_OBJ = .ListObjects("Wallpaper_Selection").DataBodyRange.Value2
            .Visible = xlSheetVeryHidden
            .SetBackgroundPicture fileName:=vbNullString
            
        End With
        
        With WorksheetFunction
            Their_HUB = Wall_Path & .VLookup("Their_HUB", Variant_OBJ, 2, 0)
            Their_Weekly = Wall_Path & .VLookup("Their_Weekly", Variant_OBJ, 2, 0)
        End With
        
        With WallP
            .Add Array(Weekly, Their_Weekly)
            .Add Array(HUB, Their_HUB)
        End With
        
        For X = 1 To WallP.count
            
            Variant_OBJ = WallP(X) 'an array
            
            If FileOrFolderExists(CStr(Variant_OBJ(1))) Then
                Variant_OBJ(0).SetBackgroundPicture fileName:=Variant_OBJ(1)
            Else
                'MsgBox "Wallpaper not found for " & Variant_OBJ(0).name
                Variant_OBJ(0).SetBackgroundPicture fileName:=vbNullString
            End If
            
        Next X
        
        For Each obj In Array(QueryT)
            obj.Visible = xlSheetVeryHidden
        Next obj
        
        'wallpaperChangeTimer.DPrint
        
    End If

End Sub
Public Sub Creator_Version()
Attribute Creator_Version.VB_Description = "Locked macro."
Attribute Creator_Version.VB_ProcData.VB_Invoke_Func = " \n14"
'======================================================================================================
'Hides/Shows certain shapes and worksheets
'======================================================================================================

    Dim Variant_OBJ() As Variant, Wall_Path As String, X As Byte, T As Byte, obj As Shape, _
    MY_HUB As String, My_Weekly As String, My_Variables As String, WallP As New Collection, Workbook_Saved As Boolean
    
    If UUID Then
    
        Workbook_Saved = ThisWorkbook.Saved
        
        useCreatorWallpapers = True
       
        Wall_Path = Environ("USERPROFILE") & "\Desktop\Wallpapers\"
        
        Variant_OBJ = Array(HUB, Variable_Sheet, Weekly)
         
         For X = LBound(Variant_OBJ) To UBound(Variant_OBJ)
         
            For Each obj In Variant_OBJ(X).Shapes
            
                With obj
            
                    Select Case LCase(.name)
                        Case "make_macros_visible"
                            .Visible = True
                        Case "macro_check"
                        
                        Case "macros", "patch", "wallpaper_items"
                        
                            .Visible = False
                            
                        Case "object_group"
                            
                            For T = 1 To .GroupItems.count 'make everything but Email visible
                                
                                With .GroupItems(T)
                                
                                    Select Case .name
                                    
                                        Case "Feedback", "DN_List", "Database_Paths", "Disclaimer", "DropBox Folder", "Diagnostic", "Donate"
                                        
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
                        
            Next obj
            
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
        
        For X = 1 To WallP.count
            
            Variant_OBJ = WallP(X)
            
            If FileOrFolderExists(CStr(Variant_OBJ(1))) Then
                Variant_OBJ(0).SetBackgroundPicture fileName:=Variant_OBJ(1)
            Else
                MsgBox "Wallpaper not found for " & Variant_OBJ(0).name
            End If
            
        Next X
        
        ThisWorkbook.Saved = Workbook_Saved
    
    End If

End Sub
Public Sub Worksheet_Protection_Toggle(Optional This_Sheet As Worksheet, Optional Allow_Color_Change As Boolean = True, Optional Manual_Trigger As Boolean = True)

    Dim Sheet_Pass As String, fileNumber As Long, Path As String, INTC As Interior, SHP As Shape, _
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
            
            Path = Environ("OneDriveConsumer") & "\C.O.T Password.txt" 'path of file containing password
            
            fileNumber = FreeFile
            Open Path For Input As #fileNumber                 'open text file and load delmited string to variable
                .Password_M = Input(LOF(fileNumber), #fileNumber)
                .Password_M = Split(.Password_M, Chr(44))(0)   'password will be first item in the string
            Close #fileNumber
            
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

    MsgBox "The protection status of worksheet: " & This_Sheet.name & " couldn't be changed."
    Exit Sub

Password_File_Not_Found:

    MsgBox "Password File not found in OneDrive folder. Worksheets were not locked. If Password has been forgotten check DropBox or Onedrive."

End Sub
Public Sub Close_Workbook() 'CTRL+W
Attribute Close_Workbook.VB_Description = "CTRL+W"
Attribute Close_Workbook.VB_ProcData.VB_Invoke_Func = "w\n14"

    Dim Stored_DTA_UPD_Time As Date, Stored_WB_UPD_Time As Date, SV As Variant, WBN As String, saved_state As Boolean
    
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
            WBN = "'" & .name & "'!"
        End With
        
        SV = Variable_Sheet.ListObjects("Saved_Variables").DataBodyRange.Value2
            
        With Variable_Sheet
            Stored_DTA_UPD_Time = .Range("DropBox_Date_Query_Time")
            Stored_WB_UPD_Time = .Range("Data_Retrieval_Time")
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
    
    Set Stored_WB_UPD_RNG = Variable_Sheet.Range("DropBox_Date_Query_Time")
    
    Schedule = Now + TimeSerial(0, 10, 0)                'Schedule this procedure to run every 10 minutes
    
    Stored_WB_UPD_RNG = Schedule                         'Save value to range
    
    Application.OnTime Schedule, "'" & ThisWorkbook.name & "'!Update_Check"   'Schedule Check
    
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
    
    Application.OnTime Schedule, "'" & ThisWorkbook.name & "'!Update_Check", Schedule:=False

End Sub
Private Sub Windows_Update_Check()

    Dim WinHttpReq As Object, Workbook_Version As Date, URL As String, HTML As Object, splitChr As String, X As Byte
    
    #If DatabaseFile Then
        X = 1
        URL = "https://www.dropbox.com/s/8xgmlc2mfmwt032/Current_Version.txt?dl=0"
        splitChr = ":"
    #Else
        
        splitChr = ","
        URL = "https://www.dropbox.com/s/78l4v2gp99ggp1g/Date_Check.txt?dl=0"
        
        X = Application.Match(ReturnReportType, Array("L", "D", "T"), 0) - 1
    
        If Not combined_workbook() Then X = X + 3
    
    #End If
        
    URL = Replace(URL, "www.dropbox.com", "dl.dropboxusercontent.com")
       
    Set HTML = CreateObject("htmlFile")
    
    Set WinHttpReq = CreateObject("MSXML2.XMLHTTP")
        
    With WinHttpReq
    
        .Open "GET", URL, False 'File is a URL/web page: False means that it has to make the connection before moving on
        .send         'File is the URL of the file or webpage
    
        HTML.Body.innerHTML = .responseText
    
    End With
    
    Workbook_Version = Range("Workbook_Update_Version").value
    
    If Workbook_Version < CDate(Split(HTML.Body.FirstChild.data, splitChr)(X)) Then
        Workbook_Is_Outdated = True
        Update_File.Show 'Array elements in order [L,D,T]
    End If
     
End Sub
Sub MAC_Update_Check(Optional QT As QueryTable)
'======================================================================================================
'Checks if a folder on dropbox has number greater than the last creator save with a querytable
'If true then display the update Userform
'======================================================================================================
    
    Dim URL As String, X As Byte, Workbook_Version As Date, File_Type As String, _
    Query_L As QueryTable, splitChr As String

    #If DatabaseFile Then
        splitChr = ":"
        X = 1
        URL = "https://www.dropbox.com/s/8xgmlc2mfmwt032/Current_Version.txt?dl=0"
    #Else
        splitChr = ","
        X = Application.Match(ReturnReportType, Array("L", "D", "T"), 0) - 1
        If Not combined_workbook() Then X = X + 3
        
        URL = "https://www.dropbox.com/s/78l4v2gp99ggp1g/Date_Check.txt?dl=0"
    #End If
        
    URL = Replace(URL, "www.dropbox.com", "dl.dropboxusercontent.com")

    For Each Query_L In QueryT.QueryTables
        If InStr(1, Query_L.name, "MAC_Creator_Update_Check") > 0 Then
            Set QT = Query_L
        End If
    Next Query_L
    
    If QT Is Nothing Then 'create Query_Table if it doesn't exist
    
        Set QT = QueryT.QueryTables.Add("TEXT;" & URL, Destination:=QueryT.Range("A1"))
    
        With QT
            .RefreshStyle = xlOverwriteCells
            .BackgroundQuery = True
            .name = "MAC_Creator_Update_Check"
            .WorkbookConnection.name = "Creator Update Checks {MAC}"
            .AdjustColumnWidth = False
            .SaveData = False
        End With
    
    End If
    
    'Query_EVNT.HookUpQueryTable QT, "MAC_Update_Check", ThisWorkbook, Variable_Sheet, True, Weekly
                                '0                1                   2            3           4        5
    QT.Refresh False 'refresh in background

    Workbook_Version = Range("Workbook_Update_Version")
    
    With QT.ResultRange 'Ran after Query has finished Refreshing
        
        If Workbook_Version < CDate(Split(.Cells(1, 1), splitChr)(X)) Then
            Workbook_Is_Outdated = True
            Update_File.Show
            
        End If
        
        .ClearContents
        
    End With

End Sub
Private Sub Unlock_Project()

    Dim G As Long, Path As String, RR As Range, WshShell As Object, PWD As String, saved_state As Boolean
            
    #If Mac Then
    
        Exit Sub
    
    #Else
        
        If Not UUID Then Exit Sub
        
        #If DatabaseFile Then
            Const PWD_Target As Byte = 6
        #Else
            Const PWD_Target As Byte = 5
        #End If
        
        With Variable_Sheet.ListObjects("Saved_Variables").DataBodyRange.columns(1)
        
            Set RR = .Cells(WorksheetFunction.Match("Unlock_Project_Toggle", .Value2, 0), 2)
            
        End With
    
        If RR.Value2 = False Then
            
            saved_state = ThisWorkbook.Saved
            
            Set WshShell = CreateObject("WScript.Shell")
            
            With ThisWorkbook
                
                G = FreeFile
                
                Path = Environ("OneDriveConsumer") & "\C.O.T Password.txt"
                
                If Not FileOrFolderExists(Path) Then Exit Sub
                
                Open Path For Input As #G
                    PWD = Input(LOF(G), #G)
                    PWD = Split(PWD, Chr(44))(PWD_Target)
                Close #G
                    
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
            
            ThisWorkbook.Saved = saved_state
            
        Else
            Remove_Images
            'ThisWorkbook.VBProject.VBE.MainWindow.Visible = True
            
        End If
        
    #End If

End Sub
Sub Schedule_Data_Update(Optional Workbook_Open_EVNT As Boolean = False)

'======================================================================================================
'Checks if new data is available and schedules the process to be run again at the next data release
'======================================================================================================

    Dim INTE_D As Date, nextCftcUpdateTime As Date, saved_state As Boolean, unscheduleCftcUpdateTime As Date, WBN As String
    
    Dim UserForm_OB As Object, Stored_DTA_UPD_RNG As Range, Automatic_Checkbox As CheckBox
            
    Dim releaseScheduleHasBeenQueried As Boolean
    
    If Workbook_Open_EVNT = True Then
        On Error GoTo Ask_For_Auto_Scheduling_Permissions
    Else
        On Error GoTo Default_Disable_Scheduling
    End If
    
    Set Automatic_Checkbox = Weekly.Shapes("Auto-U-CHKBX").OLEFormat.Object
    
    If Automatic_Checkbox.value <> 1 Then 'If user doesn't want to auto-schedule and retrieve
        If Workbook_Open_EVNT Then ThisWorkbook.Saved = True
        Exit Sub
    End If
    
Check_If_Workbook_Is_Outdated:
    
    If Actions.Workbook_Is_Outdated = True Then 'Cancel auto data refresh check & Turn off Auto-Update CheckBox
        
        On Error Resume Next
        
        Automatic_Checkbox.value = -4146
        
        MsgBox "Automatic workbook data updates and scheduling have been terminated due to a workbook update availble on DropBox." & vbNewLine & _
               vbNewLine & _
               vbNewLine & _
               "It is highly encouraged that you download the latest version of this workbook."
               
        Exit Sub
        
    End If
    
Retrieve_Info_Ranges:     On Error GoTo 0
    
    With Variable_Sheet
        
         Set Stored_DTA_UPD_RNG = .Range("Data_Retrieval_Time")
    
        .Range("Triggered_Data_Schedule").Value2 = True 'record that this SUB has been called
        
         releaseScheduleHasBeenQueried = .Range("Release_Schedule_Queried").Value2  'Determined by macros in Query Table module
        
    End With
    
    On Error Resume Next
    
    unscheduleCftcUpdateTime = Stored_DTA_UPD_RNG.Value2  'Recorded time for the next CFTC Update
    
    With ThisWorkbook
    
        WBN = "'" & .name & "'!"
        
        If Not Workbook_Open_EVNT Then saved_state = .Saved
        
        Call New_Data_Query(Scheduled_Retrieval:=True, Overwrite_All_Data:=False)
    
Scheduling_Next_Update:
    
        On Error Resume Next 'Unschedule any stored dates...possibly redundant but dosn't hurt
                
        Application.OnTime unscheduleCftcUpdateTime, Procedure:=WBN & "Schedule_Data_Update", Schedule:=False
        
        If releaseScheduleHasBeenQueried = True Then
            
            nextCftcUpdateTime = CFTC_Release_Dates(Find_Latest_Release:=False) 'Date and Time to schedule next data update check in Local Time.
                
            If Now > nextCftcUpdateTime Then  'if current local time exceeds local time of the next stored release then update the release schedule
    
                Application.Run WBN & "Time_Zones_Refresh" 'Updates Release Schedule Tab;e
                nextCftcUpdateTime = CFTC_Release_Dates(False)  'Returns Local Time for next update
                
            End If
                
            nextCftcUpdateTime = nextCftcUpdateTime + TimeSerial(0, 1, 0)
            
            If nextCftcUpdateTime <> TimeSerial(0, 0, 0) And nextCftcUpdateTime > Now Then
                
                Application.OnTime EarliestTime:=nextCftcUpdateTime, _
                                    Procedure:=WBN & "Schedule_Data_Update", _
                                    Schedule:=True
                                  
                Stored_DTA_UPD_RNG.value = nextCftcUpdateTime
                
            End If
            
        End If
        
        If Workbook_Open_EVNT = True And Not Data_Retrieval.Data_Updated_Successfully Then
            .Saved = True   'Data wasn't updated so no changes to save state needed
        Else
            .Saved = saved_state  'Save state is what it was when this procedure was originally called
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

'======================================================================================================
'Edits a text file so that it holds the last saved date and time
'======================================================================================================

    Dim Path As String, fileNumber As Byte, FileN As String, Update_Range As Range, _
    X As Byte, newString As String, Update As Date
    
    If Not ThisWorkbook.Last_Used_Sheet Is Nothing And IsCreator Then 'Only to be ran while saving by me
        
        Update = Now
        
        #If DatabaseFile Then
        
            FileN = "Current_Version.txt"
            Path = Environ("OneDriveConsumer") & "\COT Workbooks\Database Version\" & FileN
            newString = "Workbook Version:" & Update
        #Else
            
            Dim storedDateValues() As String
            
            FileN = "Date_Check.txt"
            Path = Environ("OneDriveConsumer") & "\COT Workbooks\" & FileN
            
            X = Application.Match(ReturnReportType, Array("L", "D", "T"), 0) - 1 '-1 adjusts for split function used below
            
            If Not combined_workbook() Then X = X + 3 '-- offset by 3 to get index of Futures Only Workbook
            
            Path = Environ("OneDriveConsumer") & "\COT Workbooks\" & FileN
            
            fileNumber = FreeFile
            
            Open Path For Input As #fileNumber
                FileN = Input(LOF(fileNumber), #fileNumber)
            Close #fileNumber
            
            storedDateValues = Split(FileN, ",")
            
            storedDateValues(X) = Update
            
            newString = Join(storedDateValues, ",")
        
        #End If
        
        fileNumber = FreeFile
        
        Open Path For Output As #fileNumber 'Join array elements together and write back to text file
            
            Print #fileNumber, newString
        
        Close #fileNumber
    
        Range("Workbook_Update_Version").Value2 = Update 'Update saved Last_Saved_Time and date within workbook
     
    End If
    
End Sub
Public Sub Save_Workbooks()
Attribute Save_Workbooks.VB_Description = "Rus the Custom Save macro in all workbooks if available."
Attribute Save_Workbooks.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim Valid_Workbooks As New Collection, Z As Long, Saved_STR As String, Active_WB As Workbook
    
    Saved_STR = "The following workbooks were saved >" & vbNewLine & vbNewLine
    
    Set Active_WB = ActiveWorkbook
    
    On Error GoTo Save_Error
    
    For Z = 1 To Workbooks.count
    
        With Workbooks(Z)
        
            If Not .name Like "PERSONAL.*" Then
            
                Application.EnableEvents = False
                
                Application.Run "'" & .name & "'!Custom_Save"
                
                Valid_Workbooks.Add Workbooks(Z)
                
            End If
            
        End With
        
Resume_Workbook_Loop:
    
    Next Z
    
    Active_WB.Activate
    
    Application.EnableEvents = True
    
    On Error GoTo 0
    
    For Z = 1 To Valid_Workbooks.count
        
        If TypeOf Valid_Workbooks(Z) Is Workbook Then
        
            With Valid_Workbooks(Z)
                If .Saved = False Then .Saved = True
                Saved_STR = Saved_STR & vbTab & .name & vbNewLine
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
    
    Dim Item  As Variant, FileN As String
    
    With Application
    
        .DisplayAlerts = False
        
        Call Before_Save(Enable_Events_Toggle:=False)  'Enable events is turned off here
    
        With ThisWorkbook
        
            .RemovePersonalInformation = True
    
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
                
                FileN = Application.GetSaveAsFilename
                
               If FileN <> "FALSE" Then .SaveAs FileN
                
            End If
            
        End With
    
        Call After_Save 'Enable events is turned on here
    
        .DisplayAlerts = True
    
    End With

End Sub
Private Sub Before_Save(Enable_Events_Toggle As Boolean)

    Dim WBN As String, Saved_Variables_RNG As Range, Saved_Variables() As Variant, _
    X As Long, Creator As Boolean, saveTimer As TimedTask
        
    With Application
        .ScreenUpdating = False: .EnableEvents = False
    End With
    
    With ThisWorkbook
        WBN = "'" & .name & "'!"
        Set .Last_Used_Sheet = .ActiveSheet
    End With
    
    With HUB
    
        .Shapes("Macro_Check").Visible = True
        .Disable_ActiveX_Events = True
        .Shapes("Diagnostic").Visible = False
        .Shapes("DN_List").Visible = False
        
        If Not ThisWorkbook.ActiveSheet Is HUB Then .Activate
        
    End With
    
    Creator = UUID
    
    If Creator Then
    
        If HUB.ProtectContents = False Then Call Worksheet_Protection_Toggle(HUB, True, False)
        
        With ThisWorkbook.Event_Storage
        
            On Error Resume Next 'Remove timer in case of code debuging and it wasn't removed
            .Remove Save_Timer_Key
            
            Set saveTimer = New TimedTask
            
            saveTimer.Start ThisWorkbook.name & " (" & Now & ") ~ Save Event"
            
            .Add saveTimer, Save_Timer_Key 'Add back to Collection
            
            On Error GoTo 0
            
        End With
        
        Call Remove_Images(executeAsPartOfSaveEvent:=True)   'Make Workbook SFW
    
        Call Update_Date_Text_File(Creator)  'Update last saved date and time in text file..Text File will be uploaded to DropBox
        
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
        
        #If DatabaseFile Then
            If Creator Then Variable_Sheet.Range("Enable_Favorites").value = False
        #End If
        
        On Error GoTo 0
        
    End With
    
    For X = LBound(Saved_Variables, 1) To UBound(Saved_Variables, 1) 'Change certain values to false based on name
    
        Select Case Saved_Variables(X, 1)
        
            Case "Release Schedule Queried", "Triggered Data Schedule", "Unlock_Project_Toggle"
                
                 Saved_Variables(X, 2) = False         'Set range to boolean location
                
        End Select
        
    Next X
    
    Saved_Variables_RNG.columns(2).Value2 = Application.Index(Saved_Variables, 0, 2) 'overwrite the saved variable range in column 2
    
    Application.EnableEvents = Enable_Events_Toggle 'Turn back on to allow After_Save if not running custom_save macro

End Sub
Private Sub After_Save()

    Dim Misc As Variant, WBN As String, Creator As Boolean ', Remove_Item As Boolean
    
    WBN = "'" & ThisWorkbook.name & "'!"
    
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
            
            On Error GoTo Finished_Creator_Specified_Events
            
            If useCreatorWallpapers Then Application.Run WBN & "Creator_Version"
            
            .Item(Save_Timer_Key).DPrint
    
            .Remove Save_Timer_Key
    
        End If
    
Finished_Creator_Specified_Events:     On Error GoTo 0
    
    End With
    
    If Not Creator Then
    
        With HUB
        
            .Shapes("Diagnostic").Visible = True
            
            #If DatabaseFile Then
                If Range("Github_Version").value = True Then .Shapes("DN_List").Visible = True
            #Else
                .Shapes("DN_List").Visible = True
            #End If
            
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

Private Sub Show_Chart_Settings()
    Chart_Settings.Show
End Sub
Private Sub Adjust_Dash_Shapes()
    
    Dim FUT As Shape, FutOpt As Shape, wd As Double, DashWs As Worksheet
    
    Set DashWs = ThisWorkbook.ActiveSheet
    
    With DashWs
        Set FUT = .Shapes("FUT only")
        wd = (.Range("C1:E1").Width - 10) / 2
        Set FutOpt = .Shapes("FUT+OPT")
    End With
    
    With FUT
        .Top = 0
        .Left = DashWs.Range("c1").Left
        .Height = DashWs.Range("c1").Height
        .Width = wd
        .OLEFormat.Object.value = 1
    End With
    
    With FutOpt
        .Top = 0
        .Left = FUT.Left + FUT.Width + 10
        .Height = FUT.Height
        .Width = wd
        .OLEFormat.Object.value = -4146
    End With
    
    With DashWs.Shapes("Options")
        .Left = FUT.Left
        .Width = 2 * wd + 10
        .Height = FUT.Height + 5
        .Top = FUT.Top
    End With
    
    With DashWs.Shapes("Generate Dash")
        .Top = FUT.Top
        .Left = FutOpt.Left + FutOpt.Width + 10
        .Height = FUT.Height
    End With
    
End Sub

Private Sub ff()
Dim QT As Variant

    For Each QT In QueryT.QueryTables
         'Debug.Print QT.name
         With QT
            .WorkbookConnection.Delete
            .Delete
        End With
    Next
    
'    For Each QT In ThisWorkbook.Connections
'         Debug.Print QT.name
'         'QT.Delete
'
'         If QT.name Like "jun7-fc8e*" Then QT.Delete
'    Next


End Sub

