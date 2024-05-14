Attribute VB_Name = "Actions"
Private HUB_Date_Color As Long
Private Workbook_Is_Outdated As Boolean

Public CustomCloseActivated As Boolean

Private Const Saved_Variables_Key As String = "Saved_Variables"
Private Const Save_Timer_Key As String = "Save Events Timer"
Private useCreatorWallpapers As Boolean

'Ary = Application.Index(Range("A1:G1000").value2, Evaluate("row(1:200)"), Array(4, 7, 1))
Private Sub EndWorksheetTimedEvents()
    
    Dim nextCftcCheckTime  As Date, nextNewVersionAvailableCheck As Date, WBN As String
    
    WBN = "'" & ThisWorkbook.name & "'!"
    
    With Variable_Sheet
        nextCftcCheckTime = .Range("DropBox_Date_Query_Time")
        nextNewVersionAvailableCheck = .Range("Data_Retrieval_Time")
    End With
        
    With Application
        
        On Error Resume Next 'Unschedule OnTime events
        
        .Run "MT.CancelSwapColorOfCheckBox"
        .OnTime nextNewVersionAvailableCheck, WBN & "Update_Check", Schedule:=False
        .OnTime nextCftcCheckTime, WBN & "Schedule_Data_Update", Schedule:=False
        .StatusBar = vbNullString
        
        On Error GoTo 0
        
    End With
        
End Sub
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
                    
                    Select Case LCase$(.name)
                    
                        Case "macro_check"
                        
                            If Not ThisWorkbook.ActiveSheetBeforeSaving Is Nothing And Not .Visible Then .Visible = True
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
        
        #If DatabaseFile Then
            ClientAvn.Visible = xlSheetHidden
            ReversalCharts.Visible = xlSheetHidden
        #End If
        
    End If

End Sub
Public Sub Creator_Version()
Attribute Creator_Version.VB_Description = "Adjusts Workbook for creator adjustments."
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
            
                    Select Case LCase$(.name)
                        Case "make_macros_visible"
                            .Visible = True
                        Case "macro_check"
                        
                        Case "macros", "patch", "wallpaper_items"
                        
                            .Visible = False
                            
                        Case "object_group"
                            
                            For T = 1 To .GroupItems.count 'make everything but Email visible
                                
                                With .GroupItems(T)
                                
                                    Select Case .name
                                    
                                        Case "Disclaimer", "Feedback", "DN_List", "Database_Paths", "Disclaimer", "DropBox Folder", "Diagnostic", "Donate"
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
        
        If LenB(.Password_M) = 0 And Creator Then
        
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

    Custom_Close False

End Sub

Public Function Custom_Close(closeWorkbookEventActive As Boolean) As Boolean

    Dim doesUserWantToSave As Long, MSG As String, userStillWantsToClose As Boolean, cancelStatus As Boolean
    
    CustomCloseActivated = True
    
    If ThisWorkbook.Saved = False Then 'If there are unsaved changes
    
        MSG = "Do you want to save changes for this workbook?"
    
        doesUserWantToSave = MsgBox(MSG, vbYesNoCancel)
        
        Select Case doesUserWantToSave
        
            Case vbYes, vbNo
                
                userStillWantsToClose = True
                     
                If doesUserWantToSave = vbYes Then
                    'False so that After_Save event isn't executed
                    Call Before_Save(Enable_Events_Toggle:=False)
                                
                    With Application
                    
                        .DisplayAlerts = False
                        
                        ThisWorkbook.Save                   'Before/After save events will not be executed
                        
                        .DisplayAlerts = True
                        
                        .EnableEvents = True
                        
                   End With
                    
                End If
                
            Case Else 'If user has pressed cancel
                
                cancelStatus = True
                
        End Select
    
    ElseIf ThisWorkbook.Saved = True Then
        
        userStillWantsToClose = True
        
    End If
    
    If userStillWantsToClose Then 'True as long as cancel or X button aren't clicked
                
        Re_Enable
        Application.StatusBar = vbNullString
    
        With ThisWorkbook
            'if closing workbook with CTRL+W instead of button click
            If Not closeWorkbookEventActive Then
                .Close
            End If
        End With
    
    End If
    
    Custom_Close = cancelStatus
    CustomCloseActivated = False
    
End Function

Private Sub Update_Check()

    Dim Stored_WB_UPD_RNG As Range, Schedule As Date, Error_STR As String
    
    Set Stored_WB_UPD_RNG = Variable_Sheet.Range("DropBox_Date_Query_Time")
    
    Schedule = Now + TimeSerial(0, 10, 0)                'Schedule this procedure to run every 10 minutes
    
    Stored_WB_UPD_RNG = Schedule                         'Save value to range
    
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
            
    Application.OnTime Schedule, "'" & ThisWorkbook.name & "'!Update_Check"
    
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
    
        If Not IsWorkbookForFuturesAndOptions() Then X = X + 3
    
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
        If Not IsWorkbookForFuturesAndOptions() Then X = X + 3
        
        URL = "https://www.dropbox.com/s/78l4v2gp99ggp1g/Date_Check.txt?dl=0"
    #End If
        
    URL = Replace(URL, "www.dropbox.com", "dl.dropboxusercontent.com")

    For Each Query_L In QueryT.QueryTables
        If InStrB(1, Query_L.name, "MAC_Creator_Update_Check") > 0 Then
            Set QT = Query_L
            Exit For
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

    Workbook_Version = Range("Workbook_Update_Version").Value2
    
    With QT.ResultRange 'Ran after Query has finished Refreshing
        
        If Workbook_Version < CDate(Split(.Cells(1, 1).Value2, splitChr)(X)) Then
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
            Dim response As String
            response = InputBox("1 : Remove images" & vbNewLine & vbNewLine & "2 : Creator version", "Choose macro")
            
            If response = "1" Then
                Remove_Images
            ElseIf response = "2" Then
                Creator_Version
            End If
            
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
    
    If Automatic_Checkbox.value <> xlOn Then 'If user doesn't want to auto-schedule and retrieve
        If Workbook_Open_EVNT Then ThisWorkbook.Saved = True
        Exit Sub
    End If
    
Check_If_Workbook_Is_Outdated:
    
    If Actions.Workbook_Is_Outdated = True Then 'Cancel auto data refresh check & Turn off Auto-Update CheckBox
        
        On Error Resume Next
        
        Automatic_Checkbox.value = xlOff
        
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
                
            nextCftcUpdateTime = nextCftcUpdateTime + TimeSerial(0, 1, 30)
            
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
    X As Byte, newString As String, update As Date, dateStr As String
        
    If Not ThisWorkbook.ActiveSheetBeforeSaving Is Nothing And IsCreator Then 'Only to be ran while saving by me
        
        update = Now
        dateStr = Format(update, "dd-MMM-yyyy")
        #If DatabaseFile Then
            FileN = "Current_Version.txt"
            Path = Environ("OneDriveConsumer") & "\COT Workbooks\Database Version\" & FileN
            newString = "Workbook Version:" & dateStr
        #Else
            
            Dim storedDateValues() As String
            
            FileN = "Date_Check.txt"
            Path = Environ("OneDriveConsumer") & "\COT Workbooks\" & FileN
            
            X = Application.Match(ReturnReportType, Array("L", "D", "T"), 0) - 1 '-1 adjusts for split function used below
            
            If Not IsWorkbookForFuturesAndOptions() Then X = X + 3 '-- offset by 3 to get index of Futures Only Workbook
            
            Path = Environ("OneDriveConsumer") & "\COT Workbooks\" & FileN
            
            fileNumber = FreeFile
            
            Open Path For Input As #fileNumber
                FileN = Input(LOF(fileNumber), #fileNumber)
            Close #fileNumber
            
            storedDateValues = Split(FileN, ",")
            
            storedDateValues(X) = dateStr
            
            newString = Join(storedDateValues, ",")
        
        #End If
        
        fileNumber = FreeFile
        
        Open Path For Output As #fileNumber 'Join array elements together and write back to text file
            Print #fileNumber, newString
        Close #fileNumber
    
        Range("Workbook_Update_Version").Value2 = update 'Update saved Last_Saved_Time and date within workbook
     
    End If
    
End Sub
Public Sub Save_Workbooks()
Attribute Save_Workbooks.VB_Description = "Saves all workbooks that have a Custom_Save sub.\r\n"
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
    MsgBox ("Unable to save " & Workbooks(Z).fullName)
    
    Resume Resume_Workbook_Loop

End Sub
Sub Custom_SaveAS()
    Save_Workbook savingAsDifferentWorkbook:=True
End Sub
Sub Custom_Save()
    Save_Workbook savingAsDifferentWorkbook:=False
End Sub
Private Sub Save_Workbook(Optional savingAsDifferentWorkbook As Boolean = False)

    Re_Enable
    
    Dim Item  As Variant, FileN As String
    
    With Application
        
        .StatusBar = "[" & ThisWorkbook.name & "] Saving using Save_Workbook macro."
        .DisplayAlerts = False
        'Do before save actions and turn off events.
        Call Before_Save(Enable_Events_Toggle:=False)
    
        With ThisWorkbook
        
            .RemovePersonalInformation = True
    
            If Not savingAsDifferentWorkbook Then
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
        .StatusBar = vbNullString
        
    End With

End Sub
Private Sub Before_Save(Enable_Events_Toggle As Boolean)

    Dim WBN As String, Creator As Boolean, saveTimer As TimedTask, rngVar As Range, Item As Variant
    ' Move to hub first to unschedule any procedures that can be unscheduled via an event.
    Application.ScreenUpdating = False
    
    With ThisWorkbook
        WBN = "'" & .name & "'!"
        Set .ActiveSheetBeforeSaving = .ActiveSheet
    End With
    
    With HUB
        .Shapes("Macro_Check").Visible = True
        .Shapes("Diagnostic").Visible = False
        .Shapes("DN_List").Visible = False
        If Not ThisWorkbook.ActiveSheet Is HUB Then .Activate
    End With
    
    Creator = UUID
    
    Application.EnableEvents = False
    
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
        
        Call Remove_Images(executeAsPartOfSaveEvent:=True)
    
        Call Update_Date_Text_File(Creator)  'Update last saved date and time in text file..Text File will be uploaded to DropBox
        On Error GoTo Handle_Compile
        Application.VBE.CommandBars.FindControl(Type:=msoControlButton, ID:=578).Execute 'Compile the project
        
    End If
    
    Dim valuesToEditBeforeSave As New Collection
    
    On Error GoTo 0
    
    With valuesToEditBeforeSave
        For Each Item In Array("Triggered_Project_Unlock", "Triggered_Data_Schedule", "Release_Schedule_Queried")
            Set rngVar = Range(Item)
            .Add Array(rngVar, rngVar.value)
            rngVar.value = False
        Next Item
    End With
    
    With ThisWorkbook.Event_Storage
    
        On Error Resume Next
        .Remove Saved_Variables_Key
        .Add valuesToEditBeforeSave, Saved_Variables_Key
        On Error GoTo 0
        
        #If DatabaseFile Then
            ' This is done so that the file saves in a state in which end users will initially see all contracts
            ' when using the Contract_Selection Userform.
            If Creator Then Variable_Sheet.Range("Enable_Favorites").Value2 = False
        #End If
        
    End With
    
    With Application
        'Turn back on to allow After_Save if not running custom_save macro
        .EnableEvents = Enable_Events_Toggle
        '.DisplayAlerts = False
    End With
    
    Exit Sub

Handle_Compile:
    
    If Err.Number = -2147467259 Then
        ' Already compiled.
        Resume Next
    Else
        Err.Raise Err.Number
    End If
    
End Sub
Private Sub After_Save()

    Dim Misc As Variant, WBN As String, Creator As Boolean, workbookState As Boolean ', Remove_Item As Boolean
    
    With ThisWorkbook
        WBN = "'" & .name & "'!"
        workbookState = .Saved
    End With
    
    Application.EnableEvents = False
    
    HUB.Shapes("Macro_Check").Visible = False 'Turns this textbox back off if macros are enabled
    
    With ThisWorkbook.Event_Storage
    
        On Error Resume Next
            For Each Misc In .Item(Saved_Variables_Key)
                Misc(0).Value2 = Misc(1)
            Next Misc
            
            .Remove Saved_Variables_Key
        On Error GoTo 0
        
        Creator = UUID
        
        If Creator Then
            
            #If DatabaseFile Then
                HUB.Shapes("Database_Paths").Visible = False
            #End If
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
                If Range("Github_Version").Value2 = True Then .Shapes("DN_List").Visible = True
            #Else
                .Shapes("DN_List").Visible = True
            #End If
            
        End With
    
    End If
    
    With Application
        .EnableEvents = True
        With ThisWorkbook
            If Not .ActiveSheetBeforeSaving Is Nothing Then
                .ActiveSheetBeforeSaving.Activate
                Set .ActiveSheetBeforeSaving = Nothing
            End If
            .Saved = workbookState
        End With
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
        .OLEFormat.Object.value = xlOff
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

Private Sub DeleteAllQueryTablesOnQueryTSheet()
    
    Dim QT As Variant

    For Each QT In QueryT.QueryTables
         'Debug.Print QT.name
         With QT
            'If Not .WorkbookConnection Is Nothing Then .WorkbookConnection.Delete
            '.Delete
            Debug.Print QT.name
        End With
    Next
    
'    For Each QT In ThisWorkbook.Connections
'         Debug.Print QT.name
'         'QT.Delete
'
'         If QT.name Like "jun7-fc8e*" Then QT.Delete
'    Next


End Sub
Public Sub Change_Background() 'For use on the HUB worksheet
    
    Dim fNameAndPath As Variant, WP As Range, SplitN() As String, ZZ As Long, _
    File_Path As String, File_Content As String
    
    If UUID Then
        
        Select Case ThisWorkbook.ActiveSheet.name
        
            Case HUB.name, Weekly.name, Variable_Sheet.name
                Background_Change.Show
            Case Else
                GoTo Normal_Change
                
        End Select
        
    Else
    
Normal_Change:
    
        fNameAndPath = Application.GetOpenFilename(, Title:="Select Background_Image. If no image is selected, background will not be changed")
        
        If Not fNameAndPath = "False" Then
            On Error GoTo Invalid_Image
            ActiveSheet.SetBackgroundPicture fileName:=fNameAndPath
        End If
        
    End If
    
    Exit Sub

Invalid_Image:

    MsgBox "An error occured while attempting to apply the selected file."

End Sub
Sub ToTheHub()
Attribute ToTheHub.VB_Description = "Takes the user to the HUB worksheet."
Attribute ToTheHub.VB_ProcData.VB_Invoke_Func = " \n14"
     HUB.Activate
End Sub
Private Sub Workbook_Information_Userform()
    Workbook_Information.Show
End Sub

#If Not DatabaseFile Then

    Sub Navigation_Userform()
    
        Dim UserForm_OB As Object
        
        For Each UserForm_OB In VBA.UserForms
        
            If UserForm_OB.name = "Navigation" Then
        
                Unload UserForm_OB
                Exit Sub
            End If
        
        Next UserForm_OB
    
        navigation.Show
    
    End Sub
        
    Private Sub Column_Visibility_Form()
        Column_Visibility.Show
    End Sub
    
    Sub PopulateListBoxes(updateHub As Boolean, updateCharts As Boolean, updateForm As Boolean, Optional formComboBox As Object)
    
        Dim WorksheetNameCLCTN As New Collection, WS As Worksheet, _
        validCountBasic As Integer, wsKeys() As Variant, contractKeys() As Variant, validContractCount As Integer
        
        ReDim wsKeys(1 To ThisWorkbook.Worksheets.count)
        ReDim contractKeys(1 To ThisWorkbook.Worksheets.count)
        
        For Each WS In ThisWorkbook.Worksheets
        
            Select Case WS.name
            
                Case HUB.name, Weekly.name, Variable_Sheet.name, QueryT.name, MAC_SH.name, Symbols.name
                
                Case Else
                
                    validCountBasic = validCountBasic + 1
                    wsKeys(validCountBasic) = WS.name
                    
                    Set Tb = ReturnCftcTable(WS)
                    
                    If Not Tb Is Nothing Then
                        validContractCount = validContractCount + 1
                        contractKeys(validContractCount) = WS.name
                    End If
                                                
            End Select
        Next WS
        
        ReDim Preserve wsKeys(1 To validCountBasic)
        ReDim Preserve contractKeys(1 To validContractCount)
        
        Call Quicksort(wsKeys, LBound(wsKeys), UBound(wsKeys))
        Call Quicksort(contractKeys, LBound(contractKeys), UBound(contractKeys))
        
        #If Not Mac Then
            If updateHub Then
                HUB.Sheet_Selection.List = wsKeys
            End If
            
            If updateCharts Then
                Chart_Sheet.Sheet_Selection.List = contractKeys
            End If
        #End If
        
        If updateForm Then
            formComboBox.List = wsKeys
        End If
        
    End Sub
    Sub To_Charts()
        Chart_Sheet.Activate
    End Sub
    Public Sub Reset_UsedRange()
        '===========================================================================================
        'Reset each worksheets usedrange if there is a valid Table on the worksheet
        'Valid Table designate by having CFTC_Market_Code somewhere in its header row
        'Anything to the Right or Below this table will be deleted
        '===========================================================================================
        Dim HN As Collection, LRO As Range, LCO As Range, I As Long, TBL_RNG As Range, Worksheet_TB As Object, _
        Column_Total As Long, Row_Total As Long, UR_LastCell As Range, TB_Last_Cell As Range ', WSL As Range
        
        With Application 'Store all valid tables in an array
            Set HN = ContractDetails
        End With
        
        For I = 1 To HN.count
        
            Set TBL_RNG = HN(I).TableSource.Range       'Entire range of table
            Set Worksheet_TB = TBL_RNG.Parent 'Worksheet where table is found
            
            With Worksheet_TB '{Must be typed as object to fool the compiler when resetting the Used Range]
        
                With TBL_RNG 'Find the Bottom Right cell of the table
                    Set TB_Last_Cell = .Cells(.Rows.count, .columns.count)
                End With
                
                With .UsedRange 'Find the Bottom right cell of the Used Range
                    Set UR_LastCell = .Cells(.Rows.count, .columns.count)
                End With
                
                If UR_LastCell.Address <> TB_Last_Cell.Address Then
                
                    'If UR_LastCell AND TB_Last_Cell don't refer to the same cell
                    
                    With TB_Last_Cell
                        Set LRO = .Offset(1, 0) 'last row of table offset by 1
                        Set LCO = .Offset(0, 1) 'last column of table offset by 1
                    End With
                    
                    If UR_LastCell.column <> TB_Last_Cell.column And UR_LastCell.row = TB_Last_Cell.row Then
                        'Delete excess columns if columns are different but rows are the same
                        
                        .Range(LCO, UR_LastCell).EntireColumn.Delete  'Delete excess columns
                        
                    ElseIf UR_LastCell.column = TB_Last_Cell.column And UR_LastCell.row <> TB_Last_Cell.row Then
                        'Delete excess rows if rows are different but columns are the same
                        
                        .Range(LRO, UR_LastCell).EntireRow.Delete 'Delete exess rows
                        
                    ElseIf UR_LastCell.column <> TB_Last_Cell.column And UR_LastCell.row <> TB_Last_Cell.row Then
                        'if rows and columns are different
                        
                        .Range(LRO, UR_LastCell).EntireRow.Delete 'Delete excess usedrange
                        .Range(LCO, UR_LastCell).EntireColumn.Delete
                        
                    End If
                
                    .UsedRange 'reset usedrange
                    
                End If
            
            End With
               
        Next I
    
    End Sub
    
    Public Sub Autofit_Columns()
    
        Dim Tb As Long, TBR As Range
        
        With Application
            .ScreenUpdating = False
            Set Valid_Table_Info = ContractDetails
        End With
        
        For Tb = 1 To Valid_Table_Info.count
            Set TBR = Valid_Table_Info(Tb).TableSource.Range
            TBR.columns.AutoFit
        Next Tb
    
        Application.ScreenUpdating = True
     
    End Sub
    Public Sub Copy_Formats_From_ActiveSheet()
    
        Dim Valid_Table_Info As Collection, I As Long, TS As Worksheet, ASH As Worksheet, Target_TableR As Range, OT As ListObject, _
        Hidden_Collection As New Collection, HC As Range, T As Long, Original_Hidden_Collection As New Collection
        
        With Application
        
           .ScreenUpdating = False
        
        On Error GoTo Invalid_Function
        
            Set Valid_Table_Info = ContractDetails
        
        On Error GoTo 0
        
        End With
        
        Set ASH = ThisWorkbook.ActiveSheet
        
        On Error Resume Next
        
        Set OT = ReturnCftcTable(ASH)
        
        If Err.Number <> 0 Then
            MsgBox "No table with 'CFTC_Contract_Market_Code' found in table header ranges on the activesheet."
            End
        Else
            On Error GoTo 0
        End If
        
        With OT
                
            For Each HC In .HeaderRowRange.Cells 'Loop cells in the header and check hidden property
                If HC.EntireColumn.Hidden = True Then Original_Hidden_Collection.Add HC
            Next HC
            
            With .DataBodyRange
                .EntireColumn.Hidden = False ' unhide any hidden cells
                .Copy
            End With
            
        End With
        
        For I = 1 To Valid_Table_Info.count
        
            Set Target_TableR = Valid_Table_Info(I).TableSource.DataBodyRange  'databodyrange of the target table
          
            If Not Target_TableR.Parent Is ASH Then 'if the worksheet objects aren't the same
                       
                With Hidden_Collection 'store range objects of hidden columns inside a collection
                    
                    For Each HC In Valid_Table_Info(I)(1).HeaderRowRange.Cells 'Loop cells in the header and check hidden property
                        If HC.EntireColumn.Hidden = True Then .Add HC
                    Next HC
                    
                End With
                
                With Target_TableR
                    
                    .EntireColumn.Hidden = False 'unhide any hidden cells
                    
                    .FormatConditions.Delete 'remove formats from table
                    
                    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                    SkipBlanks:=False, Transpose:=False 'paste formats from ASH
                    
                End With
                
                With Hidden_Collection
                
                    If .count > 0 Then 'if at least 1 column was hidden then reapply hidden roperty to specified column
                    
                        For T = 1 To .count
                            Hidden_Collection(T).EntireColumn.Hidden = True
                        Next T
                        
                        Set Hidden_Collection = Nothing 'empty the collection
                    
                    End If
                    
                End With
                
            End If
            
        Next I
        
        With Original_Hidden_Collection
        
            If .count > 0 Then 'if at least 1 column was hidden then reapply hidden roperty to specified column
            
                For T = 1 To .count
                    Original_Hidden_Collection(T).EntireColumn.Hidden = True
                Next T
                
                Set Original_Hidden_Collection = Nothing 'empty the collection
            
            End If
            
        End With
        
        With Application
        
            .ScreenUpdating = True
            .CutCopyMode = False
            
        End With
        
        Exit Sub
    
Invalid_Function:
        MsgBox "Function Get_Worksheet_Info is unavailable. This Macro is intended for files created by MoshiM." & vbNewLine & vbNewLine & _
        "If you see this message and one of my files is the Active Workbook then please contact me."
    
    End Sub
    Public Sub Copy_Formulas_From_Active_Sheet()
        
        Dim Valid_Table_Info As Collection, Tb As ListObject, Source_TB_RNG As Range, _
        I As Long, Valid_Table As Boolean
        
        Set Valid_Table_Info = ContractDetails
        
        For Each Tb In ThisWorkbook.ActiveSheet.ListObjects 'Find the Listobject on the activesheet within the array
            
            For I = 1 To Valid_Table_Info.count
                
                If Valid_Table_Info(I).TableSource Is Tb Then
                
                    Valid_Table = True
                    Exit For
                    
                End If
                
            Next I
            
            If Valid_Table = True Then Exit For
            
        Next Tb
        
        If Valid_Table = False Then GoTo Active_Sheet_is_Invalid
        
        Set Source_TB_RNG = Valid_Table_Info(I).TableSource.DataBodyRange
        
        Dim Formula_Collection As New Collection, Cell As Range, Item As Variant
        
        For Each Cell In Source_TB_RNG.Rows(Source_TB_RNG.Rows.count).Cells
        
            With Cell
                If Left$(.Formula, 1) = "=" Then Formula_Collection.Add Array(.Formula, .column)
            End With
            
        Next Cell
        
        With Application
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
        End With
        
        For I = 1 To Valid_Table_Info.count 'loop all listobjects contained within the array
            
            Set Tb = Valid_Table_Info(I).TableSource
            
            If Not Tb Is Source_TB_RNG.ListObject Then 'if not the table that is being copied from
            
                'With TB.DataBodyRange 'Take formulas from collection and apply
    
                    For Each Item In Formula_Collection
                        Tb.ListColumns(Item(1)).DataBodyRange.Formula = Item(0)
                        '.Cells(.Rows.Count, Item(1)).Formula = Item(0)
                    Next
    
                'End With
               
            End If
            
        Next I
        
        Re_Enable
    
    Set Formula_Collection = Nothing
    
    Exit Sub
    
Active_Sheet_is_Invalid:
    
        MsgBox "You are trying to copy data formulas from an invalid worksheet"
        
        Application.Calculation = xlCalculationAutomatic
        
    End Sub

    Public Sub Copy_Valid_Data_Headers()
    
        Dim Headers() As Variant, Tb As ListObject, WS As Worksheet, Table_Info As Collection, I As Long
        
        Set Tb = ReturnCftcTable(ActiveSheet)
        
        If Not Tb Is Nothing Then
        
            With Application 'Store all valid tables in an array
                Set Table_Info = ContractDetails
            End With
        
            Headers = Tb.HeaderRowRange.Value2
        
            For I = 1 To Table_Info.count
        
                If Not Table_Info(I).TableSource Is Tb Then
        
                    Table_Info(I).TableSource.HeaderRowRange.Resize(1, UBound(Headers, 2)) = Headers
        
                End If
        
            Next I
        
        End If
    
    End Sub
    Sub deleteDataGreaterThanDate()
    
        Dim rowsToDeleteCount As Long, tblRange As Range, CC As Variant, RR As Range, Tb As ListObject
        
        For Each CC In ContractDetails
                    
            Set tblRange = CC.TableSource.DataBodyRange
            
            Set RR = tblRange.columns(1)
            
            Set RR = RR.Find(Format((DateSerial(2023, 6, 6)), "yyyy-mm-dd"), , xlValues, xlWhole)
            
            If Not RR Is Nothing Then
                                
                Set Tb = CC.TableSource
                
                With tblRange.Parent
                    
                    .Range(RR.Offset(1), .Cells(tblRange.Rows.count + 1, tblRange.columns.count)).ClearContents
                    Tb.Resize Range(CC.TableSource.Range.Cells(1, 1), .Cells(RR.row, tblRange.columns.count))
                End With
                
            End If
                    
        Next CC
    
    End Sub
#Else
    Public Sub Export_Data_Userform()
Attribute Export_Data_Userform.VB_Description = "Launches a userform to export data in a format accepted by the AmiBroker platform."
Attribute Export_Data_Userform.VB_ProcData.VB_Invoke_Func = " \n14"
        Export_Data.Show
    End Sub
#End If


