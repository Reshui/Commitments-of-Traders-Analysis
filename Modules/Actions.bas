Attribute VB_Name = "Actions"
Private HUB_Date_Color As Long
Private Workbook_Is_Outdated As Boolean

Public CustomCloseActivated As Boolean

Private Const Saved_Variables_Key$ = "Saved_Variables"
Private Const Save_Timer_Key$ = "Save Events Timer"
Private Const EventErrorKey$ = "Event_Error"

'Ary = Application.Index(Range("A1:G1000").value2, Evaluate("row(1:200)"), Array(4, 7, 1))

Option Explicit
Private Sub EndWorksheetTimedEvents()
    
    Dim nextCftcCheckTime  As Date, nextNewVersionAvailableCheck As Date, WBN$
    
    WBN = "'" & ThisWorkbook.Name & "'!"
    
    With Variable_Sheet
        nextCftcCheckTime = .Range("Data_Retrieval_Time").Value2
        nextNewVersionAvailableCheck = .Range("DropBox_Date_Query_Time").Value2
    End With
        
    With Application
    
        On Error Resume Next
        
        .Run "MT.CancelSwapColorOfCheckBox"
        
        If nextNewVersionAvailableCheck <> 0 Then
            .OnTime nextNewVersionAvailableCheck, WBN & "CheckForDropBoxUpdate", Schedule:=False
        End If
        
        If nextCftcCheckTime <> 0 Then
            .OnTime nextCftcCheckTime, WBN & "Schedule_Data_Update", Schedule:=False
        End If
        
        On Error GoTo 0
        .StatusBar = vbNullString
        
    End With
        
End Sub
'Private Sub Run_These_Key_Binds()
'
'    Dim Key_Bind$(), saved_state As Boolean, Procedure$(), x As Byte, WBN$ ', Saved_State As Boolean
'
'    With ThisWorkbook
'        saved_state = .Saved
'
'        Key_Bind = Split("^b,^s,^w", ",")
'
'        Procedure = Split("ToTheHub,Custom_Save,Close_Workbook", ",")
'
'        WBN = "'" & ThisWorkbook.Name & "'!"
'
'        With Application
'            For x = LBound(Key_Bind) To UBound(Key_Bind)
'                .OnKey Key_Bind(x), WBN & Procedure(x)
'            Next x
'        End With
'
'        .Saved = saved_state
'    End With
'
'End Sub
'Private Sub Remove_Key_Binds()
'
'    Dim Key_Bind$(), x As Byte, saved_state As Boolean
'
'    With ThisWorkbook
'        saved_state = .Saved
'        Key_Bind = Split("^b,^s,^w", ",")
'
'        With Application
'            For x = LBound(Key_Bind) To UBound(Key_Bind)
'                .OnKey Key_Bind(x)
'            Next x
'        End With
'
'        .Saved = saved_state
'    End With
'
'End Sub
Public Sub Remove_Images(Optional executeAsPartOfSaveEvent As Boolean = False)
'======================================================================================================
'Hides/Shows certain worksheets or images
'======================================================================================================

    Dim Variant_OBJ() As Variant, Wall_Path$, x As Byte, _
    hubImageForClients$, weeklyImageForClients$, WallP As New Collection, _
    obj As Variant, Shape_Group As GroupShapes, AnT As Shape, Variant_OBJ_Names$ ', wallpaperChangeTimer As TimedTask
    
    If IsOnCreatorComputer Then
        
        'Set wallpaperChangeTimer = New TimedTask:wallpaperChangeTimer.Start "Change to Non-Creator wallpapers."
    
        Wall_Path = Environ$("USERPROFILE") & "\Desktop\Wallpapers\"
            
        Variant_OBJ = Array(HUB, Variable_Sheet, Weekly)
            
        For x = LBound(Variant_OBJ) To UBound(Variant_OBJ)
          
            For Each obj In Variant_OBJ(x).Shapes
            
                With obj
                    
                    Select Case LCase$(.Name)
                    
                        Case "macro_check"
                        
                            If Not ThisWorkbook.ActiveSheetBeforeSaving Is Nothing And Not .Visible Then .Visible = True
                            .ZOrder (msoBringToFront)
                            
                            'If saving then make this shape visible
                        Case "object_group"
                            
                            Set Shape_Group = .GroupItems
                            
                            For Each AnT In Shape_Group
                            
                                With AnT
                                    Select Case .Name
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
                            If Not Variant_OBJ(x) Is Weekly Then .Visible = False
                    End Select
                    
                End With
                
            Next obj
        
        Next x
        
        With Variable_Sheet
            Variant_OBJ = .ListObjects("Wallpaper_Selection").DataBodyRange.Value2
            .Visible = xlSheetVeryHidden
            .SetBackgroundPicture fileName:=vbNullString
            If Not executeAsPartOfSaveEvent Then .Range("CreatorActiveState").Value2 = False
        End With
        
        With WorksheetFunction
            hubImageForClients = Wall_Path & .VLookup("Their_HUB", Variant_OBJ, 2, 0)
            weeklyImageForClients = Wall_Path & .VLookup("Their_Weekly", Variant_OBJ, 2, 0)
        End With
        
        With WallP
            .Add Array(Weekly, weeklyImageForClients)
            .Add Array(HUB, hubImageForClients)
        End With
        
        For x = 1 To WallP.count
            
            Variant_OBJ = WallP(x) 'an array
            
            If FileOrFolderExists(CStr(Variant_OBJ(1))) Then
                Variant_OBJ(0).SetBackgroundPicture fileName:=Variant_OBJ(1)
            Else
                'MsgBox "Wallpaper not found for " & Variant_OBJ(0).name
                Variant_OBJ(0).SetBackgroundPicture fileName:=vbNullString
            End If
            
        Next x
        
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
'======================================================================================================
'Hides/Shows certain shapes and worksheets
'======================================================================================================

    Dim Variant_OBJ() As Variant, Wall_Path$, x As Byte, T As Byte, obj As Shape, _
    MY_HUB$, My_Weekly$, My_Variables$, WallP As New Collection, Workbook_Saved As Boolean
    
    If IsOnCreatorComputer Then
    
        Workbook_Saved = ThisWorkbook.Saved

        Wall_Path = Environ$("USERPROFILE") & "\Desktop\Wallpapers\"
        
        Variant_OBJ = Array(HUB, Variable_Sheet, Weekly)
         
         For x = LBound(Variant_OBJ) To UBound(Variant_OBJ)
         
            For Each obj In Variant_OBJ(x).Shapes
            
                With obj
            
                    Select Case LCase$(.Name)
                        Case "make_macros_visible"
                            .Visible = True
                        Case "macro_check"
                        Case "macros", "patch", "wallpaper_items"
                            .Visible = False
                        Case "object_group"
                            
                            For T = 1 To .GroupItems.count 'make everything but Email visible
                                
                                With .GroupItems(T)
                                    Select Case .Name
                                        Case "Disclaimer", "Feedback", "DN_List", "Database_Paths", "Disclaimer", "DropBox Folder", "Diagnostic", "Donate"
                                            .Visible = False
                                            
                                        Case Else
                                            .Visible = True
                                    End Select
                                End With
                                
                            Next T
                        
                        Case Else
                            If Not Variant_OBJ(x) Is Weekly Then .Visible = False
                    End Select
                
                End With
                        
            Next obj
            
        Next x
        
        With Variable_Sheet 'load wallpaper strings into array and make worksheet visible
            Variant_OBJ = .ListObjects("Wallpaper_Selection").DataBodyRange.Value2
            .Visible = xlSheetVisible
            .Range("CreatorActiveState").Value2 = True
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
        
        For x = 1 To WallP.count
            Variant_OBJ = WallP(x)
            
            If FileOrFolderExists(CStr(Variant_OBJ(1))) Then
                Variant_OBJ(0).SetBackgroundPicture fileName:=Variant_OBJ(1)
            Else
                MsgBox "Wallpaper not found for " & Variant_OBJ(0).Name
            End If
        Next x
        
        ThisWorkbook.Saved = Workbook_Saved
        
    End If

End Sub
Public Sub Worksheet_Protection_Toggle(Optional sheetToToggleProtection As Worksheet, Optional Allow_Color_Change As Boolean = True, Optional Manual_Trigger As Boolean = True)

    Dim INTC As Interior, SHP As Shape, PWD$, savedState As Boolean, _
    HUB_Color_Change As Boolean, Creator As Boolean, creatorProperties As Collection, inputSheetIsHub As Boolean
    
    Const Alternate_Password$ = "F84?59D87~$[]\=<ApPle>###43"
    
    Creator = IsOnCreatorComputer
        
    If sheetToToggleProtection Is Nothing Then Set sheetToToggleProtection = HUB
    
    inputSheetIsHub = sheetToToggleProtection Is HUB
    
    If Not Creator And inputSheetIsHub Then Exit Sub
        
    With ThisWorkbook
    
        savedState = .Saved
        
        If Manual_Trigger And Creator And inputSheetIsHub And ActiveWorkbook Is ThisWorkbook Then
            HUB.Shapes("My_Date").TopLeftCell.offset(1, 0).Select 'In case I need design mode
        End If
        
        If LenB(.Password_M) = 0 And Creator Then
            Set creatorProperties = GetCreatorPasswordsAndCodes()
            .Password_M = creatorProperties("HUB_PASSWORD")
        End If
        
        If inputSheetIsHub And Creator Then
            PWD = .Password_M
        ElseIf Not inputSheetIsHub Then
            PWD = Alternate_Password
        End If
    
    End With
        
    On Error GoTo Worksheet_Protection_Change_Error
    
    With sheetToToggleProtection
    
        If inputSheetIsHub And Creator And Allow_Color_Change = True Then
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
                 
            .Protect password:=PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                    UserInterfaceOnly:=False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
                    AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows:=True, _
                    AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, AllowDeletingRows:=True, _
                    AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
                 
        End If
        
    End With
    
    If inputSheetIsHub Then ThisWorkbook.Saved = savedState
    
    Exit Sub

Worksheet_Protection_Change_Error:

    MsgBox "The protection status of worksheet: " & sheetToToggleProtection.Name & " couldn't be changed."
    Exit Sub

Password_File_Not_Found:

    MsgBox "Password File not found in OneDrive folder. Worksheets were not locked. If Password has been forgotten check DropBox or Onedrive."

End Sub
Public Sub Close_Workbook() 'CTRL+W
Attribute Close_Workbook.VB_Description = "Closes the workbook in a manner tha suppresses warningis about personal info being removed from the workbook."
Attribute Close_Workbook.VB_ProcData.VB_Invoke_Func = "w\n14"

    Custom_Close False

End Sub

Public Function Custom_Close(closeWorkbookEventActive As Boolean) As Boolean

    Dim doesUserWantToSave As Long, MSG$, userStillWantsToClose As Boolean, cancelStatus As Boolean
    
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

Private Sub CheckForDropBoxUpdate()

    Dim Stored_WB_UPD_RNG As Range, Schedule As Date, Error_STR$
    
    Set Stored_WB_UPD_RNG = Variable_Sheet.Range("DropBox_Date_Query_Time")
    
    Schedule = Now + TimeSerial(0, 10, 0)                'Schedule this procedure to run every 10 minutes
    
    Stored_WB_UPD_RNG = Schedule                         'Save value to range
    
    #If Mac Then
        On Error GoTo Date_Check_Error
        MAC_CheckForDropBoxUpdate     'Refresh and check in background with QueryTable
    #Else
        On Error Resume Next
        Windows_CheckForDropBoxUpdate 'Refresh with a HTTP Request
        
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo Date_Check_Error
            MAC_CheckForDropBoxUpdate
        End If
        
    #End If
            
    Application.OnTime Schedule, "'" & ThisWorkbook.Name & "'!CheckForDropBoxUpdate"
    
    Exit Sub

Date_Check_Error:
 
    Error_STR = "An error occurred while checking for Feature/Macro Updates or you aren't connected to the internet." & vbNewLine & _
    "If you have a stable internet connection and continue to see this error then contact me at MoshiM_UC@outlook.com"
    
    With HoldError(Err)
        On Error Resume Next
        Application.OnTime Schedule, "'" & ThisWorkbook.Name & "'!CheckForDropBoxUpdate", Schedule:=False
        PropagateError .HeldError, "CheckForDropBoxUpdate", Error_STR
    End With

End Sub
Private Sub Windows_CheckForDropBoxUpdate()

    Dim Workbook_Version As Date, _
    url$, HTML As Object, splitChr$, x As Byte, splitLimit As Long
    
    #If DatabaseFile Then
        x = 1
        url = "https://www.dropbox.com/s/8xgmlc2mfmwt032/Current_Version.txt?dl=0"
        splitChr = ":"
        splitLimit = 2
    #Else
        
        splitChr = ","
        splitLimit = -1
        url = "https://www.dropbox.com/s/78l4v2gp99ggp1g/Date_Check.txt?dl=0"
        
        x = -1 + Application.Match(ReturnReportType, Array("L", "D", "T"), 0) + IIf(IsWorkbookForFuturesAndOptions(), 0, 3)
    
    #End If
        
    url = Replace(url, "www.dropbox.com", "dl.dropboxusercontent.com")
       
    Set HTML = CreateObject("htmlFile")
        
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", url, False 'File is a URL/web page: False means that it has to make the connection before moving on
        .send         'File is the URL of the file or webpage
    
        HTML.Body.innerHTML = .responseText
    End With
    
    Workbook_Version = Range("Workbook_Update_Version").Value2
    
    If Workbook_Version < CDate(Split(HTML.Body.FirstChild.data, splitChr, splitLimit)(x)) Then
        Workbook_Is_Outdated = True
        Update_File.Show
    End If
     
End Sub
Sub MAC_CheckForDropBoxUpdate(Optional QT As QueryTable)
'======================================================================================================
'Checks if a folder on dropbox has number greater than the last creator save with a querytable
'If true then display the update Userform
'======================================================================================================
    
    Dim url$, x As Byte, Workbook_Version As Date, File_Type$, _
    Query_L As QueryTable, splitChr$

    #If DatabaseFile Then
        splitChr = ":"
        x = 1
        url = "https://www.dropbox.com/s/8xgmlc2mfmwt032/Current_Version.txt?dl=0"
    #Else
        splitChr = ","
        x = Application.Match(ReturnReportType, Array("L", "D", "T"), 0) - 1
        If Not IsWorkbookForFuturesAndOptions() Then x = x + 3
        
        url = "https://www.dropbox.com/s/78l4v2gp99ggp1g/Date_Check.txt?dl=0"
    #End If
        
    url = Replace(url, "www.dropbox.com", "dl.dropboxusercontent.com")

    For Each Query_L In QueryT.QueryTables
        If InStrB(1, Query_L.Name, "MAC_Creator_CheckForDropBoxUpdate") > 0 Then
            Set QT = Query_L
            Exit For
        End If
    Next Query_L
    
    If QT Is Nothing Then 'create Query_Table if it doesn't exist
    
        Set QT = QueryT.QueryTables.Add("TEXT;" & url, Destination:=QueryT.Range("A1"))
    
        With QT
            .RefreshStyle = xlOverwriteCells
            .BackgroundQuery = True
            .Name = "MAC_Creator_CheckForDropBoxUpdate"
            .WorkbookConnection.Name = "Creator Update Checks {MAC}"
            .AdjustColumnWidth = False
            .SaveData = False
        End With
    
    End If
    
    'Query_EVNT.HookUpQueryTable QT, "MAC_CheckForDropBoxUpdate", ThisWorkbook, Variable_Sheet, True, Weekly
                                '0                1                   2            3           4        5
    QT.Refresh False 'refresh in background

    Workbook_Version = Range("Workbook_Update_Version").Value2
    
    With QT.ResultRange 'Ran after Query has finished Refreshing
        
        If Workbook_Version < CDate(Split(.Cells(1, 1).Value2, splitChr)(x)) Then
            Workbook_Is_Outdated = True
            Update_File.Show
        End If
        
        .ClearContents
        
    End With

End Sub
Private Sub Unlock_Project(Optional unlockingFromPersonal As Boolean = False, Optional delayTime As Date)
            
    #If Mac Then
        Exit Sub
    #Else
        
        Dim projectUnlockCalledRange As Range, loadedPassword$, saved_state As Boolean, iCount As Byte, replacements$()
        
        If Not IsOnCreatorComputer Then
            Exit Sub
        ElseIf Not ThisWorkbook Is ActiveWorkbook Then
            ThisWorkbook.ActiveSheet.Activate
        End If
        
        #If DatabaseFile Then
            Const PWD_Target$ = "COT_DB_PASSWORD"
        #Else
            Const PWD_Target$ = "COT_BASIC_PASSWORD"
        #End If
        
        Set projectUnlockCalledRange = Variable_Sheet.Range("Triggered_Project_Unlock")

        If projectUnlockCalledRange.Value2 = False Then
            
            saved_state = ThisWorkbook.Saved
            
            loadedPassword = GetCreatorPasswordsAndCodes(PWD_Target)
            replacements = Split("+,^,%,~,(,)", ",")
            For iCount = LBound(replacements) To UBound(replacements)
                loadedPassword = Replace$(loadedPassword, replacements(iCount), "{" & replacements(iCount) & "}")
            Next iCount
            'application.VBE.CommandBars.FindControls(
            On Error GoTo Catch_FailedUnlock
            With CreateObject("WScript.Shell")
            
                .SendKeys "%l", True                       'ALT L   Developer Tab
                
                .SendKeys "c", True                        'C       View Code for worksheet
                
                .SendKeys loadedPassword, True                        'Supply  Password
                
                .SendKeys "{ENTER}", True                  'Submit
                
            End With
            
            ThisWorkbook.VBProject.VBE.MainWindow.Visible = False 'Close Editor
            
            If unlockingFromPersonal And delayTime > 0 Then Application.Wait (Now + delayTime)
            projectUnlockCalledRange.Value2 = True
            
            ThisWorkbook.Saved = saved_state
            
        ElseIf Not unlockingFromPersonal Then
            Dim response$
            response = InputBox("1 : Remove images" & vbNewLine & vbNewLine & "2 : Creator version", "Choose macro")
            
            If response = "1" Then
                Remove_Images
            ElseIf response = "2" Then
                Creator_Version
            End If
            
            'ThisWorkbook.VBProject.VBE.MainWindow.Visible = True
        End If
        
    #End If
    Exit Sub
Catch_FailedUnlock:
    MsgBox "Unlock failed."
End Sub
Sub Schedule_Data_Update(Optional Workbook_Open_EVNT As Boolean = False)

'======================================================================================================
'Checks if new data is available and schedules the process to be run again at the next data release
'======================================================================================================

    Dim nextCftcUpdateTime As Date, saved_state As Boolean, unscheduleCftcUpdateTime As Date, WBN$
    
    Dim Stored_DTA_UPD_RNG As Range, Automatic_Checkbox As CheckBox, releaseScheduleHasBeenQueried As Boolean

    If Workbook_Open_EVNT Then
        On Error GoTo Ask_For_Auto_Scheduling_Permissions
    Else
        On Error GoTo Default_Disable_Scheduling
    End If
    
Check_If_Workbook_Is_Outdated:

    Set Automatic_Checkbox = Weekly.Shapes("Auto-U-CHKBX").OLEFormat.Object
    
    If Automatic_Checkbox.value = xlOff Then
        'If user doesn't want to auto-schedule and retrieve
        If Workbook_Open_EVNT Then ThisWorkbook.Saved = True
        Exit Sub
    ElseIf Actions.Workbook_Is_Outdated = True Then 'Cancel auto data refresh check & Turn off Auto-Update CheckBox
        On Error Resume Next
        Automatic_Checkbox.value = xlOff
        MsgBox "Automatic workbook data updates and scheduling have been terminated due to a workbook update availble on DropBox." & vbNewLine & _
               vbNewLine & _
               vbNewLine & _
               "It is highly encouraged that you download the latest version of this workbook."
        Exit Sub
    End If

Retrieve_Info_Ranges:

    On Error GoTo 0
    
    With Variable_Sheet
         Set Stored_DTA_UPD_RNG = .Range("Data_Retrieval_Time")
        .Range("Triggered_Data_Schedule").Value2 = True
         releaseScheduleHasBeenQueried = .Range("Release_Schedule_Queried").Value2
    End With
    
    unscheduleCftcUpdateTime = Stored_DTA_UPD_RNG.Value2  'Recorded time for the next CFTC Update
    
    With ThisWorkbook
    
        WBN = "'" & .Name & "'!"
        
        If Not Workbook_Open_EVNT Then saved_state = .Saved
        
        On Error GoTo Catch_RetrievalError
        Call New_Data_Query(Scheduled_Retrieval:=True, Overwrite_All_Data:=False, IsWorbookOpenEvent:=Workbook_Open_EVNT)
    
Scheduling_Next_Update:
    
        On Error Resume Next
        ' Unschedule this script if it is already slated to run.
        Application.OnTime unscheduleCftcUpdateTime, Procedure:=WBN & "Schedule_Data_Update", Schedule:=False
        Err.Clear
        
        If releaseScheduleHasBeenQueried Then
            
            nextCftcUpdateTime = CFTC_Release_Dates(Find_Latest_Release:=False) 'Date and Time to schedule next data update check in Local Time.
                
            If Now > nextCftcUpdateTime Then  'if current local time exceeds local time of the next stored release then update the release schedule
                Application.Run WBN & "RefreshTimeZoneTable" 'Updates Release Schedule Tab;e
                nextCftcUpdateTime = CFTC_Release_Dates(False)  'Returns Local Time for next update
            End If
            
            If nextCftcUpdateTime <> TimeSerial(0, 0, 0) And nextCftcUpdateTime > Now Then
                
                nextCftcUpdateTime = nextCftcUpdateTime + TimeSerial(0, 1, 10)
                
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
    
    Err.Clear
    
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
Catch_RetrievalError:
    Resume Scheduling_Next_Update
Propagate:
    PropagateError Err, "Schedule_Data_Update"
End Sub
Private Sub Update_Date_Text_File(IsCreator As Boolean)

'======================================================================================================
'Edits a text file so that it holds the last saved date and time
'======================================================================================================

    Dim Path$, fileNumber As Byte, FileN$, Update_Range As Range, _
    x As Byte, newString$, update As Date, dateStr$
        
    If Not ThisWorkbook.ActiveSheetBeforeSaving Is Nothing And IsCreator Then 'Only to be ran while saving by me
        
        update = Now
        dateStr = Format(update, "dd-MMM-yyyy")
        #If DatabaseFile Then
            FileN = "Current_Version.txt"
            Path = Environ$("OneDriveConsumer") & "\COT Workbooks\Database Version\" & FileN
            newString = "Workbook Version:" & dateStr
        #Else
            
            Dim storedDateValues$()
            
            FileN = "Date_Check.txt"
            Path = Environ$("OneDriveConsumer") & "\COT Workbooks\" & FileN
            
            x = Application.Match(ReturnReportType, Array("L", "D", "T"), 0) - 1 '-1 adjusts for split function used below
            
            If Not IsWorkbookForFuturesAndOptions() Then x = x + 3 '-- offset by 3 to get index of Futures Only Workbook
            
            Path = Environ$("OneDriveConsumer") & "\COT Workbooks\" & FileN
            
            fileNumber = FreeFile
            
            Open Path For Input As #fileNumber
                FileN = Input(LOF(fileNumber), #fileNumber)
            Close #fileNumber
            
            storedDateValues = Split(FileN, ",")
            
            storedDateValues(x) = dateStr
            
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
Attribute Save_Workbooks.VB_Description = "Saves all workbooks that have a Custom_Save macro."
Attribute Save_Workbooks.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim Valid_Workbooks As New Collection, Z As Long, Saved_STR$, Active_WB As Workbook
    
    Saved_STR = "The following workbooks were saved >" & vbNewLine & vbNewLine
    
    Set Active_WB = ActiveWorkbook
    
    On Error GoTo Save_Error
    
    For Z = 1 To Workbooks.count
    
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
    
    For Z = 1 To Valid_Workbooks.count
        
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
    MsgBox ("Unable to save " & Workbooks(Z).fullName)
    
    Resume Resume_Workbook_Loop

End Sub
Sub Custom_SaveAS(Optional fileName As String)
    Save_Workbook savingAsDifferentWorkbook:=True, fileName:=fileName
End Sub
Sub Custom_Save()
Attribute Custom_Save.VB_Description = "Saves the workbook without warnings.\r\n"
Attribute Custom_Save.VB_ProcData.VB_Invoke_Func = "s\n14"
    Save_Workbook savingAsDifferentWorkbook:=False
End Sub
Private Sub Save_Workbook(Optional savingAsDifferentWorkbook As Boolean = False, Optional fileName As String)

    Re_Enable
    
    Dim Item  As Variant
    
    On Error GoTo DisplayErr
    
    With Application
        
        .StatusBar = "[" & ThisWorkbook.Name & "] Saving using Save_Workbook macro."
        .DisplayAlerts = False
        'Do before save actions and turn off events.
        Call Before_Save(Enable_Events_Toggle:=False)
    
        With ThisWorkbook
        
            .RemovePersonalInformation = True
    
            If Not savingAsDifferentWorkbook Then
                .Save
            Else
            
                If IsOnCreatorComputer Then
                    On Error Resume Next
                    With Variable_Sheet
                        For Each Item In Array("Holding", "Temp", "Unbound")
                            .Shapes(Item).Delete
                        Next
                    End With
                    On Error GoTo 0
                End If
                
                If LenB(fileName) = 0 Then
                    fileName = Application.GetSaveAsFilename
                End If
                
                If fileName <> "FALSE" Then .SaveAs fileName
                
            End If
            
        End With
    
        Call After_Save 'Enable events is turned on here
        .DisplayAlerts = True
        .StatusBar = vbNullString
        
    End With
    Exit Sub
DisplayErr:
    DisplayErr Err, "Save_Workbook"
End Sub
Private Sub Before_Save(Enable_Events_Toggle As Boolean)

    Dim WBN$, Creator As Boolean, saveTimer As TimedTask, rngVar As Range, Item As Variant
    ' Move to hub first to unschedule any procedures that can be unscheduled via an event.
    Application.ScreenUpdating = False
    
    Const procedureName$ = "Before_Save"
    
    On Error GoTo PropagateErr
     
    With ThisWorkbook
        WBN = "'" & .Name & "'!"
        Set .ActiveSheetBeforeSaving = .ActiveSheet
    End With
    
    With HUB
        .Shapes("Macro_Check").Visible = True
        .Shapes("Diagnostic").Visible = False
        .Shapes("DN_List").Visible = False
        If Not ThisWorkbook.ActiveSheet Is HUB Then .Activate
    End With
    
    Creator = IsOnCreatorComputer
    
    Application.EnableEvents = False
    
    If Creator Then
    
        If HUB.ProtectContents = False Then Call Worksheet_Protection_Toggle(HUB, True, False)
        'Remove timer in case of code debuging and it wasn't removed
        On Error Resume Next
        With ThisWorkbook.Event_Storage
        
            .Remove Save_Timer_Key
            
            Set saveTimer = New TimedTask
            
            saveTimer.Start ThisWorkbook.Name & " (" & Now & ") ~ Save Event"
            
            .Add saveTimer, Save_Timer_Key 'Add back to Collection
            
        End With
        On Error GoTo PropagateErr
        
        Call Remove_Images(executeAsPartOfSaveEvent:=True)
    
        Call Update_Date_Text_File(Creator)  'Update last saved date and time in text file..Text File will be uploaded to DropBox
        On Error GoTo Handle_Compile
        Application.VBE.CommandBars.FindControl(Type:=msoControlButton, ID:=578).Execute 'Compile the project
        
    End If
    
    Dim valuesToEditBeforeSave As New Collection
    
    On Error GoTo PropagateErr
    
    With valuesToEditBeforeSave
        For Each Item In Array("Triggered_Project_Unlock", "Triggered_Data_Schedule", "Release_Schedule_Queried", "CreatorActiveState")
            Set rngVar = Variable_Sheet.Range(Item)
            .Add Array(rngVar, rngVar.value)
            rngVar.value = False
        Next Item
    End With
    
    With ThisWorkbook.Event_Storage
    
        On Error Resume Next
        .Remove Saved_Variables_Key
        .Add valuesToEditBeforeSave, Saved_Variables_Key
        On Error GoTo PropagateErr
        
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
        PropagateError Err, procedureName, "Compile error in project."
    End If
PropagateErr:
    PropagateError Err, procedureName
End Sub
Private Sub After_Save()

    Dim Misc As Variant, WBN$, Creator As Boolean, workbookState As Boolean ', Remove_Item As Boolean
    Const procedureName$ = "After_Save"
    
    On Error GoTo PropagateErr
    
    With ThisWorkbook
        WBN = "'" & .Name & "'!"
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
        
        On Error GoTo PropagateErr
        
        Creator = IsOnCreatorComputer
        
        If Creator Then
            
            #If DatabaseFile Then
                HUB.Shapes("Database_Paths").Visible = False
            #End If
            On Error GoTo Finished_Creator_Specified_Events
            
            If Variable_Sheet.Range("CreatorActiveState").Value2 = True Then
                Run_This ThisWorkbook, "Creator_Version"
            End If
            
            .Item(Save_Timer_Key).DPrint
            .Remove Save_Timer_Key
    
        End If
Finished_Creator_Specified_Events:
    End With
    
    On Error GoTo PropagateErr
    
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
        
    With ThisWorkbook
        If Not .ActiveSheetBeforeSaving Is Nothing Then
            .ActiveSheetBeforeSaving.Activate
            Set .ActiveSheetBeforeSaving = Nothing
        End If
        .Saved = workbookState
    End With
        
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
    Exit Sub
PropagateErr:
    PropagateError Err, procedureName
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
    
    Dim QT As QueryTable

    For Each QT In QueryT.QueryTables
         'Debug.Print QT.name
         With QT
            'If Not .WorkbookConnection Is Nothing Then .WorkbookConnection.Delete
            '.Delete
            Debug.Print QT.Name
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
Attribute Change_Background.VB_Description = "Changes the background for the active worksheet."
Attribute Change_Background.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim fNameAndPath As Variant, WP As Range, SplitN$(), ZZ As Long, _
    File_Path$, File_Content$
    
    If IsOnCreatorComputer Then
        
        Select Case ThisWorkbook.ActiveSheet.Name
        
            Case HUB.Name, Weekly.Name, Variable_Sheet.Name
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
Attribute ToTheHub.VB_ProcData.VB_Invoke_Func = "b\n14"
     HUB.Activate
End Sub
Private Sub Workbook_Information_Userform()
    Workbook_Information.Show
End Sub
Private Sub WorkbookOpenEvent()

    Dim Creator As Boolean, Error_STR$, TT As Byte, rngVar As Range, _
    Field_Count As Byte, databaseConnectedWorkbook As Boolean, nameOfRangesToAlter$()
    
    Dim openEventTimer As New TimedTask, disableDataCheck As Boolean, gitHubVersion As Boolean, eventErrors As New Collection
       
    Const donatorInfoShapeName$ = "DN_List", clickToDonateShapeName$ = "Donate", _
        gitHubRangeName$ = "Github_Version", diagnosticShapeName$ = "Diagnostic"
    
    Call IncreasePerformance

    #If DatabaseFile And Not Mac Then
        databaseConnectedWorkbook = True
        
        On Error GoTo Catch_MissingDatabase
        Run_This ThisWorkbook, "FindDatabasePathInSameFolder"
        
Check_If_Exe_Available:

        With Variable_Sheet
        
            gitHubVersion = .Range(gitHubRangeName).Value2
            
            On Error GoTo C_EXE_Not_Defined
            With .Range("CSharp_Exe")
                If Not FileOrFolderExists(.Value2) Then
                    .Value2 = Empty
                End If
            End With
            
        End With
    #ElseIf Not DatabaseFile Then
        PopulateListBoxes True, True, False
    #End If
    
Start_Event_Timer:

    On Error Resume Next
    openEventTimer.Start ThisWorkbook.Name & " workbook Open event"

    ' Update ranges with default start values.
    nameOfRangesToAlter = Split("Triggered_Project_Unlock,Triggered_Data_Schedule,Release_Schedule_Queried,DropBox_Date_Query_Time,Data_Retrieval_Time", ",")
    
    On Error GoTo AttemptNextDefaultAssignment
    For TT = LBound(nameOfRangesToAlter) To UBound(nameOfRangesToAlter)
    
        Set rngVar = Variable_Sheet.Range(nameOfRangesToAlter(TT))
        
        Select Case nameOfRangesToAlter(TT)
            Case "DropBox_Date_Query_Time", "Data_Retrieval_Time"
                rngVar.value = Empty
            Case Else
                rngVar.value = False
        End Select
        Field_Count = Field_Count + 1
AttemptNextDefaultAssignment: On Error GoTo -1
    Next TT
    
    Set rngVar = Nothing
    
    On Error Resume Next
    'Determine if on creator computer..also update range
    Creator = IsOnCreatorComputer
    
    If Creator And Field_Count <> UBound(nameOfRangesToAlter) - LBound(nameOfRangesToAlter) + 1 Then
        MsgBox "1 or more fields weren't found when setting default settings on the Variable Sheet during the workbook open events."
    End If
    
    With HUB
        
        .Visible = xlSheetVisible
        'Turns off disclaimer box if macros are on
        .Shapes("Macro_Check").Visible = False
        'Makes donator count invisIble
        .Shapes(donatorInfoShapeName).Visible = False
        
        #If Mac And DatabaseFile Then
            GateMacAccessToWorkbook
        #End If
        
        If Not Creator Then
            
            .Shapes(diagnosticShapeName).Visible = True
            
            #If DatabaseFile Then
                If Not gitHubVersion Then .Shapes(clickToDonateShapeName).Visible = True
            #Else
                .Shapes(clickToDonateShapeName).Visible = True
            #End If
            
            If Not gitHubVersion Then
                'Show Donator dollar amount and # of donators
                Donators QueryT, .Shapes(donatorInfoShapeName)
                'Check if a new workbook version is available
            End If

        ElseIf databaseConnectedWorkbook Then
            .Shapes("Database_Paths").Visible = False
        End If
    
    End With
    
    With Weekly.Shapes("Test_Toggle").OLEFormat.Object
        If .value = xlOn Then 'Turn Test Mode off if its on
            .value = xlOff
            Application.Run "Weekly.Test_Toggle"
        End If
    End With
    
    With Application
    
        .Run "HUB.Range_Zoom"
        
        On Error GoTo Log_ERR_Resume_Next
        If Not Creator Then
            If Not gitHubVersion Then Call CheckForDropBoxUpdate
            Call UploadStats
        End If
        
        On Error Resume Next
        Application.Run "Query_Tables.RefreshTimeZoneTable", eventErrors
        Err.Clear
    End With
    
    If Not disableDataCheck Then Call Schedule_Data_Update(Workbook_Open_EVNT:=True)
    
Build_Error_String:
    
    With eventErrors 'Load all error messages into a singular string
            
        If .count > 0 Then
            On Error GoTo Next_EVNT_Item
            For TT = .count To 1 Step -1
                Error_STR = .Item(TT) & vbNewLine & vbNewLine & Error_STR
Next_EVNT_Item:
                On Error GoTo -1
            Next TT
            If LenB(Error_STR) > 0 Then MsgBox Error_STR, Title:="Error Message"
        End If
    
    End With
    
    On Error GoTo 0
Finally:

    openEventTimer.DPrint
    Call Re_Enable
    
    Exit Sub
    
    #If DatabaseFile Then
Catch_MissingDatabase:
        disableDataCheck = True
        Resume Check_If_Exe_Available
    #End If
    
C_EXE_Not_Defined:
    DisplayErr Err, "WorkbookOpenEvent", "Variable_Sheet.Range ('CSharp_Exe') not defined."
    Resume Start_Event_Timer
    
Log_ERR_Resume_Next:
    eventErrors.Add Err.Description
    Resume Next
    
Catch_GeneralError:
    DisplayErr Err, "WorkbookOpenEvent"
    Resume Next
End Sub
Private Sub UploadStats()
    
    Dim apiResponse$, postUpload$, workbookVersion$, success As Boolean, apiValuesByName As Object, postData$
    
    On Error GoTo Catch_POST_FAILED
    Const ApiURL$ = "https://ipapi.co/json/", ip$ = "ip", city$ = "city", country$ = "country_name", region$ = "region"
    
    Const FormURL$ = "https://docs.google.com/forms/d/e/1FAIpQLSfDB8cfBFZFcPf15tnaxuq6OStkmRYm4VlZjYWE8PEvm6qhFA/formResponse"
    
    #If DatabaseFile Then
        workbookVersion = "Database"
    #Else
        workbookVersion = ReturnReportType & "_" & IIf(IsWorkbookForFuturesAndOptions, "Combined", "FUT")
    #End If
    
    On Error GoTo Catch_GET_FAILED
    apiResponse = HttpGet(ApiURL, success)
                   
    If success Then
    
        On Error GoTo CATCH_JSON_PARSER_FAILURE
        
        Set apiValuesByName = Parse_Json_String(apiResponse)
        
        postData = "&entry.227122838=" & apiValuesByName.Item(city) & _
            "&entry.1815706984=" & apiValuesByName.Item(region) & _
            "&entry.55364550=" & apiValuesByName.Item(country) & _
            "&entry.1590825643=" & apiValuesByName.Item(ip) & _
            "&entry.1917934143=" & workbookVersion & _
            "&entry.1144002976=" & Format$(Variable_Sheet.Range("Workbook_Update_Version").Value2, "yyyy-mm-dd hh:mm:ss")
        
        On Error GoTo Catch_POST_FAILED

        HttpPost FormURL, postData, True
                
    End If
    
Finally:
    Exit Sub
Catch_GET_FAILED:
    PropagateError Err, "Stats", "GET failed."
Catch_POST_FAILED:
    PropagateError Err, "Stats", "POST failed."
CATCH_JSON_PARSER_FAILURE:
    PropagateError Err, "Stats", "JSON Parser failed."
End Sub

#If Not DatabaseFile Then

    Sub Navigation_Userform()
    
        Dim UserForm_OB As Object
        
        For Each UserForm_OB In VBA.UserForms
            If UserForm_OB.Name = "Navigation" Then
                Unload UserForm_OB
                Exit Sub
            End If
        Next UserForm_OB
    
        Navigation.Show
    
    End Sub
        
    Private Sub Column_Visibility_Form()
        Column_Visibility.Show
    End Sub
    
    Sub PopulateListBoxes(updateHub As Boolean, updateCharts As Boolean, updateForm As Boolean, Optional formComboBox As Object)
    
        Dim WorksheetNameCLCTN As New Collection, WS As Worksheet, _
        validCountBasic As Long, wsKeys() As Variant, contractKeys() As Variant, validContractCount As Long
        
        ReDim wsKeys(1 To ThisWorkbook.Worksheets.count)
        ReDim contractKeys(1 To ThisWorkbook.Worksheets.count)
        
        For Each WS In ThisWorkbook.Worksheets
        
            Select Case WS.Name
            
                Case HUB.Name, Weekly.Name, Variable_Sheet.Name, QueryT.Name, MAC_SH.Name, Symbols.Name
                
                Case Else
                    validCountBasic = validCountBasic + 1
                    With WS
                        wsKeys(validCountBasic) = .Name
                    
                        If Not ReturnCftcTable(WS) Is Nothing Then
                            validContractCount = validContractCount + 1
                            contractKeys(validContractCount) = .Name
                        End If
                    End With
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
        Dim HN As Collection, LRO As Range, LCO As Range, i As Long, TBL_RNG As Range, Worksheet_TB As Object, _
        Column_Total As Long, Row_Total As Long, UR_LastCell As Range, TB_Last_Cell As Range ', WSL As Range
        
        With Application 'Store all valid tables in an array
            Set HN = GetAvailableContractInfo
        End With
        
        For i = 1 To HN.count
        
            Set TBL_RNG = HN(i).TableSource.Range       'Entire range of table
            Set Worksheet_TB = TBL_RNG.parent 'Worksheet where table is found
            
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
                        Set LRO = .offset(1, 0) 'last row of table offset by 1
                        Set LCO = .offset(0, 1) 'last column of table offset by 1
                    End With
                    
                    If UR_LastCell.Column <> TB_Last_Cell.Column And UR_LastCell.row = TB_Last_Cell.row Then
                        'Delete excess columns if columns are different but rows are the same
                        
                        .Range(LCO, UR_LastCell).EntireColumn.Delete  'Delete excess columns
                        
                    ElseIf UR_LastCell.Column = TB_Last_Cell.Column And UR_LastCell.row <> TB_Last_Cell.row Then
                        'Delete excess rows if rows are different but columns are the same
                        
                        .Range(LRO, UR_LastCell).EntireRow.Delete 'Delete exess rows
                        
                    ElseIf UR_LastCell.Column <> TB_Last_Cell.Column And UR_LastCell.row <> TB_Last_Cell.row Then
                        'if rows and columns are different
                        
                        .Range(LRO, UR_LastCell).EntireRow.Delete 'Delete excess usedrange
                        .Range(LCO, UR_LastCell).EntireColumn.Delete
                        
                    End If
                
                    .UsedRange 'reset usedrange
                    
                End If
            
            End With
               
        Next i
    
    End Sub
    
    Public Sub Autofit_Columns()
    
        Dim Tb As Long, TBR As Range, Valid_Table_Info As Collection
        
        With Application
            .ScreenUpdating = False
            Set Valid_Table_Info = GetAvailableContractInfo
        End With
        
        For Tb = 1 To Valid_Table_Info.count
            Set TBR = Valid_Table_Info(Tb).TableSource.Range
            TBR.columns.AutoFit
        Next Tb
    
        Application.ScreenUpdating = True
     
    End Sub
    Public Sub Copy_Formats_From_ActiveSheet()
    
        Dim Valid_Table_Info As Collection, i As Long, TS As Worksheet, ASH As Worksheet, Target_TableR As Range, OT As ListObject, _
        Hidden_Collection As New Collection, HC As Range, T As Long, Original_Hidden_Collection As New Collection
        
        With Application
        
           .ScreenUpdating = False
        
        On Error GoTo Invalid_Function
        
            Set Valid_Table_Info = GetAvailableContractInfo
        
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
        
        For i = 1 To Valid_Table_Info.count
        
            Set Target_TableR = Valid_Table_Info(i).TableSource.DataBodyRange  'databodyrange of the target table
          
            If Not Target_TableR.parent Is ASH Then 'if the worksheet objects aren't the same
                       
                With Hidden_Collection 'store range objects of hidden columns inside a collection
                    
                    For Each HC In Valid_Table_Info(i)(1).HeaderRowRange.Cells 'Loop cells in the header and check hidden property
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
            
        Next i
        
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
        i As Long, CC As ContractInfo
        
        Set Valid_Table_Info = GetAvailableContractInfo
        
        For Each Tb In ThisWorkbook.ActiveSheet.ListObjects 'Find the Listobject on the activesheet within the array
            
            For Each CC In Valid_Table_Info
                With CC
                    If .TableSource Is Tb Then
                        Set Source_TB_RNG = .TableSource.DataBodyRange
                        Exit For
                    End If
                End With
            Next CC
            
            If Not Source_TB_RNG Is Nothing Then Exit For
            
        Next Tb
        
        If Not Source_TB_RNG Is Nothing Then GoTo Active_Sheet_is_Invalid
        
        Dim Formula_Collection As New Collection, Cell As Range, Item As Variant
        
        For Each Cell In Source_TB_RNG.Rows(1).Cells
        
            With Cell
                If Left$(.Formula, 1) = "=" Then Formula_Collection.Add Array(.Formula, .Column - Source_TB_RNG.Column + 1)
            End With
            
        Next Cell
        
        With Application
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
        End With
        
        For Each CC In Valid_Table_Info
            
            Set Tb = CC.TableSource
            
            If Not Tb Is Source_TB_RNG.ListObject Then 'if not the table that is being copied from
            
                'With TB.DataBodyRange 'Take formulas from collection and apply
    
                    For Each Item In Formula_Collection
                        Tb.ListColumns(Item(1) - Source_TB_RNG.Column + 1).DataBodyRange.Formula = Item(0)
                        '.Cells(.Rows.Count, Item(1)).Formula = Item(0)
                    Next
    
                'End With
               
            End If
            
        Next CC
Finally:
        Re_Enable
        Set Formula_Collection = Nothing
    
        Exit Sub
    
Active_Sheet_is_Invalid:
    
        MsgBox "You are trying to copy data formulas from an invalid worksheet"
        Resume Finally
        
    End Sub

    Public Sub Copy_Valid_Data_Headers()
    
        Dim Headers() As Variant, Tb As ListObject, WS As Worksheet, Table_Info As Collection, i As Long
        
        Set Tb = ReturnCftcTable(ActiveSheet)
        
        If Not Tb Is Nothing Then
        
            With Application 'Store all valid tables in an array
                Set Table_Info = GetAvailableContractInfo
            End With
        
            Headers = Tb.HeaderRowRange.Value2
        
            For i = 1 To Table_Info.count
        
                If Not Table_Info(i).TableSource Is Tb Then
        
                    Table_Info(i).TableSource.HeaderRowRange.Resize(1, UBound(Headers, 2)) = Headers
        
                End If
        
            Next i
        
        End If
    
    End Sub
    Private Sub DeleteDataGreaterThanDate()
    
        Dim rowsToDeleteCount As Long, tblRange As Range, CC As ContractInfo, _
        RR As Range, Tb As ListObject, minDateToKeep As Date
        minDateToKeep = CDate(InputBox("yyyy-mm-dd"))
        For Each CC In GetAvailableContractInfo
                    
            Set tblRange = CC.TableSource.DataBodyRange
            
            Set RR = tblRange.columns(1)
            
            Set RR = RR.Find(Format(minDateToKeep, "yyyy-mm-dd"), , xlValues, xlWhole)
            
            If Not RR Is Nothing Then
                                
                Set Tb = CC.TableSource
                
                With tblRange.parent
                    .Range(RR.offset(1), .Cells(tblRange.Rows.count + 1, tblRange.columns.count)).ClearContents
                    Tb.Resize Range(CC.TableSource.Range.Cells(1, 1), .Cells(RR.row, tblRange.columns.count))
                End With
                
            End If
                    
        Next CC
        
        With Variable_Sheet
            .Range("Last_Updated_CFTC").Value2 = minDateToKeep
            On Error Resume Next
            .Range("Last_Updated_ICE").Value2 = minDateToKeep
        End With
        
    End Sub
#Else
    Public Sub Export_Data_Userform()
Attribute Export_Data_Userform.VB_Description = "Launches a UserForm fo export data in a form acceptable to AmiBroker."
Attribute Export_Data_Userform.VB_ProcData.VB_Invoke_Func = " \n14"
        Export_Data.Show
    End Sub
#End If




