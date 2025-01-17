Attribute VB_Name = "Actions"
Private HUB_Date_Color As Long
Private Workbook_Is_Outdated As Boolean
Public CustomCloseActivated As Boolean

Private Const Saved_Variables_Key$ = "Saved_Variables"
Private Const SaveEventTimerKey$ = "Save Events Timer"
Private Const SaveEventDurationKey$ = "Save Duration"

Private Const EventErrorKey$ = "Event_Error"

'Ary = Application.Index(Range("A1:G1000").value2, Evaluate("row(1:200)"), Array(4, 7, 1))

Option Explicit
Private Sub EndWorksheetTimedEvents()
'===================================================================================================================
    'Summary: Unschedulues certain procedures so that the workbook doesn't reopen on its own.
'===================================================================================================================
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

    Dim Variant_OBJ() As Variant, wallpaperFolderPath$, x As Byte, _
    hubImageForClients$, weeklyImageForClients$, fileByWorksheetName As Object, _
    outerShape As Shape, innerShape As Shape, Variant_OBJ_Names$, worksheetName As Variant ', savedState As Boolean
    
    If IsOnCreatorComputer Then
        'savedState = ThisWorkbook.Saved
        wallpaperFolderPath = Environ$("USERPROFILE") & "\Desktop\Wallpapers\"
          
        Variant_OBJ = Array(HUB, Variable_Sheet, Weekly)
            
        For x = LBound(Variant_OBJ) To UBound(Variant_OBJ)
            For Each outerShape In Variant_OBJ(x).Shapes
                With outerShape
                    Select Case LCase$(.Name)
                        Case "object_group"
                            For Each innerShape In .GroupItems
                                With innerShape
                                    Select Case LCase$(.Name)
                                        Case "dn_list", "diagnostic", "donate"
                                            .Visible = False
                                        Case "macro_check"
                                            If Not executeAsPartOfSaveEvent Then .Visible = False
                                            .ZOrder msoBringToFront
                                        Case Else
                                            .Visible = True
                                    End Select
                                End With
                            Next innerShape
                        Case "temp", "unbound", "holding", "wallpaper_items"
                            .Visible = False
                        Case "sql_server_permission", "c#_permissions"
                            .Visible = True
                        Case Else
                            If Not Variant_OBJ(x) Is Weekly Then .Visible = False
                    End Select
                End With
            Next outerShape
        Next x
        
        With Variable_Sheet
            .Visible = xlSheetVeryHidden
            If Not executeAsPartOfSaveEvent Then .Range("CreatorActiveState").Value2 = False
            .SetBackgroundPicture fileName:=vbNullString
        End With
        
        Set fileByWorksheetName = CreateObject("Scripting.Dictionary")
        
        With GetWallpapersJSON()
            fileByWorksheetName.Add Weekly.Name, .Item("Their_Weekly")
            fileByWorksheetName.Add HUB.Name, .Item("Their_HUB")
        End With
        
        Dim wallpaperPath$
        
        With fileByWorksheetName
            For Each worksheetName In .keys
                wallpaperPath = wallpaperFolderPath & .Item(worksheetName)
                ThisWorkbook.Worksheets(worksheetName).SetBackgroundPicture fileName:=IIf(FileOrFolderExists(wallpaperPath), wallpaperPath, vbNullString)
            Next worksheetName
        End With
        
        QueryT.Visible = xlSheetVeryHidden

        #If DatabaseFile Then
            ClientAvn.Visible = xlSheetHidden
            ReversalCharts.Visible = xlSheetHidden
        #End If
        'ThisWorkbook.Saved = savedState
    End If

End Sub
Public Sub Creator_Version()
Attribute Creator_Version.VB_Description = "IF all checks are cleared then the Creator flag will be switched on."
Attribute Creator_Version.VB_ProcData.VB_Invoke_Func = " \n14"
'======================================================================================================
'Hides/Shows certain shapes and worksheets
'======================================================================================================

    Dim Variant_OBJ() As Variant, wallpaperFolderPath$, x As Byte, shp As Shape, _
    fileByWorksheetName As Object, Workbook_Saved As Boolean, wallpaperPath$, innerShape As Shape, worksheetName As Variant
    
    On Error GoTo Propagate
    If IsOnCreatorComputer() Then
            
        Workbook_Saved = ThisWorkbook.Saved
        
        Variant_OBJ = Array(HUB, Variable_Sheet, Weekly)
         
         For x = LBound(Variant_OBJ) To UBound(Variant_OBJ)
            For Each shp In Variant_OBJ(x).Shapes
                With shp
                    Select Case LCase$(.Name)
                        Case "make_macros_visible"
                            .Visible = True
                        Case "macro_check"
                            .Visible = False
                        Case "macros", "patch", "wallpaper_items"
                            .Visible = False
                        Case "object_group"
                            
                            For Each innerShape In .GroupItems 'make everything but Email visible
                                With innerShape
                                    Select Case .Name
                                        Case "Macro_Check", "Disclaimer", "Feedback", "DN_List", "Database_Paths", "Disclaimer", "DropBox Folder", "Diagnostic", "Donate"
                                            .Visible = False
                                        Case Else
                                            .Visible = True
                                    End Select
                                End With
                            Next innerShape
                        Case "sql_server_permission", "c#_permissions"
                            .Visible = True
                        Case Else
                            If Not Variant_OBJ(x) Is Weekly Then .Visible = False
                    End Select
                End With
            Next shp
        Next x
        
        With Variable_Sheet
            .Range("CreatorActiveState").Value2 = True
            .Visible = xlSheetVisible
        End With
        
        Set fileByWorksheetName = CreateObject("Scripting.Dictionary")
        
        With GetWallpapersJSON()
            fileByWorksheetName.Add Weekly.Name, .Item("My_Weekly")
            fileByWorksheetName.Add HUB.Name, .Item("My_HUB")
            fileByWorksheetName.Add Variable_Sheet.Name, .Item("This_Sheet")
        End With
        
        wallpaperFolderPath = Environ$("USERPROFILE") & "\Desktop\Wallpapers\"
        
        With fileByWorksheetName
            For Each worksheetName In .keys
                wallpaperPath = wallpaperFolderPath & .Item(worksheetName)
                
                If FileOrFolderExists(wallpaperPath) Then
                    ThisWorkbook.Worksheets(worksheetName).SetBackgroundPicture fileName:=wallpaperPath
                Else
                    MsgBox "Wallpaper not found for " & worksheetName
                End If
            Next worksheetName
        End With
        
        ThisWorkbook.Saved = Workbook_Saved
    End If
    
    Exit Sub
Propagate:
    PropagateError Err, "Creator_Version"
End Sub
Private Function GetWallpapersJSON() As Object
    
    Dim jsonPath$, jsonSerializer As New JsonParserB
    
    On Error GoTo Propagate
    
    jsonPath = Environ$("USERPROFILE") & "\Desktop\Wallpapers\COT_Wallpapers.json"
    Set GetWallpapersJSON = jsonSerializer.Deserialize(TxtMethods(jsonPath, True, False, False)).Item(GetWorkbookJsonWallpaperKey())

    Exit Function
Propagate:
    PropagateError Err, "GetWallpapersJSON"
End Function

Public Sub Worksheet_Protection_Toggle(Optional sheetToToggleProtection As Worksheet, Optional Allow_Color_Change As Boolean = True, Optional Manual_Trigger As Boolean = True)

    Dim INTC As Interior, shp As Shape, PWD$, savedState As Boolean, _
    HUB_Color_Change As Boolean, Creator As Boolean, creatorProperties As Object, inputSheetIsHub As Boolean
    
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
            Set shp = .Shapes("My_Date")
            Set INTC = .Range("A1").Interior
            HUB_Color_Change = True
        End If
        
        If .ProtectContents = True Then 'If the worksheet is protected
        
            .Unprotect PWD
            
            If HUB_Color_Change = True Then 'If not saving
            
                With shp.Fill.ForeColor
                    HUB_Date_Color = .RGB   'Store color in Variable
                    INTC.Color = .RGB       'Store current color on worksheet
                    .RGB = RGB(88, 143, 33) 'Change to Green Color
                End With
                    
            End If
                
        ElseIf .ProtectContents = False Then 'If worksheet is not protected
        
            If HUB_Color_Change = True Then 'if not saving
             
                With INTC
                
                    If HUB_Date_Color <> .Color And .Color <> RGB(255, 255, 255) Then
                    
                        shp.Fill.ForeColor.RGB = .Color  'Apply stored color in range back to Shape
                    
                    ElseIf HUB_Date_Color <> RGB(255, 255, 255) Then
                        
                        shp.Fill.ForeColor.RGB = HUB_Date_Color 'use color stored in variable
                    
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
        'if closing workbook with CTRL+W instead of button click
        If Not closeWorkbookEventActive Then ThisWorkbook.Close
    End If
    
    Custom_Close = cancelStatus
    CustomCloseActivated = False
    
End Function

Private Sub CheckForDropBoxUpdate()
'================================================================================================================================
    'Summary: Queries a DropBox file to determine if an update for the workbook is available then an update userform is launched.
    'Note:
    '   - This sub will schedule itself to run once every 10 minutes.
'================================================================================================================================
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

    Dim workbookVersion As Date, url$
        
    url = "https://www.dropbox.com/scl/fi/zao1xn7oexc4gypf8jfml/ReleaseVersion.json?rlkey=6q8t04c24mh1u9g4taq8lxoyv&st=p459w1px&dl=0"
    
    url = Replace$(url, "www.dropbox.com", "dl.dropboxusercontent.com")
    
    Dim result$, retrievedJSON As Object, jp As New JsonParserB, serverDate As Date

    If TryGetRequest(url, result) Then
        
        Set retrievedJSON = jp.Deserialize(result, True, False, False)
        #If DatabaseFile Then
            serverDate = retrievedJSON.Item("Database")
        #Else
            serverDate = retrievedJSON.Item(IIf(IsWorkbookForFuturesAndOptions, "Futures & Options", "Futures Only")).Item(ReturnReportType())
        #End If
        workbookVersion = Variable_Sheet.Range("Workbook_Update_Version").Value2
         
        If workbookVersion < serverDate Then
            Workbook_Is_Outdated = True
            Update_File.Show
        End If
        
    Else
        MsgBox "Unable to check for DropBox updates." & vbNewLine & "There may be an update available or you aren't connected to the internet."
    End If
     
End Sub
Private Sub MAC_CheckForDropBoxUpdate()
'======================================================================================================
'Checks if a folder on dropbox has number greater than the last creator save with a querytable
'If true then display the update Userform
'======================================================================================================
    
    Dim url$, workbookVersion As Date, QT As QueryTable, foundQuery As Boolean
    
    url = "https://www.dropbox.com/scl/fi/zao1xn7oexc4gypf8jfml/ReleaseVersion.json?rlkey=6q8t04c24mh1u9g4taq8lxoyv&st=p459w1px&dl=0"
    url = Replace$(url, "www.dropbox.com", "dl.dropboxusercontent.com")

    For Each QT In QueryT.QueryTables
        If InStrB(QT.Name, "MAC_Creator_CheckForDropBoxUpdate") <> 0 Then
            foundQuery = True
            Exit For
        End If
    Next QT
    
    If Not foundQuery Then
    
        Set QT = QueryT.QueryTables.Add("TEXT;" & url, Destination:=QueryT.Range("A1"))
    
        With QT
            .RefreshStyle = xlOverwriteCells
            .Name = "MAC_Creator_CheckForDropBoxUpdate"
            .WorkbookConnection.Name = "Creator Update Checks {MAC}"
            .AdjustColumnWidth = False
            .SaveData = False
        End With
    
    End If

    QT.Refresh False

    workbookVersion = Variable_Sheet.Range("Workbook_Update_Version").Value2
    
    Dim jsonResult As Object, jp As New JsonParserB, serverDate As Date
    
    With QT.ResultRange
    
        Set jsonResult = jp.Deserialize(Join(WorksheetFunction.Transpose(.Value2), vbNullString), True, False, False)
        
        #If DatabaseFile Then
            serverDate = jsonResult.Item("Database")
        #Else
            serverDate = jsonResult.Item(IIf(IsWorkbookForFuturesAndOptions, "Futures & Options", "Futures Only")).Item(ReturnReportType())
        #End If
        
        If workbookVersion < serverDate Then
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
            response = InputBox("1 : Remove images" & vbNewLine & "2 : Creator version" & vbNewLine & "3 : Release Workbook", "Choose macro")
            
            If response = "1" Then
                Remove_Images
            ElseIf response = "2" Then
                Creator_Version
            ElseIf response = "3" Then
                UpdateReleaseVersion
            End If
            
            'ThisWorkbook.VBProject.VBE.MainWindow.Visible = True
        End If
        
    #End If
    Exit Sub
Catch_FailedUnlock:
    MsgBox "Unlock failed."
End Sub
Sub Schedule_Data_Update(Optional Workbook_Open_EVNT As Boolean = False, Optional profiler As TimedTask)
'===================================================================================================================
    'Summary: Checks if new data is available and schedules the process to be run again at the next data release.
    'Inputs:
    '   Workbook_Open_EVNT - Set to True if running from the workbook open event.
    '   profiler - Optional TimedTask used to profile tasks within New_Data_Query().
'===================================================================================================================
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
        Call New_Data_Query(Scheduled_Retrieval:=True, Overwrite_All_Data:=False, IsWorbookOpenEvent:=Workbook_Open_EVNT, workbookEventProfiler:=profiler)
    
Scheduling_Next_Update:
    
        On Error Resume Next
        ' Unschedule this script if it is already slated to run.
        Application.OnTime unscheduleCftcUpdateTime, Procedure:=WBN & "Schedule_Data_Update", Schedule:=False
        Err.Clear
        
        If releaseScheduleHasBeenQueried Then
            
            nextCftcUpdateTime = CFTC_Release_Dates(False, True) 'Date and Time to schedule next data update check in Local Time.
            ' If current local time exceeds local time of the next stored release then update the release schedule.
            If Now > nextCftcUpdateTime Then
                ' Next update time couldn't be acquired, query source for updated dates..
                Application.Run WBN & "RefreshTimeZoneTable"
                ' Get time of next release.
                nextCftcUpdateTime = CFTC_Release_Dates(Find_Latest_Release:=False, convertToLocalTime:=True)
            End If
            
            If nextCftcUpdateTime > Now Then
                ' Schedule for 1:33:30 PM Friday.
                nextCftcUpdateTime = nextCftcUpdateTime + TimeSerial(0, 3, 30)
                
                Application.OnTime EarliestTime:=nextCftcUpdateTime, Procedure:=WBN & "Schedule_Data_Update", Schedule:=True
                ' Store DateTime so it can be used to unschedule on workbook close.
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

    If MsgBox("Auto-Scheduling and Retrieval checkbox couldn't be located." & String$(2, vbNewLine) & _
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
Private Sub UpdateReleaseVersion(Optional saveWorkbook As Boolean = True)
'======================================================================================================
'Edits a text file so that it holds the last saved date and time
'======================================================================================================
    #If Not Mac Then
        Dim fileNumber&, currentUtcTime As Date
        
        On Error GoTo Propagate
        
        If IsOnCreatorComputer Then
                           
            Dim jsonPath$, savedDates As Object, jp As New JsonParserB, formattedDate$, json$
    
            jsonPath = Environ$("OneDriveConsumer") & "\COT Workbooks\ReleaseVersion.json"
            
            If FileOrFolderExists(jsonPath) Then
                fileNumber = FreeFile
                Open jsonPath For Input As #fileNumber
                    json = Input(LOF(fileNumber), #fileNumber)
                    If LenB(json) <> 0 Then Set savedDates = jp.Deserialize(json, False)
                Close #fileNumber
            End If
                                
            currentUtcTime = GetUtcTime
            formattedDate = Format$(currentUtcTime, "yyyy-MM-ddTHH:mm:ssZ")
            
            If savedDates Is Nothing Then Set savedDates = CreateObject("Scripting.Dictionary")
            
            #If DatabaseFile Then
                savedDates("Database") = formattedDate
            #Else
            
                Dim secondKey$: secondKey = IIf(IsWorkbookForFuturesAndOptions, "Futures & Options", "Futures Only")
                
                If Not savedDates.Exists(secondKey) Then
                    Set savedDates(secondKey) = CreateObject("Scripting.Dictionary")
                End If
                
                savedDates(secondKey)(ReturnReportType()) = formattedDate
            
            #End If
            
            fileNumber = FreeFile
            Open jsonPath For Output As #fileNumber
            Print #fileNumber, jp.Serialize(savedDates, False, True)
            Close #fileNumber
            
            Variable_Sheet.Range("Workbook_Update_Version").Value2 = currentUtcTime 'Update saved Last_Saved_Time and date within workbook
            
            If saveWorkbook Then Custom_Save
        End If
    #End If
    
    Exit Sub
Propagate:
    PropagateError Err, "UpdateReleaseVersion"
End Sub

Public Sub Save_Workbooks()
Attribute Save_Workbooks.VB_Description = "Saves all workbooks that have a Custom_Save macro."
Attribute Save_Workbooks.VB_ProcData.VB_Invoke_Func = " \n14"
'================================================================================================
' Saves all currently open workbooks that have a Custom_Save macro.
'================================================================================================
    Dim Valid_Workbooks As New Collection, Saved_STR$, wb As Workbook, savedState As Boolean, markForRelease As Boolean
    
    Saved_STR = "The following workbooks were saved >" & String$(2, vbNewLine)
    
    If IsOnCreatorComputer Then
        markForRelease = MsgBox("Mark for release?", vbYesNo) = vbYes
    End If
    
    On Error GoTo Save_Error
    
    For Each wb In Workbooks
        With wb
            If Not .Name Like "PERSONAL.*" Then
                With Application
                    .ScreenUpdating = False
                    .EnableEvents = False
                End With
                
                If markForRelease Then
                    Application.Run "'" & .Name & "'!UpdateReleaseVersion"
                Else
                    Application.Run "'" & .Name & "'!Custom_Save"
                End If
                
                If .Saved Then
                    Valid_Workbooks.Add wb
                Else
                    MsgBox ("Unable to save " & .fullName)
                End If
            End If
        End With
Resume_Workbook_Loop:
    Next wb
        
    On Error GoTo 0
    Saved_STR = "Saved " & Valid_Workbooks.Count & " workbooks." & String$(2, vbNewLine)
    For Each wb In Valid_Workbooks
        With wb
            .Saved = True
            Saved_STR = Saved_STR & vbTab & .Name & vbNewLine
        End With
    Next wb
    
    MsgBox Saved_STR
    Application.EnableEvents = True
    Exit Sub

Save_Error:
    MsgBox ("Unable to save " & wb.fullName)
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
    
    Dim savedState As Boolean, i As Byte, shapeNames$()
    
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
                    shapeNames = Split("Holding,Temp,Unbound", ",")
                    With Variable_Sheet
                        For i = LBound(shapeNames) To UBound(shapeNames)
                            .Shapes(shapeNames(i)).Delete
                        Next i
                    End With
                    Erase shapeNames
                    On Error GoTo 0
                End If
                
                If LenB(fileName) = 0 Then
                    fileName = Application.GetSaveAsFilename
                End If
                
                If fileName <> "FALSE" Then .SaveAs fileName
            End If
            savedState = .Saved
        End With
        '
        'Enable events is turned on within After_Save
        Call After_Save
        .DisplayAlerts = True
        Courtesy
        ThisWorkbook.Saved = savedState
    End With
    
    Exit Sub
DisplayErr:
    DisplayErr Err, "Save_Workbook"
End Sub
Private Sub Before_Save(Enable_Events_Toggle As Boolean)

    Dim onCreatorMachine As Boolean, saveTimer As TimedTask, appProperties As Collection
    
    Const ProcedureName$ = "Before_Save", switchVersionTask$ = "Switch user display"
    
    On Error GoTo PropagateErr
    
    With ThisWorkbook
        Set .ActiveSheetBeforeSaving = .ActiveSheet
    End With
    
    onCreatorMachine = IsOnCreatorComputer
    
    Set appProperties = DisableApplicationProperties(disableEvents:=True, disableAutoCalculations:=False, disableScreenUpdating:=True)
    
    If onCreatorMachine Then
        
        On Error GoTo Catch_Compilation_Error
        Application.VBE.CommandBars.FindControl(Type:=msoControlButton, ID:=578).Execute
        
        Set saveTimer = New TimedTask
        
        With saveTimer
            .Start ThisWorkbook.Name & " (" & Now & ") ~ Save Event"
            On Error GoTo PropagateErr
            
            With .StartSubTask(switchVersionTask)
                Call Remove_Images(executeAsPartOfSaveEvent:=True)
                .EndTask
            End With
        End With
        
        If HUB.ProtectContents = False Then Call Worksheet_Protection_Toggle(HUB, True, False)
        
        With ThisWorkbook.Event_Storage
            On Error Resume Next
            .Remove SaveEventTimerKey
            On Error GoTo PropagateErr
            .Add saveTimer, SaveEventTimerKey
        End With
        
        #If DatabaseFile Then
            ' This is done so that the file saves in a state in which end users will initially see all contracts
            ' when using the Contract_Selection Userform.
            Variable_Sheet.Range("Enable_Favorites").Value2 = False
        #End If
        
    End If
    
    With HUB
        .Shapes("Macro_Check").Visible = True
        .Shapes("Diagnostic").Visible = False
        .Shapes("DN_List").Visible = False
        Application.EnableEvents = True
        If Not ThisWorkbook.ActiveSheet Is HUB Then .Activate
        Application.EnableEvents = False
    End With
    
    Dim valuesToEditBeforeSave As New Collection, rngVar As Range, i As Byte, rangeNames$()
    
    rangeNames = Split("Triggered_Project_Unlock,Triggered_Data_Schedule,Release_Schedule_Queried,CreatorActiveState,SqlServerName", ",")
    
    With valuesToEditBeforeSave
        On Error GoTo Catch_Range_Not_Found
        For i = LBound(rangeNames) To UBound(rangeNames)
            Set rngVar = Variable_Sheet.Range(rangeNames(i))
            .Add Array(rngVar, rngVar.Value2), rangeNames(i)
            
            If rangeNames(i) = "SqlServerName" Then
                If onCreatorMachine Then
                    rngVar.Value2 = Empty
                Else
                    .Remove rangeNames(i)
                End If
            Else
                rngVar.Value2 = False
            End If
Next_Range_Name:
        Next i
        On Error GoTo PropagateErr
    End With
    
    With ThisWorkbook.Event_Storage
        On Error Resume Next
        .Remove Saved_Variables_Key
        .Add valuesToEditBeforeSave, Saved_Variables_Key
    End With
    
    On Error GoTo PropagateErr
    'Turn back on to allow After_Save if not running custom_save macro
     Application.EnableEvents = Enable_Events_Toggle

    If onCreatorMachine Then saveTimer.StartSubTask SaveEventDurationKey
    
    Exit Sub

Catch_Compilation_Error:
    If Err.Number = -2147467259 Then
        ' Already compiled.
        Resume Next
    Else
        Err.Description = "Compile error in project."
        GoTo PropagateErr
    End If
Catch_Range_Not_Found:
    Dim displayMsg As Boolean: displayMsg = True
    #If Not DatabaseFile Then
        'SQL Server variables only apply for the database version.
        If LCase$(rangeNames(i)) Like "*sql*" Then
            displayMsg = False
        End If
    #End If
    If displayMsg Then MsgBox "Couldn't find named range: Variable_Sheet." & rangeNames(i)
    Resume Next_Range_Name
PropagateErr:
    If Not appProperties Is Nothing Then EnableApplicationProperties appProperties
    PropagateError Err, ProcedureName
End Sub
Private Sub After_Save()

    Dim rangeAndValueArray As Variant, Creator As Boolean, workbookState As Boolean
    Const ProcedureName$ = "After_Save"
    
    On Error GoTo PropagateErr
    
    Application.EnableEvents = False
    
    With ThisWorkbook
        
        workbookState = .Saved
        On Error GoTo PropagateErr
        Creator = IsOnCreatorComputer()
        With .Event_Storage
                        
            If HasKey(ThisWorkbook.Event_Storage, Saved_Variables_Key) Then
                On Error Resume Next
                For Each rangeAndValueArray In .Item(Saved_Variables_Key)
                    rangeAndValueArray(0).Value2 = rangeAndValueArray(1)
                Next rangeAndValueArray
                .Remove Saved_Variables_Key
            End If
            
            If Creator Then
                With .Item(SaveEventTimerKey)
                    On Error GoTo Print_SaveEvntTimes
                    .StopSubTask SaveEventDurationKey
                    If Variable_Sheet.Range("CreatorActiveState").Value2 = True Then
                        With .StartSubTask("Enable Creator Version")
                            Run_This ThisWorkbook, "Creator_Version"
                            .EndTask
                        End With
                    End If
Print_SaveEvntTimes:
                    .DPrint
                End With
                On Error Resume Next
                .Remove SaveEventTimerKey
            End If
            
        End With
        
        On Error GoTo PropagateErr
        If Not .ActiveSheetBeforeSaving Is Nothing Then
            .ActiveSheetBeforeSaving.Activate
            Set .ActiveSheetBeforeSaving = Nothing
        End If
        
    End With
    
    With HUB
        .Shapes("Macro_Check").Visible = False 'Turns this textbox back off if macros are enabled
        If Not Creator Then
            .Shapes("Diagnostic").Visible = True
            #If DatabaseFile Then
                If Variable_Sheet.Range("Github_Version").Value2 = True Then .Shapes("DN_List").Visible = True
            #Else
                .Shapes("DN_List").Visible = True
            #End If
        Else
            #If DatabaseFile Then
                .Shapes("Database_Paths").Visible = False
            #End If
        End If
    End With
        
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
    ThisWorkbook.Saved = workbookState
    
    Exit Sub
PropagateErr:
    PropagateError Err, ProcedureName
End Sub

Private Sub Show_Chart_Settings()
    Chart_Settings.Show
End Sub
'Private Sub Adjust_Dash_Shapes()
'
'    Dim FUT As Shape, FutOpt As Shape, wd As Double, DashWs As Worksheet
'
'    Set DashWs = ThisWorkbook.ActiveSheet
'
'    With DashWs
'        Set FUT = .Shapes("FUT only")
'        wd = (.Range("C1:E1").Width - 10) / 2
'        Set FutOpt = .Shapes("FUT+OPT")
'    End With
'
'    With FUT
'        .Top = 0
'        .Left = DashWs.Range("c1").Left
'        .Height = DashWs.Range("c1").Height
'        .Width = wd
'        .OLEFormat.Object.value = 1
'    End With
'
'    With FutOpt
'        .Top = 0
'        .Left = FUT.Left + FUT.Width + 10
'        .Height = FUT.Height
'        .Width = wd
'        .OLEFormat.Object.value = xlOff
'    End With
'
'    With DashWs.Shapes("Options")
'        .Left = FUT.Left
'        .Width = 2 * wd + 10
'        .Height = FUT.Height + 5
'        .Top = FUT.Top
'    End With
'
'    With DashWs.Shapes("Generate Dash")
'        .Top = FUT.Top
'        .Left = FutOpt.Left + FutOpt.Width + 10
'        .Height = FUT.Height
'    End With
'
'End Sub

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
Public Sub SelectWorkbooksToOpen()
Attribute SelectWorkbooksToOpen.VB_Description = "Opens all selecteed workbooks within the current application."
Attribute SelectWorkbooksToOpen.VB_ProcData.VB_Invoke_Func = " \n14"
'===================================================================================================================
    'Summary: Prompts the user for Excel workbooks to open and opens them.
'===================================================================================================================
    Dim wbNames() As Variant, i As Byte
    On Error GoTo UserQuit
    wbNames = Application.GetOpenFilename("Excel Workbooks (*.xls*),*.xls*", , "Select workbooks to open", , True)
    
    If IsArrayAllocated(wbNames) Then
        On Error Resume Next
        For i = LBound(wbNames) To UBound(wbNames)
            Workbooks.Open wbNames(i)
        Next i
        Err.Clear
    End If
    Exit Sub
UserQuit:
    'PropagateError Err, "SelectWorkbooksToOpen"
End Sub
Private Sub WorkbookOpenEvent()
'===================================================================================================================
    'Summary: Executes all necessary workbook open events.
'===================================================================================================================
    Dim onCreatorMachine As Boolean, Error_STR$, iCount As Byte, _
    Field_Count As Byte, databaseConnectedWorkbook As Boolean, nameOfRangesToAlter$()
    
    Dim openEventTimer As New TimedTask, disableDataCheck As Boolean, isWorkBookFromGitHub As Boolean, eventErrors As New Collection
       
    Const donatorInfoShapeName$ = "DN_List", clickToDonateShapeName$ = "Donate", gitHubRangeName$ = "Github_Version", _
                                    diagnosticShapeName$ = "Diagnostic", sqlServerName$ = "SqlServerName"
    
    #If Mac And DatabaseFile Then
        GateMacAccessToWorkbook
    #End If
    
    Call IncreasePerformance
    
    On Error GoTo Catch_OperatingSystemRange_404
    
    Variable_Sheet.Range("OperatingSystem").Value2 = Evaluate("=INFO(""system"")")
    
    On Error GoTo 0
    
    onCreatorMachine = IsOnCreatorComputer
    
    #If DatabaseFile And Not Mac Then
    
        databaseConnectedWorkbook = True
        On Error GoTo Catch_MissingDatabase
        ' Determine if any database files have been moved or can no longer be found.
        Run_This ThisWorkbook, "Database_Interactions.FindDatabasePathInSameFolder"
        
Check_If_Exe_Available:
        On Error GoTo 0
        
        With Variable_Sheet
            If onCreatorMachine Then .Range(sqlServerName).Value2 = GetCreatorPasswordsAndCodes("SQL_Connection_String")

            isWorkBookFromGitHub = .Range(gitHubRangeName).Value2
            On Error GoTo C_EXE_Not_Defined
            
            With .Range("CSharp_Exe")
                If Not FileOrFolderExists(.Value2) Then .Value2 = Empty
            End With
        End With
        
    #ElseIf Not DatabaseFile Then
        PopulateListBoxes True, True, False
    #End If
Start_Event_Timer:

    On Error Resume Next

    openEventTimer.Start "[" & ThisWorkbook.Name & "] Workbook Open event"
    
    ' Update ranges with default values.
    nameOfRangesToAlter = Split("Triggered_Project_Unlock,Triggered_Data_Schedule,Release_Schedule_Queried,DropBox_Date_Query_Time,Data_Retrieval_Time", ",")
    
    On Error GoTo AttemptNextDefaultAssignment
    For iCount = LBound(nameOfRangesToAlter) To UBound(nameOfRangesToAlter)
        
        With Variable_Sheet.Range(nameOfRangesToAlter(iCount))
            Select Case nameOfRangesToAlter(iCount)
                Case "DropBox_Date_Query_Time", "Data_Retrieval_Time"
                    .Value2 = Empty
                Case Else
                    .Value2 = False
            End Select
        End With
        
        Field_Count = Field_Count + 1
AttemptNextDefaultAssignment:
        If Err.Number <> 0 Then On Error GoTo -1
    Next iCount

    On Error Resume Next
    
    If onCreatorMachine And Field_Count <> UBound(nameOfRangesToAlter) - LBound(nameOfRangesToAlter) + 1 Then
        MsgBox "1 or more fields weren't found when setting default settings on the Variable Sheet during the workbook open events."
    End If
    
    With HUB
        .Visible = xlSheetVisible
        'Turns off disclaimer box if macros are on
        .Shapes("Macro_Check").Visible = False
        'Makes donator count invisIble
        .Shapes(donatorInfoShapeName).Visible = False
        
        If Not onCreatorMachine Then
            .Shapes(diagnosticShapeName).Visible = True
            
            #If DatabaseFile Then
                If Not isWorkBookFromGitHub Then .Shapes(clickToDonateShapeName).Visible = True
            #Else
                .Shapes(clickToDonateShapeName).Visible = True
            #End If
            
            If Not isWorkBookFromGitHub Then
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
        If Not onCreatorMachine Then
            If Not isWorkBookFromGitHub Then Call CheckForDropBoxUpdate
            Call UploadStats
        End If
        
        On Error Resume Next
        Application.Run "Query_Tables.RefreshTimeZoneTable", eventErrors
        Err.Clear
    End With
    
    If Not disableDataCheck Then Call Schedule_Data_Update(Workbook_Open_EVNT:=True, profiler:=openEventTimer)
    
Build_Error_String:
    With eventErrors 'Load all error messages into a singular string
        If .Count > 0 Then
            On Error GoTo Next_EVNT_Item
            For iCount = .Count To 1 Step -1
                Error_STR = .Item(iCount) & String$(2, vbNewLine) & Error_STR
Next_EVNT_Item: On Error GoTo -1
            Next iCount
            If LenB(Error_STR) <> 0 Then MsgBox Error_STR, Title:="Error Message"
        End If
    End With
    
    On Error GoTo 0
Finally:
    openEventTimer.DPrint
    Call Re_Enable
    If Not onCreatorMachine And Err.Number = 0 Then Courtesy
    Exit Sub
    
#If DatabaseFile Then
Catch_MissingDatabase:
    disableDataCheck = True
    Resume Check_If_Exe_Available
#End If

Catch_OperatingSystemRange_404:

    Dim addedRow As ListRow
        
    Set addedRow = Variable_Sheet.ListObjects("Saved_Variables").ListRows.Add(, False)
    
    With addedRow.Range
        .Value2 = Array("Operating System", Evaluate("=INFO(""system"")"))
        .Cells(1, 2).Name = "OperatingSystem"
    End With
    
    Resume Next
    
C_EXE_Not_Defined:
    DisplayErr Err, "WorkbookOpenEvent", "Variable_Sheet.Range('CSharp_Exe') not defined."
    Resume Start_Event_Timer
    
Log_ERR_Resume_Next:
    eventErrors.Add Err.Description
    Resume Next
    
Catch_GeneralError:
    DisplayErr Err, "WorkbookOpenEvent"
    Resume Next
End Sub
Private Sub UploadStats()
'===================================================================================================================
    'Summary: Uploads anonymous user info to a google spread sheet.
'===================================================================================================================
    Dim apiResponse$, postUpload$, workbookVersion$, postData$, jp As New JsonParserB
    
    On Error GoTo Catch_POST_FAILED
    Const apiUrl$ = "https://ipapi.co/json/", ip$ = "ip", city$ = "city", country$ = "country_name", region$ = "region"
    
    Const FormURL$ = "https://docs.google.com/forms/d/e/1FAIpQLSfDB8cfBFZFcPf15tnaxuq6OStkmRYm4VlZjYWE8PEvm6qhFA/formResponse"
    
    #If DatabaseFile Then
        workbookVersion = "Database"
    #Else
        workbookVersion = ReturnReportType & "_" & IIf(IsWorkbookForFuturesAndOptions, "Combined", "FUT")
    #End If
    
    On Error GoTo Catch_GET_FAILED
 
    If TryGetRequest(apiUrl, apiResponse) Then
    
        On Error GoTo CATCH_JSON_PARSER_FAILURE
        
        With jp.Deserialize(apiResponse)
            On Error GoTo MissingKey
            postData = "&entry.227122838=" & .Item(city) & _
                        "&entry.1815706984=" & .Item(region) & _
                        "&entry.55364550=" & .Item(country) & _
                        "&entry.1590825643=" & .Item(ip) & _
                        "&entry.1917934143=" & workbookVersion & _
                        "&entry.1144002976=" & Format$(Variable_Sheet.Range("Workbook_Update_Version").Value2, "yyyy-mm-dd hh:mm:ss")
        End With
        
        On Error GoTo Catch_POST_FAILED
        HttpPost FormURL, postData, True
                
    End If
    
    Exit Sub
Catch_GET_FAILED:
    PropagateError Err, "Stats", "GET failed."
Catch_POST_FAILED:
    PropagateError Err, "Stats", "POST failed."
CATCH_JSON_PARSER_FAILURE:
    PropagateError Err, "Stats", "JSON Parser failed."
MissingKey:
    PropagateError Err, "Stats", "JSON key not found."
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
    '===================================================================================================================
        'Summary: Checks if new data is available and schedules the process to be run again at the next data release.
        'Inputs:
        '   updateHub - Set to True if you want to update the ComboBox on the HUB worksheet.
        '   updateCharts - Set to True if you want to update the ComboBox on the Charts worksheet.
        '   formComboBox - If not nothing then this combobox will be filled with available sheet names.
    '===================================================================================================================
        Dim WorksheetNameCLCTN As New Collection, WS As Worksheet, _
        validCountBasic As Long, wsKeys() As Variant, contractKeys() As Variant, validContractCount As Long
        
        ReDim wsKeys(1 To ThisWorkbook.Worksheets.Count)
        ReDim contractKeys(1 To ThisWorkbook.Worksheets.Count)
        
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
        
        For i = 1 To HN.Count
        
            Set TBL_RNG = HN(i).TableSource.Range       'Entire range of table
            Set Worksheet_TB = TBL_RNG.Parent 'Worksheet where table is found
            
            With Worksheet_TB '{Must be typed as object to fool the compiler when resetting the Used Range]
        
                With TBL_RNG 'Find the Bottom Right cell of the table
                    Set TB_Last_Cell = .Cells(.Rows.Count, .columns.Count)
                End With
                
                With .UsedRange 'Find the Bottom right cell of the Used Range
                    Set UR_LastCell = .Cells(.Rows.Count, .columns.Count)
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
        
        For Tb = 1 To Valid_Table_Info.Count
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
        
        For i = 1 To Valid_Table_Info.Count
        
            Set Target_TableR = Valid_Table_Info(i).TableSource.DataBodyRange  'databodyrange of the target table
          
            If Not Target_TableR.Parent Is ASH Then 'if the worksheet objects aren't the same
                       
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
                    If .Count > 0 Then 'if at least 1 column was hidden then reapply hidden roperty to specified column
                        For T = 1 To .Count
                            Hidden_Collection(T).EntireColumn.Hidden = True
                        Next T
                        Set Hidden_Collection = Nothing 'empty the collection
                    End If
                End With
                
            End If
            
        Next i
        
        With Original_Hidden_Collection
            If .Count > 0 Then 'if at least 1 column was hidden then reapply hidden roperty to specified column
                For T = 1 To .Count
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
        MsgBox "Function Get_Worksheet_Info is unavailable. This Macro is intended for files created by MoshiM." & String$(2, vbNewLine) & _
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
    
        Dim Headers() As Variant, Tb As ListObject, availableContracts As Collection, CI As ContractInfo
        
        Set Tb = ReturnCftcTable(ThisWorkbook.ActiveSheet)
        
        If Not Tb Is Nothing Then
            Set availableContracts = GetAvailableContractInfo
            Headers = Tb.HeaderRowRange.Value2
        
            For Each CI In availableContracts
                With CI
                    If Not .TableSource Is Tb Then
                        .TableSource.HeaderRowRange.Resize(1, UBound(Headers, 2)) = Headers
                    End If
                End With
            Next CI
        End If
    
    End Sub
    Private Sub DeleteDataGreaterThanDate()
    
        Dim rowsToDeleteCount As Long, tblRange As Range, CC As ContractInfo, _
        rr As Range, Tb As ListObject, minDateToKeep As Date
        minDateToKeep = CDate(InputBox("yyyy-mm-dd"))
        For Each CC In GetAvailableContractInfo
                    
            Set tblRange = CC.TableSource.DataBodyRange
            
            Set rr = tblRange.columns(1)
            
            Set rr = rr.Find(Format(minDateToKeep, "yyyy-mm-dd"), , xlValues, xlWhole)
            
            If Not rr Is Nothing Then
                Set Tb = CC.TableSource
                
                With tblRange.Parent
                    .Range(rr.offset(1), .Cells(tblRange.Rows.Count + 1, tblRange.columns.Count)).ClearContents
                    Tb.Resize Range(CC.TableSource.Range.Cells(1, 1), .Cells(rr.row, tblRange.columns.Count))
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




