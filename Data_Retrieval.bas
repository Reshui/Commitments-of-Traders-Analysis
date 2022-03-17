Attribute VB_Name = "Data_Retrieval"

Public Const TypeF As String = "L"
Public Const date_column As Long = 3
Public Retrieval_Halted_For_User_Interaction As Boolean
Public Data_Updated_Successfully As Boolean

Public Running_Weekly_Retrieval As Boolean

Public Enum Append_Type
    Add_To_Old = 1
    Multiple_1d = 2
    multiple_2d = 3
End Enum

Public Enum Data_Identifier
    Block_Data = 14
    Old_Data = 44
    Weekly_Data = 33
End Enum

'Public Debug_Timer As Double

Option Explicit

Sub New_Data_Query(Optional Scheduled_Retrieval As Boolean = False)

'======================================================================================================
'Retrieves the latest data and outputs it to the worksheet if available
'======================================================================================================

    Dim Last_Update_CFTC As Long, Start_Time As Double, INTD_Timer As Double, CFTC_Incoming_Data_Date As Long, ICE_Incoming_Data_Date As Long

    Dim Debug_Mode As Boolean, Time_Record As String, Mac_User As Boolean, CFTC_Data() As Variant, report As Variant, ICE_Data() As Variant, Historical_Query() As Variant

    Dim DBM_Weekly_Retrieval As Boolean, DBM_Historical_Retrieval As Boolean, price_column As Long, combined_workbook_bool As Variant

    Dim update_workbook_tables As Boolean, Last_Update_ICE As Long ', All_Available_Contracts() As Variant
    
    Dim Download_CFTC As Boolean, Download_ICE As Boolean, Check_ICE As Boolean, new_data_found As Boolean
    
    Dim DataBase_Not_Found_CLCTN As New Collection, Symbol_Info As Collection, Data_CLCTN As Collection, Legacy_Combined_Data As Boolean
    
    Const contract_code_column As Long = 4, legacy_initial As String = "L"
    
    Retrieval_Halted_For_User_Interaction = False
    
    Running_Weekly_Retrieval = True
    
    #If Mac Then
        Mac_User = True
    #End If

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Set Symbol_Info = Get_Price_Symbols
    
    On Error GoTo Deny_Debug_Mode
    
    If Weekly.Shapes("Test_Toggle").OLEFormat.Object.value = 1 Then Debug_Mode = True 'Determine if Debug status
    
    If Debug_Mode = True Then
    
        If MsgBox("Test Weekly Data Retrieval ?", vbYesNo, "Choose what to debug") = vbYes Then
            DBM_Weekly_Retrieval = True
        ElseIf MsgBox("Test Multi-Week Historical Retrieval ?", vbYesNo, "Choose what to debug") = vbYes Then
            DBM_Historical_Retrieval = True
        End If
    
    End If
    
Retrieve_Latest_Data:

    On Error GoTo 0

    For Each report In Array(legacy_initial, "D", "T") 'Legacy data must be retrieved first so that price data only needs to be retrieved once
        '
        Start_Time = Timer
    
        For Each combined_workbook_bool In Array(True, False) 'True must be first so that price data can be retrieved for futures only data
            
            INTD_Timer = Timer 'Start Data Retrieval Timer
            
            If combined_workbook_bool And report = legacy_initial Then
                Legacy_Combined_Data = True
            Else
                Legacy_Combined_Data = False
            End If
            
            Last_Update_CFTC = Latest_Date(Report_Type:=CStr(report), combined_wb_bool:=CBool(combined_workbook_bool), ICE_Query:=False)   'The date the data was last sorted for
            
            If Database_Interactions.DataBase_Not_Found = True Then
                
                With DataBase_Not_Found_CLCTN
                    
                    .Add Evaluate("VLOOKUP(""" & report & """,Report_Abbreviation,2,FALSE)"), "T Name"
                    .Add "Missing database for " & IIf(.Item("T Name") = "TFF", "Traders in Financial Futures", .Item("T Name")) & " in range " & Assign_Linked_Data_Sheet(CStr(report)).Range("Database_Path").Address(, , , True)
                    .Remove "T Name"
                    
                End With
                
                If Legacy_Combined_Data Then
                    'The Legacy_Combined data is the only one for which price data is queried.
                    GoTo Exit_Procedure
                Else
                    Exit For
                End If
                
            End If
            
            If CFTC_Incoming_Data_Date = 0 Or Last_Update_CFTC < CFTC_Incoming_Data_Date Or Debug_Mode Then
                'Initially only data for legacy combined gets retrieved
                'All other versions have their most recent database date queried to see if it is less than this value
                CFTC_Data = HTTP_Weekly_Data(Last_Update_CFTC, Auto_Retrieval:=Scheduled_Retrieval, Combined_Version:=CBool(combined_workbook_bool), Report_Type:=CStr(report), DebugMD:=DBM_Weekly_Retrieval)
                If CFTC_Incoming_Data_Date = 0 Then CFTC_Incoming_Data_Date = CFTC_Data(1, date_column)
            Else
                'Currently on Non-Legacy Combined Data and no new data is available, Button clicked manully and not in debug mode
                Exit For
            End If

            If report = "D" Then
            
                On Error GoTo ICE_Retrieval_Failed
                
                Check_ICE = True
                Last_Update_ICE = Latest_Date(Report_Type:=CStr(report), combined_wb_bool:=CBool(combined_workbook_bool), ICE_Query:=True) 'The date the data was last sorted for
                
                ICE_Data = Weekly_ICE(Last_Update_ICE, Mac_User, CStr(report), CBool(combined_workbook_bool), CDate(CFTC_Data(1, date_column)))
                
                ICE_Incoming_Data_Date = ICE_Data(1, date_column)
            Else
                Check_ICE = False
            End If

Finished_ICE:

            price_column = UBound(CFTC_Data, 2) + 1
            
            Time_Record = "[" & report & "]-Data Retrieval: " & Round(Timer - INTD_Timer, 2) & " seconds. " & vbNewLine

            INTD_Timer = Timer 'Start timing how long it takes to sort & calculate data
        
            If Not Debug_Mode And CFTC_Incoming_Data_Date = Last_Update_CFTC And (Not Check_ICE Or (Check_ICE And ICE_Incoming_Data_Date = Last_Update_ICE)) Then
                    
                If Legacy_Combined_Data Then
                    GoTo Exit_Procedure
                Else
                    GoTo Next_Combined_Value
                End If
                
            ElseIf CFTC_Incoming_Data_Date - Last_Update_CFTC > 7 Or (Check_ICE And ICE_Incoming_Data_Date - Last_Update_ICE > 7) Or DBM_Historical_Retrieval Then
                
                If CFTC_Incoming_Data_Date - Last_Update_CFTC > 7 Or DBM_Historical_Retrieval Then
                    Download_CFTC = True
                Else
                    Download_CFTC = False
                End If
                
                If Check_ICE And (ICE_Incoming_Data_Date - Last_Update_ICE > 7 Or DBM_Historical_Retrieval) Then
                    Download_ICE = True
                Else
                    Download_ICE = False
                End If
                
                Historical_Query = Missing_Data(retrieve_combined_data:=CBool(combined_workbook_bool), _
                    ICE_Data:=ICE_Data, CFTC_Data:=CFTC_Data, _
                    Download_ICE_Data:=Download_ICE, Download_CFTC_Data:=Download_CFTC, _
                    Report_Type:=CStr(report), _
                    CFTC_Last_Updated_Day:=Last_Update_CFTC, ICE_Last_Updated_Day:=Last_Update_ICE, _
                    DebugMD:=DBM_Historical_Retrieval)  'Will download missing data and overwrite the current array
                    
                If Not (Download_ICE And Download_CFTC) Then
                    
                    If Download_CFTC And (Check_ICE And ICE_Incoming_Data_Date - Last_Update_ICE > 0) Then
                        'Determine if the most recently queried Ice Data needs to be added
                        Set Data_CLCTN = New Collection
                        
                        With Data_CLCTN
                            .Add Historical_Query
                            .Add ICE_Data
                        End With
                        
                        Historical_Query = Multi_Week_Addition(Data_CLCTN, multiple_2d)
                        
                    ElseIf Download_ICE And CFTC_Incoming_Data_Date - Last_Update_CFTC > 0 Then
                        'Determine if CFTC data needs to be added

                        Set Data_CLCTN = New Collection
                        
                        With Data_CLCTN
                            .Add Historical_Query
                            .Add CFTC_Data
                        End With
                        
                        Historical_Query = Multi_Week_Addition(Data_CLCTN, multiple_2d)
                        
                    End If
                    
                End If
                
                'If Retrieval_Halted_For_User_Interaction Then Exit Sub
                             
                update_workbook_tables = True
                new_data_found = True
                
            ElseIf (CFTC_Incoming_Data_Date - Last_Update_CFTC) > 0 Or (Check_ICE And ICE_Incoming_Data_Date - Last_Update_ICE > 0) Or Debug_Mode = True Then  'If just a 1 week difference
                
                Set Data_CLCTN = New Collection
                
                With Data_CLCTN
                
                    If Check_ICE And (ICE_Incoming_Data_Date - Last_Update_ICE > 0 Or Debug_Mode) Then
                        .Add ICE_Data
                    End If
                    
                    If CFTC_Incoming_Data_Date - Last_Update_CFTC > 0 Or Debug_Mode Then
                        .Add CFTC_Data
                    End If
                    
                    If .Count = 1 Then
                        Historical_Query = .Item(1)
                    ElseIf .Count = 2 Then
                        Historical_Query = Multi_Week_Addition(Data_CLCTN, multiple_2d)
                    Else
                        Exit For
                    End If
                    
                End With
                
                update_workbook_tables = True
                new_data_found = True
                
            End If
        
            If new_data_found = True Then
        
                ReDim Preserve Historical_Query(LBound(Historical_Query, 1) To UBound(Historical_Query, 1), LBound(Historical_Query, 2) To price_column)  'Expand for calculations
                
                Call Block_Query(Query:=Historical_Query, Code_Column:=contract_code_column, price_column:=price_column, Report_Type:=CStr(report), processing_combined_data:=CBool(combined_workbook_bool), Symbol_Info:=Symbol_Info)

                Time_Record = Time_Record & "[" & report & "]-Loop Time: " & Round(Timer - INTD_Timer, 2) & " seconds." & vbNewLine
                
                new_data_found = False
                
                If combined_workbook_bool = True Then Call save_recent_contract_identifiers(CStr(report))
                
            End If
            
            Debug.Print Time_Record & "[" & report & "]-Time to completion: " & Round(Timer - Start_Time, 2) & " seconds. " & Now & vbNewLine

Next_Combined_Value:

        Next combined_workbook_bool

Next_Report_Release_Type:
    
    Next report
      
Exit_Procedure:
    
    If update_workbook_tables Then
        
        Data_Updated_Successfully = True
        
        If CFTC_Incoming_Data_Date > Range("Most_Recently_Queried_Date").value Then
            Update_Text CFTC_Incoming_Data_Date   'Update Text Boxes "My_Date" on the HUB and Weekly worksheets.
            Range("Most_Recently_Queried_Date").value = CFTC_Incoming_Data_Date
        End If
        
        If Not Debug_Mode Then
            If Not Scheduled_Retrieval Then HUB.Activate 'If ran manually then bring the User to the HUB
            Courtesy                                     'Change Status Bar_Message
        End If
        
    ElseIf DataBase_Not_Found_CLCTN.Count > 0 And Not Scheduled_Retrieval Then

        For Each report In DataBase_Not_Found_CLCTN
            MsgBox report
        Next report
        
    ElseIf Not Scheduled_Retrieval Then
    
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
    
    End If
    
    Set Symbol_Info = Nothing
    
    Running_Weekly_Retrieval = False
    
    Re_Enable
    
    Exit Sub

Deny_Debug_Mode:
    
    Debug_Mode = False
    
    Resume Retrieve_Latest_Data

ICE_Retrieval_Failed:
    
    Check_ICE = False
    Resume Finished_ICE
    
End Sub
Private Sub Block_Query(ByRef Query, Code_Column As Long, price_column As Long, Report_Type As String, processing_combined_data As Boolean, Symbol_Info As Collection)

'======================================================================================================
'Subroutine takes a properly formatted array and outputs all contracts to where they need to go.
'Overwrite_Worksheet =True when all data on a worksheet needs to be replaced
'======================================================================================================

Dim X As Long, C As Long, w As Long, WS_Data() As Variant, T As Long, total_count As Long

Dim Block() As Variant, Contract_CLCTN As New Collection, contract_code As String, retrieve_price_data As Boolean

'=======================================================================================================================
'Progress Bar variables
Dim Progress_CHK As CheckBox, Bar_Increment As Double, Progress_Control As MSForms.Label, Percent_Mod As Long

    Set Progress_CHK = Weekly.Shapes("Progress_CHKBX").OLEFormat.Object
    
    If processing_combined_data And Report_Type = "L" Then
    
        If Progress_CHK.value = 1 Then 'Show Progress Bar if Toggled on
            Call Progress_Bar_Custom_Initialize(Progress_Control, Bar_Increment, UBound(Query, 1), Percent_Mod)
        End If
    
Block_Query_Main_Function:     On Error GoTo 0
    
        Set Progress_CHK = Nothing
    
        ReDim Block(1 To UBound(Query, 2))
        
        For X = LBound(Query, 1) To UBound(Query, 1) 'Add contracts to their own collection for grouping
                                                     'Array should be date sorted
            For C = LBound(Query, 2) To UBound(Query, 2)
                Block(C) = Query(X, C)
            Next C
            
            If Not HasKey(Contract_CLCTN, CStr(Block(Code_Column))) Then
                Contract_CLCTN.Add New Collection, Block(Code_Column)
            End If
            
            Contract_CLCTN(Block(Code_Column)).Add Block
            
        Next X
    
        total_count = Contract_CLCTN.Count
    
        For C = Contract_CLCTN.Count To 1 Step -1   'Loop list of wanted Contract Codes
            
            T = T + 1
            On Error GoTo Contract_Code_Is_Missing
            
            With Contract_CLCTN(C) 'Retrieve contract data from the selected Collection
                
                ReDim Block(1 To .Count, 1 To UBound(Query, 2))
                
                For X = LBound(Block, 1) To UBound(Block, 1) 'Fill in array data
                    WS_Data = .Item(X)
                    For w = LBound(WS_Data) To UBound(WS_Data)
                        Block(X, w) = WS_Data(w)
                    Next w
                Next X
        
            End With
            
            contract_code = Block(1, Code_Column)
            
            Contract_CLCTN.Remove contract_code 'Removes 1 row arrays within the collection
            
            If HasKey(Symbol_Info, contract_code) Then
                Call Retrieve_Tuesdays_CLose(Block, price_column, Symbol_Info(contract_code)) 'Gets Price_Info
            End If
            
            Contract_CLCTN.Add Block, contract_code
            
Progress_Bar_Actions:
        
            If Not Progress_Control Is Nothing Then 'IF progress bar is active
            
                If T = total_count Then 'if on the last loop
                    Unload Progress_Bar
                ElseIf T Mod Percent_Mod = 0 Then 'else update length if condition is met
                    Increment_Progress Label:=Progress_Control, New_Width:=T * Bar_Increment, Loop_Percentage:=T / total_count
                End If
                
            End If
            
        Next C
        
        Query = Multi_Week_Addition(Contract_CLCTN, multiple_2d)
        Set Contract_CLCTN = Nothing
        
    End If

    On Error GoTo Database_Update_Error
    
    Call Update_DataBase(data_array:=Query, combined_wb_bool:=processing_combined_data, Report_Type:=Report_Type)
    
    Exit Sub

Progress_Checkbox_Missing:

    Set Progress_Control = Nothing
    Resume Block_Query_Main_Function
     
Contract_Code_Is_Missing:
    Resume Progress_Bar_Actions
    
Database_Update_Error:

    MsgBox "Unhandled Error in database update step."
    
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
Private Function Missing_Data(ByRef CFTC_Data As Variant, ByVal CFTC_Last_Updated_Day As Long, ByVal ICE_Last_Updated_Day As Long, ByRef ICE_Data As Variant, Report_Type As String, retrieve_combined_data As Boolean, Download_ICE_Data As Boolean, Download_CFTC_Data As Boolean, Optional DebugMD As Boolean = False) As Variant 'Should change to function; Block will find the amount of missed time and download appropriate files
'======================================================================================================
'Determines which files need to be downloaded for when multiple weeks of data have been missed
'======================================================================================================
Dim File_CLCTN As New Collection, Y As Long, MacB As Boolean, OBJ As Object, _
Hyperlink_RNG As Range, T As Long, new_data As New Collection

#If Mac Then
    MacB = True
#End If

If DebugMD Then
    If Not MacB Then If MsgBox("Do you want to test MAC OS data retrieval?", vbYesNo) = vbYes Then MacB = True
    If Download_CFTC_Data Then CFTC_Last_Updated_Day = DateAdd("yyyy", -5, CFTC_Last_Updated_Day)
    If Download_ICE_Data Then ICE_Last_Updated_Day = DateAdd("yyyy", -5, ICE_Last_Updated_Day)
End If

Application.DisplayAlerts = False

If Download_ICE_Data And Download_CFTC_Data And CFTC_Last_Updated_Day = ICE_Last_Updated_Day Then

        Retrieve_Historical_Workbooks _
            Path_CLCTN:=File_CLCTN, _
            ICE_Contracts:=True, _
            CFTC_Contracts:=True, _
            Mac_User:=MacB, _
            Report_Type:=Report_Type, _
            combined_data_version:=retrieve_combined_data, _
            ICE_Start_Date:=CDate(ICE_Last_Updated_Day), ICE_End_Date:=CDate(ICE_Data(1, date_column)), _
            CFTC_Start_Date:=CDate(CFTC_Last_Updated_Day), CFTC_End_Date:=CDate(CFTC_Data(1, date_column))
        
        If Not MacB Then new_data.Add Historical_Parse(File_CLCTN, Combined_Version:=retrieve_combined_data, Report_Type:=Report_Type, Yearly_C:=True, After_This_Date:=ICE_Last_Updated_Day, Kill_Previous_Workbook:=DebugMD)

Else

    If Download_ICE_Data Then
    
        Retrieve_Historical_Workbooks _
            Path_CLCTN:=File_CLCTN, _
            ICE_Contracts:=True, _
            CFTC_Contracts:=False, _
            Mac_User:=MacB, _
            Report_Type:=Report_Type, _
            combined_data_version:=retrieve_combined_data, _
            ICE_Start_Date:=CDate(ICE_Last_Updated_Day), _
            ICE_End_Date:=CDate(ICE_Data(1, date_column))
        
        If Not MacB Then new_data.Add Historical_Parse(File_CLCTN, Combined_Version:=retrieve_combined_data, Report_Type:=Report_Type, Yearly_C:=True, After_This_Date:=ICE_Last_Updated_Day, Kill_Previous_Workbook:=DebugMD)
    
    End If

    If Download_CFTC_Data Then
        
        Set File_CLCTN = Nothing
        
        Retrieve_Historical_Workbooks _
            Path_CLCTN:=File_CLCTN, _
            ICE_Contracts:=False, _
            CFTC_Contracts:=True, _
            Mac_User:=MacB, _
            CFTC_Start_Date:=CDate(CFTC_Last_Updated_Day), _
            CFTC_End_Date:=CDate(CFTC_Data(1, date_column)), _
            Report_Type:=Report_Type, _
            combined_data_version:=retrieve_combined_data
        
        If Not MacB Then new_data.Add Historical_Parse(File_CLCTN, Combined_Version:=retrieve_combined_data, Report_Type:=Report_Type, Yearly_C:=True, After_This_Date:=CFTC_Last_Updated_Day, Kill_Previous_Workbook:=DebugMD)
        
    End If

End If

If MacB Then            'Get user to download and open files
    
'    With MAC_SH.ListObjects("File_Paths_N").DataBodyRange
'
'        If WorksheetFunction.CountA(.Columns(2)) <> File_CLCTN.Count Then 'if there aren't the correct number of file paths
'
'            On Error Resume Next
'
'                With .Columns(1)
'
'                    If WorksheetFunction.CountA(.Cells) > 0 Then
'                        .ClearContents
'                        .ClearHyperlinks
'                    End If
'
'                End With
'
'            On Error GoTo 0
'
'            .ListObject.Resize .Offset(-1, 0).Resize(File_CLCTN.Count + 1, 2)
'
'            For T = 1 To File_CLCTN.Count 'Create hyperlinks
'
'                Set Hyperlink_RNG = .Cells(T, 1)
'
'                MAC_SH.Hyperlinks.Add Hyperlink_RNG, File_CLCTN(T)(0), TextToDisplay:=File_CLCTN(T)(1)
'
'            Next T
'
'            With .Worksheet
'                .Visible = xlSheetVisible
'                .Activate
'            End With
'
'            Re_Enable
'
'            Retrieval_Halted_For_User_Interaction = True 'this will become false again if the Sort_Letter sub is run again
'
'            Exit Function 'it is preferential that The Sorting algo is exited all together
'
'        Else
'
'            For T = 1 To File_CLCTN.Count
'                File_CLCTN.Remove (T)
'                File_CLCTN.Add .Cells(T, 2).value, , Before:=T
'            Next T
'
'            #If MAC_OFFICE_VERSION >= 15 Then
'
'                Dim fileAccessGranted As Boolean
'
'                fileAccessGranted = GrantAccessToMultipleFiles(Application.Transpose(.Columns(2).Value2))
'
'                If fileAccessGranted = False Then
'                    Retrieval_Halted_For_User_Interaction = True
'                    Exit Function
'                End If
'
'            #End If
'
'            .ClearContents
'
'            If MAC_SH.Hyperlinks.Count > 0 Then .Columns(1).ClearHyperlinks
'
'            MAC_SH.Visible = xlSheetVeryHidden
'
'        End If
'
'    End With
    
End If

Application.DisplayAlerts = True

If new_data.Count = 1 Then
    Missing_Data = new_data(1)
ElseIf new_data.Count > 1 Then
    Missing_Data = Multi_Week_Addition(new_data, multiple_2d)
End If

End Function
Private Function Weekly_ICE(LAst_Updated As Long, isMAC As Boolean, Report_Type As String, Retrieve_Combined As Boolean, Most_Recent_CFTC_Date As Date) As Variant

Dim Path_CLCTN As New Collection

Dim ICE_URL As String

If isMAC Then
    On Error GoTo Exit_Sub
    'Weekly_ICE = ICE_Query
    
Else

    On Error GoTo ICE_QueryT_Retrieval
    
    With Path_CLCTN
    
        .Add Environ("TEMP") & "\" & Date & "_Weekly_ICE.csv", "ICE"
        
        ICE_URL = Get_ICE_URL(Most_Recent_CFTC_Date)
        
        If Dir(.Item(1)) = vbNullString Then Call Get_File(ICE_URL, .Item("ICE"))
        
    End With

    Weekly_ICE = Historical_Parse(Path_CLCTN, Weekly_ICE_Data:=True, After_This_Date:=LAst_Updated, Report_Type:=Report_Type, Combined_Version:=Retrieve_Combined)

End If

Exit Function

ICE_QueryT_Retrieval:
    
    Err.Clear
    
    On Error GoTo Exit_Sub
    
    'Weekly_ICE = ICE_Query
    
    On Error GoTo 0
    
Exit_Sub:

End Function

#If Not Mac Then

    Private Function Get_ICE_URL(Query_Date As Date) As String
        Get_ICE_URL = "https://www.theice.com/publicdocs/cot_report/automated/COT_" & Format(Query_Date, "ddmmyyyy") & ".csv"
    End Function
    
#End If

Sub Progress_Bar_Custom_Initialize(ByRef Bar_Object As MSForms.Label, ByRef Increment_Per_Loop As Double, ByRef Number_of_Elements As Long, ByRef Mod_Loop_value As Long)

    With Progress_Bar
        .Show
        Set Bar_Object = .Bar                            'The colored Bar
        Increment_Per_Loop = .Frame.Width / Number_of_Elements  'How much to increment the bar EACH loop when conditions are met
        Mod_Loop_value = Number_of_Elements * 0.1           'Update when the loop# is a factor of this variable
        If Mod_Loop_value = 0 Then Mod_Loop_value = 1
    End With

End Sub








