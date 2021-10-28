Attribute VB_Name = "Data_Retrieval"
Public Valid_Table_Info() As Variant
Public Const TypeF As String = "L"
Public Retrieval_Halted_For_User_Interaction As Boolean

Public Enum Append_Type
    Add_To_Old = 1
    Multiple_1d = 2
    Multiple_2d = 3
End Enum

Public Enum Data_Identifier
    Block_Data = 14
    Old_Data = 44
    Weekly_Data = 33
End Enum

'Public Debug_Timer As Double

Option Explicit

Sub Sort_L(Optional Scheduled_Retrieval As Boolean = False, Optional Queried_Data As Variant)

'
'   Retrieves new data if available and applies it to worksheet
'   If arguements are given both must be present
'


'
'
'


Dim Commercial_NTC As Long, Last_Update_CFTC As Long, CC_Number As Long, _
Start_Time As Double, INTD_Timer As Double, Incoming_Data_Date As Long

Dim Debug_Mode As Boolean, Variable_Range As Range, Time_Record As String, _
Tables_Updated As Boolean, Mac_User As Boolean

Dim WS_Data() As Variant, Valid_ContractC() As Variant, _
DBM_Weekly_Retrieval As Boolean, DBM_Historical_Retrieval As Boolean, Price_Column As Long

Dim Last_Calculated_Column As Long ', All_Available_Contracts() As Variant

Const YearIxThree As Long = 156, MonthIxSix As Long = 26

Retrieval_Halted_For_User_Interaction = False

Start_Time = Timer

#If Mac Then
    Mac_User = True
#End If

With Application

    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    '.DisplayStatusBar = False
    .EnableEvents = False
    
    Valid_Table_Info = .Run("'" & ThisWorkbook.Name & "'!Get_Worksheet_Info")  'Update Valid contract list & store in array
    Valid_ContractC = .Index(Valid_Table_Info, 0, 1)                           'Contract Codes are in the first column of the above array
    
End With

Set Variable_Range = Variable_Sheet.ListObjects("Saved_Variables").DataBodyRange 'Table of saved variables

WS_Data = Variable_Range.Value2                                                 'Load variables to an array

With WorksheetFunction 'Use array to assign values variables within code

    Last_Update_CFTC = .VLookup("Last_Updated_CFTC", WS_Data, 2, False)     'The date the data was last sorted for
    
    Last_Calculated_Column = .VLookup("Last Calculated Column", WS_Data, 2, False) 'How many columns are reserved for automatic processes
      
End With

'All_Available_Contracts = HTTP_Weekly_Data(Last_Update_CFTC, False, True, False)

Erase WS_Data

If IsMissing(Queried_Data) Then 'If an array of input data wasn't supplied then run the below function( ie: when the macro is ran manually )
    
    On Error GoTo Deny_Debug_Mode
    
    If Weekly.Shapes("Test_Toggle").OLEFormat.Object.Value = 1 Then Debug_Mode = True 'Determine if Debug status
    
    If Debug_Mode = True Then
    
       If MsgBox("Test Weekly Data Retrieval ?", vbYesNo, "Choose what to debug") = vbYes Then
       
            DBM_Weekly_Retrieval = True
            
       ElseIf MsgBox("Test Multi-Week Historical Retrieval ?", vbYesNo, "Choose what to debug") = vbYes Then
       
            DBM_Historical_Retrieval = True

       End If
       
    End If
    
Retrieve_Latest_Data: On Error GoTo 0
    
    INTD_Timer = Timer 'Start Data Retrieval Timer
    
    Queried_Data = HTTP_Weekly_Data(Last_Update_CFTC, Auto_Retrieval:=Scheduled_Retrieval, Valid_Tables_Available:=True, DebugMD:=DBM_Weekly_Retrieval)

    Time_Record = "[" & Data_Retrieval.TypeF & "]-Data Retrieval: " & Round(Timer - INTD_Timer, 2) & " seconds. " & vbNewLine

End If

Incoming_Data_Date = Queried_Data(1, 1) 'Dates are in first column of array

CC_Number = UBound(Queried_Data, 2) 'Contract Codes are in the rightmost column of the un-altered array
Price_Column = UBound(Queried_Data, 2) + 1
Commercial_NTC = UBound(Queried_Data, 2) + 3 '[Net commercial column] Start of calculated columns
    
INTD_Timer = Timer 'Start timing how long it takes to sort & calculate data

If Incoming_Data_Date = Last_Update_CFTC And Debug_Mode = False Then
    
    GoTo No_New_Data 'Will prevent user from running this SUB if no new data is available
    
ElseIf Incoming_Data_Date - Last_Update_CFTC > 7 Or DBM_Historical_Retrieval Then 'Compares dates and if more than 13 days different then data will be downloaded
    
    Queried_Data = Missing_Data(Data:=Queried_Data, _
                                Last_Updated_Day:=Last_Update_CFTC, _
                                DebugMD:=DBM_Historical_Retrieval) 'Will download missing data and overwrite Array
    
    If Retrieval_Halted_For_User_Interaction Then Exit Sub
    
    ReDim Preserve Queried_Data(1 To UBound(Queried_Data, 1), 1 To Last_Calculated_Column)
    
    Call Block_Query(Query:=Queried_Data, DL_Reference:=Commercial_NTC, Codes:=Valid_ContractC, Time1:=YearIxThree, Time2:=MonthIxSix, Code_Column:=CC_Number, Price_Column:=Price_Column)
    
    Tables_Updated = True
    
ElseIf (Incoming_Data_Date - Last_Update_CFTC) > 0 Or Debug_Mode = True Then 'If just a 1 week difference

    ReDim Preserve Queried_Data(1 To UBound(Queried_Data, 1), 1 To Last_Calculated_Column)  'Expand for calculations
    
    Call Block_Query(Query:=Queried_Data, DL_Reference:=Commercial_NTC, Codes:=Valid_ContractC, Time1:=YearIxThree, Time2:=MonthIxSix, Code_Column:=CC_Number, Price_Column:=Price_Column)
    
    Tables_Updated = True
    
End If

Clean_Up: 'This section is ran if everything was successful.

    If Tables_Updated Then

        Time_Record = Time_Record & "[" & Data_Retrieval.TypeF & "]-Loop Time: " & Round(Timer - INTD_Timer, 2) & " seconds." & vbNewLine
        
        If Incoming_Data_Date > Last_Update_CFTC Then
        
            Update_Text Incoming_Data_Date   'Update Text Boxes "My_Date" on the HUB and Weekly worksheets.
            
            With Variable_Range.Columns(1)   'Update worksheet value for Last_Updated_CFTC
                .Cells(WorksheetFunction.Match("Last_Updated_CFTC", .Value2, 0), 2).Value2 = Incoming_Data_Date
            End With
            
        End If
        
        If Not Debug_Mode Then
            If Not Scheduled_Retrieval Then HUB.Activate 'If ran manually then bring the User to the HUB
            Courtesy                                     'Change Status Bar_Message
        End If
    
    End If
    
Erase_Objects: 'Ran regardless of successful parsing

    Re_Enable

    Erase Queried_Data
    Erase Valid_Table_Info
    Erase Valid_ContractC
    Set Variable_Range = Nothing
    
    Debug.Print Time_Record & "[" & Data_Retrieval.TypeF & "]-Time to completion: " & Round(Timer - Start_Time, 2) & " seconds. " & Now & vbNewLine
    
Exit Sub
    
No_New_Data:

    MsgBox "Data retrieval was successful." & _
    vbNewLine & _
    vbNewLine & _
        "The next release is scheduled for " & vbNewLine & vbTab & Format(CFTC_Release_Dates(False), "[$-x-sysdate]dddd, mmmm dd, yyyy") & " 3:30 PM Eastern time." & _
    vbNewLine & _
    vbNewLine & _
        "Enabling Test Mode will allow you to continue, but only the most recent data will be applied to the workbook if " & _
        "it's more recent than the data in the bottom row of each data table. " & _
    vbNewLine & _
    vbNewLine & _
        "Otherwise, try again after new data has been released. Check the release schedule for more information.", , Title:="New data is unavailable."

    Application.StatusBar = vbNullString
    
    GoTo Erase_Objects

Deny_Debug_Mode:
    
    Debug_Mode = False
    
    Resume Retrieve_Latest_Data
    
End Sub
Private Sub Block_Query(ByRef Query, _
            DL_Reference As Long, _
            Codes, _
            Time1 As Long, Time2 As Long, Code_Column As Long, _
            Price_Column As Long, Optional Overwrite_Worksheet As Boolean = False)

'
'Used to parse contract data for multiple or single week data
'Overwrite_Worksheet is for overwriting data that already exists on a worksheet
'

Dim X As Long, Contract_Block_Count As Long, C As Long, Row_Iterator As Long, w As Long, _
WS_Data() As Variant, Block() As Variant, Table_Range As Range, Contract_CLCTN As New Collection, Codes_In_Query() As Variant, Rows_Ordered_Old_To_New As Boolean

'=======================================================================================================================
'Progress Bar variables
Dim Progress_CHK As CheckBox, Bar_Increment As Double, Progress_Control As MSForms.Label, Percent_Mod As Long

Dim Current_Filters() As Variant, Error_Occurred As Boolean, Error_Count As Long

On Error GoTo Progress_Checkbox_Missing

Set Progress_CHK = Weekly.Shapes("Progress_CHKBX").OLEFormat.Object

If Progress_CHK.Value = 1 Then 'Show Progress Bar if Toggled on
    Call Progress_Bar_Custom_Initialize(Progress_Control, Bar_Increment, UBound(Codes, 1), Percent_Mod)
End If

Block_Query_Main_Function: On Error GoTo 0

Set Progress_CHK = Nothing
                      
With Application
    Codes_In_Query = .Transpose(.Index(Query, 0, Code_Column))
End With

On Error GoTo Error_Occurred_Toggle

For C = LBound(Codes, 1) To UBound(Codes, 1)                         'Loop list of wanted Contract Codes
    
    Contract_Block_Count = 0                         'Restart variable at 0

    For X = LBound(Codes_In_Query) To UBound(Codes_In_Query)             'Ensure filter worked properly.increment variable by 1 each time if match
        If LCase(Codes_In_Query(X)) = LCase(Codes(C, 1)) Then Contract_Block_Count = Contract_Block_Count + 1
    Next X
    
    If Contract_Block_Count > 0 Then 'Group all contracts that match the wanted Contract Code into a single array
           
        Row_Iterator = 1

        ReDim Block(1 To Contract_Block_Count, 1 To UBound(Query, 2))

        For X = LBound(Query, 1) To UBound(Query, 1)                     'Group contracts that match a certain contract code

            If Query(X, Code_Column) = Codes(C, 1) Then    'If current row matches the current wanted contract code

                For w = LBound(Block, 2) To UBound(Block, 2)            'Copy values to Block Array
                    Block(Row_Iterator, w) = Query(X, w)
                Next w
                
                If Row_Iterator = Contract_Block_Count Then Exit For 'Exit loop if all rows have been found
                
                Row_Iterator = Row_Iterator + 1

            End If

        Next X
                    
        Set Table_Range = Valid_Table_Info(WorksheetFunction.Match(Codes(C, 1), Codes, 0), 4).DataBodyRange
        'Set Reference to Target Table
        If Not Overwrite_Worksheet Then 'Combine New data with Old
        
            '---------------------------------------------------------
            Rows_Ordered_Old_To_New = Detect_Old_To_New(Table_Range.ListObject, 1)
            
            Call ChangeFilters(Table_Range.ListObject, Current_Filters)
            
            Call Sort_Table(Table_Range.ListObject, 1, xlAscending)
            
            WS_Data = Table_Range.Value2 'Old data already on worksheet
            '---------------------------------------------------------
            
            If UBound(WS_Data, 2) > UBound(Query, 2) Then ReDim Preserve WS_Data(1 To UBound(WS_Data, 1), 1 To UBound(Query, 2))
        
            With Contract_CLCTN
                .Add Array(Old_Data, WS_Data), "Old" 'Previous Worksheet Data[] |ORDER OF ADDITION IS IMPORTANT      <<<<<<<<<<<
                .Add Array(Block_Data, Block), "Block" 'Current row from Query_Range
            End With
            
            Block = Multi_Week_Addition(Contract_CLCTN, Append_Type.Add_To_Old) 'adds the contents of the NEW array TO the contents of the OLD
            
            Set Contract_CLCTN = Nothing
            
        End If
                
        Block = Multi_Calculations(Block, Contract_Block_Count, DL_Reference, Time1, Time2)
            
        Call Retrieve_Tuesdays_CLose(Block, Price_Column, Code_Column, Valid_Table_Info) 'Gets Price_Info

        If Not Overwrite_Worksheet Then
        
            Call Paste_To_Range(Data_Input:=Block, Sheet_Data:=WS_Data, Table_DataB_RNG:=Table_Range, Overwrite_Data:=Overwrite_Worksheet) 'Paste to bottom of table
            
            If Rows_Ordered_Old_To_New = False Then Call Sort_Table(Table_Range.ListObject, 1, xlDescending)
        
            Call RestoreFilters(Table_Range.ListObject, Current_Filters)
            
        Else
        
            Call Paste_To_Range(Data_Input:=Block, Table_DataB_RNG:=Table_Range, Overwrite_Data:=Overwrite_Worksheet)  'Paste to bottom of table
        
        End If
        
        Erase Block
        Erase WS_Data
        Set Table_Range = Nothing
        
    End If
    
Progress_Bar_Actions:

    If Not Progress_Control Is Nothing Then 'IF progress bar is active
    
        If C = UBound(Codes, 1) Then 'if on the last loop
            
            Unload Progress_Bar
            
        ElseIf C Mod Percent_Mod = 0 Then 'else update length if condition is met
            Increment_Progress Label:=Progress_Control, New_Width:=C * Bar_Increment, Loop_Percentage:=C / UBound(Codes, 1)
        End If
        
    End If
    
Next C

If Error_Occurred = True Then
    MsgBox "An error occurred while applying data to the worksheet during subroutinne Block_Query for " & Error_Count & " of " & UBound(Codes, 1) & " contracts."
End If

Exit Sub

Progress_Checkbox_Missing:

    Set Progress_Control = Nothing
    
    Resume Block_Query_Main_Function
    
Error_Occurred_Toggle:

    Error_Occurred = True
    
    Error_Count = Error_Count + 1
    Erase Block
    Erase WS_Data
    Set Table_Range = Nothing
    Set Contract_CLCTN = Nothing
    
    Resume Progress_Bar_Actions
    
End Sub
Sub New_Contract()
Attribute New_Contract.VB_Description = "Adds a contract to the worksheet. Requires you to know the CFTC contract code."
Attribute New_Contract.VB_ProcData.VB_Invoke_Func = " \n14"
    
'
'Prompt user for a CFTC contract code to add to the workbook if not present
'
Dim Contract_Code As String, C_Name As String, Calculated_Columns_Start As Long, LAst_Updated As Long, Contract_Data() As Variant, _
File_Collection As New Collection, Variable_Range As Range, New_Sheet As Worksheet, MacB As Boolean, _
Last_Calculated_Column As Long, Price_Column As Long, Contract_Code_CLMN As Long

#If Mac Then
    MsgBox "This macro is unavailable to your operating system."
    MacB = True
    Exit Sub
#End If

Set Variable_Range = Variable_Sheet.ListObjects("Saved_Variables").DataBodyRange

With Application

    .DisplayAlerts = False
    .EnableEvents = False
    .ScreenUpdating = False
    .DisplayStatusBar = True

    Valid_Table_Info = .Run("'" & ThisWorkbook.Name & "'!Get_Worksheet_Info") 'load data into array and apply to spreadsheet
    
    LAst_Updated = .VLookup("Last_Updated_CFTC", Variable_Range.Value2, 2, False) 'Last updated Date
    
    Price_Column = .CountIf(Variable_Sheet.ListObjects("User_Selected_Columns").DataBodyRange.Columns(2), True) + 1
    
End With

Retrieve_Historical_Workbooks _
            Path_CLCTN:=File_Collection, _
            ICE_Contracts:=False, _
            CFTC_Contracts:=True, _
            Mac_User:=MacB, _
            CFTC_Start_Date:=DateSerial(2017, 1, 1), _
            CFTC_End_Date:=CDate(LAst_Updated), _
            Historical_Archive_Download:=True
            
With Application
    
    .DisplayAlerts = True

    .StatusBar = "Download Complete. Input Contract Code"

End With

'Call Historical_TXT_Compilation(File_Collection, False)

'Stop

Do While True 'Will always be true

Input_Code:

    Contract_Code = Application.InputBox("Please supply the CFTC Contract Code of the contract you want")
        
    If Contract_Code = "False" Or Contract_Code = vbNullString Then
        
        If MsgBox("No input recieved. Would you like to try again?", vbYesNo) = vbYes Then
            GoTo Input_Code
        Else
            Call Re_Enable
            Application.StatusBar = vbNullString
            Exit Sub
        End If
        
    End If
            
    With Application
    
        If IsError(.Match(Contract_Code, .Index(Valid_Table_Info, 0, 1), 0)) = True Then
            'If the contract Code is not found within the array Valid_Table_Info
            Exit Do
        
        Else
            'C_Name = the worksheet name of where the Contract Code is found
            C_Name = Valid_Table_Info(.Match(Contract_Code, .Index(Valid_Table_Info, 0, 1), 0), 3)
            
            If MsgBox("Selected contract already exists within this workbook on worksheet : " & C_Name & vbNewLine & _
            "Would you like to try again with a different Contract Code?", vbYesNo, "Please choose") _
                    = vbNo Then
                        
                Call Re_Enable: Exit Sub
                
            End If

        End If
    
    End With
            
Loop

Application.StatusBar = "Creating intermediary workbook."

Calculated_Columns_Start = Price_Column + 2
Last_Calculated_Column = WorksheetFunction.VLookup("Last Calculated Column", Variable_Range.Value2, 2, False)

Contract_Data = Historical_Parse(File_Collection, Contract_Code:=Contract_Code, Specified_Contract:=True)

Contract_Code_CLMN = UBound(Contract_Data, 2)

Contract_Code = Contract_Data(1, UBound(Contract_Data, 2))

ReDim Preserve Contract_Data(LBound(Contract_Data, 1) To UBound(Contract_Data, 1), LBound(Contract_Data, 2) To Last_Calculated_Column)

Select Case TypeF

    Case "L"
        Contract_Data = Application.Run("'" & ThisWorkbook.Name & "'!Multi_Calculations", Contract_Data, UBound(Contract_Data, 1), Calculated_Columns_Start, 156, 26)
    Case "D"
        Contract_Data = Application.Run("'" & ThisWorkbook.Name & "'!Multi_Calculations", Contract_Data, UBound(Contract_Data, 1), Calculated_Columns_Start, 156, 26)
    Case "T"
        Contract_Data = Application.Run("'" & ThisWorkbook.Name & "'!Multi_Calculations", Contract_Data, UBound(Contract_Data, 1), Calculated_Columns_Start, 156, 26, 52)

End Select
    
Call Retrieve_Tuesdays_CLose(Contract_Data, Price_Column, Contract_Code_CLMN, Valid_Table_Info) 'Gets Price_Info
    
With ThisWorkbook

    Set New_Sheet = .Worksheets.Add                'Create a new Worksheet

    With New_Sheet                                 'Format certain columns
        
        .Columns(1).NumberFormat = "yyyy-mm-dd"
        .Columns(Contract_Code_CLMN).NumberFormat = "@"
        .Columns(Contract_Code_CLMN + 1).NumberFormat = "0.0000"
        
        Select Case TypeF
            Case "L"
               File_Collection.Add Array(9, 13, 14, 16, 17, 18, 19), "Percentage Format"
            Case "D"
                File_Collection.Add Array(10, 11, 12), "Percentage Format"
            Case "T"
                File_Collection.Add Array(5), "Percentage Format"
        End Select
        
        For LAst_Updated = 0 To UBound(File_Collection("Percentage Format"))
            .Columns(Calculated_Columns_Start + File_Collection("Percentage Format")(LAst_Updated)).NumberFormat = "0%"
        Next LAst_Updated
        
    End With

End With

With Symbols.ListObjects("Symbols_TBL")
    
    If IsError(Application.Match(Contract_Code, .DataBodyRange.Columns(1).Value, 0)) Then
        .ListRows.Add.Range.Value = Array(Contract_Code, New_Sheet.Name, Empty, Empty, Empty)
    Else
        Call Retrieve_Tuesdays_CLose(Contract_Data, Price_Column, Contract_Code_CLMN, Valid_Table_Info)
    End If
    
End With

Call Paste_To_Range(Target_Sheet:=New_Sheet, Sheet_Data:=Contract_Data, Historical_Paste:=True)

New_Sheet.ListObjects(1).Name = "CFTC_" & Replace(Contract_Code, "+", ".")

MsgBox "Don't forget to add a price symbol if available to the symbols worksheet for your chosen contract. Afterwords, run the recalculate worksheet macro to update price data.", , "Further instructions"

Set File_Collection = Nothing

Call Courtesy

Re_Enable
    
End Sub
Private Sub Update_Text(ByVal New_Date As Long)

'
'Updates Date shapes on the HUB and Weekly Retrieval worksheets
'

Dim AN1 As String, TWS() As Variant

AN1 = "[v] " & MonthName(Month(New_Date)) & " " & Day(New_Date) & ", " & Year(New_Date)

On Error Resume Next

TWS = Array(Weekly, HUB)

For New_Date = LBound(TWS) To UBound(TWS)

    TWS(New_Date).Shapes("My_Date").TextFrame.Characters.Text = AN1
        
Next New_Date

End Sub
Private Function Missing_Data(ByVal Data, ByVal Last_Updated_Day As Long, Optional DebugMD As Boolean = False) As Variant  'Should change to function; Block will find the amount of missed time and download appropriate files

'
'Downloads missing contract weeks if on Windows automatically
'If on MAC prompts user to download files and supply paths on a different worksheet
'
Dim File_CLCTN As New Collection, Y As Long, MacB As Boolean, OBJ As Object, Hyperlink_RNG As Range, T As Long

#If Mac Then
    MacB = True
#End If

If DebugMD Then

    If Not MacB Then If MsgBox("Do you want to test MAC OS data retrieval?", vbYesNo) = vbYes Then MacB = True
    
    Last_Updated_Day = DateAdd("yyyy", -5, Last_Updated_Day)

End If

Application.DisplayAlerts = False

Retrieve_Historical_Workbooks _
            Path_CLCTN:=File_CLCTN, _
            ICE_Contracts:=False, CFTC_Contracts:=True, _
            Mac_User:=MacB, _
            CFTC_Start_Date:=CDate(Last_Updated_Day), _
            CFTC_End_Date:=CDate(Data(1, 1))
    
If MacB Then            'Get user to download and open files
    
    With MAC_SH.ListObjects("File_Paths_N").DataBodyRange

        If WorksheetFunction.CountA(.Columns(2)) <> File_CLCTN.Count Then 'if there aren't the correct number of file paths
            
            On Error Resume Next
                
                With .Columns(1)
                
                    If WorksheetFunction.CountA(.Cells) > 0 Then
                        .ClearContents
                        .ClearHyperlinks
                    End If
                    
                End With
                
            On Error GoTo 0
            
            .ListObject.Resize .Offset(-1, 0).Resize(File_CLCTN.Count + 1, 2)
            
            For T = 1 To File_CLCTN.Count 'Create hyperlinks

                Set Hyperlink_RNG = .Cells(T, 1)

                MAC_SH.Hyperlinks.Add Hyperlink_RNG, File_CLCTN(T)(0), TextToDisplay:=File_CLCTN(T)(1)

            Next T
            
            With .Worksheet
                .Visible = xlSheetVisible
                .Activate
            End With
            
            Re_Enable

            Retrieval_Halted_For_User_Interaction = True 'this will become false again if the Sort_Letter sub is run again

            Exit Function 'it is preferential that The Sorting algo is exited all together

        Else
            
            For T = 1 To File_CLCTN.Count

                File_CLCTN.Remove (T)

                File_CLCTN.Add .Cells(T, 2).Value, , Before:=T
                
            Next T
            
            #If MAC_OFFICE_VERSION >= 15 Then
                
                Dim fileAccessGranted As Boolean
                
                fileAccessGranted = GrantAccessToMultipleFiles(Application.Transpose(.Columns(2).Value2))
                
                If fileAccessGranted = False Then
                    Retrieval_Halted_For_User_Interaction = True
                    Exit Function
                End If
                
            #End If
                
            .ClearContents

            If MAC_SH.Hyperlinks.Count > 0 Then .Columns(1).ClearHyperlinks

            MAC_SH.Visible = xlSheetVeryHidden

        End If

    End With
    
'    With Request_Files 'Userform Initalize_Event ran here
'
'        .Link_Collections File_CLCTN 'Add items/urls to listbox and resize form controls
'
'        Application.EnableEvents = True
'
'            .Show vbModeless
'
'        Application.EnableEvents = False
'
'        Do Until .Exit_Userform = True
'            DoEvents
'        Loop
'
'    End With
'
'    For Each OBJ In VBA.UserForms       'Closing Userfrom if it's still open [IF closed vis submission control]
'        If OBJ.Name = "Request_Files" Then
'            Unload Request_Files
'            Exit For
'        End If
'    Next OBJ
'
'    With ThisWorkbook
'        Set File_CLCTN = .Event_Storage("MAC-XL-WB")
'        .Event_Storage.Remove "MAC-XL-WB"
'    End With
'
'    If File_CLCTN.Count = 0 Then
'
'        MsgBox "There are 1 or more missing validly named Excel Workbook." & vbCrLf & vbCrLf & _
'               "Aborting Historical Data Retrieval."
'
'               Re_Enable
'
'               End
'
'    End If
     
End If

Application.DisplayAlerts = True

Missing_Data = Historical_Parse(File_CLCTN, Yearly_C:=True, After_This_Date:=Last_Updated_Day, Kill_Previous_Workbook:=DebugMD)

End Function
Private Function Multi_Calculations(AR1 As Variant, Weeks_Missed As Long, CommercialC As Long, _
Time1 As Long, Time2 As Long) As Variant

'
'Does calculations for certain fields to the right of raw data
'

Dim X As Long, Y As Long, n As Long, Start As Long, Finish As Long, INTE_B() As Variant, Z As Long

Start = UBound(AR1, 1) - (Weeks_Missed - 1)
Finish = UBound(AR1, 1)

    'Time1 is Year3,Time2 is Month6

On Error Resume Next

    For X = Start To Finish
        
        For Y = 0 To 2 'Commercial Net,Non-Commercial Net,Non-Reportable
        
            n = Array(7, 4, 11)(Y)
            
            AR1(X, CommercialC + Y) = AR1(X, n) - AR1(X, n + 1)
            
        Next Y
        
        AR1(X, CommercialC + 20) = AR1(X, 27) - AR1(X, 28) 'net %OI Commercial
        AR1(X, CommercialC + 21) = AR1(X, 24) - AR1(X, 25) 'net %OI Non-Commercial
        
        AR1(X, CommercialC + 9) = AR1(X, CommercialC) / (AR1(X, 3) - AR1(X, 6))     'Commercial/OI

        If AR1(X, 4) > 0 Or AR1(X, 5) > 0 Then

            AR1(X, CommercialC + 13) = AR1(X, 4) / (AR1(X, 4) + AR1(X, 5))      'NC Long%

            AR1(X, CommercialC + 14) = 1 - AR1(X, CommercialC + 13)             'NC Short%

        End If

        If X >= 2 Then

            AR1(X, CommercialC + 15) = AR1(X, CommercialC) - AR1(X - 1, CommercialC) 'Commercial Net Change

            For Y = 7 To 8  'Commercial Gross Long % Change & Commercial Gross Short % Change

                If AR1(X - 1, Y) > 0 Then
    
                    AR1(X, CommercialC + 9 + Y) = (AR1(X, Y) - AR1(X - 1, Y)) / AR1(X - 1, Y)
                    
                End If

            Next Y

        End If

        If AR1(X, 7) > 0 Or AR1(X, 8) > 0 Then

            AR1(X, CommercialC + 18) = AR1(X, 7) / (AR1(X, 7) + AR1(X, 8))                  'Commercial Long %

            AR1(X, CommercialC + 19) = 1 - AR1(X, CommercialC + 18)                         'Commercial Short %

        End If

    Next X

On Error GoTo 0

    If UBound(AR1, 1) > Time1 Then '3Y index
        
        For Y = 0 To 3
        
            INTE_B = Stochastic_Calculations(CommercialC + Array(0, 2, 9, 1)(Y), Time1, AR1, Weeks_Missed)
            
            n = Array(3, 5, 11, 4)(Y) + CommercialC         'used to calculate column number
            Z = 1 'finish-x
            For X = Start To Finish
                AR1(X, n) = INTE_B(Z)                       '[0]Commercial index 3Y  [1]Non-Reportable 3Y   < values of Y
                Z = Z + 1                                   '[2] Willco3Y            [3] Non-Commerical 3Y
            Next X
            
            Erase INTE_B
            
        Next Y

    End If

    If UBound(AR1, 1) > Time2 Then '6M index
        
        For Y = 0 To 3
            
            INTE_B = Stochastic_Calculations(CommercialC + Array(0, 2, 9, 1)(Y), Time2, AR1, Weeks_Missed)
            
            n = Array(6, 8, 10, 7)(Y) + CommercialC ' used to calculate column number
            Z = 1
            For X = Start To Finish
                AR1(X, n) = INTE_B(Z)   '[0]Commerical 6M [1]Non-Reportable 6M
                Z = Z + 1               '[2]WillCo6M      [3]Non Commercial 6M
            Next X
            
            Erase INTE_B
            
        Next Y

    End If
    
    n = CommercialC + 11 'Willco 3Y Column
    Y = n + 1            'movement index column

    For X = Start To Finish 'First Missed to most recent do Movement Index Calculations

        If X > Time1 + 6 Then

            AR1(X, Y) = AR1(X, n) - AR1(X - 6, n)
    
        End If

    Next X

    'The below code block is for adding only the missing data to the output array
    n = 1

    ReDim INTE_B(1 To Weeks_Missed, 1 To UBound(AR1, 2))

    For X = Start To Finish 'populate each row sequentially

        For Y = 1 To UBound(AR1, 2)
        
            INTE_B(n, Y) = AR1(X, Y)
            
        Next Y

        n = n + 1

    Next X
    
    Multi_Calculations = INTE_B

End Function
Private Sub Workbook_Data_Market_Conversion()


'
'Converts data between Futures Only and + Options or to redownload all data
'


Dim Historical_Data_Workbooks As New Collection, Combined_Markets As Boolean, _
Current_Data_Date As Date, Futures_Only_Data As Variant, Reference_Column As Long, _
ICE_Contracts_Needed As Boolean, URL As String, Query_Formula As String, Last_Calculated_Column As Long, Code_Column As Long

Dim New_File_Name As String, TempA As Variant, Combined_Markets_RNG As Range, _
Save_Collection As New Collection, User_Input As Long, Price_Column As Long

#If Mac Then
    Exit Sub
#End If

If ThisWorkbook.Saved = False Then
    
    If MsgBox("Workbook must be saved before preceding. Would you like to save now?", vbYesNo) = vbYes Then
        Application.Run "'" & ThisWorkbook.Name & "'!Custom_Save"
    Else
        Exit Sub
    End If
    
End If

With Application
    .EnableEvents = False
    .ScreenUpdating = False
End With

'HUB.Activate

With Variable_Sheet.ListObjects("Saved_Variables").DataBodyRange                      'Change Combined Workbook status
    
    Set Combined_Markets_RNG = .Cells(WorksheetFunction.Match("Combined Workbook", .Columns(1), 0), 2)
    
    On Error GoTo No_Selection
    
    User_Input = CLng(Application.InputBox("Enter 1 to overwrite with Futures and Options Combined data." & vbNewLine & vbNewLine & "2 for Futures Only."))
    
    On Error GoTo 0
    
    If User_Input = 1 Then
    
        Combined_Markets = True
        
    ElseIf User_Input = 2 Then
        
        Combined_Markets = False
    
    Else
        
        GoTo No_Selection
        
    End If
    
    Combined_Markets_RNG.Value2 = Combined_Markets  'Change Workbook Type
    
    Current_Data_Date = .Cells(WorksheetFunction.Match("Last_Updated_CFTC", .Columns(1), 0), 2).Value 'Date for most recently released data
    
    Last_Calculated_Column = WorksheetFunction.VLookup("Last Calculated Column", .Value2, 2, False)
    
    Price_Column = WorksheetFunction.CountIf(Variable_Sheet.ListObjects("User_Selected_Columns").DataBodyRange.Columns(2), True) + 1
    
    Reference_Column = Price_Column + 2
    
End With

Select Case Data_Retrieval.TypeF

    Case "L":
    
        If Combined_Markets Then
            URL = "https://www.cftc.gov/dea/newcot/deacom.txt"
        Else
            URL = "https://www.cftc.gov/dea/newcot/deafut.txt"
        End If
        
    Case "D":
        
        ICE_Contracts_Needed = True
        
        If Combined_Markets Then
            URL = "https://www.cftc.gov/dea/newcot/c_disagg.txt"
        Else
            URL = "https://www.cftc.gov/dea/newcot/f_disagg.txt"
        End If
        
    Case "T":
    
        'Reference_Column = 58
        
        If Combined_Markets Then
            URL = "https://www.cftc.gov/dea/newcot/FinComWk.txt"
        Else
            URL = "https://www.cftc.gov/dea/newcot/FinFutWk.txt"
        End If
        
End Select

Data_Retrieval.Valid_Table_Info = Application.Run("'" & ThisWorkbook.Name & "'!Get_Worksheet_Info")

Call Retrieve_Historical_Workbooks(Historical_Data_Workbooks, ICE_Contracts_Needed, CFTC_Contracts:=True, Mac_User:=False, CFTC_Start_Date:=DateSerial(2017, 1, 1), CFTC_End_Date:=Current_Data_Date, ICE_Start_Date:=DateSerial(2011, 1, 1), ICE_End_Date:=Current_Data_Date, Historical_Archive_Download:=True)

Futures_Only_Data = Historical_Parse( _
        File_CLCTN:=Historical_Data_Workbooks, _
        Contract_Code:=vbNullString, _
        After_This_Date:=0, _
        Kill_Previous_Workbook:=False, _
        Yearly_C:=False, _
        Specified_Contract:=False, _
        Weekly_ICE_Data:=False, _
        CFTC_TXT:=False, _
        Parse_All_Data:=True) 'Compile and filter historical data

Code_Column = UBound(Futures_Only_Data, 2)

ReDim Preserve Futures_Only_Data(1 To UBound(Futures_Only_Data, 1), 1 To Last_Calculated_Column)

Select Case Data_Retrieval.TypeF

    Case "L"
    
        Call Block_Query(Futures_Only_Data, Reference_Column, Application.Index(Valid_Table_Info, 0, 1), 156, 26, Code_Column, Overwrite_Worksheet:=True, Price_Column:=Price_Column)
    
    Case "D"
         
        'Call Block_Query(Futures_Only_Data, Reference_Column, Application.Index(Valid_Table_Info, 0, 1), 156, 26,  Code_Column,  True, Price_Column)

    Case "T"
    
'        Call Block_Query(Query:=Futures_Only_Data, _
'        DL_Reference:=Reference_Column, _
'        Codes:=Application.Index(Valid_Table_Info, 0, 1), _
'        Time1:=156, Time2:=26, Time3:=52, _
'        Code_Column:=Code_Column, _
'        Overwrite_Worksheet:=True, _
'        Price_Column:=Price_Column)

End Select

'Create Array Subsections and place on worksheets

'New_File_Name = Application.GetSaveAsFilename

With ThisWorkbook
    
    TempA = Split(.Queries("Weekly").Formula, Chr(34), 3) 'SPlit with quotation mark
    
    TempA(1) = URL
    
    .Queries("Weekly").Formula = Join(TempA, Chr(34))
    
    Application.Run "'" & .Name & "'!Before_Save", False 'True will allow AfterSave event
'
'    Set Save_Collection = ThisWorkbook.Event_Storage
'
'    On Error Resume Next
'
'    If Dir(New_File_Name) <> vbNullString And New_File_Name <> ThisWorkbook.FullName Then Kill New_File_Name
'
'    If Err.Number <> 0 Then Err.Clear
'
'    .SaveAs New_File_Name, xlExcel12, ConflictResolution:=xlUserResolution 'Save workbook to new location
'
'    If Err.Number = 0 Then
'        MsgBox "Succesfully saved"
'        On Error GoTo 0
'
'        Set ThisWorkbook.Event_Storage = Save_Collection
'
'        Application.Run "'" & .Name & "'!After_Save"
'
'    Else
'
'        Err.Clear
'
'        .Save
'
'        If Err.Number = 0 Then
'
'            MsgBox "Succesfully saved"
'            On Error GoTo 0
'            Set ThisWorkbook.Event_Storage = Save_Collection
'            Application.Run "'" & .Name & "'!After_Save"
'
'        End If
'
'    End If
    
End With

Exit_Stuff:

Re_Enable

MsgBox "Finished Conversion."

Exit Sub

No_Selection:

        MsgBox "No selection made. Exiting procedures."
        Re_Enable
'Application.Run "'" & ThisWorkbook.Name & "'!After_Save" 'True will allow AfterSave event

End Sub

Sub Progress_Bar_Custom_Initialize(ByRef Bar_Object As MSForms.Label, ByRef Increment_Per_Loop As Double, ByRef Number_of_Elements As Long, ByRef Mod_Loop_value As Long)

    With Progress_Bar
        .Show
        Set Bar_Object = .Bar                            'The colored Bar
        Increment_Per_Loop = .Frame.Width / Number_of_Elements  'How much to increment the bar EACH loop when conditions are met
        Mod_Loop_value = Number_of_Elements * 0.1           'Update when the loop# is a factor of this variable
        If Mod_Loop_value = 0 Then Mod_Loop_value = 1
    End With

End Sub
