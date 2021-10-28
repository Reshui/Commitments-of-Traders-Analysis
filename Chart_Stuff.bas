Attribute VB_Name = "Chart_Stuff"
Public Current_Table_Source As ListObject
Option Explicit


'Sub Delete_Bottom_Row_of_Data_Tables()
'
'Dim KK() As Variant, J As Long, Current_Filters() As Variant, TB As ListObject
'
'KK = Get_Worksheet_Info
'
'For J = 1 To 2
'
'    Set TB = KK(J, 4)
'    Call ChangeFilters(TB, Current_Filters)
'
'    KK(J, 4).ListRows(TB.ListRows.Count).Delete
'
'    Call RestoreFilters(TB, Current_Filters)
'
'Next J
'
'End Sub

Private Function Get_Worksheet_Info() As Variant 'retrieve destination info

Dim i As Long, r As Long, This_C As New Collection, Contract_WS_Name As String, Keys_A() As String, _
TWOA() As Variant, TB As ListObject, SH As Worksheet, EVNT As Boolean, Code As String, Item As Variant, WSN As String, Contract_Name As String

Dim SymbolA() As Variant, Current_Symbol As String, Exceptions() As Variant ', Contract_Code_Column As Long

'Dim Start_Time As Double: Start_Time = Timer

Const Locator_STR As String = "CFTC_Contract_Market_Code"

SymbolA = Symbols.ListObjects("Symbols_TBL").DataBodyRange.Value

With Application
     EVNT = .EnableEvents
            .EnableEvents = False
End With

'Contract_Code_Column = Application.CountIf(Variable_Sheet.ListObjects("User_Selected_Columns").DataBodyRange.Columns(2), True)

Exceptions = Array("B", "Cocoa", "G", "RC", "Wheat", "W")

On Error GoTo Invalid_Worksheet_Table

For Each SH In ThisWorkbook.Worksheets
    
    Select Case True
    
        Case SH Is Weekly, SH Is HUB, SH Is Variable_Sheet, SH Is Chart_Sheet, SH Is QueryT, _
             SH Is Symbols, SH Is MAC_SH, SH Is Dashboard_V1
        
        Case Else

            Set TB = CFTC_Table(ThisWorkbook, SH, Code) 'Finds a table on the specified sheet with Locator_STR in the header row
    
            If Not TB Is Nothing Then 'TB is found based on if Locator_STR is in the the header row
                
                'SH.Columns(66).NumberFormat = "0.0000"
                
                If Len(Code) = 6 Or Not IsError(Application.Match(Code, Exceptions, 0)) Then
    
                    On Error Resume Next
                        
                        i = Application.Match(Code, Application.Index(SymbolA, 0, 1), 0) 'Row number of the queried contract code in the symbols array
                        
                        If Err.Number = 0 Then 'Decide which symbol to use based on availability
                        
                            If Not IsEmpty(SymbolA(i, 3)) Then 'Yahoo Finance
                                
                                Current_Symbol = SymbolA(i, 3)
                                
                            ElseIf Not IsEmpty(SymbolA(i, 4)) Then 'Stooq
                            
                                Current_Symbol = SymbolA(i, 4)
                            Else
                            
                                Current_Symbol = vbNullString
                                
                            End If
                        
                        Else
                            
                            Err.Clear
                            
                        End If
                        
                    On Error GoTo Invalid_Worksheet_Table
                    
                    With SH
                    
                        Contract_WS_Name = .Name '....LCASE used for A-Z sorting of worksheet names
                        
                        'Add information array to collection
                        
                        This_C.Add Array(Code, .Index, Contract_WS_Name, TB, Current_Symbol), LCase(Contract_WS_Name)
                        
                        
                        'add worksheet name to string for sorting purposees later
                        WSN = IIf(WSN = vbNullString, LCase(Contract_WS_Name), WSN & "," & LCase(Contract_WS_Name))
    
                        Contract_WS_Name = vbNullString
                        Code = vbNullString
                        Current_Symbol = vbNullString
                    
                    End With
                    
                End If
                
                Set TB = Nothing
                
            End If
            
    End Select
    
Invalid_Worksheet_Table: If Err.Number <> 0 Then On Error GoTo -1
    
Next SH

Set SH = Nothing

On Error GoTo No_Data_Available

With This_C
    
    Keys_A = Split(WSN, ",")

    Call Quicksort(Keys_A, LBound(Keys_A), UBound(Keys_A))

    ReDim TWOA(1 To UBound(Keys_A) + 1, 1 To UBound(.Item(1)) + 1) 'store all arrays in a single array
    
    For i = 1 To UBound(TWOA)                   'Loop Collection items
    
        Item = .Item(Keys_A(i - 1))             'Reference array stored within Collection
        
        For r = LBound(TWOA, 2) To UBound(TWOA, 2)
            
            If r <> 4 Then        'If not supplying table to the Array[Columns 1-3,5]
                
                TWOA(i, r) = Item(r - 1)
            
            ElseIf r = 4 Then
                
                Set TWOA(i, r) = Item(r - 1) '[Column 4] is a table object
            
            End If
            
        Next r
    
    Next i
    
End With

Erase Keys_A:  Erase Item:  Set This_C = Nothing

Set TB = Variable_Sheet.ListObjects("Table_WSN")

With TB 'place array on worksheet

    With .DataBodyRange
    
        i = .Rows.Count
        
        If i > UBound(TWOA, 1) Then .ClearContents 'clear contents if there will be too many rows in the table

        .Resize(UBound(TWOA, 1), 3).Value2 = TWOA 'write array to worksheet [-1 since the 4th column contains table objects that shouldn't be written to the worksheet

    End With

    If i > UBound(TWOA, 1) Then .Resize .Range.Resize(UBound(TWOA, 1) + 1, 3) 'Resize to fit current region if table has too many rows
        
End With

'Set TB = Symbols.ListObjects("Symbols_TBL")
'
'With TB
'
'    With .DataBodyRange
'
'        I = 0
'
'        .ClearContents
'
'        For r = 1 To UBound(TWOA, 2)
'
'            Select Case r
'
'                Case 1, 3, 5
'
'                   I = I + 1
'
'                   .Columns(I).Resize(UBound(TWOA, 1), 1).Value = Application.Index(TWOA, 0, r)
'
'            End Select
'
'        Next r
'
'    End With
'
'End With

Get_Worksheet_Info = TWOA
Application.EnableEvents = EVNT

'Debug.Print Timer - Start_Time

Exit Function

No_Data_Available:

    MsgBox "Worksheet identifiers couldn't be loaded during the Get_Worksheet_Info function." & vbNewLine & vbNewLine & _
           "Please email MoshiM_UC@outlook.com or submit a bug report." & vbNewLine & vbNewLine & _
           "Further code execution will be halted."
           
    Re_Enable
    
    End
    
End Function
Public Sub Update_List(Optional All_Sheets_Combobox As ComboBox, Optional Charts_ComboBox As ComboBox, Optional Exclude_HUB_CB As Boolean = False)

Dim Worksheet_Names As Variant, All_Sheets As New Dictionary, SH As Worksheet, _
WSNA() As Variant, WSN As String, Z As Long, Target_is_Navigation As Boolean, UserForm_OB As Object

With Application 'load worksheet names that have CFTC codes on them to an array
    Worksheet_Names = .Index(.Run("'" & ThisWorkbook.Name & "'!Get_Worksheet_Info"), 0, 3)
End With
    
#If Mac Then

#Else 'Populate Charts Combobox

    If All_Sheets_Combobox Is Nothing Then Set All_Sheets_Combobox = HUB.OLEObjects("Sheet_Selection").Object
    
    ' HUB.Shapes("Worksheet_Selector").OLEFormat.Object ' 'Default this combobox to the one on the HUB
    'HUB.Shapes("Worksheet_Selector").controlformat.list=
    
    If Charts_ComboBox Is Nothing And Exclude_HUB_CB = False Then 'optional control boolean
    
        Set Charts_ComboBox = Chart_Sheet.OLEObjects("Sheet_Selection").Object
    
        With Charts_ComboBox
            
            .List = Worksheet_Names: Erase Worksheet_Names 'Use the sorted array of valid Worksheets to populte the list.
                
            .AddItem "{Select a Worksheet}" 'Add this to the bottom.
            
        End With
    
    End If

#End If

With All_Sheets '<<---Dictionary Object--Populate Combobox that directs the user towards all other Worksheets

    For Each SH In ThisWorkbook.Worksheets 'loop through worksheets and store worksheet names in a dictionary
       
        WSN = SH.Name
        
        Select Case True
        
            Case SH Is QueryT, SH Is MAC_SH
            
            Case SH Is Variable_Sheet
                
                If UUID Then .Add LCase(WSN), WSN 'only add the variable sheet to the list if on my computer
            
            Case Else
            
                 .Add LCase(WSN), WSN 'Lower case version is used as a key for sorting purposes
            
        End Select
       
    Next SH

    WSNA = .Keys 'Used for sorting
            
    Quicksort WSNA, LBound(WSNA), UBound(WSNA) 'Sort A-Z
                                         
    For Z = LBound(WSNA) To UBound(WSNA) 'The quicksort function has trouble with capital letters so this method is used
        WSNA(Z) = .Item(WSNA(Z))         'Overwrite array elements with the corresponding dictionary item based on key
    Next Z
    
End With
        
With All_Sheets_Combobox

    .List = WSNA
     'check if navigation is open  and if it is then check if navigation is the target
    For Each UserForm_OB In VBA.UserForms
        
        If UserForm_OB.Name = "Navigation" Then
            
            If All_Sheets_Combobox Is Navigation.S_Selection Then
                Target_is_Navigation = True
                Exit For
            End If
            
        End If
    
    Next UserForm_OB
        
    If Not Target_is_Navigation Then .AddItem "-----------------------------------------"
    
End With

Set All_Sheets = Nothing
Set All_Sheets_Combobox = Nothing
Set Charts_ComboBox = Nothing

End Sub
Private Sub Update_Charts(Optional Worksheet_Name As String) 'changing Chart Data

Dim TT As Long, AR As Range, HAT() As Variant, Date_Range As Range, _
Chart_Obj As ChartObject, Chart_Series As Series, _
Error_STR As String, Chart_Source As Worksheet, Array_Method As Boolean, _
Formula_AR() As String, Dates() As Variant, _
Area_Compilation As New Collection, Series_CLCTN As SeriesCollection, _
Series_A() As Variant, WBN As String

Dim Series_Info As New Collection, Use_Stored_Series_ID As Boolean, New_Series_Added As Boolean

Dim Use_User_Dates As Boolean, Minimum_Date As Date, Maximum_Date As Date, C1 As String, C2 As String, Use_Dashboard_V1_Dates As Boolean

Select Case Worksheet_Name
    Case Weekly.Name, Variable_Sheet.Name, HUB.Name, Chart_Sheet.Name, QueryT.Name, MAC_SH.Name, Dashboard_V1.Name, Symbols.Name
    End
End Select
'DD = Timer
#If Not Mac Then
    If Worksheet_Name = vbNullString Then Worksheet_Name = Chart_Sheet.OLEObjects("Sheet_Selection").Object.Value 'Get name from combobox
#End If

On Error Resume Next

With Chart_Sheet.ListObjects("Chart_Settings_TBL").DataBodyRange

    If .Cells(2, 2) = True Then 'If user wants text dates

        Array_Method = True

        If Err.Number <> 0 Then 'Default to using Range Methods
            Err.Clear
            Array_Method = False
        End If

    End If

    If .Cells(5, 2) = True Then 'Use date range starting at Dashboard V1 lookback period

        Use_Dashboard_V1_Dates = True

    ElseIf .Cells(1, 2) = False Then 'If the user wants to use their own dates rather than worksheet dates

        If Not IsEmpty(.Range(.Cells(3, 2), .Cells(4, 2))) Then 'if at least one date

            Minimum_Date = .Cells(3, 2)
            Maximum_Date = .Cells(4, 2)

            If CDbl(Minimum_Date) <> 0 Then C1 = ">="
            If CDbl(Maximum_Date) <> 0 Then C2 = "<="

            If (Maximum_Date < Minimum_Date) And CDbl(Maximum_Date) <> 0 Then

                MsgBox "Maximum Date cannont be less than Minimum Date. Defaulting to worksheet filters."

            Else

                Use_User_Dates = True

            End If

        End If

    End If

End With

On Error GoTo Worksheet_Not_Exists

    Set Chart_Source = ThisWorkbook.Worksheets(Worksheet_Name)

On Error GoTo 0

Set Current_Table_Source = CFTC_Table(ThisWorkbook, Chart_Source) 'List Object is returned

If Current_Table_Source Is Nothing Then GoTo No_Table

With Current_Table_Source 'Object is a valid contract table so retrieve needed info
    
    On Error GoTo Show_All_Data
    
    Set AR = .DataBodyRange.SpecialCells(xlCellTypeVisible) 'This is just to test if data is available via error checking
    
    On Error GoTo Load_Data_Error
         
    Set AR = .DataBodyRange 'Load Table Range to variable
    
    On Error GoTo 0
    
    If Use_Dashboard_V1_Dates Then 'If the user wants to use the dae range from the V1 dashboard
        C1 = ">="                  'Condition 1 set to greater than or equal to
        TT = AR.Rows.Count - Dashboard_V1.Cells(1, 2).Value + 1 'Number of data rows - Dashboard N weeks value... +1 is so that >= can apply
        If TT <= 0 Then TT = 1      'Ensures condition isn't outside the range of the table
        Minimum_Date = AR.Cells(TT, 1).Value
    End If

    If Use_User_Dates Or Use_Dashboard_V1_Dates Then

        If Len(C1) > 0 And Len(C2) > 0 Then 'If both a maximum and minimum date have been supplied

            Current_Table_Source.Range.AutoFilter _
                Field:=1, _
                Criteria1:=C1 & Minimum_Date, Operator:=xlAnd, Criteria2:=C2 & Maximum_Date

        ElseIf Len(C1) > 0 Then 'If only a minimum has been supplied

            Current_Table_Source.Range.AutoFilter _
                Field:=1, _
                Criteria1:=C1 & Minimum_Date, Operator:=xlFilterValues

        ElseIf Len(C2) > 0 Then 'If only a maximum has been supplied

            Current_Table_Source.Range.AutoFilter _
                Field:=1, _
                Criteria1:=C2 & Maximum_Date, Operator:=xlFilterValues
        Else
            .AutoFilter.ShowAllData
        End If

    End If

    Set Date_Range = AR.Columns(1)                  'Column 1 of table should hold dates

    HAT = .HeaderRowRange.Value2                    'Load headers from table to array

End With

With Date_Range.SpecialCells(xlCellTypeVisible) 'This WITH block only affects the array Dates used when Array_Dates=true and for the Experimental Indicator

    If .Areas.Count = 1 Then                        'if only one area then take directly from sheet

        Dates = WorksheetFunction.Transpose(.Value2) 'Create 1D list of dates

    Else

       For TT = 1 To .Areas.Count                   'Loop each Area and add them to the collection and then combine
           Area_Compilation.Add .Areas(TT).Value2
       Next TT

       Dates = WorksheetFunction.Transpose(Application.Run(WBN & "Multi_Week_Addition", Area_Compilation, Append_Type.Multiple_2d)) 'join areas together and transpose to 1D

       Set Area_Compilation = Nothing

    End If

End With

If Array_Method = True Then 'String dates will be used instead of generating from the range data

   For TT = LBound(Dates) To UBound(Dates)
       Dates(TT) = Format(CDate(Dates(TT)), "yyyy-mm-dd") 'Convert number dates to DATE typed variable
   Next TT

End If

C1 = vbNullString 'variable will now be used to hold Chart columns when needed
C2 = vbNullString

With Series_Info
    .Add New Collection, "Old" 'Will hold default values for previously used series
    .Add New Collection, "New" 'Will hold series info for newly added or adjusted series
End With

Series_A = Chart_Sheet.ListObjects("Series_Ok").DataBodyRange.Value 'Array that holds default values for each series

On Error Resume Next 'Used in case a series is adjusted

With Series_Info("Old")

    For TT = 1 To UBound(Series_A, 1)
    
        C1 = Series_A(TT, 1) & "_" & Series_A(TT, 2) 'Key is a composite of the chart name and the series name
        
        .Add Array(Series_A(TT, 1), Series_A(TT, 2), Series_A(TT, 3), C1), C1 'add an array of series identifiers to the collection
        
        If Err.Number <> 0 Then
            .Remove C1
            .Add Array(Series_A(TT, 1), Series_A(TT, 2), Series_A(TT, 3), C1), C1
            Err.Clear
        End If
        
    Next TT
    
    C1 = vbNullString
    Erase Series_A
    
End With

On Error GoTo 0

For Each Chart_Obj In Chart_Sheet.ChartObjects 'For each chart on the Charts Worksheet

    Use_Stored_Series_ID = False

    With Chart_Obj

        If Not .Name = "NET-OI-INDC" And Not Chart_Obj.Chart.ChartType = xlHistogram Then

            On Error Resume Next

            Set Series_CLCTN = .Chart.SeriesCollection 'Used in stead of fullseriescollection for backwards compatability
            
            If Series_CLCTN.Count = 0 Or Err.Number <> 0 Then 'If data for the chart can't be found
                
                On Error GoTo 0
                
                Use_Stored_Series_ID = True
                
                For TT = 1 To Series_Info("Old").Count 'Loop collection and add data to charts that match the current chart name
                    
                    If Series_Info("Old")(TT)(0) = .Name Then 'If equal to the current charts name
                    
                        Series_CLCTN.Add AR.Columns(Series_Info("Old")(TT)(2)) 'Add the recorded column to the chart
                        
                        Series_CLCTN(Series_CLCTN.Count).Name = Series_Info("Old")(TT)(1) 'Change the name of the added series
                        
                    End If
                    
                Next TT
                
            End If
            
            On Error GoTo 0
            
            If Use_Stored_Series_ID = False Then
            
                For Each Chart_Series In Series_CLCTN 'For each series on that chart
    
                    On Error GoTo Chart_Data_NFound
    
                    With Chart_Series
    
                        If .Name <> "USD Index Net Non-Commercial OI %" Then
    
                            Formula_AR = Split(.Formula, ":")
                            TT = Range(Split(Formula_AR(UBound(Formula_AR)), ",")(0)).Column 'Column number for the series used on the chart

                            .Values = AR.Columns(TT) 'Switch data to correspond with the same location but different table
    
                            If TT <= UBound(HAT, 2) And .Name <> HAT(1, TT) Then .Name = HAT(1, TT)  'Rename the Series if needed
    
                        End If
                        
                        .XValues = IIf(Array_Method = True, Dates, Date_Range) 'Dates: 1d Array |  Date_Range:Auto Updating Range
                        
                    End With
                    
                    C1 = .Name & "_" & Chart_Series.Name 'Create key
                    
                    If HasKey(Series_Info("Old"), C1) Then
                        If Series_Info("Old")(C1)(2) <> TT Then New_Series_Added = True 'If the column source has changed then a new series has been added
                    Else
                        New_Series_Added = True
                    End If
                    
                    On Error Resume Next
                    
                    If New_Series_Added Then
                        New_Series_Added = False
                        Series_Info("New").Add Array(.Name, Chart_Series.Name, TT, C1), C1
                    End If
                    
                    On Error GoTo 0
                    
                    C1 = vbNullString
Skip_Series:
                Next Chart_Series
            
            End If
            
            .Chart.Axes(xlCategory).TickLabels.NumberFormat = "yyyy-mm-dd"

            If .Name = "Price Chart" Then 'Adjust minimum valus to fit price range
                .Chart.Axes(xlValue).MinimumScale = Application.Min(AR.Columns(TT).SpecialCells(xlCellTypeVisible))
            End If
            
        ElseIf Chart_Obj.Chart.ChartType = xlHistogram Then

            On Error GoTo 0

            Select Case Chart_Obj.Name 'This is done by chart name since you cant query the formula or source range of the chart

                Case "Open Interest Histogram"

                    TT = 3 'OI
                    
                    Set Chart_Series = Chart_Obj.Chart.SeriesCollection(1)

                    On Error GoTo Next_Chart

                    Call Open_Interest_Histogram(Chart_Obj, TT, AR, Chart_Series)

                Case Else

'                   C1 = Split(Chart_OBJ.Chart.ChartTitle.Text, "-")(0)
'
'                   If Len(C1) > 0 And IsNumeric(C1) Then
'                       TT = CLng(C1)
'                   Else
                    GoTo Next_Chart
'                   End If

            End Select

        ElseIf .Name = "NET-OI-INDC" Then

            On Error GoTo Skip_ScatterC

            ScatterC_OI Worksheet_Name, Chart_Dates:=Dates
            
Skip_ScatterC: On Error GoTo -1
            
        End If

    End With

    With Chart_Obj.Chart.ChartTitle 'Adjust the title of he chart to shhow the dates plotted
        .Text = Split(.Text, vbTab)(0) & vbTab & "[" & CDate(WorksheetFunction.Min(Dates)) & " to " & CDate(WorksheetFunction.Max(Dates)) & "]"
    End With

Next_Chart: On Error GoTo -1

Next Chart_Obj

Set Chart_Obj = Nothing

Set AR = Nothing

Erase HAT
Erase Dates

If Series_Info("New").Count > 0 Then

    With Chart_Sheet.ListObjects("Series_Ok").DataBodyRange

        Series_A = Multi_Week_Addition(Series_Info("New"), Append_Type.Multiple_1d)

        .Cells(.Rows.Count, 1).Resize(UBound(Series_A, 1), UBound(Series_A, 2) - 1).Value = Series_A

    End With

End If

Set Series_Info = Nothing

If Error_STR <> vbNullString Then MsgBox Error_STR

Exit Sub

No_Table:

    MsgBox "A valid Table was not found on sheet " & Worksheet_Name & vbNewLine & _
    vbNewLine & "Please note that valid tables are those with a column labelled: CFTC_Contract_Market_Code"
    Exit Sub

Chart_Data_NFound:

    Error_STR = Error_STR & "No data for:" & vbNewLine & vbNewLine & _
                                Worksheet_Name & " : " & Chart_Obj.Chart.ChartTitle.Text & " __ " & Chart_Series.Name _
                                & vbNewLine

    Resume Skip_Series

Worksheet_Not_Exists:

    MsgBox "Worksheet [ " & Worksheet_Name & " ] doesn't exist."
    
    Exit Sub
    
Show_All_Data:

    Current_Table_Source.AutoFilter.ShowAllData
    
    Resume Next ' Sends program to line that loads table data into a range variable
    
Load_Data_Error:

    MsgBox ("Data could not be charted for " & Worksheet_Name)
    Exit Sub
    
End Sub
Public Sub ScatterC_OI(Worksheet_N As String, ByVal Chart_Dates As Variant)

Dim BS_Count As Long, Previous_Net As Long, Data_A() As Variant, T As Long, Z As Long, OI_Change As Long, _
Current_Net As Long, Buy_Sell_Array() As Variant, X As Long, _
INDC_Chart_Series As FullSeriesCollection, BuyN As Long, SellN As Long, Date_LNG() As Long

Const OI_Change_Column As Long = 13

ReDim Buy_Sell_Array(1 To 2, 1 To UBound(Chart_Dates))

'Chart Dates is likely an array of long Dates


'[1]-Buy
'[2]-Sell

Data_A = CFTC_Table(ThisWorkbook, ThisWorkbook.Worksheets(Worksheet_N)).DataBodyRange.Value2 'retrieve data from worksheet

Set INDC_Chart_Series = Chart_Sheet.ChartObjects("NET-OI-INDC").Chart.FullSeriesCollection 'set reference to collection of series on this chart

T = WorksheetFunction.CountIf(Variable_Sheet.ListObjects("User_Selected_Columns").DataBodyRange.Columns(2), True) + 3
'^^^^^the Column Number of the Commercial Net column within the worksheet

'Change in OI is in column 13

ReDim Date_LNG(LBound(Chart_Dates) To UBound(Chart_Dates))

If Not IsNumeric(Chart_Dates(1)) Then

    For Z = LBound(Chart_Dates) To UBound(Chart_Dates)
        Date_LNG(Z) = CLng(Chart_Dates(Z))
    Next Z
    
Else

    For Z = LBound(Chart_Dates) To UBound(Chart_Dates)
        Date_LNG(Z) = Chart_Dates(Z)
    Next Z
    
End If

On Error GoTo NET_OI_Skip

For Z = 2 To UBound(Data_A, 1) 'start on row 2 of array to avoid no data being available

   If Not IsError(Application.Match(Data_A(Z, 1), Date_LNG, 0)) Then
        '^^^^^If current date exists among current xvalues of other charts
        Current_Net = Data_A(Z, T)
        Previous_Net = Data_A(Z - 1, T)
        OI_Change = Data_A(Z, OI_Change_Column)
        BS_Count = BS_Count + 1
        
        X = X + 1
        
        If OI_Change <> 0 And Current_Net <> 0 Then
                    
            If Current_Net > Previous_Net And OI_Change < 0 Then 'Buy signal?:if the Change in Commercial Net positions
                                                               'increases and the change of OI drops
                BuyN = BuyN + 1
                
                If BuyN Mod 2 = 0 Then              'Testing for whether or not BuyN is even allows the points to
                    Buy_Sell_Array(1, BS_Count) = 0.7 'not be placed directly to the left or right of each other
                Else
                    Buy_Sell_Array(1, BS_Count) = 0.65
                End If
                
            End If
             
            If Current_Net < Previous_Net And OI_Change > 0 Then  'Sell signal?:if the Change in Commercial Net positions
                                                                  'falls and the change of OI increases
                SellN = SellN + 1
                
                If SellN Mod 2 = 0 Then
                    Buy_Sell_Array(2, BS_Count) = 0.5
                Else
                    Buy_Sell_Array(2, BS_Count) = 0.45
                End If
                
            End If
            
        End If
        
        If Not IsDate(Chart_Dates(X)) Then Chart_Dates(X) = Format(CDate(Chart_Dates(X)), "yyyy-mm-dd")
        
    End If
    
NET_OI_Skip:

If Err.Number <> 0 Then Err.Clear

Next Z

On Error Resume Next

With INDC_Chart_Series("B_Cluster")
    .Values = WorksheetFunction.Index(Buy_Sell_Array, 1, 0)
    .XValues = Chart_Dates
End With

With INDC_Chart_Series("S_Cluster")
    .Values = WorksheetFunction.Index(Buy_Sell_Array, 2, 0)
    .XValues = Chart_Dates
End With

End Sub

Private Sub Overwrite_Table_Names()

Dim Valid_Tables() As Variant, Z As Long

Valid_Tables = Get_Worksheet_Info

For Z = 1 To UBound(Valid_Tables, 1)
    
    'Valid_Tables(Z, 4).Name = Valid_Tables(Z, 1)
    
Next Z

End Sub

Private Sub Normalize_Format()

Dim gg As Variant, Z As Long, Calculated_Columns_Start As Long, E As Long, PercentA() As Variant, No_Decimal() As Variant

Calculated_Columns_Start = WorksheetFunction.CountIf(Variable_Sheet.ListObjects("User_Selected_Columns").DataBodyRange.Columns(2), True) + 3

gg = Get_Worksheet_Info

Select Case Data_Retrieval.TypeF
    Case "L"
       PercentA = Array(9, 13, 14, 16, 17, 18, 19)
    Case "D"
        PercentA = Array(10, 11, 12)
        No_Decimal = Array(9)
        
    Case "T"
        PercentA = Array(5)
End Select
        
For Z = 1 To UBound(gg, 1)  'Loop through data tables and Format certain columns
On Error GoTo Next_Z

    With gg(Z, 4).Parent
        
        .Columns(1).NumberFormat = "yyyy-mm-dd" 'Date column
        
        .Columns(Calculated_Columns_Start - 3).NumberFormat = "@" 'Contract Codes
        .Columns(Calculated_Columns_Start - 2).NumberFormat = "0.0000" 'Contract Codes
        
        For E = LBound(PercentA) To UBound(PercentA)
            
            .Columns(Calculated_Columns_Start + PercentA(E)).NumberFormat = "0%"
                
        Next E

        For E = LBound(No_Decimal) To UBound(No_Decimal)
            
            .Columns(Calculated_Columns_Start + No_Decimal(E)).NumberFormat = "0"
                
        Next E
        
    End With
    
Next_Z: On Error GoTo -1

Next Z

End Sub
Private Sub Open_Interest_Histogram(Chart_Obj As ChartObject, Index_Key As Long, DataR As Range, SS As Series)

Dim Bin_Size As Double, Histogram_Min_Value As Double, Number_of_Bins As Long, Found_Bin_Group As Boolean, _
Histogram_Info As ChartGroup, Current_Week_Value As Double, V As Long, Chart_Points As Points, Special_RNG As Range

Set Special_RNG = DataR.Columns(Index_Key).SpecialCells(xlCellTypeVisible)
            
On Error GoTo 0
    
SS.Values = DataR.Columns(Index_Key) 'Chart will only show data that is visible

Histogram_Min_Value = WorksheetFunction.Min(Special_RNG) 'Minimum of visible range

Set Histogram_Info = SS.Parent 'set this = to the chart

Bin_Size = Histogram_Info.BinWidthValue  'retrieve the the size of each bin

Number_of_Bins = Histogram_Info.BinsCountValue 'get the total number of bins/columns

Set Histogram_Info = Nothing

'Now determine which bin the most recent value is in.
Current_Week_Value = Special_RNG.End(xlDown).Value

For V = 1 To Number_of_Bins

    If Histogram_Min_Value + (Bin_Size * (V - 1)) <= Current_Week_Value And Current_Week_Value <= Histogram_Min_Value + (Bin_Size * ((V - 1) + 1)) Then
        Current_Week_Value = V
        Found_Bin_Group = True
        Exit For
    End If

Next V

If Found_Bin_Group = False Then Current_Week_Value = 0 'ensures that all bins will be turned blue

Set Chart_Points = SS.Points

For V = 1 To Chart_Points.Count 'turn the bin with the current week's value to yellow else blue

    If V = Current_Week_Value Then
        Chart_Points(V).Format.Fill.ForeColor.RGB = RGB(204, 0, 153) 'RGB(255, 208, 139)
    Else
        Chart_Points(V).Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
    End If
    
Next V

Set Chart_Points = Nothing

Found_Bin_Group = False
Current_Week_Value = 0
V = 0

End Sub
Private Sub Name_Tables()

Dim CFTC_Tables() As Variant, L As Long, Exceptions() As Variant, Market As String

CFTC_Tables = Get_Worksheet_Info

Exceptions = Array("B", "Cocoa", "G", "RC", "Wheat", "W")

For L = 1 To UBound(CFTC_Tables, 1)
    
    Market = IIf(IsError(Application.Match(CFTC_Tables(L, 1), Exceptions, 0)) = True, "CFTC_", "ICE_")
    
    CFTC_Tables(L, 4).Name = Market & Replace(CFTC_Tables(L, 1), "+", ".")
    
Next L

End Sub



