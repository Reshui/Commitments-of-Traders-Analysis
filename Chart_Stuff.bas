Attribute VB_Name = "Chart_Stuff"

Option Explicit


Public Sub Update_Charts(Current_Table_Source As ListObject, Sheet_With_Charts As Worksheet, Disable_Filtering As Boolean)

'======================================================================================================
'Edits the referenced worksheet for each series on the worksheet
'======================================================================================================


Dim TT As Long, AR As Range, HAT() As Variant, Date_Range As Range, Chart_Series As Series, Array_Method As Boolean, _
Formula_AR() As String, Dates() As Variant, Area_Compilation As Collection, Chart_Obj As ChartObject

Dim Worksheet_Name As String, Min_Date As Date, Max_Date As Date, Source_Table_Start_Column As Long, Column_Numbers As New Collection, X As Long, Y As Long

Dim Use_User_Dates As Boolean, Minimum_Date As Date, Maximum_Date As Date, C1 As String, C2 As String, Use_Dashboard_V1_Dates As Boolean, Series_Invalid_Formula As Boolean

Worksheet_Name = Current_Table_Source.Parent.Name

'Dim DD As Double: DD = Timer

With L_Charts.ListObjects("Chart_Settings_TBL").DataBodyRange

    If .Cells(2, 2) = True Then 'If user wants text dates

        Array_Method = True

        If Err.Number <> 0 Then 'Default to using Range Methods
            Err.Clear
            Array_Method = False
        End If

    End If

'    If .Cells(5, 2) = True Then 'Use date range starting at Dashboard V1 lookback period
'
'        Use_Dashboard_V1_Dates = True

    If Not Disable_Filtering And .Cells(1, 2) = False Then 'If the user wants to use their own dates rather than worksheet dates

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

With Current_Table_Source 'Object is a valid contract table so retrieve needed info
    
    Source_Table_Start_Column = .Range.Cells(1, 1).Column
    
    On Error GoTo Show_All_Data
    
    Set AR = .DataBodyRange.SpecialCells(xlCellTypeVisible) 'This is just to test if data is available via error checking
    
    'On Error GoTo Load_Data_Error
         
    'Set AR = .DataBodyRange 'Load Table Range to variable
    
    On Error GoTo 0
    
'   If Use_Dashboard_V1_Dates Then 'If the user wants to use the dae range from the V1 dashboard
'        C1 = ">="                  'Condition 1 set to greater than or equal to
'        TT = AR.Rows.Count - Dashboard_V1.Cells(1, 2).value + 1 'Number of data rows - Dashboard N weeks value... +1 is so that >= can apply
'        If TT <= 0 Then TT = 1      'Ensures condition isn't outside the range of the table
'        Minimum_Date = AR.Cells(TT, 1).value
'    End If

    If Not Disable_Filtering And Use_User_Dates Or Use_Dashboard_V1_Dates Then
        
        '.AutoFilter.ShowAllData
        
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
            
        End If

    End If
    
    On Error GoTo Exit_Chart_Update
    Set AR = .DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    Set Date_Range = AR.Columns(1)                  'Column 1 of table should hold dates
    
    Min_Date = WorksheetFunction.Min(Date_Range)
    Max_Date = WorksheetFunction.Max(Date_Range)
    
    HAT = .HeaderRowRange.Value2                    'Load headers from table to array
      
End With

With Column_Numbers
    For TT = 1 To UBound(HAT, 2)
        .Add Array(TT, HAT(1, TT)), HAT(1, TT)
    Next TT
End With

Erase HAT

On Error GoTo Show_All_Data

If Array_Method = True Then 'String dates will be used instead of generating from the range data

    With Date_Range.SpecialCells(xlCellTypeVisible) 'This WITH block only affects the array Dates used when Array_Dates=true and for the Experimental Indicator
        On Error GoTo 0
        If .Areas.Count = 1 Then                        'if only one area then take directly from sheet
            Dates = WorksheetFunction.Transpose(.Value2) 'Create 1D list of dates
        Else
        
            Set Area_Compilation = New Collection
            
            For TT = 1 To .Areas.Count                   'Loop each Area and add them to the collection and then combine
                Area_Compilation.Add .Areas(TT).Value2
            Next TT
            
            Dates = WorksheetFunction.Transpose(Multi_Week_Addition(Area_Compilation, Append_Type.multiple_2d))  'join areas together and transpose to 1D
            
            Set Area_Compilation = Nothing
    
        End If
    
    End With

   For TT = LBound(Dates) To UBound(Dates)
       Dates(TT) = Format(CDate(Dates(TT)), "yyyy-mm-dd") 'Convert number dates to DATE typed variable
   Next TT

End If

C1 = vbNullString 'variable will now be used to hold Chart columns when needed
C2 = vbNullString

On Error GoTo 0

With Sheet_With_Charts.ChartObjects

    For X = 1 To .Count 'For each chart on the Charts Worksheet
        
        Set Chart_Obj = .Item(X)
        
        With Chart_Obj
    
            If Not (.Name = "NET-OI-INDC" Or .Chart.ChartType = xlHistogram) Then
    
                '.Chart.Axes(xlCategory).TickLabels.NumberFormat = "yyyy-mm-dd"
                
                On Error Resume Next
                
                With .Chart.SeriesCollection
                
                    For Y = 1 To .Count
                    
                        Set Chart_Series = .Item(Y)
                        'Split series formula with a $ and use the second to last element to determine what column to map it to within the source table
                        With Chart_Series
                          
                            If InStr(1, .Formula, "$") = 0 Then Series_Invalid_Formula = True
                            
                            If Not Series_Invalid_Formula And Not HasKey(Column_Numbers, .Name) Then
                                
                                Formula_AR = Split(.Formula, "$")
                                TT = Sheet_With_Charts.Cells(1, Formula_AR(UBound(Formula_AR) - 1)).Column - (Source_Table_Start_Column - 1)
                                .Name = Column_Numbers(TT)(1)
                                Erase Formula_AR
                                
                            ElseIf Series_Invalid_Formula Then
                            
                                .XValues = Date_Range
                                .Values = AR.Columns(Column_Numbers(.Name)(0))
                                Series_Invalid_Formula = False
                                
                            End If
                            
'                            .XValues = Date_Range
'                            .Values = AR.Columns(Column_Numbers(.Name)(0))
                            
                        End With
Next_Regular_Series:
                    Next Y
                
                End With
                
                On Error GoTo 0
                
                If .Name = "Price Chart" Then 'Adjust minimum valus to fit price range
                                
                    TT = 1 + Evaluate("VLOOKUP(""" & Left(Current_Table_Source.Name, 1) & """,Report_Abbreviation,5,FALSE)")
                    
                    With .Chart.Axes(xlValue)
                        .MinimumScale = Application.Min(AR.Columns(TT))
                        .MaximumScale = Application.Max(AR.Columns(TT))
                    End With
                    
                End If
                
            ElseIf .Chart.ChartType = xlHistogram Then
    
                On Error GoTo 0
    
                Select Case .Name 'This is done by chart name since you cant query the formula or source range of the chart
    
                    Case "Open Interest Histogram"
    
                        TT = 3 'OI
                        
                        On Error GoTo Open_Interest_Series_Missing
                        
                        Set Chart_Series = .Chart.SeriesCollection(1)
    
                        On Error GoTo Error_In_Open_Interest_Histogra_Subroutine
    
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
    
                Call ScatterC_OI(Current_Table_Source, Chart_Dates:=Date_Range.value, Chart_Worksheet:=Sheet_With_Charts)
Skip_ScatterC:
                
            End If
            
            On Error GoTo Chart_Has_No_Title
            
            With .Chart.ChartTitle 'Adjust the title of he chart to shhow the dates plotted
                .Text = Split(.Text, vbTab)(0) & vbTab & "[" & Format(Min_Date, "yyyy-mm-dd") & " to " & Format(Max_Date, "yyyy-mm-dd") & "]"
            End With
        
        End With

Next_Chart:         'On Error GoTo -1
    
    Next X

End With

'Debug.Print Timer - DD

Exit Sub

Chart_Has_No_Title:
    Resume Next_Chart
    
Open_Interest_Series_Missing:
    
    With Chart_Obj.Chart.SeriesCollection
        .Add AR.Columns(3) ', xlRows, False, False
        Set Chart_Series = .Item(1)
    End With
    
    Resume Next
    
Show_All_Data:

    Current_Table_Source.AutoFilter.ShowAllData
    Resume  ' Sends program to line that loads table data into a range variable
    
Load_Data_Error:

    MsgBox ("Data could not be charted for " & Worksheet_Name)
    Exit Sub
    
OI_Scatter_Chart_Error:
    Resume Skip_ScatterC
    
Error_In_Open_Interest_Histogra_Subroutine:
    Resume Next_Chart
    
Exit_Chart_Update:

End Sub
Public Sub ScatterC_OI(Worksheet_Data_ListObject As ListObject, ByVal Chart_Dates As Variant, Chart_Worksheet As Worksheet)

Dim BS_Count As Long, Previous_Net As Long, Data_A() As Variant, T As Long, Z As Long, OI_Change As Long, _
Current_Net As Long, Buy_Sell_Array() As Variant, X As Long, _
INDC_Chart_Series As FullSeriesCollection, BuyN As Long, SellN As Long, Date_LNG() As Long, Report_Name As String

Const OI_Change_Column As Long = 13

ReDim Buy_Sell_Array(1 To 2, 1 To UBound(Chart_Dates))

'Chart Dates is likely an array of long Dates


'[1]-Buy
'[2]-Sell

Report_Name = Evaluate("VLOOKUP(""" & Left(Worksheet_Data_ListObject.Name, 1) & """,Report_Abbreviation,2,FALSE)")

Data_A = Worksheet_Data_ListObject.DataBodyRange.Value2 'retrieve data from worksheet

Set INDC_Chart_Series = Chart_Worksheet.ChartObjects("NET-OI-INDC").Chart.FullSeriesCollection 'set reference to collection of series on this chart

T = WorksheetFunction.CountIf(Variable_Sheet.ListObjects(Report_Name & "_User_Selected_Columns").DataBodyRange.Columns(2), True) + 3
'^^^^^the Column Number of the Commercial Net column within the worksheet

'Change in OI is in column 13

ReDim Date_LNG(LBound(Chart_Dates) To UBound(Chart_Dates))

If Not IsNumeric(Chart_Dates(1, 1)) Then

    For Z = LBound(Chart_Dates, 1) To UBound(Chart_Dates, 1)
        Date_LNG(Z) = CLng(Chart_Dates(Z, 1))
    Next Z
    
Else

    For Z = LBound(Chart_Dates, 1) To UBound(Chart_Dates, 1)
        Date_LNG(Z) = Chart_Dates(Z, 1)
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
        
        If Not IsDate(Chart_Dates(X, 1)) Then Chart_Dates(X, 1) = Format(CDate(Chart_Dates(X, 1)), "yyyy-mm-dd")
        
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
Current_Week_Value = Special_RNG.End(xlDown).value

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
Public Function Non_Equal_Arrays(AR1 As Variant, AR2 As Variant) As Boolean 'Arrays must be 1D

    Dim Y As Long
    
    If UBound(AR1) <> UBound(AR2) Then
    
        Non_Equal_Arrays = True
        Exit Function
        
    Else
    
        For Y = LBound(AR1) To UBound(AR1)
        
            If AR1(Y) <> AR2(Y) Then
                Non_Equal_Arrays = True
                Exit Function
            End If
            
        Next Y
        
    End If

End Function


