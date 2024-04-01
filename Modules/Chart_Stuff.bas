Attribute VB_Name = "Chart_Stuff"

Option Explicit


Public Sub Update_Charts(Current_Table_Source As ListObject, Sheet_With_Charts As Worksheet, Disable_Filtering As Boolean)

'======================================================================================================
'Edits the referenced worksheet for each series on the worksheet
'======================================================================================================

    Dim TT As Integer, AR As Range, HAT() As Variant, Date_Range As Range, Chart_Series As Series, Array_Method As Boolean, _
    Formula_AR() As String, Dates() As Variant, Area_Compilation As Collection, Chart_Obj As ChartObject
    
    Dim Worksheet_Name As String, Min_Date As Date, Max_Date As Date, Source_Table_Start_Column As Integer, Column_Numbers As New Collection, X As Byte, Y As Byte
    
    Dim Use_User_Dates As Boolean, minimum_date As Date, Maximum_Date As Date, C1 As String, C2 As String, Use_Dashboard_V1_Dates As Boolean, Series_Invalid_Formula As Boolean
    
    Dim updateChartsTimer As New TimedTask
    
    Const filterTableRange As String = "Filter table", calculateBoundsTimer As String = "Calculate Max and Min Date", _
    reassignColumnRangeTimer As String = "Update series ranges", scatterOiCalculation As String = "Scatter OI", _
    histogramUpdate As String = "Update Histogram", priceScaleAdjustment As String = "Price Chart Scale Adjustment", renameTitle As String = "Rename chart titles"
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
    Worksheet_Name = Current_Table_Source.Parent.name
    
    updateChartsTimer.Start Worksheet_Name & " ~ Update Charts (" & Time & ")"
    
    Dim WW As Worksheet
    
    #If DatabaseFile Then
        Set WW = L_Charts
    #Else
        Set WW = Sheet_With_Charts
    #End If
    
    'Chart_Settings_TBL13
    
    With WW.ListObjects("Chart_Settings_TBL").DataBodyRange
    
        If .Cells(2, 2) = True Then
            'If user wants text dates
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
    
                minimum_date = .Cells(3, 2)
                Maximum_Date = .Cells(4, 2)
    
                If CDbl(minimum_date) <> 0 Then C1 = ">="
                If CDbl(Maximum_Date) <> 0 Then C2 = "<="
    
                If (Maximum_Date < minimum_date) And CDbl(Maximum_Date) <> 0 Then
    
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
            
            updateChartsTimer.SubTask(filterTableRange).Start
            
            If Len(C1) > 0 And Len(C2) > 0 Then 'If both a maximum and minimum date have been supplied
    
                Current_Table_Source.Range.AutoFilter _
                    Field:=1, _
                    Criteria1:=C1 & minimum_date, Operator:=xlAnd, Criteria2:=C2 & Maximum_Date
    
            ElseIf Len(C1) > 0 Then 'If only a minimum has been supplied
    
                Current_Table_Source.Range.AutoFilter _
                    Field:=1, _
                    Criteria1:=C1 & minimum_date, Operator:=xlFilterValues
    
            ElseIf Len(C2) > 0 Then 'If only a maximum has been supplied
    
                Current_Table_Source.Range.AutoFilter _
                    Field:=1, _
                    Criteria1:=C2 & Maximum_Date, Operator:=xlFilterValues
                
            End If
            
            updateChartsTimer.SubTask(filterTableRange).EndTask
            
        End If
        
        On Error GoTo Exit_Chart_Update
        Set AR = .DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        Set Date_Range = AR.columns(1)                  'Column 1 of table should hold dates
        
        'updateChartsTimer.SubTask(calculateBoundsTimer).Start
        
        Min_Date = WorksheetFunction.Min(Date_Range)
        Max_Date = WorksheetFunction.Max(Date_Range)
        
        'updateChartsTimer.SubTask(calculateBoundsTimer).EndTask
        
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
            If .Areas.count = 1 Then                        'if only one area then take directly from sheet
                Dates = WorksheetFunction.Transpose(.Value2) 'Create 1D list of dates
            Else
            
                Set Area_Compilation = New Collection
                
                For TT = 1 To .Areas.count                   'Loop each Area and add them to the collection and then combine
                    Area_Compilation.Add .Areas(TT).Value2
                Next TT
                
                Dates = WorksheetFunction.Transpose(Multi_Week_Addition(Area_Compilation, Append_Type.Multiple_2d))  'join areas together and transpose to 1D
                
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
    
        For X = 1 To .count 'For each chart on the Charts Worksheet
            
            Set Chart_Obj = .Item(X)
            
            With Chart_Obj
        
                If Not (.name = "NET-OI-INDC" Or .Chart.ChartType = xlHistogram) Then
        
                    '.Chart.Axes(xlCategory).TickLabels.NumberFormat = "yyyy-mm-dd"
                    
                    On Error Resume Next
                    
                    With .Chart.SeriesCollection
                        
                        For Y = 1 To .count
                        
                            updateChartsTimer.SubTask(reassignColumnRangeTimer).Start
                            
                            Set Chart_Series = .Item(Y)
                            'Split series formula with a $ and use the second to last element to determine what column to map it to within the source table
                            With Chart_Series
                              
                                If InStr(1, .Formula, "$") = 0 Then Series_Invalid_Formula = True
                                
                                If Not Series_Invalid_Formula Then
                                    'And Not HasKey(Column_Numbers, .name)
                                    
                                    #If Not DatabaseFile Then
                                        Formula_AR = Split(.Formula, "$")
                                        TT = Sheet_With_Charts.Cells(1, Formula_AR(UBound(Formula_AR) - 1)).Column - (Source_Table_Start_Column - 1)
                                        
                                        .XValues = Date_Range
                                        .values = AR.columns(TT)
                                        
                                        .name = Column_Numbers(TT)(1)
                                        Erase Formula_AR
                                    #End If
                                
                                ElseIf Series_Invalid_Formula Then
                                
                                    .XValues = Date_Range
                                    .values = AR.columns(Column_Numbers(.name)(0))
                                    Series_Invalid_Formula = False
                                    
                                End If
                                
    '                            .XValues = Date_Range
    '                            .Values = AR.Columns(Column_Numbers(.Name)(0))
                                
                            End With
Next_Regular_Series:
                            updateChartsTimer.SubTask(reassignColumnRangeTimer).Pause
                            
                        Next Y
                    
                    End With
                    
                    On Error GoTo 0
                    
                    If .name = "Price Chart" Then 'Adjust minimum valus to fit price range
                        
                        'Stop
                        
                        updateChartsTimer.SubTask(priceScaleAdjustment).Start
                        #If DatabaseFile Then
                            TT = 1 + Evaluate("VLOOKUP(""" & Left(Current_Table_Source.name, 1) & """,Report_Abbreviation,5,FALSE)")
                        #Else
                            TT = 1 + WorksheetFunction.CountIf(Variable_Sheet.ListObjects(ReturnReportType & "_User_Selected_Columns").DataBodyRange.columns(2), True)
                        #End If
                        
                        With .Chart.Axes(xlValue)
                            .MinimumScale = Application.Min(AR.columns(TT))
                            .MaximumScale = Application.Max(AR.columns(TT))
                        End With
                         
                        updateChartsTimer.SubTask(priceScaleAdjustment).EndTask
                        
                    End If
                    
                ElseIf .Chart.ChartType = xlHistogram Then
        
                    On Error GoTo 0
        
                    Select Case .name 'This is done by chart name since you cant query the formula or source range of the chart
        
                        Case "Open Interest Histogram"
                            
                            updateChartsTimer.SubTask(histogramUpdate).Start
                            
                            TT = 3 'OI
                            
                            On Error GoTo Open_Interest_Series_Missing
                            
                            Set Chart_Series = .Chart.SeriesCollection(1)
        
                            On Error GoTo Error_In_Open_Interest_Histogram_Subroutine
        
                            Call Open_Interest_Histogram(Chart_Obj, TT, AR, Chart_Series, Date_Range.Cells(1) > Date_Range.Cells(2))
                            
                            updateChartsTimer.SubTask(histogramUpdate).EndTask
                            
                    End Select
                    
                ElseIf .name = "NET-OI-INDC" Then
        
                    On Error GoTo Experimental_Chart_Error
                    
                    updateChartsTimer.SubTask(scatterOiCalculation).Start
                    
                    Call ScatterC_OI(Current_Table_Source, Date_RNG:=Date_Range, Chart_Worksheet:=Sheet_With_Charts)
Skip_ScatterC:
                    updateChartsTimer.SubTask(scatterOiCalculation).EndTask
                    
                End If
                
            End With
    
Next_Chart:
        
        Next X
    
    End With
    
    With updateChartsTimer
    
        .SubTask(reassignColumnRangeTimer).EndTask
        
    '    With .SubTask(renameTitle)
    '        .Start
        With Sheet_With_Charts.Shapes("Date Display")
            .TextFrame.Characters.Text = Format(Min_Date, "yyyy-mm-dd") & " to " & Format(Max_Date, "yyyy-mm-dd")
            .Height = Sheet_With_Charts.Range("A1:A2").Height
            .Top = 0
        End With
            '.EndTask
    '    End With
        .EndTask
        
        '.DPrint
        
    End With
    
    Re_Enable
    
    Exit Sub

Chart_Has_No_Title:
    Resume Next_Chart
    
Open_Interest_Series_Missing:
    
    With Chart_Obj.Chart.SeriesCollection
        .Add AR.columns(3) ', xlRows, False, False
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
    
Error_In_Open_Interest_Histogram_Subroutine:
    Resume Next_Chart

Experimental_Chart_Error:
    Resume Skip_ScatterC
    
Exit_Chart_Update:

End Sub
Public Sub ScatterC_OI(Worksheet_Data_ListObject As ListObject, ByVal Date_RNG As Range, Chart_Worksheet As Worksheet)

    Dim BS_Count As Integer, Previous_Net As Long, Data_A() As Variant, T As Integer, Z As Integer, OI_Change As Long, _
    Current_Net As Long, Buy_Sell_Array() As Variant, _
    INDC_Chart_Series As FullSeriesCollection, BuyN As Integer, SellN As Integer, Date_LNG() As Long
    
    Dim Chart_Dates() As Variant
    
    Const OI_Change_Column As Byte = 13
    
    Chart_Dates = Date_RNG.Value2
    
    ReDim Buy_Sell_Array(1 To 2, 1 To UBound(Chart_Dates))
    
    '[1]-Buy
    '[2]-Sell
    
    #If DatabaseFile Then
        T = 3 + Evaluate("VLOOKUP(""" & Left(Worksheet_Data_ListObject.name, 1) & """,Report_Abbreviation,5,FALSE)")
    #Else
        T = 3 + Evaluate("COUNTIF(" & ReturnReportType & "_User_Selected_Columns[Wanted],TRUE)")
    #End If
    
    Data_A = Worksheet_Data_ListObject.DataBodyRange.SpecialCells(xlCellTypeVisible).Value2 'retrieve data from worksheet
    
    If Data_A(1, 1) > Data_A(2, 1) Then
        'The array needs to be reversed so it can be proccessed
        Data_A = Reverse_2D_Array(Data_A, selected_columns:=Array(1, OI_Change_Column, T))
    End If
    
    Set INDC_Chart_Series = Chart_Worksheet.ChartObjects("NET-OI-INDC").Chart.FullSeriesCollection
    
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
    
            If OI_Change <> 0 Then
                        
                If Current_Net > Previous_Net And OI_Change < 0 Then 'Buy signal?:if the Change in Commercial Net positions
                                                                     'increases and the change of OI drops
                    BuyN = BuyN + 1
                    
                    If BuyN Mod 2 = 0 Then              'Testing for whether or not BuyN is even allows the points to
                        Buy_Sell_Array(1, BS_Count) = 0.7 'not be placed directly to the left or right of each other
                    Else
                        Buy_Sell_Array(1, BS_Count) = 0.5 ' 0.65
                    End If
                    
                End If
                 
                If Current_Net < Previous_Net And OI_Change > 0 Then  'Sell signal?:if the Change in Commercial Net positions
                                                                      'falls and the change of OI increases
                    SellN = SellN + 1
                    
                    If SellN Mod 2 = 0 Then
                        Buy_Sell_Array(2, BS_Count) = 0.7 '0.5
                    Else
                        Buy_Sell_Array(2, BS_Count) = 0.5 '0.65 '0.45
                    End If
                    
                End If
                
            End If
            
        End If
        
NET_OI_Skip:
    
    If Err.Number <> 0 Then Err.Clear
    
    Next Z
    
    On Error Resume Next
    
    With INDC_Chart_Series("B_Cluster")
        .values = WorksheetFunction.Index(Buy_Sell_Array, 1, 0)
        .XValues = Date_RNG
    End With
    
    With INDC_Chart_Series("S_Cluster")
        .values = WorksheetFunction.Index(Buy_Sell_Array, 2, 0)
        .XValues = Date_RNG
    End With

End Sub

Private Sub Open_Interest_Histogram(Chart_Obj As ChartObject, Index_Key As Integer, DataR As Range, SS As Series, sortedASC As Boolean)

    Dim Bin_Size As Double, Histogram_Min_Value As Double, Number_of_Bins As Byte, Found_Bin_Group As Boolean, _
    Histogram_Info As ChartGroup, Current_Week_Value As Double, V As Byte, Chart_Points As Points, Special_RNG As Range
    
    Set Special_RNG = DataR.columns(Index_Key).SpecialCells(xlCellTypeVisible)
                
    On Error GoTo TestERR
        
    SS.values = DataR.columns(Index_Key) 'Chart will only show data that is visible
    
    Histogram_Min_Value = WorksheetFunction.Min(Special_RNG) 'Minimum of visible range
    
    Set Histogram_Info = SS.Parent 'set this = to the chart
    
    With Histogram_Info
        Bin_Size = .BinWidthValue  'retrieve the the size of each bin
        Number_of_Bins = .BinsCountValue 'get the total number of bins/columns
    End With
    
    Set Histogram_Info = Nothing
    
    'Now determine which bin the most recent value is in.
    
    With Special_RNG
        
        Current_Week_Value = .End(xlDown).value
        
        If sortedASC Then
            Current_Week_Value = .Rows(1).Value2
        Else
            Current_Week_Value = .Rows(.Rows.count).Value2
        End If
        
    End With
    
    'Current_Week_Value = Partition(Current_Week_Value, Histogram_Min_Value, Histogram_Min_Value + (Bin_Size * Number_of_Bins), Bin_Size)
    
    For V = 1 To Number_of_Bins
    
        If Histogram_Min_Value + (Bin_Size * (V - 1)) <= Current_Week_Value And Current_Week_Value <= Histogram_Min_Value + (Bin_Size * ((V - 1) + 1)) Then
            Current_Week_Value = V
            Found_Bin_Group = True
            Exit For
        End If
    
    Next V
    
    If Not Found_Bin_Group Then Current_Week_Value = 0  'ensures that all bins will be turned blue
    
    Dim currentBinColor As Long, otherBinColor As Long
    
    currentBinColor = RGB(206, 94, 139)
    otherBinColor = RGB(191, 186, 182) 'RGB(178, 178, 178) 'RGB(199, 187, 187)
    
    With SS
        Set Chart_Points = .Points
        .HasDataLabels = False
    End With
    
    For V = 1 To Chart_Points.count 'turn the bin with the current week's value to yellow else blue
    
        'On Error GoTo ghg
        
        With Chart_Points(V)
            
            With .Format
        
    '            With .Line
    '                .ForeColor.RGB = RGB(0, 0, 0)
    '                .Weight = 0.5
    '            End With
                
                If V = Current_Week_Value Then
                    .Fill.ForeColor.RGB = currentBinColor
                Else
                    .Fill.ForeColor.RGB = otherBinColor
                End If
                
            End With
            
    '        If V = Current_Week_Value Then
    '
    '            .ApplyDataLabels AutoText:=True, ShowValue:=True
    '
    '            .HasDataLabel = True
    '
    '            .DataLabel.Text = "Current Bin"
    '
    '        End If
            
        End With
        
    Next V
    
    Set Chart_Points = Nothing
    
    Exit Sub

TestERR:

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

#If DatabaseFile Then
        
    Private Sub AdjustShapesOnCharts()
        
        Dim WS As Variant, launchUf As Shape, showData As Shape, dateDp As Shape, chartSettings As Shape, vv As Variant
        
        For Each WS In Array(L_Charts, D_Charts, T_Charts)
            
'            For Each vv In WS.Shapes
'                Debug.Print vv.name
'            Next vv
            
            With WS
                Set launchUf = .Shapes("Launch Userform")
                Set showData = .Shapes("GoTo Sheet")
                Set dateDp = .Shapes("Date Display")
                Set chartSettings = .Shapes("Chart Settings")
            End With
            
            With launchUf
                .Left = 0
                .Top = 0
                .Height = WS.Range("A1:A2").Height
            End With
            
            dateDp.Left = launchUf.Left + launchUf.Width
            dateDp.Height = launchUf.Height
            dateDp.Top = 0
            
            showData.Left = dateDp.Left + dateDp.Width
            showData.Height = dateDp.Height
            showData.Top = 0
            
            
            chartSettings.Left = showData.Left + showData.Width
            chartSettings.Height = showData.Height
            chartSettings.Top = 0
            
        Next WS
        
    End Sub
        
#End If
