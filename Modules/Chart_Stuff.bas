Attribute VB_Name = "Chart_Stuff"
#Const EnableTimers = False
Option Explicit

Public Sub Update_Charts(Current_Table_Source As ListObject, Sheet_With_Charts As Worksheet, Disable_Filtering As Boolean)

'======================================================================================================
'Edits the referenced worksheet for each series on the worksheet
'======================================================================================================

    Dim iCount As Long, visibleTableDataRange As Range, tableHeaders() As Variant, Date_Range As Range, Chart_Series As Series, _
    Formula_AR$(), chartOnSheet As ChartObject
    
    Dim sourceWorksheetName$, Min_Date As Date, Max_Date As Date, Source_Table_Start_Column As Long, Column_Numbers As New Collection
    
    Dim Use_User_Dates As Boolean, minimum_date As Date, Maximum_Date As Date, inequalityConditionOne$, inequalityConditionTwo$, Use_Dashboard_V1_Dates As Boolean, isSeriesFormulaInvalid As Boolean
    
    Dim updateChartsTimer As New TimedTask, appProperties As Collection, tableSortOrder As XlSortOrder, tableSortFields As SortField
            
    Set appProperties = DisableApplicationProperties(True, False, True)
    
    'Sheet_With_Charts.Calculate
    
    sourceWorksheetName = Current_Table_Source.parent.Name
    
    #If EnableTimers Then
        Const filterTableRange$ = "Filter table", calculateBoundsTimer$ = "Calculate Max and Min Date", _
        reassignColumnRangeTimer$ = "Update series ranges", scatterOiCalculation$ = "Scatter OI", _
        histogramUpdate$ = "Update Histogram", priceScaleAdjustment$ = "Price Chart Scale Adjustment", renameTitle$ = "Rename chart titles"
            
        updateChartsTimer.Start sourceWorksheetName & " ~ Update Charts (" & Time & ")"
    #End If
    
    Dim WW As Worksheet
    
    #If DatabaseFile Then
        Set WW = L_Charts
    #Else
        Set WW = Sheet_With_Charts
    #End If
    
    'Chart_Settings_TBL13
    
    With WW.ListObjects("Chart_Settings_TBL").DataBodyRange
    
    '    If .Cells(5, 2) = True Then 'Use date range starting at Dashboard V1 lookback period
    '
    '        Use_Dashboard_V1_Dates = True
    
        If Not Disable_Filtering And .Cells(1, 2).Value2 = False Then 'If the user wants to use their own dates rather than worksheet dates
    
            If Not IsEmpty(.Range(.Cells(3, 2), .Cells(4, 2))) Then 'if at least one date
    
                minimum_date = .Cells(3, 2).Value2
                Maximum_Date = .Cells(4, 2).Value2
    
                If CDbl(minimum_date) <> 0 Then inequalityConditionOne = ">="
                If CDbl(Maximum_Date) <> 0 Then inequalityConditionTwo = "<="
    
                If (Maximum_Date < minimum_date) And CDbl(Maximum_Date) <> 0 Then
                    MsgBox "Maximum Date cannont be less than Minimum Date. Defaulting to worksheet filters."
                Else
                    Use_User_Dates = True
                End If
    
            End If
    
        End If
    
    End With
    
    With Current_Table_Source 'Object is a valid contract table so retrieve needed info
             
        Source_Table_Start_Column = .Range.Column
        
        On Error GoTo Show_All_Data
        
        Set visibleTableDataRange = .DataBodyRange.SpecialCells(xlCellTypeVisible) 'This is just to test if data is available via error checking
        
        'On Error GoTo Load_Data_Error
             
        'Set visibleTableDataRange = .DataBodyRange 'Load Table Range to variable
        
        On Error GoTo 0
        
    '   If Use_Dashboard_V1_Dates Then 'If the user wants to use the dae range from the V1 dashboard
    '        inequalityConditionOne = ">="                  'Condition 1 set to greater than or equal to
    '        iCount = visibleTableDataRange.Rows.Count - Dashboard_V1.Cells(1, 2).value2 + 1 'Number of data rows - Dashboard N weeks value... +1 is so that >= can apply
    '        If iCount <= 0 Then iCount = 1      'Ensures condition isn't outside the range of the table
    '        Minimum_Date = visibleTableDataRange.Cells(iCount, 1).value2
    '    End If
    
        If Not Disable_Filtering And Use_User_Dates Or Use_Dashboard_V1_Dates Then
            
            '.AutoFilter.ShowAllData
            #If EnableTimers Then
                updateChartsTimer.SubTask(filterTableRange).Start
            #End If
            
            If LenB(inequalityConditionOne) > 0 And LenB(inequalityConditionTwo) > 0 Then 'If both a maximum and minimum date have been supplied
    
                Current_Table_Source.Range.AutoFilter _
                    Field:=1, _
                    Criteria1:=inequalityConditionOne & minimum_date, Operator:=xlAnd, Criteria2:=inequalityConditionTwo & Maximum_Date
    
            ElseIf LenB(inequalityConditionOne) > 0 Then 'If only a minimum has been supplied
    
                Current_Table_Source.Range.AutoFilter _
                    Field:=1, _
                    Criteria1:=inequalityConditionOne & minimum_date, Operator:=xlFilterValues
    
            ElseIf LenB(inequalityConditionTwo) > 0 Then 'If only a maximum has been supplied
    
                Current_Table_Source.Range.AutoFilter _
                    Field:=1, _
                    Criteria1:=inequalityConditionTwo & Maximum_Date, Operator:=xlFilterValues
            End If
            
            #If EnableTimers Then
                updateChartsTimer.SubTask(filterTableRange).EndTask
            #End If
            
        End If
        
        On Error GoTo Catch_NoVisibleDataAvailable
            Set visibleTableDataRange = .DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        ' Column 1 of table should hold dates
        Set Date_Range = visibleTableDataRange.columns(1)

        tableSortOrder = xlAscending
        
        For Each tableSortFields In .Sort.SortFields
            With tableSortFields
                If Not Intersect(.key, Date_Range) Is Nothing Then
                    tableSortOrder = .Order
                    Exit For
                End If
            End With
        Next tableSortFields
        
        'updateChartsTimer.SubTask(calculateBoundsTimer).Start
        
        Min_Date = WorksheetFunction.Min(Date_Range)
        Max_Date = WorksheetFunction.Max(Date_Range)
        
        'updateChartsTimer.SubTask(calculateBoundsTimer).EndTask
        
        tableHeaders = .HeaderRowRange.Value2
          
    End With
    
    With Column_Numbers
        For iCount = 1 To UBound(tableHeaders, 2)
            .Add Array(iCount, tableHeaders(1, iCount)), tableHeaders(1, iCount)
        Next iCount
    End With
    
    Erase tableHeaders
    
    On Error GoTo Show_All_Data
    
    inequalityConditionOne = vbNullString 'variable will now be used to hold Chart columns when needed
    inequalityConditionTwo = vbNullString
    
    On Error GoTo 0
    
    For Each chartOnSheet In Sheet_With_Charts.ChartObjects 'For each chart on the Charts Worksheet
        
        With chartOnSheet
                    
            If Not (.Name = "NET-OI-INDC" Or .Chart.ChartType = xlHistogram) Then
    
                '.Chart.Axes(xlCategory).TickLabels.NumberFormat = "yyyy-mm-dd"
                #If Not DatabaseFile Then
                
                    On Error Resume Next
                    For Each Chart_Series In .Chart.SeriesCollection
                        
                        #If EnableTimers Then
                            updateChartsTimer.StartSubTask reassignColumnRangeTimer
                        #End If
                        
                        'Split series formula with a $ and use the second to last element to determine what column to map it to within the source table
                        With Chart_Series
                          
                            If InStrB(1, .Formula, "$") = 0 Then isSeriesFormulaInvalid = True
                            
                            If Not isSeriesFormulaInvalid Then
                                'And Not HasKey(Column_Numbers, .name)
                                
                                #If Not DatabaseFile Then
                                    Formula_AR = Split(.Formula, "$")
                                    
                                    iCount = 1 + Sheet_With_Charts.Cells(1, Formula_AR(UBound(Formula_AR) - 1)).Column - (Source_Table_Start_Column)
                                    
                                    .XValues = Date_Range
                                    .values = visibleTableDataRange.columns(iCount)
                                    
                                    .Name = Column_Numbers(iCount)(1)
                                    Erase Formula_AR
                                #End If
                            
    '                            Else Then
    '
    '                                .XValues = Date_Range
    '                                .values = visibleTableDataRange.columns(Column_Numbers(.name)(0))
    '                                isSeriesFormulaInvalid = False
                            End If
                            
    '                            .XValues = Date_Range
    '                            .Values = visibleTableDataRange.Columns(Column_Numbers(.Name)(0))
                            
                        End With
Next_Regular_Series:
                        #If EnableTimers Then
                            updateChartsTimer.SubTask(reassignColumnRangeTimer).Pause
                        #End If
                        
                    Next Chart_Series
                
                #End If
                
                On Error GoTo 0
                
                If .Name = "Price Chart" Then 'Adjust minimum valus to fit price range

                    #If EnableTimers Then
                        updateChartsTimer.SubTask(priceScaleAdjustment).Start
                    #End If
                    
                    #If DatabaseFile Then
                        iCount = 1 + Evaluate("VLOOKUP(""" & Left$(Current_Table_Source.Name, 1) & """,Report_Abbreviation,5,FALSE)")
                    #Else
                        iCount = 1 + WorksheetFunction.CountIf(GetAvailableFieldsTable(ReturnReportType()).DataBodyRange.columns(2), True)
                    #End If
                    
                    With .Chart.Axes(xlValue)
                        .MinimumScale = Application.Min(visibleTableDataRange.columns(iCount))
                        .MaximumScale = Application.Max(visibleTableDataRange.columns(iCount))
                    End With
                    
                    #If EnableTimers Then
                        updateChartsTimer.SubTask(priceScaleAdjustment).EndTask
                    #End If
                    
                ElseIf InStrB(1, LCase(.Name), "dry powder") > 0 Then
                    EditDryPowderChart chartOnSheet, tableSortOrder
                End If
                
            ElseIf .Chart.ChartType = xlHistogram Then
    
                On Error GoTo 0
    
                Select Case .Name 'This is done by chart name since you cant query the formula or source range of the chart
    
                    Case "Open Interest Histogram"
                        
                        #If EnableTimers Then
                            updateChartsTimer.SubTask(histogramUpdate).Start
                        #End If
                        
                        iCount = 3 'OI
                        
                        On Error GoTo Open_Interest_Series_Missing
                        
                        Set Chart_Series = .Chart.SeriesCollection(1)
    
                        On Error GoTo Error_In_Open_Interest_Histogram_Subroutine
    
                        Call Open_Interest_Histogram(chartOnSheet, iCount, visibleTableDataRange, Chart_Series, Date_Range.Cells(1) > Date_Range.Cells(2))
                        
                        #If EnableTimers Then
                            updateChartsTimer.SubTask(histogramUpdate).EndTask
                        #End If
                        
                End Select
                
            ElseIf .Name = "NET-OI-INDC" Then
    
                On Error GoTo Experimental_Chart_Error
                
                #If EnableTimers Then
                    updateChartsTimer.SubTask(scatterOiCalculation).Start
                #End If
                
                Call ScatterC_OI(Current_Table_Source, Date_RNG:=Date_Range, Chart_Worksheet:=Sheet_With_Charts)
Skip_ScatterC:
                #If EnableTimers Then
                    updateChartsTimer.SubTask(scatterOiCalculation).EndTask
                #End If
                
            End If
            
        End With

Next_Chart:
    
    Next chartOnSheet
    
    With updateChartsTimer
    
        #If EnableTimers Then
            .SubTask(reassignColumnRangeTimer).EndTask
        #End If
    '    With .SubTask(renameTitle)
    '        .Start
        With Sheet_With_Charts.Shapes("Date Display")
            .TextFrame.Characters.Text = Format(Min_Date, "yyyy-mm-dd") & " to " & Format(Max_Date, "yyyy-mm-dd")
            '.Height = Sheet_With_Charts.Range("A1:A2").Height
            '.Top = 0
        End With
            '.EndTask
    '    End With
        #If EnableTimers Then
            .EndTask
            .DPrint
        #End If
        
    End With
    
Finally:
    EnableApplicationProperties appProperties
    Exit Sub

Chart_Has_No_Title:
    Resume Next_Chart
    
Open_Interest_Series_Missing:
    
    With chartOnSheet.Chart.SeriesCollection
        .Add visibleTableDataRange.columns(3) ', xlRows, False, False
        Set Chart_Series = .Item(1)
    End With
    
    Resume Next
    
Show_All_Data:

    Current_Table_Source.AutoFilter.ShowAllData
    Resume  ' Sends program to line that loads table data into a range variable
    
Load_Data_Error:

    MsgBox ("Data could not be charted for " & sourceWorksheetName)
    Exit Sub
    
OI_Scatter_Chart_Error:
    Resume Skip_ScatterC
    
Error_In_Open_Interest_Histogram_Subroutine:
    Resume Next_Chart

Experimental_Chart_Error:
    Resume Skip_ScatterC
    
Catch_NoVisibleDataAvailable:
    MsgBox "No visible data available."
    Resume Finally
    
End Sub
Public Sub ScatterC_OI(Worksheet_Data_ListObject As ListObject, ByVal Date_RNG As Range, Chart_Worksheet As Worksheet)

    Dim BS_Count As Long, Previous_Net As Long, visibleDataA() As Variant, T As Long, Z As Long, OI_Change As Long, _
    Current_Net As Long, Buy_Sell_Array() As Variant, _
    INDC_Chart_Series As FullSeriesCollection, BuyN As Long, SellN As Long, Date_LNG() As Long
    
    Dim Chart_Dates() As Variant
    
    Const OI_Change_Column As Byte = 13
    
    Chart_Dates = Date_RNG.Value2
    
    ReDim Buy_Sell_Array(1 To 2, 1 To UBound(Chart_Dates))
    
    '[1]-Buy
    '[2]-Sell
    
    #If DatabaseFile Then
        T = 3 + Evaluate("VLOOKUP(""" & Left(Worksheet_Data_ListObject.Name, 1) & """,Report_Abbreviation,5,FALSE)")
    #Else
        T = 3 + Evaluate("COUNTIF(" & ReturnReportType & "_User_Selected_Columns[Wanted],TRUE)")
    #End If
    
    visibleDataA = Worksheet_Data_ListObject.DataBodyRange.SpecialCells(xlCellTypeVisible).Value2
    
    If visibleDataA(1, 1) > visibleDataA(2, 1) Then
        'The array needs to be reversed so it can be proccessed
        visibleDataA = Reverse_2D_Array(visibleDataA, selected_columns:=Array(1, OI_Change_Column, T))
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
    
    For Z = 2 To UBound(visibleDataA, 1) 'start on row 2 of array to avoid no data being available
    
       If Not IsError(Application.Match(visibleDataA(Z, 1), Date_LNG, 0)) Then
            '^^^^^If current date exists among current xvalues of other charts
            Current_Net = visibleDataA(Z, T)
            Previous_Net = visibleDataA(Z - 1, T)
            OI_Change = visibleDataA(Z, OI_Change_Column)
            BS_Count = BS_Count + 1
    
            If OI_Change <> 0 Then
                        
                If Current_Net > Previous_Net And OI_Change < 0 Then 'Buy signal?:if the Change in Commercial Net positions
                                                                     'increases and the change of OI drops
                    BuyN = BuyN + 1
                    
                    'If BuyN Mod 2 = 0 Then              'Testing for whether or not BuyN is even allows the points to
                        'Buy_Sell_Array(1, BS_Count) = 0.7 'not be placed directly to the left or right of each other
                    'Else
                        Buy_Sell_Array(1, BS_Count) = 0.5 ' 0.65
                    'End If

                ElseIf Current_Net < Previous_Net And OI_Change > 0 Then  'Sell signal?:if the Change in Commercial Net positions
                                                                      'falls and the change of OI increases
                    SellN = SellN + 1
                    
                    'If SellN Mod 2 = 0 Then
                        Buy_Sell_Array(2, BS_Count) = -0.5 '0.5
                    'Else
                        'Buy_Sell_Array(2, BS_Count) = 0.5 '0.65 '0.45
                    'End If
                    
                End If
                
            End If
            
        End If
        
NET_OI_Skip:
        If Err.Number <> 0 Then Err.Clear
    
    Next Z
    
    On Error Resume Next
    
    With INDC_Chart_Series("B_Cluster")
        .values = WorksheetFunction.index(Buy_Sell_Array, 1, 0)
        .XValues = Date_RNG
    End With
    
    With INDC_Chart_Series("S_Cluster")
        .values = WorksheetFunction.index(Buy_Sell_Array, 2, 0)
        .XValues = Date_RNG
    End With

End Sub

Private Sub Open_Interest_Histogram(Chart_Obj As ChartObject, Index_Key As Long, DataR As Range, ss As Series, sortedASC As Boolean)

    Dim Bin_Size As Double, Histogram_Min_Value As Double, Number_of_Bins As Byte, Found_Bin_Group As Boolean, _
    Histogram_Info As ChartGroup, Current_Week_Value As Double, v As Byte, Chart_Points As Points, Special_RNG As Range
    
    Set Special_RNG = DataR.columns(Index_Key).SpecialCells(xlCellTypeVisible)
                
    On Error GoTo TestERR
        
    ss.values = DataR.columns(Index_Key) 'Chart will only show data that is visible
    
    Histogram_Min_Value = WorksheetFunction.Min(Special_RNG) 'Minimum of visible range
    
    Set Histogram_Info = ss.parent 'set this = to the chart
    
    With Histogram_Info
        Bin_Size = .BinWidthValue  'retrieve the the size of each bin
        Number_of_Bins = .BinsCountValue 'get the total number of bins/columns
    End With
    
    Set Histogram_Info = Nothing
    
    'Now determine which bin the most recent value is in.
    
    With Special_RNG
        
        Current_Week_Value = .End(xlDown).Value2
        
        If sortedASC Then
            Current_Week_Value = .Rows(1).Value2
        Else
            Current_Week_Value = .Rows(.Rows.count).Value2
        End If
        
    End With
    
    'Current_Week_Value = Partition(Current_Week_Value, Histogram_Min_Value, Histogram_Min_Value + (Bin_Size * Number_of_Bins), Bin_Size)
    
    For v = 1 To Number_of_Bins
    
        If Histogram_Min_Value + (Bin_Size * (v - 1)) <= Current_Week_Value And Current_Week_Value <= Histogram_Min_Value + (Bin_Size * ((v - 1) + 1)) Then
            Current_Week_Value = v
            Found_Bin_Group = True
            Exit For
        End If
    
    Next v
    
    If Not Found_Bin_Group Then Current_Week_Value = 0  'ensures that all bins will be turned blue
    
    Dim currentBinColor As Long, otherBinColor As Long
    
    currentBinColor = RGB(206, 94, 139)
    otherBinColor = RGB(191, 186, 182) 'RGB(178, 178, 178) 'RGB(199, 187, 187)
    
    With ss
        Set Chart_Points = .Points
        .HasDataLabels = False
    End With
    
    For v = 1 To Chart_Points.count 'turn the bin with the current week's value to yellow else blue
    
        'On Error GoTo ghg
        
        With Chart_Points(v)
            
            With .Format
        
    '            With .Line
    '                .ForeColor.RGB = RGB(0, 0, 0)
    '                .Weight = 0.5
    '            End With
                
                If v = Current_Week_Value Then
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
        
    Next v
    
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
Public Sub EditDryPowderChart(chartToEdit As ChartObject, tableSortOrder As XlSortOrder)
    
    'Dim seriesLong As Series, shortSeries As Series, iPoints As Long, defaultMarkerSize As Byte, defaultMarkerColor As Long
    Dim seriesOnChart As Series, markerOne As Series
    
    Dim indexToColor As Long, seriesCount As Byte, firstMarkerAllocated As Boolean, _
    minimumTraders As Long, allocatedPossibleMin As Boolean, minTradersForSeries As Long, _
    recentTraders(1), recentValues(1)
    
    Const mostRecentSeriesName$ = "Recent Values"
    
    With chartToEdit.Chart.SeriesCollection
        Set markerOne = .Item(mostRecentSeriesName)
    End With
    
    For Each seriesOnChart In chartToEdit.Chart.SeriesCollection
                
        With seriesOnChart
            
            If UBound(.values) > 1 And .Name <> mostRecentSeriesName Then
                
                minTradersForSeries = Application.Min(.XValues)

                If Not allocatedPossibleMin Or minTradersForSeries < minimumTraders Then
                    minimumTraders = minTradersForSeries
                    allocatedPossibleMin = True
                End If
                
                indexToColor = IIf(tableSortOrder = xlAscending, UBound(.values), 1)
                
                recentTraders(seriesCount) = .XValues(indexToColor)
                recentValues(seriesCount) = .values(indexToColor)
                seriesCount = seriesCount + 1
            End If
            
        End With
        
    Next seriesOnChart
    
    With markerOne
        .values = recentValues
        .XValues = recentTraders
        '.name = "Recent Values"
    End With
    
    With chartToEdit.Chart.Axes(xlCategory)
        .MinimumScale = Application.Max(0, minimumTraders - .MajorUnit)
    End With
    
End Sub

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
