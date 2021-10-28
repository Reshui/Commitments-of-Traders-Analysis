Attribute VB_Name = "Dashboards"
Private Sub Dashboard_Basic()

Dim ValidT As Variant, J As Long, g As Long, Group_Number As Long, Output() As Variant, _
WSD() As Variant, TBL_Bottom_Row As Long

Dim Max_Date As Date, Week_Range As Long, Net_Position_Average As Double, _
Percentile_Change_Variance As Double, Net_Position_Variance As Double, Columns_Target() As Variant

Dim Item As Variant, Parsing() As Variant, NP_Percentile_Change_Average As Double, Array_Section() As Variant

Dim C_Net_Column As Long, TB As ListObject, Old_2_New As Boolean, Filters() As Variant, Queried_Row As Long

C_Net_Column = WorksheetFunction.CountIf(Variable_Sheet.ListObjects("User_Selected_Columns").DataBodyRange.Columns(2), True) + 3

Const Columns_Per_Group As Long = 4

ValidT = Application.Run("'" & ThisWorkbook.Name & "'!Get_Worksheet_Info")

Select Case Data_Retrieval.TypeF 'Columns_Per_Group * number of groups +1

    Case "L":
    
        Columns_Target = Array(C_Net_Column, C_Net_Column + 1, C_Net_Column + 2)
        
    Case "D", "T":
    
        Columns_Target = Array(C_Net_Column, C_Net_Column + 1, C_Net_Column + 2, C_Net_Column + 3, C_Net_Column + 4)

End Select

ReDim Output(1 To UBound(ValidT, 1), 1 To ((UBound(Columns_Target) + 1) * Columns_Per_Group) + 1)

Week_Range = Dashboard_V1.Cells(1, 2).Value

For J = 1 To UBound(ValidT, 1) 'For each contract in the array of valid contracts in the workbook
    
    Set TB = ValidT(J, 4)      'Set TB = to the listobject
    
    With TB
    
        Output(J, 1) = .Parent.Name 'Worksheet Name
        
        Old_2_New = Detect_Old_To_New(TB, 1)
        
        ChangeFilters TB, Filters
        
        Call Sort_Table(TB, 1, xlAscending)
        
        WSD = .DataBodyRange.Value2 'Loads worksheet data to an array
        
    End With
    
    Group_Number = 0
    
    TBL_Bottom_Row = UBound(WSD, 1)
    
    If WSD(TBL_Bottom_Row, 1) > Max_Date Then Max_Date = WSD(TBL_Bottom_Row, 1)
    
    If UBound(WSD, 1) >= Week_Range Then 'IF enough data is available
    
        For Each Item In Columns_Target
            
            ReDim Parsing(1 To Week_Range, 1 To 3)
             
            Group_Number = Group_Number + 1
            
'            Net_Position_Average = 0
'            NP_Percentile_Change_Average = 0
'            Net_Position_Variance = 0
'            Percentile_Change_Variance = 0
            
            ' Period Week Change Comparison
            On Error Resume Next
            
            For g = 1 To UBound(Parsing, 1) 'loads net positions/OI into an array
            
                Queried_Row = TBL_Bottom_Row - (g - 1)
                
                Parsing(g, 2) = WSD(Queried_Row, Item) 'net positions
                
                Parsing(g, 1) = (WSD(Queried_Row, Item) - WSD(Queried_Row - 1, Item)) / WSD(Queried_Row - 1, Item) 'net % change as decimal
                
                Parsing(g, 3) = Parsing(g, 2) / WSD(Queried_Row, 3) 'Net/OI
                
'                Net_Position_Average = Net_Position_Average + Parsing(G, 1)
'                NP_Percentile_Change_Average = NP_Percentile_Change_Average + Parsing(G, 2)
                
            Next g
            
            On Error GoTo 0
'
'            Net_Position_Average = Net_Position_Average / UBound(Parsing, 1)
'            NP_Percentile_Change_Average = NP_Percentile_Change_Average / UBound(Parsing, 1)
'
'            For G = 1 To UBound(Parsing, 1)
'                'Percentile_Change_Variance As Double, Net_Position_Variance As Double
'                Net_Position_Variance = ((Net_Position_Average - Parsing(G, 1)) ^ 2) + Net_Position_Variance
'
'                Percentile_Change_Variance = ((NP_Percentile_Change_Average - Parsing(G, 2)) ^ 2) + Percentile_Change_Variance
'
'            Next G
'
'            Net_Position_Variance = Net_Position_Variance / UBound(Parsing, 1)
'
'            Percentile_Change_Variance = Percentile_Change_Variance / UBound(Parsing, 1)
            
            For g = 1 To Columns_Per_Group
                
                Current_Column = (g + 1) + (Columns_Per_Group * Group_Number) - Columns_Per_Group
                
                If g <= 3 Then

                    With WorksheetFunction
                    
                        Array_Section = .Index(Parsing, 0, g) 'Section of array for which values will be computed
                        
                        Output(J, Current_Column) = (Parsing(UBound(Parsing, 1), g) - .Average(Array_Section)) / .StDev(Array_Section)
                        'Computed Z-Score
                    End With
                    
                ElseIf g = 4 Then
                
                    Output(J, Current_Column) = Stochastic_Calculations(CLng(Item), Week_Range, WSD, 1, Cap_Extremes:=True)(1)
                
                End If
                
            Next g
            
            Erase Array_Section
                
        Next Item
    
    End If
    
Next J

Dashboard_V1.Range("A2") = Max_Date

If Old_2_New = False Then Call Sort_Table(TB, 1, xlDescending)

RestoreFilters TB, Filters

With Dashboard_V1.ListObjects("Dashboard_V1").DataBodyRange
     
    .ClearContents
    
    .Resize(UBound(Output, 1), UBound(Output, 2)).Value2 = Output
    
    '.Columns(1).Hyperlinks.Delete

    For J = 1 To UBound(Output, 1)
        .Hyperlinks.Add .Cells(J, 1), "", "'" & Output(J, 1) & "'!A1", , Output(J, 1)
    Next J
    
End With

End Sub


