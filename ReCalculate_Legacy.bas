Attribute VB_Name = "ReCalculate_Legacy"
Sub Recalculate_Workbook_Legacy()

Dim i As Long, dd() As Variant, TB As ListObject

With Application

    dd = .Run("'" & ThisWorkbook.Name & "'!Get_Worksheet_Info")
    
    .ScreenUpdating = False
    
End With

For i = 1 To UBound(dd, 1)

On Error GoTo ski

    Set TB = dd(i, 4)
        
    Call Recalculate_Legacy_Version(TB, dd)
    
ski:

Next i

Application.ScreenUpdating = True

End Sub
Sub Recalculate_Worksheet_Legacy()

With Application

    .ScreenUpdating = False
    
        Call Recalculate_Legacy_Version(CFTC_Table(ThisWorkbook, ActiveSheet), Application.Run("'" & ThisWorkbook.Name & "'!Get_Worksheet_Info"))
        
    .ScreenUpdating = True

End With

End Sub
Public Sub Recalculate_Legacy_Version(TB As ListObject, Workbook_INfo As Variant)

Dim filterArray As Variant, Script_2_Run As String, _
Commercial_Column As Long, Last_Calculated_Column As Long, Price_Column As Long, Data_Array() As Variant, Sorted_Old_2_New As Boolean

If Not TB Is Nothing Then

    Script_2_Run = "'" & ThisWorkbook.Name & "'!Multi_Calculations"
    
    With Variable_Sheet.ListObjects("Saved_Variables").DataBodyRange
    
        Commercial_Column = WorksheetFunction.CountIf(Variable_Sheet.ListObjects("User_Selected_Columns").DataBodyRange.Columns(2), True) + 3
        
        Last_Calculated_Column = WorksheetFunction.VLookup("Last Calculated Column", .Value2, 2, False)
        
        Price_Column = Commercial_Column - 2
        
    End With
    
    With TB
        
        Sorted_Old_2_New = Detect_Old_To_New(TB, 1)
        
        ChangeFilters TB, filterArray
        
        Call Sort_Table(TB, 1, xlAscending)
        
        With .DataBodyRange
            Data_Array = .Cells(1, 1).Resize(.Rows.Count, Price_Column).Value2
        End With
        
        If UBound(Data_Array, 2) <> Last_Calculated_Column Then ReDim Preserve Data_Array(1 To UBound(Data_Array, 1), 1 To Last_Calculated_Column)
            
        Data_Array = Application.Run(Script_2_Run, Data_Array, UBound(Data_Array, 1), Commercial_Column, 156, 26)
        
        Call Retrieve_Tuesdays_CLose(Data_Array, Price_Column, Price_Column - 1, Workbook_INfo)
        
        .DataBodyRange.Value2 = Data_Array
        
        If Sorted_Old_2_New = False Then Call Sort_Table(TB, 1, xlDescending)
        
        RestoreFilters TB, filterArray
    
    End With

End If

End Sub
Public Sub ChangeFilters(w As ListObject, ByRef filterArray)

With w.AutoFilter

    With .Filters
        ReDim filterArray(1 To .Count, 1 To 3)
        For f = 1 To .Count
            With .Item(f)
                If .On Then
                    filterArray(f, 1) = .Criteria1
                    If .Operator Then
                        filterArray(f, 2) = .Operator
                        filterArray(f, 3) = .Criteria2
                    End If
                End If
            End With
        Next
    End With
    
    .ShowAllData
    
End With

End Sub
Public Sub RestoreFilters(w As ListObject, ByVal filterArray)

Dim col As Long

With w.DataBodyRange

    For col = 1 To UBound(filterArray, 1)
    
        If Not IsEmpty(filterArray(col, 1)) Then
            If filterArray(col, 2) Then
                .AutoFilter Field:=col, _
                    Criteria1:=filterArray(col, 1), _
                        Operator:=filterArray(col, 2), _
                    Criteria2:=filterArray(col, 3)
            Else
                .AutoFilter Field:=col, _
                    Criteria1:=filterArray(col, 1)
            End If
            
        End If
        
    Next

End With

End Sub






