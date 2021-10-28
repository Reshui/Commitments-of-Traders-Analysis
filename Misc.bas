Attribute VB_Name = "Misc"
'Private Sub My_Resize()
'
'    Dim Valid_Table_Info() As Variant, Item As Variant, WSC As Range
'
'    Application.ScreenUpdating = False
'
'    Valid_Table_Info = Variable_Sheet.ListObjects("Table_WSN").DataBodyRange.Columns(3).Value2
'
'
'    For Each Item In Valid_Table_Info
'
'        With ThisWorkbook.Worksheets(Item).Range("A3").ListObject
'
'
'        .DataBodyRange.Rows(.DataBodyRange.Rows.Count).ClearContents
'        .Resize .Range.CurrentRegion
'        End With
'
'    Next Item
'
'    With Application
'    '    .CutCopyMode = False
'        .ScreenUpdating = True
'    End With
'
'End Sub


Public Sub Change_Background() 'For use on the HUB worksheet
    
Dim fNameAndPath As Variant, WP As Range, SplitN() As String, ZZ As Long, _
File_Path As String, File_Content As String

If UUID Then

    If (ActiveSheet Is HUB Or ActiveSheet Is Weekly Or ActiveSheet Is Variable_Sheet) Then
        Background_Changer.Show
    Else
        GoTo Normal_Change
    End If
    
Else

Normal_Change:

    fNameAndPath = Application.GetOpenFilename(, Title:="Select Background_Image. If no image is selected, background will not be changed")
    
    If Not fNameAndPath = "False" Then
        On Error GoTo Invalid_Image
        ActiveSheet.SetBackgroundPicture Filename:=fNameAndPath
    End If
    
End If

Exit Sub

Invalid_Image:

MsgBox "An error occured while attempting to apply the selected file."

End Sub
Sub ToTheHub()
Attribute ToTheHub.VB_Description = "CTRL+B"
Attribute ToTheHub.VB_ProcData.VB_Invoke_Func = "b\n14"
     HUB.Activate
End Sub
Sub To_Charts()
Attribute To_Charts.VB_Description = "CTRL+SHIFT+B"
Attribute To_Charts.VB_ProcData.VB_Invoke_Func = "B\n14"
    Chart_Sheet.Activate
End Sub
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
Sub Column_Visibility_Form()
Attribute Column_Visibility_Form.VB_Description = "Opens Userform to hide or show workbook columns."
Attribute Column_Visibility_Form.VB_ProcData.VB_Invoke_Func = " \n14"

    Column_Visibility.Show

End Sub
Public Sub Hide_Workbooks()

Dim WB As Workbook

For Each WB In Application.Workbooks
    If Not WB Is ActiveWorkbook Then WB.Windows(1).Visible = False
Next WB

End Sub
Public Sub Show_Workbooks()

Dim WB As Workbook

For Each WB In Application.Workbooks
    WB.Windows(1).Visible = True
Next WB

End Sub
Sub Copy_Formats_From_ActiveSheet()
Attribute Copy_Formats_From_ActiveSheet.VB_Description = "Copies conditional formats from the current worksheet to all other worksheets that host CFTC data."
Attribute Copy_Formats_From_ActiveSheet.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Valid_Table_Info As Variant, i As Long, TS As Worksheet, ASH As Worksheet, Target_TableR As Range, OT As ListObject, _
Hidden_Collection As New Collection, HC As Range, T As Long, Original_Hidden_Collection As New Collection

With Application

   .ScreenUpdating = False

On Error GoTo Invalid_Function

    Valid_Table_Info = .Run("'" & ActiveWorkbook.Name & "'!Get_Worksheet_Info")

On Error GoTo 0

End With

Set ASH = ActiveWorkbook.ActiveSheet

On Error Resume Next

Set OT = CFTC_Table(ActiveWorkbook, ASH)

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

For i = LBound(Valid_Table_Info, 1) To UBound(Valid_Table_Info, 1)

    Set Target_TableR = Valid_Table_Info(i, 4).DataBodyRange 'databodyrange of the target table
  
    If Not Target_TableR.Parent Is ASH Then 'if the worksheet objects aren't the same
               
        With Hidden_Collection 'store range objects of hidden columns inside a collection
            
            For Each HC In Valid_Table_Info(i, 4).HeaderRowRange.Cells 'Loop cells in the header and check hidden property
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
    MsgBox "Function Get_Worksheet_Info is unavailable. This Macro is intended for files created by MoshiM." & vbNewLine & vbNewLine & _
    "If you see this message and one of my files is the Active Workbook then please contact me."

End Sub
Public Sub Reset_UsedRange()
'===========================================================================================
'Reset each worksheets usedrange if there is a valid Table on the worksheet
'Valid Table designate by having CFTC_Market_Code somewhere in its header row
'Anything to the Right or Below this table will be deleted
'===========================================================================================
Dim HN As Variant, LRO As Range, LCO As Range, i As Long, TBL_RNG As Range, Worksheet_TB As Object, _
Column_Total As Long, Row_Total As Long, UR_LastCell As Range, TB_Last_Cell As Range ', WSL As Range

With Application 'Store all valid tables in an array
    HN = .Run("'" & ThisWorkbook.Name & "'!Get_Worksheet_Info")
End With

For i = 1 To UBound(HN, 1)

    Set TBL_RNG = HN(i, 4).Range      'Entire range of table
    Set Worksheet_TB = TBL_RNG.Parent 'Worksheet where table is found
    
    With Worksheet_TB '{Must be typed as object to fool the compiler when resetting the Used Range]

        With TBL_RNG 'Find the Bottom Right cell of the table
            Set TB_Last_Cell = .Cells(.Rows.Count, .Columns.Count)
        End With
        
        With .UsedRange 'Find the Bottom right cell of the Used Range
            Set UR_LastCell = .Cells(.Rows.Count, .Columns.Count)
        End With
        
        If UR_LastCell.Address <> TB_Last_Cell.Address Then
        
            'If UR_LastCell AND TB_Last_Cell don't refer to the same cell
            
            With TB_Last_Cell
                Set LRO = .Offset(1, 0) 'last row of table offset by 1
                Set LCO = .Offset(0, 1) 'last column of table offset by 1
            End With
            
            If UR_LastCell.Column <> TB_Last_Cell.Column And UR_LastCell.Row = TB_Last_Cell.Row Then
                'Delete excess columns if columns are different but rows are the same
                
                .Range(LCO, UR_LastCell).EntireColumn.Delete  'Delete excess columns
                
            ElseIf UR_LastCell.Column = TB_Last_Cell.Column And UR_LastCell.Row <> TB_Last_Cell.Row Then
                'Delete excess rows if rows are different but columns are the same
                
                .Range(LRO, UR_LastCell).EntireRow.Delete 'Delete exess rows
                
            ElseIf UR_LastCell.Column <> TB_Last_Cell.Column And UR_LastCell.Row <> TB_Last_Cell.Row Then
                'if rows and columns are different
                
                .Range(LRO, UR_LastCell).EntireRow.Delete 'Delete excess usedrange
                .Range(LCO, UR_LastCell).EntireColumn.Delete
                
            End If
        
            .UsedRange 'reset usedrange
            
        End If
    
    End With
       
Next i

End Sub
Public Sub Copy_Valid_Data_Headers()
Attribute Copy_Valid_Data_Headers.VB_Description = "Copies header names to other worksheets."
Attribute Copy_Valid_Data_Headers.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Headers() As Variant, TB As ListObject, WS As Worksheet, Table_Info() As Variant, i As Long

Set TB = CFTC_Table(ThisWorkbook, ActiveSheet)

If Not TB Is Nothing Then

    With Application 'Store all valid tables in an array
        Table_Info = .Run("'" & ThisWorkbook.Name & "'!Get_Worksheet_Info")
    End With

    Headers = TB.HeaderRowRange.Value2

    For i = 1 To UBound(Table_Info)

        If Not Table_Info(i, 4) Is TB Then

            Table_Info(i, 4).HeaderRowRange.Resize(1, UBound(Headers, 2)) = Headers

        End If

    Next i

End If

End Sub
Sub Autofit_Columns()

Dim TB As Long, TBR As Range

With Application
    .ScreenUpdating = False
    Valid_Table_Info = .Run("'" & ActiveWorkbook.Name & "'!Get_Worksheet_Info")
End With

For TB = LBound(Valid_Table_Info) To UBound(Valid_Table_Info)
    Set TBR = Valid_Table_Info(TB, 4).Range
    TBR.Columns.AutoFit
Next TB

 Application.ScreenUpdating = True
 
End Sub
Private Sub Workbook_Information_Userform()

    Workbook_Information.Show

End Sub
Sub Copy_Formulas_From_Active_Sheet()
Attribute Copy_Formulas_From_Active_Sheet.VB_Description = "Copies formulas from the activesheet."
Attribute Copy_Formulas_From_Active_Sheet.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim Valid_Table_Info() As Variant, TB As ListObject, Source_TB_RNG As Range, _
    i As Long, Valid_Table As Boolean
    
    Valid_Table_Info = Application.Run("'" & ActiveWorkbook.Name & "'!Get_Worksheet_Info")
    
    For Each TB In ActiveSheet.ListObjects 'Find the Listobject on the activesheet within the array
        
        For i = 1 To UBound(Valid_Table_Info, 1)
            
            If Valid_Table_Info(i, 4) Is TB Then
            
                Valid_Table = True
                Exit For
                
            End If
            
        Next i
        
        If Valid_Table = True Then Exit For
        
    Next TB
    
    If Valid_Table = False Then GoTo Active_Sheet_is_Invalid
    
    Set Source_TB_RNG = Valid_Table_Info(i, 4).DataBodyRange
    
    Dim Formula_Collection As New Collection, Cell As Range, Item As Variant
    
    For Each Cell In Source_TB_RNG.Rows(Source_TB_RNG.Rows.Count).Cells
    
        With Cell
            If Left$(.Formula, 1) = "=" Then Formula_Collection.Add Array(.Formula, .Column)
        End With
        
    Next Cell
    
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    
    For i = 1 To UBound(Valid_Table_Info, 1) 'loop all listobjects contained within the array
        
        Set TB = Valid_Table_Info(i, 4)
        
        If Not TB Is Source_TB_RNG.ListObject Then 'if not the table that is being copied from
            
            With TB.DataBodyRange 'Take formulas from collection and apply
                
                For Each Item In Formula_Collection
                    .Cells(.Rows.Count, Item(1)).Formula = Item(0)
                Next
                
            End With
           
        End If
        
    Next i
    
    Re_Enable

Set Formula_Collection = Nothing

Exit Sub

Active_Sheet_is_Invalid:

    MsgBox "You are trying to copy data formulas from an invalid worksheet"
    
    Application.Calculation = xlCalculationAutomatic
    
End Sub

'Sub FreezeF()
''
'' FreezeF Macro

'For Each WS In ThisWorkbook.Worksheets
'
'If WS.Index > 4 And WS.Name <> QueryT.Name Then
'
'    WS.Activate
'    ActiveWindow.FreezePanes = False
'    WS.Range("B2").Select
'    ActiveWindow.FreezePanes = True
'End If
'
'Next WS
'
'End Sub


'Sub Filter_Table_Dates()
''
'' Filter_Dates Macro
''
'Dim TB As ListObject, WS As Worksheet
'
'For Each WS In ThisWorkbook.Worksheets
'
'    Set TB = CFTC_Table(ThisWorkbook, WS)
'
'    If Not TB Is Nothing Then
'
'        TB.Range.AutoFilter Field:=1, Criteria1:= _
'            ">01-01-2015", Operator:=xlAnd
'    End If
'
'
'Next WS
'
'End Sub


'Sub Turn_Text_White()
''
'For Each WS In ThisWorkbook.Worksheets
'    If WS.Index > 4 And WS.Name <> QueryT.Name Then
'
'        With WS.ListObjects(1).Range.Cells(1, 1).Font
'            .ThemeColor = xlThemeColorDark1
'            .TintAndShade = 0
'        End With
'
'    End If
'Next
'End Sub

'Sub ff()
'
'Dim Wind As Window, WS As Worksheet
'
'For Each WS In ThisWorkbook.Worksheets
'
'    WS.Activate
'
'    ActiveWindow.Zoom = 70
'
'Next
'
'End Sub





