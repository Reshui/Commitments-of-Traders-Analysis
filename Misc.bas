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
Public Sub Open_Contract_Selection()
Attribute Open_Contract_Selection.VB_Description = "Opens a Userform to select a contract or report type "
Attribute Open_Contract_Selection.VB_ProcData.VB_Invoke_Func = "C\n14"
On Error Resume Next
    Contract_Selection.Show

End Sub

Public Sub Change_Background() 'For use on the HUB worksheet
Attribute Change_Background.VB_Description = "Changes the background for the currently active worksheet."
Attribute Change_Background.VB_ProcData.VB_Invoke_Func = " \n14"
    
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
'Sub Navigation_Userform()
'
'Dim UserForm_OB As Object
'
'For Each UserForm_OB In VBA.UserForms
'
'    If UserForm_OB.Name = "Navigation" Then
'
'        Unload UserForm_OB
'        Exit Sub
'    End If
'
'Next UserForm_OB
'
'    Navigation.Show
'
'End Sub


'Private Sub Column_Visibility_Form()
'
'    Column_Visibility.Show
'
'End Sub
Public Sub Hide_Workbooks()
Attribute Hide_Workbooks.VB_Description = "Hides all visible workbooks excluding the currently active one."
Attribute Hide_Workbooks.VB_ProcData.VB_Invoke_Func = " \n14"

Dim WB As Workbook

For Each WB In Application.Workbooks
    If Not WB Is ActiveWorkbook Then WB.Windows(1).Visible = False
Next WB

End Sub
Public Sub Show_Workbooks()
Attribute Show_Workbooks.VB_Description = "Unhides hidden workbooks."
Attribute Show_Workbooks.VB_ProcData.VB_Invoke_Func = " \n14"

Dim WB As Workbook

For Each WB In Application.Workbooks
    WB.Windows(1).Visible = True
Next WB

End Sub
Public Sub Reset_Worksheet_UsedRange(TBL_RNG As Range)
Attribute Reset_Worksheet_UsedRange.VB_Description = "Under Development"
Attribute Reset_Worksheet_UsedRange.VB_ProcData.VB_Invoke_Func = " \n14"
'===========================================================================================
'Reset each worksheets usedrange if there is a valid Table on the worksheet
'Valid Table designate by having CFTC_Market_Code somewhere in its header row
'Anything to the Right or Below this table will be deleted
'===========================================================================================
Dim LRO As Range, LCO As Range, Worksheet_TB As Object, C1 As String, C2 As String, _
Row_Total As Long, UR_LastCell As Range, TB_Last_Cell As Range ', WSL As Range

    Set Worksheet_TB = TBL_RNG.Parent 'Worksheet where table is found
    
    With Worksheet_TB '{Must be typed as object to fool the compiler when resetting the Used Range]

        With TBL_RNG 'Find the Bottom Right cell of the table
            Set TB_Last_Cell = .Cells(.Rows.Count, .Columns.Count)
        End With
        
        With .UsedRange 'Find the Bottom right cell of the Used Range
            Set UR_LastCell = .Cells(.Rows.Count, .Columns.Count)
        End With
        
        If Intersect(UR_LastCell, TB_Last_Cell) Is Nothing Then
        
            'If UR_LastCell AND TB_Last_Cell don't refer to the same cell
            
            With TB_Last_Cell
                Set LRO = .Offset(1, 0) 'last row of table offset by 1
                Set LCO = .Offset(0, 1) 'last column of table offset by 1
            End With
            
            C2 = UR_LastCell.Address
            
            If UR_LastCell.Column <> TB_Last_Cell.Column And UR_LastCell.Row <> TB_Last_Cell.Row Then
                'if rows and columns are different
                
                C1 = LRO.Address
                .Range(C1, C2).EntireRow.Delete 'Delete excess usedrange
                
                C1 = LCO.Address
                .Range(C1, C2).EntireColumn.Delete
                
            ElseIf UR_LastCell.Column <> TB_Last_Cell.Column And UR_LastCell.Row = TB_Last_Cell.Row Then
                'Delete excess columns if columns are different but rows are the same
                C1 = LCO.Address
                .Range(C1, C2).EntireColumn.Delete  'Delete excess columns
                
            ElseIf UR_LastCell.Column = TB_Last_Cell.Column And UR_LastCell.Row <> TB_Last_Cell.Row Then
                'Delete excess rows if rows are different but columns are the same
                C1 = LRO.Address
                .Range(C1, C2).EntireRow.Delete 'Delete exess rows
            End If
        
            .UsedRange 'reset usedrange
            
        End If
    
    End With

End Sub
Private Sub Workbook_Information_Userform()

    Workbook_Information.Show

End Sub

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





