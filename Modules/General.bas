Attribute VB_Name = "General"

Option Explicit

Sub Re_Enable()
Attribute Re_Enable.VB_Description = "Use this macro if the screen stops working or events fail to fire."
Attribute Re_Enable.VB_ProcData.VB_Invoke_Func = " \n14"

    With Application
        If .EnableEvents = False Then .EnableEvents = True
        If .Calculation <> xlCalculationAutomatic Then .Calculation = xlCalculationAutomatic
        If .DisplayStatusBar = False Then .DisplayStatusBar = True
        If .ScreenUpdating = False Then .ScreenUpdating = True
    End With

End Sub
Sub IncreasePerformance()
Attribute IncreasePerformance.VB_Description = "Turns off screen updating, sets calculations to manual and turns off events."
Attribute IncreasePerformance.VB_ProcData.VB_Invoke_Func = " \n14"

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    ThisWorkbook.ActiveSheet.DisplayPageBreaks = False
    
End Sub

Public Sub Hide_Workbooks()
'===================================================================================================================
    'Purpose: Hides all workbooks except for the currently active one.
    'Inputs:
    'Outputs:
'===================================================================================================================
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
    Dim LRO As Range, LCO As Range, Worksheet_TB As Object, C1$, C2$, _
    Row_Total As Long, UR_LastCell As Range, TB_Last_Cell As Range ', WSL As Range

    Set Worksheet_TB = TBL_RNG.parent 'Worksheet where table is found
    
    With Worksheet_TB '{Must be typed as object to fool the compiler when resetting the Used Range]

        With TBL_RNG 'Find the Bottom Right cell of the table
            Set TB_Last_Cell = .Cells(.Rows.count, .columns.count)
        End With
        
        With .UsedRange 'Find the Bottom right cell of the Used Range
            Set UR_LastCell = .Cells(.Rows.count, .columns.count)
        End With
        
        If Intersect(UR_LastCell, TB_Last_Cell) Is Nothing Then
        
            'If UR_LastCell AND TB_Last_Cell don't refer to the same cell
            
            With TB_Last_Cell
                Set LRO = .offset(1, 0) 'last row of table offset by 1
                Set LCO = .offset(0, 1) 'last column of table offset by 1
            End With
            
            C2 = UR_LastCell.Address
            
            If UR_LastCell.Column <> TB_Last_Cell.Column And UR_LastCell.row <> TB_Last_Cell.row Then
                'if rows and columns are different
                
                C1 = LRO.Address
                .Range(C1, C2).EntireRow.Delete 'Delete excess usedrange
                
                C1 = LCO.Address
                .Range(C1, C2).EntireColumn.Delete
                
            ElseIf UR_LastCell.Column <> TB_Last_Cell.Column And UR_LastCell.row = TB_Last_Cell.row Then
                'Delete excess columns if columns are different but rows are the same
                C1 = LCO.Address
                .Range(C1, C2).EntireColumn.Delete  'Delete excess columns
                
            ElseIf UR_LastCell.Column = TB_Last_Cell.Column And UR_LastCell.row <> TB_Last_Cell.row Then
                'Delete excess rows if rows are different but columns are the same
                C1 = LRO.Address
                .Range(C1, C2).EntireRow.Delete 'Delete exess rows
            End If
        
            .UsedRange 'reset usedrange
            
        End If
    
    End With

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
Sub Remove_Worksheet_Formatting()
Attribute Remove_Worksheet_Formatting.VB_Description = "Removes all conditional formatting from the active worksheet."
Attribute Remove_Worksheet_Formatting.VB_ProcData.VB_Invoke_Func = " \n14"
'===================================================================================================================
    'Purpose: Deletes conditional formatting from the currently active worksheet.
    'Inputs:
    'Outputs: Stores the current time on the Variable_Sheet along with the local time on the running environment.
    'Note: Keyboard shortcut > CTRL+SHIFT+X
'===================================================================================================================
    Cells.FormatConditions.Delete
End Sub
Sub ZoomToRange(ByRef ZoomThisRange As Range, ByVal PreserveRows As Boolean, WB As Workbook)

    Application.ScreenUpdating = False

    Dim Wind As Window

    Set Wind = ActiveWindow

    Application.Goto ZoomThisRange.Cells(1, 1), True

    With ZoomThisRange
        If PreserveRows = True Then
            .Resize(.Rows.count, 1).Select
        Else
            .Resize(1, .columns.count).Select
        End If
    End With

    With Wind
        .Zoom = True
        .VisibleRange.Cells(1, 1).Select
    End With

    If Not WB.ActiveSheetBeforeSaving Is Nothing And IsOnCreatorComputer Then 'accounting for if the variable has not been declared for normal use
        'do nothing
    Else
        Application.ScreenUpdating = True
    End If

End Sub
Public Sub ClearRegionBeneathTable(dataTable As ListObject)
    
    Dim lastUsedRowNumber As Long, appProperties As Collection

    With dataTable.parent.UsedRange
        lastUsedRowNumber = .Cells(.Rows.count, 1).row
    End With
    
    'Clear rows below the table
    
    With dataTable.DataBodyRange
        With .Rows(.Rows.count).offset(1)
        
            If lastUsedRowNumber >= .row Then
                Set appProperties = DisableApplicationProperties(True, False, True)
                
                With .Resize(lastUsedRowNumber - .row + 1, .columns.count)
                
                    .ClearContents

                    With .Interior
                        .Pattern = xlNone
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    
                    .Borders(xlDiagonalDown).LineStyle = xlNone
                    .Borders(xlDiagonalUp).LineStyle = xlNone
                    .Borders(xlEdgeLeft).LineStyle = xlNone
                    .Borders(xlEdgeTop).LineStyle = xlNone
                    .Borders(xlEdgeBottom).LineStyle = xlNone
                    .Borders(xlEdgeRight).LineStyle = xlNone
                    .Borders(xlInsideVertical).LineStyle = xlNone
                    .Borders(xlInsideHorizontal).LineStyle = xlNone
                    
                End With
                
                EnableApplicationProperties appProperties
                
            End If
            
        End With
        
    End With
    
End Sub
Public Sub ChangeFilters(queriedTable As ListObject, ByRef filterArray)
'===================================================================================================================
    'Purpose: Loads filters into filterArray and clears from queriedTable.
    'Inputs: queriedTable - ListObject that will have its filters removed and stored.
    '        filterArray -  Array that will store removed filters.
    'Note:
'===================================================================================================================
    Dim f As Long

    With queriedTable.AutoFilter

        With .Filters
            ReDim filterArray(1 To .count, 1 To 3)
            On Error GoTo Show_Data
            For f = 1 To .count
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
Show_Data:
        .ShowAllData
        
    End With

End Sub
Public Sub RestoreFilters(tableOBJ As ListObject, ByVal filterArray)
'===================================================================================================================
    'Purpose: uses filterArray to reapply filters.
    'Inputs: tableOBJ - ListObject that has filters applied to it.
    '        filterArray - array generated from ChangeFilters().
    'Outputs:
'===================================================================================================================
    Dim col As Long

    With tableOBJ.DataBodyRange
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
Private Sub ConvertAllNamedRangesToWorksheetScopeOnWorksheet()
'=============================================================================================
'   Scopes all named ranges on the current sheet to the worksheet rather than to the workbook.
'=============================================================================================
    Dim nm As Name, workbookActiveSheet As Worksheet, rangeNameRefersTo$, nameOfRange$
    'worksheet scope MT!hello
    'workbook scope hello
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
    Set workbookActiveSheet = ThisWorkbook.ActiveSheet
    
    On Error GoTo Attempt_Next_Name
    
    For Each nm In ThisWorkbook.names
        With nm
            If .RefersToRange.parent Is workbookActiveSheet And InStrB(1, nm.Name, workbookActiveSheet.Name & "!") <> 1 Then
                rangeNameRefersTo = .RefersTo
                nameOfRange = .Name
                .Delete
                workbookActiveSheet.names.Add workbookActiveSheet.Name & "!" & nameOfRange, rangeNameRefersTo
            End If
        End With
        
Attempt_Next_Name:
        On Error GoTo -1
    Next nm
    
    Re_Enable
    
End Sub

Public Sub Run_This(WB As Workbook, ScriptN$)
    Application.Run "'" & WB.Name & "'!" & ScriptN
End Sub

Public Function DisableApplicationProperties(disableEvents As Boolean, disableAutoCalculations As Boolean, disableScreenUpdating As Boolean) As Collection
    
    Dim values As New Collection
    
    With Application
        
        If disableEvents And .EnableEvents = True Then
            values.Add .EnableEvents, "Events"
            .EnableEvents = False
        End If
            
        If disableAutoCalculations And .Calculation <> xlCalculationManual Then
            values.Add .Calculation, "Calc"
            .Calculation = xlCalculationManual
        End If
        
        If disableScreenUpdating And .ScreenUpdating = True Then
            values.Add .ScreenUpdating, "Screen"
            .ScreenUpdating = False
        End If
                        
    End With
    
    Set DisableApplicationProperties = values
    
End Function
Public Sub EnableApplicationProperties(values As Collection)

    If Not values Is Nothing Then
        If values.count > 0 Then
            With Application
                On Error Resume Next
                .Calculation = values("Calc")
                .EnableEvents = values("Events")
                .ScreenUpdating = values("Screen")
                If Err.Number <> 0 Then Err.Clear
            End With
        End If
    End If
    
End Sub
Public Sub ResizeTableBasedOnColumn(lo As ListObject, columnToMatchLastUsedRow As Range)
'====================================================================================================================================
'   Summary: Resizes a Listobject so that its last row is in the same row as the last used row in the column represented bycolumnToMatchLastUsedRow
'====================================================================================================================================
    Dim bottomInColumn As Range, isCellEmpty As Boolean, newBottom As Range, _
    rowsToKeepCount As Long, worksheetWithTable As Worksheet, givenTableColumn As Boolean
    
    If columnToMatchLastUsedRow.columns.count > 1 Then
        MsgBox "Input range has should have only a single column"
        Exit Sub
    End If
    
    With lo
        
        Set worksheetWithTable = .parent
        
        With .DataBodyRange
                                            
            If Intersect(lo.DataBodyRange, columnToMatchLastUsedRow) Is Nothing Then
                Set bottomInColumn = worksheetWithTable.Cells(worksheetWithTable.Rows.count, columnToMatchLastUsedRow.Column).End(xlUp)
            Else
                Set bottomInColumn = .Cells(.Rows.count, columnToMatchLastUsedRow.Column - .Column + 1)
                givenTableColumn = True
            End If
            
            isCellEmpty = IsEmpty(bottomInColumn.Value2)
            
            If ((givenTableColumn And isCellEmpty) Or (Not givenTableColumn And Not isCellEmpty)) And bottomInColumn.row > lo.Range.row Then
                
                Set newBottom = IIf(givenTableColumn, bottomInColumn.End(xlUp), bottomInColumn)
            
                If Not newBottom.row = lo.Range.row And Not newBottom.row = .Rows(.Rows.count).row Then
                    rowsToKeepCount = newBottom.row - lo.Range.row + 1
                    lo.Resize lo.Range.Resize(rowsToKeepCount, .columns.count)
                End If
            
            End If
            
        End With
        
    End With

End Sub
Public Function GetNumbers(inputColumn As Variant) As Variant
'=============================================================================================
'Summary:   Finds the first whole or decimal number contained within each cell of inputRange.
'Output:    An array containing the first number of each cell.
'=============================================================================================
    Dim data() As Variant, iColumn As Byte, iRow As Long, outputA() As Variant, stringBytes() As Byte, _
    byteIndex As Byte, currentNumberBytes() As Byte, validCharacter As Boolean, cursorIndex As Byte, _
    numberAsString$, skipCharacter As Boolean, inputIsRange As Boolean, startLocation As Byte
    
    #Const useTimer = False
    
    #If useTimer Then
        Dim numberTimer As New TimedTask
        numberTimer.Start "Get Numbers"
    #End If
    
    If TypeName(inputColumn) = "Range" Then
        data = inputColumn.Value2
        inputIsRange = True
        'Debug.Print inputColumn.Parent.name
    ElseIf IsArray(inputColumn) Then
        data = inputColumn
    Else
        Exit Function
    End If
    
    ReDim outputA(LBound(data, 1) To UBound(data, 1), LBound(data, 2) To UBound(data, 2))
    
    For iRow = LBound(data, 1) To UBound(data, 1)
        
        For iColumn = LBound(data, 2) To UBound(data, 2)
        
            If Not IsEmpty(data(iRow, iColumn)) And Not IsNull(data(iRow, iColumn)) Then
                'This step converts strings to a byte array
                stringBytes = data(iRow, iColumn)
                ' Sized to fit the possibility that data(iRow,iColumn) is just a number.
                ReDim currentNumberBytes(LBound(stringBytes) To UBound(stringBytes))
                ' cursorIndex is the current index within currentNumberBytes to write to if a valid character is found.
                cursorIndex = 0
                ' Every second byte is a 0 so skip over it.
                                
                startLocation = InStrB(1, data(iRow, iColumn), "$")
                'If a dollar sign is found then startLocation is initialized base 1 and is therefore off by 1.
                If startLocation > 0 Then startLocation = startLocation + 1
                
                For byteIndex = startLocation To UBound(stringBytes) Step 2
                        
                    skipCharacter = False
                    
                    Select Case stringBytes(byteIndex)
                        ' 0 through 9
                        Case 48 To 57
                            validCharacter = True
                        Case 44, 46
                        'Comma, Period
                            On Error GoTo IndexOutOfRange
                            'Ensure that it is sandwiched between 2 numbers
                            If IsCharCodeNumber(currentNumberBytes(cursorIndex - 2)) And IsCharCodeNumber(stringBytes(byteIndex + 2)) And IsCharCodeNumber(stringBytes(byteIndex + 2)) Then
                                 validCharacter = True
                                 'If comma
                                 If stringBytes(byteIndex) = 44 Then skipCharacter = True
                            End If
                        Case 95
                            'minus sign
                            On Error GoTo IndexOutOfRange
                            If IsCharCodeNumber(stringBytes(byteIndex + 2)) And Not IsCharCodeNumber(currentNumberBytes(cursorIndex - 2)) Then
                                validCharacter = True
                            End If
                        Case Else
                            validCharacter = False
                    End Select
Next_Byte_Index:
                    On Error GoTo 0
                    
                    If Not skipCharacter Then
                        If validCharacter Then
                            
                            currentNumberBytes(cursorIndex) = stringBytes(byteIndex)
                            cursorIndex = cursorIndex + 2
                            validCharacter = False
                            
                        ElseIf cursorIndex > 0 Then
                            ' Tests if valid bytes have already been found.
                            Exit For
                        End If
                    End If
                    
                Next byteIndex
                      
                If cursorIndex > 0 Then
                    ReDim Preserve currentNumberBytes(0 To cursorIndex - 1)
                    numberAsString = currentNumberBytes
                    On Error GoTo StoreStringInOutputA
                    outputA(iRow, iColumn) = numberAsString * 1
                End If
                
                On Error GoTo 0
                
            End If
            
        Next iColumn
        
    Next iRow
    
    GetNumbers = outputA
    
    #If useTimer Then
        With numberTimer
            .EndTask
            Debug.Print .ToString
        End With
    #End If
    
    Exit Function
    
IndexOutOfRange:
    Resume Next_Byte_Index
StoreStringInOutputA:
    'outputA(iRow, iColumn) = numberAsString
    Resume Next
End Function
Private Function IsCharCodeNumber(value As Byte) As Boolean
    Select Case value
        Case 48 To 57
            IsCharCodeNumber = True
    End Select
End Function
Public Function Change_Delimiter_Not_Between_Quotes(ByRef Current_String$, ByVal Delimiter$, Optional ByVal Changed_Delimiter$ = ">�") As Variant
    
    'returns a 0 based array
        
    Dim String_Array$(), x As Long, Right_CHR$

    If InStrB(1, Current_String, Chr(34)) = 0 Then 'if there are no quotation marks then split with the supplied delimiter
        Change_Delimiter_Not_Between_Quotes = Split(Current_String, Delimiter)
        Exit Function
    End If
    
    Right_CHR = Right(Changed_Delimiter, 1) 'RightMost character in at least 2 character string that will be used as a replacement delimiter

    'Replace ALL quotation marks with the ChangedDelimiter[Quotation mark] EX: " --> $+
    Current_String = Replace(Current_String, Chr(34), Changed_Delimiter)

    String_Array = Split(Current_String, Left(Changed_Delimiter, 1))
    '1st character of Changed_Delimiter will be used to delimit a new array
    'element [0] will be an empty string if the first value in the delmited string begins with a Quotation mark.
    
    For x = LBound(String_Array) To UBound(String_Array) 'loop all elements of the array

        If Left(String_Array(x), 1) = Right_CHR And Not Left(String_Array(x), 2) = Right_CHR & Delimiter Then
            'If the string contains a valid comma
            'Checked by if [the First character is the 2nd Character in the Changed Delimiter] and the 2nd character isn't the delimiter
            'Then offset the string by 1 character to remove the 2nd portion of the changed Delimiter
            String_Array(x) = Right(String_Array(x), Len(String_Array(x)) - 1)
        
        Else
            If Left(String_Array(x), 1) = Right_CHR Then 'If 1st character = 2nd portion of the Changed Delimiter
                                                         'Then offset string by 1 and then repalce all [Delimiter]
                String_Array(x) = Replace$(Right(String_Array(x), Len(String_Array(x)) - 1), Delimiter, Changed_Delimiter)
            
            Else 'Just replace
                String_Array(x) = Replace$(String_Array(x), Delimiter, Changed_Delimiter)
            End If
        End If
        
    Next x
    'Join the Array elements back together {Do not add another delimiter] and split with the changed Delimiter
    Change_Delimiter_Not_Between_Quotes = Split(Join(String_Array), Changed_Delimiter)
    
    Erase String_Array
End Function
Public Function entUnZip1File(ByVal strZipFilename As Variant, ByVal strDstDir As Variant, ByVal strFilename As Variant) 'Opens zip file
                                                'path of file     path of Folder containing file              name of specified file within .zip file
    On Error GoTo Propagate
    Const glngcCopyHereDisplayProgressBox = 256
    '
    Dim intOptions, objShell, objSource, objTarget As Object
    '
    ' Create the required Shell objects
    Set objShell = CreateObject("Shell.Application")
    '
    ' Create a reference to the files and folders in the ZIP file
    Set objSource = objShell.Namespace(strZipFilename).items.Item(strFilename)
    '
    ' Create a reference to the target folder
    Set objTarget = objShell.Namespace(strDstDir)
    '
    intOptions = glngcCopyHereDisplayProgressBox
    '
    ' UnZIP the files
    objTarget.CopyHere objSource, intOptions
    '
    ' Release the objects
    Set objSource = Nothing
    Set objTarget = Nothing
    Set objShell = Nothing
    '
    entUnZip1File = 1
    Exit Function
Propagate:
    PropagateError Err, "entUnZip1File"
End Function
Public Sub Quicksort(ByRef vArray As Variant, arrLbound As Long, arrUbound As Long)
    'Sorts a one-dimensional VBA array from smallest to largest
    'using a very fast quicksort algorithm variant.
    Dim pivotVal As Variant
    Dim vSwap    As Variant
    Dim Temporary_Low   As Long
    Dim Temporary_High    As Long
    
    Temporary_Low = arrLbound

    Temporary_High = arrUbound

    pivotVal = vArray((arrLbound + arrUbound) \ 2) 'The element in the middle of the array
    
    While (Temporary_Low <= Temporary_High) 'divide

    While (vArray(Temporary_Low) < pivotVal And Temporary_Low < arrUbound)
        
        Temporary_Low = Temporary_Low + 1
        
    Wend
    
    While (pivotVal < vArray(Temporary_High) And Temporary_High > arrLbound)
        Temporary_High = Temporary_High - 1
    Wend
    
    If (Temporary_Low <= Temporary_High) Then
        vSwap = vArray(Temporary_Low)
        vArray(Temporary_Low) = vArray(Temporary_High)
        vArray(Temporary_High) = vSwap
        Temporary_Low = Temporary_Low + 1
        Temporary_High = Temporary_High - 1
    End If
    
    Wend
 
  If (arrLbound < Temporary_High) Then Quicksort vArray, arrLbound, Temporary_High 'conquer
  If (Temporary_Low < arrUbound) Then Quicksort vArray, Temporary_Low, arrUbound 'conquer
  
End Sub
Public Sub DownloadFile(fileUrl$, SaveFilePathAndName$)

    Dim WinHttpReq As Object
    
    On Error GoTo Propogate
    Set WinHttpReq = CreateObject("Msxml2.ServerXMLHTTP")
    
    With WinHttpReq
        .Open "GET", fileUrl, False
        .send
        If .Status = 200 Then
            With CreateObject("ADODB.Stream")
                .Open
                ' 1 for Binary data, 2 for text
                .Type = 1
                .write WinHttpReq.responseBody
                .SaveToFile SaveFilePathAndName, 2 ' 1 = no overwrite, 2 = overwrite
                .Close
            End With
        End If
    End With
    
'AppleScript:
'set u to "http://download.finance.yahoo.com/d/quotes.csv?s=AAPL&f=sl1d1t1c1ohgv&e=.csv"
'do shell script "curl -L -s " & File & " > ~/desktop/quotes.csv"
    Exit Sub
Propogate:
    PropagateError Err, "DownloadFile"
End Sub
Public Function CombineArraysInCollection(My_CLCTN As Collection, howToCombine As Append_Type) As Variant 'adds the contents of the NEW array TO the contents of the OLD
  
'===================================================================================================================
    'Purpose: Combines multiple 1D or 2D arrays.
    'Inputs:   My_CLCTN - Collection object that contains arrays to combine.
    '          howToCombine - An enum to tell the function what sort of combination to do.
    'Outputs: A 2D array of combined data.
'===================================================================================================================
    Dim finalColumnIndex As Long, x As Long, finalRowIndex As Long, UB1 As Long, UB2 As Long, Worksheet_Data() As Variant, _
    Item As Variant, Old() As Variant, Block() As Variant, Latest() As Variant, Not_Old As Byte, Is_Old As Byte
       
    'Dim Addition_Timer As Double: Addition_Timer = Time
    On Error GoTo Propagate
    With My_CLCTN
        'check the boundaries of the elements to create the array
        Select Case howToCombine
    
            Case Append_Type.Multiple_1d 'Array Elements are 1D | single rows |  "Historical_Parse"
    
                UB1 = .count 'The number of items in the dictionary will be the number of rows in the final array
    
                For Each Item In My_CLCTN 'loop through each item in the row and find the max number of columns
                    finalRowIndex = UBound(Item) + 1 - LBound(Item) 'Number of Columns if 1 based
                    If finalRowIndex > UB2 Then UB2 = finalRowIndex
                Next Item
                
            Case Append_Type.Multiple_2d
                'Indeterminate number of  2D[1-Based] arrays to be joined.
                For Each Item In My_CLCTN
                    UB1 = UBound(Item, 1) + UB1
                    finalRowIndex = UBound(Item, 2)
                    If finalRowIndex > UB2 Then UB2 = finalRowIndex
                Next Item
    
            Case Append_Type.Add_To_Old
                
                Not_Old = 1
                
                Do Until .Item(Not_Old)(0) <> ArrayID.Old_Data
                    Not_Old = Not_Old + 1
                Loop
                
                Is_Old = IIf(Not_Old = 1, 2, 1)
                Old = .Item(Is_Old)(1)
                finalRowIndex = UBound(Old, 2)
                
                Select Case .Item(Not_Old)(0)         'Number designating array type
                
                    Case ArrayID.Weekly_Data  'This key is used for when sotring weekly data
                    
                        Latest = .Item(Not_Old)(1)
                        finalColumnIndex = UBound(Latest)              'Number of columns in the 1-based 1D array
                        UB1 = UBound(Old, 1) + 1 ' +1 Since there will be only 1 row of additional weekly data
                    
                    Case ArrayID.Block_Data  'This key is used if several weeks have passed
                                                            'This will be a 2d array
                        Block = .Item(Not_Old)(1)
                        finalColumnIndex = UBound(Block, 2)
                        UB1 = UBound(Old, 1) + UBound(Block, 1)
                    
                End Select
                
                If finalRowIndex >= finalColumnIndex Then 'Determing number of columns to size the array with
                    UB2 = finalRowIndex    'S= # of Columns in the older data
                Else           'T= # of Columns in the new data
                    UB2 = finalColumnIndex
                End If
    
        End Select
        
        ReDim Worksheet_Data(1 To UB1, 1 To UB2)
        
        finalRowIndex = 1
        
        For Each Item In My_CLCTN
            
            Select Case howToCombine
    
                Case Append_Type.Multiple_1d 'All items in Collection are 1D
                    
                    For finalColumnIndex = LBound(Item) To UBound(Item)
                        
                        Worksheet_Data(finalRowIndex, finalColumnIndex + 1 - LBound(Item)) = Item(finalColumnIndex)
    
                    Next finalColumnIndex
    
                    finalRowIndex = finalRowIndex + 1
                    
                Case Append_Type.Multiple_2d 'Adding Multiple 2D arrays together
        
                        x = 1
                        
                        For finalRowIndex = finalRowIndex To UBound(Item, 1) + (finalRowIndex - 1)
                            For finalColumnIndex = LBound(Item, 2) To UBound(Item, 2)
                                Worksheet_Data(finalRowIndex, finalColumnIndex) = Item(x, finalColumnIndex)
                            Next finalColumnIndex
                            x = x + 1
                        Next finalRowIndex
                
                Case Append_Type.Add_To_Old 'Adding new Data to a 2D Array..Block is 2D...Latest is 1D
                                            
                    Select Case Item(0)                 'Key of item
    
                        Case ArrayID.Old_Data 'Current Historical data on Worksheet
                            
                            For finalRowIndex = LBound(Worksheet_Data, 1) To UBound(Old, 1)
                                For finalColumnIndex = LBound(Old, 2) To UBound(Old, 2)
                                    Worksheet_Data(finalRowIndex, finalColumnIndex) = Old(finalRowIndex, finalColumnIndex)
                                Next finalColumnIndex
                            Next finalRowIndex
                            
                        Case ArrayID.Block_Data '<--2D Array used when adding to arrays together where order is important
                        
                            x = 1
                            
                            For finalRowIndex = UBound(Worksheet_Data, 1) - UBound(Block, 1) + 1 To UBound(Worksheet_Data, 1)
                                For finalColumnIndex = LBound(Block, 2) To UBound(Block, 2)
                                    Worksheet_Data(finalRowIndex, finalColumnIndex) = Block(x, finalColumnIndex)
                                Next finalColumnIndex
                                x = x + 1
                            Next finalRowIndex
                            
                        Case ArrayID.Weekly_Data  '1 based 1D "WEEKLY" array
                                          '"OLD" is run first so S is already at the correct incremented value
                            finalRowIndex = UBound(Worksheet_Data, 1)
                            
                            For finalColumnIndex = LBound(Latest) To UBound(Latest)
                                Worksheet_Data(finalRowIndex, finalColumnIndex) = Latest(finalColumnIndex) 'worksheet data is 1 based 2D while Item is 1 BASED 1D
                            Next finalColumnIndex
                                          
                    End Select
    
            End Select
            
        Next Item
    
    End With
    
    CombineArraysInCollection = Worksheet_Data
    Exit Function
Propagate:
    PropagateError Err, "CombineArraysInCollection"

End Function

Public Function IsArrayAllocated(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayAllocated
' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
' sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
' been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
' allocated.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is just the reverse of IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim N As Long, isAllocated As Boolean
    On Error GoTo Exit_Procedure
    
    ' if Arr is not an array, return FALSE and get out.
    If IsArray(Arr) = False Then
        IsArrayAllocated = False
        Exit Function
    End If
    
    ' Attempt to get the UBound of the array. If the array has not been allocated,
    ' an error will occur. Test Err.Number to see if an error occurred.
    N = UBound(Arr, 1)

    ''''''''''''''''''''''''''''''''''''''
    ' Under some circumstances, if an array
    ' is not allocated, Err.Number will be
    ' 0. To acccomodate this case, we test
    ' whether LBound <= Ubound. If this
    ' is True, the array is allocated. Otherwise,
    ' the array is not allocated.
    '''''''''''''''''''''''''''''''''''''''
    ' no error. array has been allocated.
    isAllocated = (LBound(Arr) <= UBound(Arr))

    IsArrayAllocated = isAllocated
Exit_Procedure:
End Function
Public Function Reverse_2D_Array(ByVal data As Variant, Optional ByRef selected_columns As Variant) As Variant

    Dim x As Long, Y As Long, temp(1 To 2) As Variant, Projected_Row As Long
    
    Dim LB2 As Byte, UB2 As Long, Z As Long

    If isMissing(selected_columns) Then
        LB2 = LBound(data, 2)
        UB2 = UBound(data, 2)
    Else
        LB2 = LBound(selected_columns)
        UB2 = UBound(selected_columns)
    End If
    
    For x = LBound(data, 1) To UBound(data, 1)
            
        Projected_Row = UBound(data, 1) - (x - LBound(data, 1))
        
        If Projected_Row <= x Then Exit For
        
        For Y = LB2 To UB2
            
            If isMissing(selected_columns) Then
                Z = Y
            Else
                Z = selected_columns(Y)
            End If
            
            temp(1) = data(x, Z)
            temp(2) = data(Projected_Row, Z)
            
            data(x, Z) = temp(2)
            data(Projected_Row, Z) = temp(1)
            
        Next Y

    Next x

    Reverse_2D_Array = data

End Function
Public Function TransposeData(ByRef inputA As Variant, Optional convertNullToZero As Boolean = False) As Variant()
'===================================================================================================================
    'Purpose: Transposes the inputted inputA array.
    'Inputs: inputA - Array to transpose.
    '        convertNullToZero - If true then null values will be converted to 0.
    'Outputs: A transposed 2D array.
'===================================================================================================================
    Dim iRow As Long, iColumn As Byte, output() As Variant, baseZeroAddition As Byte
    
    On Error GoTo Propogate
    
    If LBound(inputA, 2) = 0 Then baseZeroAddition = 1
    
    ReDim output(1 To UBound(inputA, 2) + baseZeroAddition, 1 To UBound(inputA, 1) + baseZeroAddition)
    
    For iColumn = LBound(inputA, 1) To UBound(inputA, 1)
        For iRow = LBound(inputA, 2) To UBound(inputA, 2)
        
            If Not convertNullToZero Or Not IsNull(inputA(iColumn, iRow)) Then
                output(iRow + baseZeroAddition, iColumn + baseZeroAddition) = inputA(iColumn, iRow)
            Else
                output(iRow + baseZeroAddition, iColumn + baseZeroAddition) = 0
            End If
                    
        Next iRow
    Next iColumn
    
    TransposeData = output
    Exit Function
Propogate:
    PropagateError Err, "TransposeData"
End Function
Public Sub PropagateError(e As ErrObject, procedureName$, Optional moreDetails$ = vbNullString)
    
    Dim firstPropagation As Boolean, sourceParts$()
    
    Const delim$ = ": "
    
    With e
        AddParentToErrSource e, procedureName
        If LenB(moreDetails) > 0 Then .Description = moreDetails & vbNewLine & .Description
        .Raise .Number, .source, .Description
    End With
    
End Sub
Public Sub DisplayErr(errorToDisplay As ErrObject, methodName$, Optional descriptionAddenum$ = vbNullString)
    
    Dim sourceParts$(), message$
    Const delim$ = ": "

    With HoldError(errorToDisplay)
        On Error GoTo ErrorDisplayFailed
        With .HeldError
            If .Number <> 0 Then
                
                AddParentToErrSource Err, methodName
                
                If LenB(descriptionAddenum) > 0 Then .Description = descriptionAddenum & vbNewLine & .Description
                
                sourceParts = Split(.source, delim, 2)
                
                message = "An error occured." & vbNewLine & _
                "Description: " & .Description & vbNewLine & _
                "Number: " & .Number & vbNewLine & _
                "Path: " & vbNewLine & sourceParts(UBound(sourceParts)) & vbNewLine & vbNewLine & _
                "Contact email: MoshiM_UC@outlook.com"
                
                MsgBox message, Title:=sourceParts(0) & " Error Message."
                Debug.Print message
            End If
        End With
    End With
    
    Exit Sub
ErrorDisplayFailed:
    MsgBox "Error in DisplayErr(). " & Err.Description
End Sub
Public Sub AddParentToErrSource(e As ErrObject, parentName$)
    
    Dim sourceParts$()
    Const delim$ = ": "
    
    With e
    
        If InStrB(1, .source, delim) = 0 Then
            .source = "[" & .source & "]" & delim & parentName
        Else
            sourceParts = Split(.source, delim, 2)
            sourceParts(1) = parentName & " > " & sourceParts(1)
            .source = Join(sourceParts, delim)
        End If
                    
    End With
    
End Sub
Public Function HoldError(e As ErrObject) As StoredError
'======================================================================================
'Summary:Stores details of a given error.
'Intended use: Allows you to propogate an error and engage additional error handling if needed.
'======================================================================================
    Dim errorToStore As New StoredError
    
    errorToStore.Constructor e
    Set HoldError = errorToStore
    
End Function
Public Function FileOrFolderExists(FileOrFolderstr$) As Boolean
'Ron de Bruin : 1-Feb-2019
'Function to test whether a file or folder exist on a Mac in office 2011 and up
'Uses AppleScript to avoid the problem with long names in Office 2011,
'limit is max 32 characters including the extension in 2011.
    Dim ScriptToCheckFileFolder$
    Dim TestStr$
    
    If FileOrFolderstr = vbNullString Then Exit Function
    
    #If Not Mac Then
        FileOrFolderExists = LenB(Dir$(FileOrFolderstr, vbDirectory)) > 0
        Exit Function
    #End If
    
    If val(Application.Version) < 15 Then
        ScriptToCheckFileFolder = "tell application " & QuotedForm("System Events") & _
        "to return exists disk item (" & QuotedForm(FileOrFolderstr) & " as string)"
        FileOrFolderExists = MacScript(ScriptToCheckFileFolder)
    Else
        On Error Resume Next
        TestStr = Dir$(FileOrFolderstr & "*", vbDirectory)
        On Error GoTo 0
        If LenB(TestStr) > 0 Then FileOrFolderExists = True
    End If

End Function
Public Function QuotedForm(ByRef Item, Optional Enclosing_CHR$ = """") As Variant

    Dim Z As Long, subArray As Variant, subArrayIndex As Byte
    
    If IsArray(Item) Then
    
        For Z = LBound(Item) To UBound(Item)
        
            If IsArray(Item(Z)) Then
            
                subArray = Item(Z)
                
                For subArrayIndex = LBound(subArray) To UBound(subArray)
                    If Not subArray(subArrayIndex) Like Enclosing_CHR & "*" & Enclosing_CHR Then
                        subArray(subArrayIndex) = Enclosing_CHR & subArray(subArrayIndex) & Enclosing_CHR
                    End If
                Next subArrayIndex
                    
                Item(Z) = subArray
                
            Else
                If Not Item(Z) Like Enclosing_CHR & "*" & Enclosing_CHR Then Item(Z) = Enclosing_CHR & Item(Z) & Enclosing_CHR
            End If
                     
        Next Z
        
    Else
        If Not Item Like Enclosing_CHR & "*" & Enclosing_CHR Then Item = Enclosing_CHR & Item & Enclosing_CHR
    End If
    
    QuotedForm = Item
    
End Function
Function HttpGet(sUrl As String, ByRef success As Boolean) As String
    
    Dim onMac As Boolean
    
    On Error GoTo Failure
    
    #If Mac Then
        Dim shellCMD$, lExitCode&
        onMac = True
        shellCMD = "curl " & Chr(34) & sUrl & Chr(34)
        HttpGet = ExecuteShellCommandMAC(shellCMD, lExitCode)
        success = (lExitCode = 0)
    #Else
        With CreateObject("MSXML2.XMLHTTP")
            .Open "GET", sUrl, False
            '.setRequestHeader "Content-Type", "text/html; charset=utf-8"
            .setRequestHeader "Accept-Language", "en-US,en;q=0.9"
            .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0"
            .send
            If .Status = 200 Then
                HttpGet = .responseText
                success = True
            End If
        End With
    #End If

    Exit Function
    
Failure:
    PropagateError Err, "HttpGet", "OS: " & IIf(onMac, "Using MAC", "Windows")
End Function

Public Function HttpPost(url$, postData$, Optional urlEncoded As Boolean) As Boolean
    
    Dim onMac As Boolean
    
    On Error GoTo Failure
    
    #If Mac Then
        Dim shellCommand$, lExitCode&: onMac = True
        shellCommand = "curl -X POST -d """ & postData & """ """ & url & Chr(34)
        ExecuteShellCommandMAC shellCommand, lExitCode
        HttpPost = (lExitCode = 0)
    #Else
        'Shell "curl --fail -X POST -d """ & postData & """ """ & url & Chr(34)
        With CreateObject("MSXML2.XMLHTTP")
            .Open "POST", url, False
            If urlEncoded Then .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .setRequestHeader "Content-Length", Len(postData)
            .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0"
            .send postData
            HttpPost = (.Status = 200)
        End With
    #End If
    
    Exit Function
Failure:
    PropagateError Err, "HttpPost", "OS: " & IIf(onMac, "Using MAC", "Windows")
End Function

Public Function Parse_Json_String(query$) As Object
    Dim JP As New JsonParserB
    Set Parse_Json_String = JP.ParseJsonString(query)
End Function

Public Function Parse_Json_Bytes(jsonBytes() As Byte) As Object
    Dim JP As New JsonParserB
    Set Parse_Json_Bytes = JP.ParseJsonBytes(jsonBytes)
End Function
