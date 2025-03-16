Attribute VB_Name = "General"
#If Not Mac Then
    #If VBA7 Then
        Public Declare PtrSafe Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (tzInfo As TimeZoneInformation, localTime As SystemTime, returnedTimeUTC As SystemTime) As LongPtr
        Public Declare PtrSafe Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (tzInfo As TimeZoneInformation, localTime As SystemTime, returnedLocalTime As SystemTime) As LongPtr
        Public Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (tzInfo As TimeZoneInformation) As LongPtr
        Public Declare PtrSafe Function GetSystemTime Lib "kernel32" (returnedUTC As SystemTime) As LongPtr
    #Else
        Public Declare Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (tzInfo As TimeZoneInformation, localTime As SystemTime, returnedTimeUTC As SystemTime) As Long
        Public Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (tzInfo As TimeZoneInformation, localTime As SystemTime, returnedLocalTimeUTC As SystemTime) As Long
        Public Declare Function GetTimeZoneInformation Lib "kernel32" (tzInfo As TimeZoneInformation) As Long
        Public Declare Function GetSystemTime Lib "kernel32" (returnedUTC As SystemTime) As Long
    #End If
#End If

Public Type SystemTime
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type TimeZoneInformation
    Bias As Long 'The bias is the difference, in minutes, between Coordinated Universal Time (UTC) and local time.
    standardName(31) As Integer
    StandardDate As SystemTime
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SystemTime
    DaylightBias As Long
    'TimeZoneKeyName(127) As Integer
    'DynamicDaylightTimeDisabled As Boolean
End Type
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
    
    Dim wbSaved As Boolean
    
    On Error Resume Next
    
    With ThisWorkbook
        wbSaved = .Saved
        HUB.Shapes("Macro_Check").Visible = False
        .Saved = wbSaved
    End With
    
    On Error GoTo 0
    
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
Attribute Hide_Workbooks.VB_Description = "Hides all workbooks except the currently active one."
Attribute Hide_Workbooks.VB_ProcData.VB_Invoke_Func = " \n14"
'===================================================================================================================
    'Summary: Hides all workbooks except for the currently active one.
'===================================================================================================================
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If Not wb Is ActiveWorkbook Then wb.Windows(1).Visible = False
    Next wb

End Sub
Public Sub Show_Workbooks()
Attribute Show_Workbooks.VB_Description = "Unhides hidden workbooks."
Attribute Show_Workbooks.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim wb As Workbook

    For Each wb In Application.Workbooks
        wb.Windows(1).Visible = True
    Next wb

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

    Set Worksheet_TB = TBL_RNG.Parent 'Worksheet where table is found
    
    With Worksheet_TB '{Must be typed as object to fool the compiler when resetting the Used Range]

        With TBL_RNG 'Find the Bottom Right cell of the table
            Set TB_Last_Cell = .Cells(.Rows.Count, .columns.Count)
        End With
        
        With .UsedRange 'Find the Bottom right cell of the Used Range
            Set UR_LastCell = .Cells(.Rows.Count, .columns.Count)
        End With
        
        If Intersect(UR_LastCell, TB_Last_Cell) Is Nothing Then
        
            'If UR_LastCell AND TB_Last_Cell don't refer to the same cell
            
            With TB_Last_Cell
                Set LRO = .Offset(1, 0) 'last row of table offset by 1
                Set LCO = .Offset(0, 1) 'last column of table offset by 1
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

Sub Remove_Worksheet_Formatting()
Attribute Remove_Worksheet_Formatting.VB_Description = "Removes all conditional formatting from the active worksheet."
Attribute Remove_Worksheet_Formatting.VB_ProcData.VB_Invoke_Func = " \n14"
'===================================================================================================================
    'Summary: Deletes conditional formatting from the currently active worksheet.
    'Outputs: Stores the current time on the Variable_Sheet along with the local time on the running environment.
    'Note: Keyboard shortcut > CTRL+SHIFT+X
'===================================================================================================================
    ThisWorkbook.ActiveSheet.Cells.FormatConditions.Delete
End Sub
Sub ZoomToRange(ByRef ZoomThisRange As Range, ByVal PreserveRows As Boolean, wb As Workbook)

    Application.ScreenUpdating = False

    Dim Wind As Window

    Set Wind = ActiveWindow

    Application.GoTo ZoomThisRange.Cells(1, 1), True

    With ZoomThisRange
        If PreserveRows = True Then
            .Resize(.Rows.Count, 1).Select
        Else
            .Resize(1, .columns.Count).Select
        End If
    End With

    With Wind
        .Zoom = True
        .VisibleRange.Cells(1, 1).Select
    End With

    If Not wb.ActiveSheetBeforeSaving Is Nothing And IsCreatorActiveUser Then 'accounting for if the variable has not been declared for normal use
        'do nothing
    Else
        Application.ScreenUpdating = True
    End If

End Sub
Public Sub ClearRegionBeneathTable(dataTable As ListObject)
'===================================================================================================================
   'Summary: Clears the region underneath [dataTable]
   'Inputs:
   '    dataTable - ListObject that will have the area underneath it cleared.
'===================================================================================================================
    Dim lastUsedRowNumber As Long, appProperties As Collection

    With dataTable.Parent.UsedRange
        lastUsedRowNumber = .Cells(.Rows.Count, 1).row
    End With
    
    'Clear rows below the table

    With dataTable.DataBodyRange
        With .Rows(.Rows.Count).Offset(1)
            If lastUsedRowNumber >= .row Then
                Set appProperties = DisableApplicationProperties(True, False, True)
                With .Resize(lastUsedRowNumber - .row + 1, .columns.Count)
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
    'Summary: Loads filters into filterArray and clears from queriedTable.
    'Inputs: queriedTable - ListObject that will have its filters removed and stored.
    '        filterArray -  Array that will store removed filters.
'===================================================================================================================
    Dim F As Long, tableFilter As Filter

    With queriedTable.AutoFilter
        On Error GoTo Show_Data
        ReDim filterArray(1 To .Filters.Count, 1 To 3)
        
        For Each tableFilter In .Filters
            With tableFilter
                F = F + 1
                If .On Then
                    filterArray(F, 1) = .Criteria1
                    If .Operator Then
                        filterArray(F, 2) = .Operator
                        filterArray(F, 3) = .Criteria2
                    End If
                End If
            End With
        Next
Show_Data:
        .ShowAllData
    End With

End Sub
Public Sub RestoreFilters(tableOBJ As ListObject, ByRef filterArray)
'===================================================================================================================
    'Summary: uses filterArray to reapply filters.
    'Inputs: tableOBJ - ListObject that has filters applied to it.
    '        filterArray - array generated from ChangeFilters().
'===================================================================================================================
    Dim col As Long

    With tableOBJ.DataBodyRange
        For col = LBound(filterArray, 1) To UBound(filterArray, 1)
            If Not IsEmpty(filterArray(col, 1)) Then
                If filterArray(col, 2) Then
                    .AutoFilter Field:=col, _
                        Criteria1:=filterArray(col, 1), _
                        Operator:=filterArray(col, 2), _
                        Criteria2:=filterArray(col, 3)
                Else
                    .AutoFilter Field:=col, Criteria1:=filterArray(col, 1)
                End If
            End If
        Next col
    End With

End Sub
Private Sub ConvertAllNamedRangesToWorksheetScopeOnWorksheet()
'=============================================================================================
'   Scopes all named ranges on the current sheet to the worksheet rather than to the workbook.
'=============================================================================================
    Dim nm As name, workbookActiveSheet As Worksheet, rangeNameRefersTo$, nameOfRange$
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
            If .RefersToRange.Parent Is workbookActiveSheet And InStrB(1, nm.name, workbookActiveSheet.name & "!") <> 1 Then
                rangeNameRefersTo = .RefersTo
                nameOfRange = .name
                .Delete
                workbookActiveSheet.names.Add workbookActiveSheet.name & "!" & nameOfRange, rangeNameRefersTo
            End If
        End With
Attempt_Next_Name:
        On Error GoTo -1
    Next nm
    
    Re_Enable
    
End Sub

Public Sub Run_This(wb As Workbook, ScriptN$)
    On Error GoTo Propagate
    Application.Run "'" & wb.name & "'!" & ScriptN
    Exit Sub
Propagate:
    PropagateError Err, "Run_This"
End Sub

Public Function DisableApplicationProperties(disableEvents As Boolean, disableAutoCalculations As Boolean, disableScreenUpdating As Boolean) As Collection
    
    Dim applicationProperties As New Collection
    
    With Application
        
        If disableEvents Then applicationProperties.Add .EnableEvents, "Events": .EnableEvents = False
        
        If disableAutoCalculations Then applicationProperties.Add .Calculation, "Calc": .Calculation = xlCalculationManual
        
        If disableScreenUpdating Then applicationProperties.Add .ScreenUpdating, "Screen": .ScreenUpdating = False
                        
    End With
    
    Set DisableApplicationProperties = applicationProperties
    
End Function
Public Sub EnableApplicationProperties(values As Collection)
    
    Dim savedError As StoredError
    
    If Not values Is Nothing Then
        If values.Count > 0 Then
            If Err.Number <> 0 Then Set savedError = HoldError(Err)
            With Application
                On Error Resume Next
                .Calculation = values("Calc")
                .EnableEvents = values("Events")
                .ScreenUpdating = values("Screen")
                If Err.Number <> 0 Then Err.Clear
            End With
            If Not savedError Is Nothing Then Err = savedError.HeldError
        End If
    End If
    
End Sub
Public Sub ResizeTableBasedOnColumn(LO As ListObject, columnToMatchLastUsedRow As Range)
'====================================================================================================================================
'   Summary: Resizes a Listobject so that its last row is in the same row as the last used row in the column represented bycolumnToMatchLastUsedRow
'====================================================================================================================================
    Dim bottomInColumn As Range, isCellEmpty As Boolean, newBottom As Range, _
    rowsToKeepCount As Long, worksheetWithTable As Worksheet, givenTableColumn As Boolean
    
    If columnToMatchLastUsedRow.columns.Count > 1 Then
        MsgBox "Input range has should have only a single column"
        Exit Sub
    End If
    
    With LO
        Set worksheetWithTable = .Parent
        With .DataBodyRange
            If Intersect(LO.DataBodyRange, columnToMatchLastUsedRow) Is Nothing Then
                Set bottomInColumn = worksheetWithTable.Cells(worksheetWithTable.Rows.Count, columnToMatchLastUsedRow.Column).End(xlUp)
            Else
                Set bottomInColumn = .Cells(.Rows.Count, columnToMatchLastUsedRow.Column - .Column + 1)
                givenTableColumn = True
            End If
            
            isCellEmpty = IsEmpty(bottomInColumn.Value2)
            
            If ((givenTableColumn And isCellEmpty) Or (Not givenTableColumn And Not isCellEmpty)) And bottomInColumn.row > LO.Range.row Then
                Set newBottom = IIf(givenTableColumn, bottomInColumn.End(xlUp), bottomInColumn)
            
                If Not newBottom.row = LO.Range.row And Not newBottom.row = .Rows(.Rows.Count).row Then
                    rowsToKeepCount = newBottom.row - LO.Range.row + 1
                    LO.Resize LO.Range.Resize(rowsToKeepCount, .columns.Count)
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
    Dim data() As Variant, iColumn As Long, iRow As Long, outputA() As Variant, stringBytes() As Byte, _
    byteIndex As Long, currentNumberBytes() As Byte, validCharacter As Boolean, cursorIndex As Long, _
    numberAsString$, skipCharacter As Boolean, inputIsRange As Boolean, startLocation As Long
    
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
    Set objSource = objShell.Namespace(strZipFilename).Items.item(strFilename)
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
    PropagateError Err, "DownloadFile", "Failed to download " & fileUrl & " and save to " & SaveFilePathAndName
End Sub
Public Function CombineArraysInCollection(My_CLCTN As Collection, howToCombine As Append_Type) As Variant 'adds the contents of the NEW array TO the contents of the OLD
  
'===================================================================================================================
    'Summary: Combines multiple 1D or 2D arrays.
    'Inputs:   My_CLCTN - Collection object that contains arrays to combine.
    '          howToCombine - An enum to tell the function what sort of combination to do.
    'Returns: A 2D array of combined data.
'===================================================================================================================
    Dim finalColumnIndex As Long, X As Long, finalRowIndex As Long, UB1 As Long, UB2 As Long, Worksheet_Data() As Variant, _
    item As Variant, Old() As Variant, Block() As Variant, Latest() As Variant, Not_Old As Long, Is_Old As Long
       
    'Dim Addition_Timer As Double: Addition_Timer = Time
    On Error GoTo Propagate
    With My_CLCTN
        'check the boundaries of the elements to create the array
        Select Case howToCombine
    
            Case Append_Type.Multiple_1d 'Array Elements are 1D | single rows |  "Historical_Parse"
    
                UB1 = .Count 'The number of items in the dictionary will be the number of rows in the final array
    
                For Each item In My_CLCTN 'loop through each item in the row and find the max number of columns
                    finalRowIndex = UBound(item) + 1 - LBound(item) 'Number of Columns if 1 based
                    If finalRowIndex > UB2 Then UB2 = finalRowIndex
                Next item
                
            Case Append_Type.Multiple_2d
                'Indeterminate number of  2D[1-Based] arrays to be joined.
                For Each item In My_CLCTN
                    UB1 = UBound(item, 1) + UB1
                    finalRowIndex = UBound(item, 2)
                    If finalRowIndex > UB2 Then UB2 = finalRowIndex
                Next item
    
            Case Append_Type.Add_To_Old
                
                Not_Old = 1
                
                Do Until .item(Not_Old)(0) <> ArrayID.Old_Data
                    Not_Old = Not_Old + 1
                Loop
                
                Is_Old = IIf(Not_Old = 1, 2, 1)
                Old = .item(Is_Old)(1)
                finalRowIndex = UBound(Old, 2)
                
                Select Case .item(Not_Old)(0)         'Number designating array type
                
                    Case ArrayID.Weekly_Data  'This key is used for when sotring weekly data
                    
                        Latest = .item(Not_Old)(1)
                        finalColumnIndex = UBound(Latest)              'Number of columns in the 1-based 1D array
                        UB1 = UBound(Old, 1) + 1 ' +1 Since there will be only 1 row of additional weekly data
                    
                    Case ArrayID.Block_Data  'This key is used if several weeks have passed
                                                            'This will be a 2d array
                        Block = .item(Not_Old)(1)
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
        
        For Each item In My_CLCTN
            Select Case howToCombine
                Case Append_Type.Multiple_1d 'All items in Collection are 1D
                    
                    For finalColumnIndex = LBound(item) To UBound(item)
                        Worksheet_Data(finalRowIndex, finalColumnIndex + 1 - LBound(item)) = item(finalColumnIndex)
                    Next finalColumnIndex
    
                    finalRowIndex = finalRowIndex + 1
                    
                Case Append_Type.Multiple_2d 'Adding Multiple 2D arrays together
        
                        X = 1
                        
                        For finalRowIndex = finalRowIndex To UBound(item, 1) + (finalRowIndex - 1)
                            For finalColumnIndex = LBound(item, 2) To UBound(item, 2)
                                Worksheet_Data(finalRowIndex, finalColumnIndex) = item(X, finalColumnIndex)
                            Next finalColumnIndex
                            X = X + 1
                        Next finalRowIndex
                
                Case Append_Type.Add_To_Old 'Adding new Data to a 2D Array.Block is 2D..Latest is 1D
                                            
                    Select Case item(0)                 'Key of item
    
                        Case ArrayID.Old_Data 'Current Historical data on Worksheet
                            
                            For finalRowIndex = LBound(Worksheet_Data, 1) To UBound(Old, 1)
                                For finalColumnIndex = LBound(Old, 2) To UBound(Old, 2)
                                    Worksheet_Data(finalRowIndex, finalColumnIndex) = Old(finalRowIndex, finalColumnIndex)
                                Next finalColumnIndex
                            Next finalRowIndex
                            
                        Case ArrayID.Block_Data '<--2D Array used when adding to arrays together where order is important
                        
                            X = 1
                            
                            For finalRowIndex = UBound(Worksheet_Data, 1) - UBound(Block, 1) + 1 To UBound(Worksheet_Data, 1)
                                For finalColumnIndex = LBound(Block, 2) To UBound(Block, 2)
                                    Worksheet_Data(finalRowIndex, finalColumnIndex) = Block(X, finalColumnIndex)
                                Next finalColumnIndex
                                X = X + 1
                            Next finalRowIndex
                            
                        Case ArrayID.Weekly_Data  '1 based 1D "WEEKLY" array
                                          '"OLD" is run first so S is already at the correct incremented value
                            finalRowIndex = UBound(Worksheet_Data, 1)
                            
                            For finalColumnIndex = LBound(Latest) To UBound(Latest)
                                Worksheet_Data(finalRowIndex, finalColumnIndex) = Latest(finalColumnIndex) 'worksheet data is 1 based 2D while Item is 1 BASED 1D
                            Next finalColumnIndex
                    End Select
            End Select
        Next item
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

    Dim n As Long
    On Error GoTo Exit_Procedure
    
    ' if Arr is not an array, return FALSE and get out.
    If IsArray(Arr) = False Then
        IsArrayAllocated = False
        Exit Function
    End If
    
    ' Attempt to get the UBound of the array. If the array has not been allocated,
    ' an error will occur. Test Err.Number to see if an error occurred.
    n = UBound(Arr, 1)

    ''''''''''''''''''''''''''''''''''''''
    ' Under some circumstances, if an array
    ' is not allocated, Err.Number will be
    ' 0. To acccomodate this case, we test
    ' whether LBound <= Ubound. If this
    ' is True, the array is allocated. Otherwise,
    ' the array is not allocated.
    '''''''''''''''''''''''''''''''''''''''
    ' no error. array has been allocated.
    IsArrayAllocated = (LBound(Arr) <= UBound(Arr))
    
Exit_Procedure:
End Function
Public Function Reverse_2D_Array(ByVal data As Variant, Optional ByRef selected_columns As Variant) As Variant

    Dim X As Long, Y As Long, temp(1 To 2) As Variant, Projected_Row As Long
    
    Dim LB2 As Long, UB2 As Long, Z As Long

    If IsMissing(selected_columns) Then
        LB2 = LBound(data, 2)
        UB2 = UBound(data, 2)
    Else
        LB2 = LBound(selected_columns)
        UB2 = UBound(selected_columns)
    End If
    
    For X = LBound(data, 1) To UBound(data, 1)
            
        Projected_Row = UBound(data, 1) - (X - LBound(data, 1))
        
        If Projected_Row <= X Then Exit For
        
        For Y = LB2 To UB2
            
            If IsMissing(selected_columns) Then
                Z = Y
            Else
                Z = selected_columns(Y)
            End If
            
            temp(1) = data(X, Z)
            temp(2) = data(Projected_Row, Z)
            
            data(X, Z) = temp(2)
            data(Projected_Row, Z) = temp(1)
            
        Next Y

    Next X

    Reverse_2D_Array = data

End Function
Public Function TransposeData(ByRef inputA As Variant, Optional convertNullToZero As Boolean = False) As Variant()
'===================================================================================================================
    'Summary: Transposes the inputted inputA array.
    'Inputs: inputA - Array to transpose.
    '        convertNullToZero - If true then null values will be converted to 0.
    'Returns: A transposed 2D array.
'===================================================================================================================
    Dim iRow&, iColumn&, output() As Variant, baseZeroAddition As Long
    
    On Error GoTo Propogate
    
    If LBound(inputA, 2) = 0 Then baseZeroAddition = 1
    
    ReDim output(1 To UBound(inputA, 2) + baseZeroAddition, 1 To UBound(inputA, 1) + baseZeroAddition)
    
    For iColumn = LBound(inputA, 1) To UBound(inputA, 1)
        For iRow = LBound(inputA, 2) To UBound(inputA, 2)
            
            If IsNull(inputA(iColumn, iRow)) Then
                If convertNullToZero Then output(iRow + baseZeroAddition, iColumn + baseZeroAddition) = 0
            Else
                output(iRow + baseZeroAddition, iColumn + baseZeroAddition) = inputA(iColumn, iRow)
            End If
                    
        Next iRow
    Next iColumn
    
    TransposeData = output
    Exit Function
Propogate:
    PropagateError Err, "TransposeData"
End Function
Public Sub PropagateError(e As ErrObject, callingProcedureName$, Optional moreDetails$ = vbNullString)
'===================================================================================================================
'Summary: Propagates an error.
'Inputs:
'   e - Error object.
'   callingProcedureName - Name of procedure that is propagating the error.
'   moreDetails - Optional description addenum.
'===================================================================================================================
    With e
        AddParentToErrSource e, callingProcedureName
        If LenB(moreDetails) <> 0 Then AppendErrorDescription e, moreDetails
        .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
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
                
                If LenB(descriptionAddenum) <> 0 Then AppendErrorDescription errorToDisplay, descriptionAddenum
                
                sourceParts = Split(.Source, delim, 2)
                
                message = "Description: " & .Description & vbNewLine & _
                "Number: " & .Number & vbNewLine & _
                "Path: " & vbNewLine & sourceParts(UBound(sourceParts)) & String$(2, vbNewLine) & _
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
Public Sub AppendErrorDescription(e As ErrObject, descriptionDetail$)
'===================================================================================================================
'Summary: Appends a user supplied description to [e.Description].
'Inputs:
'   e - Error object.
'   descriptionDetails - Description to append.
'===================================================================================================================
    With e
        .Description = descriptionDetail & vbNewLine & " >   " & .Description
    End With
End Sub
Public Sub AddParentToErrSource(e As ErrObject, parentName$)
'===================================================================================================================
'Summary: Appends [parentName] to the Source property of the provided error object [e]
'Inputs: e - General ErrorObject Err
'        parentName - Name of the procedure to add to the error source.
'Example: [parentName] > Apple > Biscuit > Chocolate
'===================================================================================================================
    Dim sourceParts$()
    Const delim$ = ": "
    
    With e
        ' If the delimiter isn't found then this Method is being called for the first time.
        If InStrB(1, .Source, delim) = 0 Then
            If .Source = Left$(Replace$(ThisWorkbook.name, "-", "_"), InStrRev(ThisWorkbook.name, ".") - 1) Then
                .Source = "[" & .Source & "]" & delim & parentName
            Else
                .Source = delim & parentName & " > " & .Source
            End If
        Else
            sourceParts = Split(.Source, delim, 2)
            sourceParts(1) = parentName & " > " & sourceParts(1)
            .Source = Join(sourceParts, delim)
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
    
    If LenB(FileOrFolderstr) = 0 Then Exit Function
    
    #If Not Mac Then
        FileOrFolderExists = LenB(Dir$(FileOrFolderstr, vbDirectory)) <> 0
        Exit Function
    #End If
    
    If Val(Application.Version) < 15 Then
        ScriptToCheckFileFolder = "tell application " & QuotedForm("System Events") & _
        "to return exists disk item (" & QuotedForm(FileOrFolderstr) & " as string)"
        FileOrFolderExists = MacScript(ScriptToCheckFileFolder)
    Else
        On Error Resume Next
        TestStr = Dir$(FileOrFolderstr & "*", vbDirectory)
        On Error GoTo 0
        If LenB(TestStr) <> 0 Then FileOrFolderExists = True
    End If

End Function
Public Function QuotedForm(ByRef item, Optional Enclosing_CHR$ = """", Optional ensureAddition As Boolean = False) As Variant

    Dim Z As Long, subArray As Variant, subArrayIndex As Long
    
    If IsArray(item) Then
        For Z = LBound(item) To UBound(item)
            If IsArray(item(Z)) Then
                subArray = item(Z)
                
                For subArrayIndex = LBound(subArray) To UBound(subArray)
                    If ensureAddition Or Not subArray(subArrayIndex) Like Enclosing_CHR & "*" & Enclosing_CHR Then
                        subArray(subArrayIndex) = Enclosing_CHR & subArray(subArrayIndex) & Enclosing_CHR
                    End If
                Next subArrayIndex
                    
                item(Z) = subArray
            Else
                If Not item(Z) Like Enclosing_CHR & "*" & Enclosing_CHR Then item(Z) = Enclosing_CHR & item(Z) & Enclosing_CHR
            End If
        Next Z
    Else
        If ensureAddition Or (Not item Like Enclosing_CHR & "*" & Enclosing_CHR) Then item = Enclosing_CHR & item & Enclosing_CHR
    End If
    
    QuotedForm = item
    
End Function
Function TryGetRequest(sUrl As String, ByRef httpResponse$) As Boolean
'===========================================================================================================
'   Summary: Sends a GET request to [sUrl].
'   Returns: True if success; otherwise, False.
'   Parameters:
'       [sUrl] - Url to send a request to.
'       [httpResponse] - String variable that will hold the response if successful.
'===========================================================================================================
    Dim onMac As Boolean
    
    On Error GoTo Failure

    #If Mac Then
        Dim shellCMD$, lExitCode&
        onMac = True
        shellCMD = "curl " & Chr(34) & sUrl & Chr(34)
        httpResponse = ExecuteShellCommandMAC(shellCMD, lExitCode)
        TryGetRequest = (lExitCode = 0)
    #Else
        With CreateObject("MSXML2.XMLHTTP")
            .Open "GET", sUrl, False
            '.setRequestHeader "Content-Type", "text/html; charset=utf-8"
            .setRequestHeader "Accept-Language", "en-US,en;q=0.9"
            .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0"
            .send
            If .Status = 200 Then
                httpResponse = .responseText
                TryGetRequest = True
            End If
        End With
    #End If

    Exit Function
    
Failure:
    PropagateError Err, "TryGetRequest", "OS: " & IIf(onMac, "Using MAC", "Windows")
End Function

Public Function HttpPost(url$, postData$, Optional urlEncoded As Boolean) As Boolean
'===========================================================================================================
'   Summary: Sends a HTTP POST request to [Url].
'   Returns: True if success; otherwise, False.
'   Parameters:
'       [Url] - Url to send a request to.
'       [postData] - Data to send.
'       [urlEncoded] True id [postData[ is url encoded. Example: &bb=7&aa=8&uu=100
'===========================================================================================================
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
            '.setTimeouts 60, 60, 30, 30
            .send postData
            HttpPost = (.Status = 200)
        End With
    #End If
    
    Exit Function
Failure:
    PropagateError Err, "HttpPost", "OS: " & IIf(onMac, "MAC", "Windows")
End Function
Public Function HasKey(col As Collection, key$) As Boolean
'===================================================================================================================
    'Summary: Determines if a given collection has a specific key.
    'Inputs: col - Collection to check.
    '        Key - key to check col for.
    'Returns: True or false.
'===================================================================================================================

    Dim v As Boolean
    
    On Error GoTo Exit_Function
    v = IsObject(col.item(key))
    HasKey = Not IsEmpty(v)

Exit_Function:
    'The key doesn't exist.
End Function
Public Function ConvertSystemTimeToDate(timeToConvert As SystemTime) As Date
    With timeToConvert
        ConvertSystemTimeToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function
Public Function ConvertDateToSystemTime(timeToConvert As Date) As SystemTime
    With ConvertDateToSystemTime
        .wYear = Year(timeToConvert)
        .wMonth = Month(timeToConvert)
        .wDay = Day(timeToConvert)
        .wHour = Hour(timeToConvert)
        .wMinute = Minute(timeToConvert)
        .wSecond = Second(timeToConvert)
    End With
End Function
Public Function ConvertLocalToUTC(convertedDate As Date) As Date
    #If Mac Then
        'ConvertLocalDatetimeToUTC = utc_ConvertDate(convertedDate, True)
    #Else
        Dim tz As TimeZoneInformation, utcTime As SystemTime
        Call GetTimeZoneInformation(tz)
        Call TzSpecificLocalTimeToSystemTime(tz, ConvertDateToSystemTime(convertedDate), utcTime)
        ConvertLocalToUTC = ConvertSystemTimeToDate(utcTime)
    #End If
End Function
Public Function ConvertGmtWithTimeZone(ByVal gmtDateTime As Date, utcOffsetHrs As Long) As Date

    Dim dstStart As Date, dstEnd As Date
    
    'EDT starts on 2nd Sunday in March 2AM
    dstStart = DateSerial(Year(gmtDateTime), 3, 14) + TimeSerial(2, 0, 0)
    dstStart = dstStart - (Weekday(dstStart) - 1)
    
    'EDT ends on 1st Sunday in November 2AM
    dstEnd = DateSerial(Year(gmtDateTime), 11, 7) + TimeSerial(1, 0, 0)
    dstEnd = dstEnd - (Weekday(dstEnd) - 1)
    
    ConvertGmtWithTimeZone = DateAdd("h", utcOffsetHrs, gmtDateTime)
    
    If ConvertGmtWithTimeZone >= dstStart And ConvertGmtWithTimeZone <= dstEnd Then
        ConvertGmtWithTimeZone = DateAdd("h", 1, ConvertGmtWithTimeZone)
    End If
        
End Function
Public Function TxtMethods(targetFile$, getInput As Boolean, appendText As Boolean, overwriteTXT As Boolean, Optional newText$) As String
'=======================================================================
'This function is used to interface with .txt files.
'=======================================================================
    Dim fileNumber As Long
    
    On Error GoTo Propagate
    fileNumber = FreeFile

    If getInput Then
        Open targetFile For Input As #fileNumber
    ElseIf appendText Then
        Open targetFile For Append As #fileNumber
    ElseIf overwriteTXT Then
        Open targetFile For Output As #fileNumber
    End If
    
    If getInput Then
        TxtMethods = Input(LOF(fileNumber), #fileNumber)
    Else
        Print #fileNumber, newText
    End If
    
    Close fileNumber
    Exit Function
Propagate:
    PropagateError Err, "TxtMethods"
End Function

#If Not Mac Then
    Function CreateFolderRecursive(path As String) As Boolean
        ' https://stackoverflow.com/a/50818079
        Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    
        'If the path exists as a file, the function fails.
        If FSO.FileExists(path) Then
            CreateFolderRecursive = False
            Exit Function
        End If
    
        'If the path already exists as a folder, don't do anything and return success.
        If FSO.FolderExists(path) Then
            CreateFolderRecursive = True
            Exit Function
        End If
    
        'recursively create the parent folder, then if successful create the top folder.
        If CreateFolderRecursive(FSO.GetParentFolderName(path)) Then
            CreateFolderRecursive = Not FSO.CreateFolder(path) Is Nothing
        Else
            CreateFolderRecursive = False
        End If
        
    End Function
#End If
Public Function GetIpJSON() As Object
'===================================================================================================================
    'Summary: Gets basic computer/internet information from an api in JSON form.
'===================================================================================================================
    Dim apiResponse$, jp As New JsonParserB
    
    Const apiUrl$ = "https://ipapi.co/json/"
    
    On Error GoTo Catch_GET_FAILED
 
    If TryGetRequest(apiUrl, apiResponse) Then
        On Error GoTo CATCH_JSON_PARSER_FAILURE
        Set GetIpJSON = jp.Deserialize(apiResponse)
    Else
        Err.Raise vbObjectError + 799, "GetIpJSON", "Failed to retrieve data from verification URL."
    End If
    
    Exit Function
    
Catch_GET_FAILED:
    PropagateError Err, "GetIpJSON", "GET failed."
CATCH_JSON_PARSER_FAILURE:
    PropagateError Err, "GetIpJSON", "Unable to parse JSON."
End Function
