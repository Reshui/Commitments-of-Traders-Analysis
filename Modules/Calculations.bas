Attribute VB_Name = "Calculations"
Option Explicit
Public Function Stochastic_Calculations(Column_Number_Reference As Long, indexedWeekCount As Long, _
                                    sourceData As Variant, Missing_Weeks As Long, sourceDates() As Date, Optional Cap_Extremes As Boolean = False, Optional dateColumn As Long = 1) As Variant()
'===================================================================================================================
    'Summary: Calculates Stochastic values for values found in sourceData.
    'Inputs:    Column_NUmber_Reference - The column within sourceData that stochastic values will be calculated for.
    '           indexedWeekCount - The number of weeks used in each calculation.
    '           sourceData - Input data used to generate calculations.
    '           Missing_Weeks - Maximum number of calculations that will be done.
    '           Cap_Extremes - If true then the values will be contained to region between 0 and 100.
    '           sourceDates - An array of dates sorted in ascending order.
'===================================================================================================================

    Dim iRow As Long, currentRow&, stochValues() As Variant, totalRows&, initialRowToCalculate&, _
    Minimum As Double, Maximum As Double, dateStartIndex&, dateInPast As Date
    
    On Error GoTo Propagate
    
    totalRows = UBound(sourceData, 1)
    initialRowToCalculate = totalRows - (Missing_Weeks - 1)
    
    ReDim stochValues(1 To Missing_Weeks)
    
    For currentRow = initialRowToCalculate To totalRows

        dateInPast = DateAdd("ww", -indexedWeekCount, sourceDates(currentRow))
        
        ' Find smallest value >= dateInPast
        For dateStartIndex = IIf(currentRow >= indexedWeekCount, currentRow - (indexedWeekCount - 1), LBound(sourceDates)) To currentRow
            If sourceDates(dateStartIndex) >= dateInPast Then Exit For
        Next dateStartIndex
        
        If dateStartIndex > currentRow Then Exit For
        
        If DateDiff("ww", sourceDates(dateStartIndex), sourceDates(currentRow)) >= (indexedWeekCount * 0.8) Then
        
            For iRow = dateStartIndex To IIf(Cap_Extremes, currentRow, currentRow - 1)
                If iRow = dateStartIndex Then
                    Minimum = sourceData(iRow, Column_Number_Reference)
                    Maximum = Minimum
                ElseIf sourceData(iRow, Column_Number_Reference) < Minimum Then
                    Minimum = sourceData(iRow, Column_Number_Reference)
                ElseIf sourceData(iRow, Column_Number_Reference) > Maximum Then
                    Maximum = sourceData(iRow, Column_Number_Reference)
                End If
            Next iRow
            'Stochastic calculation
            If Maximum <> Minimum Then
                stochValues(Missing_Weeks - (totalRows - currentRow)) = CLng(((sourceData(currentRow, Column_Number_Reference) - Minimum) / (Maximum - Minimum)) * 100)
            End If
        End If
    Next currentRow
    
    Stochastic_Calculations = stochValues
    Exit Function
Propagate:
    PropagateError Err, "Stochastic_Calculations"
End Function
Public Function Legacy_Multi_Calculations(ByRef sourceData() As Variant, weeksToCalculateCount As Long, commercialNetColumn As Byte, _
    Time1 As Long, Time2 As Long) As Variant()
    '======================================================================================================
    'Legacy Calculations for calculated columns
    '======================================================================================================
    Dim iRow As Long, iCount As Long, n As Long, Start As Long, Finish As Long, INTE_B() As Variant, _
    Z As Long, outputA() As Variant, sourceDates() As Date, outputRow&
    
    On Error GoTo Propogate
    
    Start = UBound(sourceData, 1) - (weeksToCalculateCount - 1)
    Finish = UBound(sourceData, 1)
    
    'Time1 is Year3,Time2 is Month6
    ReDim outputA(1 To weeksToCalculateCount, 1 To UBound(sourceData, 2))
    outputRow = LBound(outputA, 1)
    For iRow = Start To Finish
        
        For iCount = 0 To 2 'Commercial Net,Non-Commercial Net,Non-Reportable
            n = Array(7, 4, 11)(iCount)
            sourceData(iRow, commercialNetColumn + iCount) = sourceData(iRow, n) - sourceData(iRow, n + 1)
        Next iCount
        
        outputA(outputRow, commercialNetColumn + 20) = sourceData(iRow, 27) - sourceData(iRow, 28) 'net %OI Commercial
        outputA(outputRow, commercialNetColumn + 21) = sourceData(iRow, 24) - sourceData(iRow, 25) 'net %OI Non-Commercial
        'Commercial Net/OI
        If sourceData(iRow, 3) <> 0 Then sourceData(iRow, commercialNetColumn + 9) = sourceData(iRow, commercialNetColumn) / sourceData(iRow, 3)

        If sourceData(iRow, 4) > 0 Or sourceData(iRow, 5) > 0 Then
            'NC Long%
            outputA(outputRow, commercialNetColumn + 13) = sourceData(iRow, 4) / (sourceData(iRow, 4) + sourceData(iRow, 5))
            'NC Short%
            outputA(outputRow, commercialNetColumn + 14) = 1 - outputA(outputRow, commercialNetColumn + 13)
        End If

        If iRow >= 2 Then
            'Commercial Net Change
            outputA(outputRow, commercialNetColumn + 15) = sourceData(iRow, commercialNetColumn) - sourceData(iRow - 1, commercialNetColumn)
            'Commercial Gross Long % Change, Commercial Gross Short % Change
            For iCount = 7 To 8
                If sourceData(iRow - 1, iCount) > 0 Then
                    outputA(outputRow, commercialNetColumn + 9 + iCount) = (sourceData(iRow, iCount) - sourceData(iRow - 1, iCount)) / sourceData(iRow - 1, iCount)
                End If
            Next iCount
        End If

        If sourceData(iRow, 7) > 0 Or sourceData(iRow, 8) > 0 Then
            'Commercial Long %
            outputA(outputRow, commercialNetColumn + 18) = sourceData(iRow, 7) / (sourceData(iRow, 7) + sourceData(iRow, 8))
            'Commercial Short %
            outputA(outputRow, commercialNetColumn + 19) = 1 - outputA(outputRow, commercialNetColumn + 18)
        End If
        outputRow = outputRow + 1
    Next iRow
    
    ReDim sourceDates(LBound(sourceData, 1) To UBound(sourceData, 1))
    
    For iRow = LBound(sourceData, 1) To UBound(sourceData, 1)
        sourceDates(iRow) = sourceData(iRow, 1)
    Next iRow
    
    Dim willCo3YIndexColumn&, movementIndexColumn&
    
    willCo3YIndexColumn = commercialNetColumn + 11
    
    If UBound(sourceData, 1) > Time1 Then
        'Calculate Three year index.
        For iCount = 0 To 3
        
            INTE_B = Stochastic_Calculations(commercialNetColumn + Array(0, 2, 9, 1)(iCount), Time1, sourceData, weeksToCalculateCount, sourceDates, Cap_Extremes:=True)
            '[0]Commercial index 3Y  [1]Non-Reportable 3Y   < values of iCount
            '[2] Willco3Y            [3] Non-Commerical 3Y
            n = Array(3, 5, 11, 4)(iCount) + commercialNetColumn
            outputRow = LBound(outputA, 1)
            For iRow = Start To Finish
                If n = willCo3YIndexColumn Then
                    sourceData(iRow, n) = INTE_B(outputRow)
                Else
                    outputA(outputRow, n) = INTE_B(outputRow)
                End If
                outputRow = outputRow + 1
            Next iRow
            
            Erase INTE_B
            
        Next iCount

    End If
    
    If UBound(sourceData, 1) > Time2 Then '6M index
        For iCount = 0 To 3
            
            INTE_B = Stochastic_Calculations(commercialNetColumn + Array(0, 2, 9, 1)(iCount), Time2, sourceData, weeksToCalculateCount, sourceDates, Cap_Extremes:=True)
            
            n = Array(6, 8, 10, 7)(iCount) + commercialNetColumn ' used to calculate column number
            
            outputRow = LBound(outputA, 1)
            
            For iRow = Start To Finish
                outputA(outputRow, n) = INTE_B(outputRow)    '[0]Commerical 6M [1]Non-Reportable 6M
                outputRow = outputRow + 1                    '[2]WillCo6M      [3]Non Commercial 6M
            Next iRow
            
            Erase INTE_B
            
        Next iCount
    End If
    
    movementIndexColumn = willCo3YIndexColumn + 1

    'The below code block is for adding only the missing data to the output array
    outputRow = LBound(outputA, 1)

    For iRow = Start To Finish 'populate each row sequentially
        For iCount = LBound(sourceData, 2) To UBound(sourceData, 2)
            If IsEmpty(outputA(outputRow, iCount)) Then
                
                If iCount = movementIndexColumn And (iRow - 6) >= LBound(sourceData, 1) Then
                    outputA(outputRow, iCount) = sourceData(iRow, willCo3YIndexColumn) - sourceData(iRow - 6, willCo3YIndexColumn)
                Else
                    outputA(outputRow, iCount) = sourceData(iRow, iCount)
                End If
                
            End If
        Next iCount
        outputRow = outputRow + 1
    Next iRow
    
    Legacy_Multi_Calculations = outputA
    Exit Function
Propogate:
    PropagateError Err, "Legacy_Multi_Calculations"
End Function
Public Function Disaggregated_Multi_Calculations(ByRef sourceData() As Variant, weeksToCalculateCount As Long, ByVal producerNetColumn As Byte, Time1 As Long, Time2 As Long) As Variant()

    Dim contractCodeColumn As Byte, iRow As Long, outputA() As Variant, _
    iCount As Byte, openInterest As Long, Start As Long, Finish As Long, INTE_B() As Variant, _
    Z As Long, columnIndexByte As Byte, isIceData As Boolean, sourceDates() As Date
    
    'Time1 is Year3,Time2 is Month6
    On Error GoTo Propogate
    
    contractCodeColumn = producerNetColumn - 3
    
    Start = UBound(sourceData, 1) - (weeksToCalculateCount - 1) '-1 to incorpotate all missed weeks
    Finish = UBound(sourceData, 1)
    isIceData = InStrB(1, LCase$(sourceData(LBound(sourceData, 1), 2)), "ice") = 1
    
    For iRow = Start To Finish
    
        For iCount = 0 To 3 'Producer Net , Swap Net , Managed Net , Other Net
            columnIndexByte = Array(4, 6, 9, 12)(iCount)
            sourceData(iRow, producerNetColumn + iCount) = sourceData(iRow, columnIndexByte) - sourceData(iRow, columnIndexByte + 1)
        Next iCount
        
        sourceData(iRow, producerNetColumn + 4) = sourceData(iRow, producerNetColumn) + sourceData(iRow, producerNetColumn + 1) 'Commercial Net
        
        If iRow >= 2 Then
            For iCount = 0 To 2 'Producer Net Change,Swap Net Change,Commercial Net Change
                columnIndexByte = Array(0, 1, 4)(iCount)
                sourceData(iRow, producerNetColumn + 20 + iCount) = sourceData(iRow, producerNetColumn + columnIndexByte) - sourceData(iRow - 1, producerNetColumn + columnIndexByte)
            Next iCount
        
            For iCount = 6 To 7 'Commercial Long/Short Change
                sourceData(iRow, producerNetColumn + 17 + iCount) = (sourceData(iRow, iCount - 2) + sourceData(iRow, iCount)) - (sourceData(iRow - 1, iCount - 2) + sourceData(iRow - 1, iCount))
            Next iCount
        End If
        
        If sourceData(iRow, 3) <> 0 And Not IsNull(sourceData(iRow, 3)) Then
            openInterest = sourceData(iRow, 3)
            'Producer/OI
            sourceData(iRow, producerNetColumn + 11) = sourceData(iRow, producerNetColumn) / openInterest
            'Swap/OI
            sourceData(iRow, producerNetColumn + 12) = sourceData(iRow, producerNetColumn + 1) / openInterest
        End If
        
        'Commercial/OI
        sourceData(iRow, producerNetColumn + 10) = sourceData(iRow, producerNetColumn + 11) + sourceData(iRow, producerNetColumn + 12)
        
        If isIceData Then
            ' Calculate changes.
            If iRow > 1 Then
                For columnIndexByte = 3 To 18
                    Select Case columnIndexByte
                        Case 15, 16
                        Case Else
                            sourceData(iRow, columnIndexByte + 16) = sourceData(iRow, columnIndexByte) - sourceData(iRow - 1, columnIndexByte)
                    End Select
                Next columnIndexByte
            End If
            'Below are the calculations for total reportable positions[OI-Non Reportable for position.]
            sourceData(iRow, 15) = sourceData(iRow, 3) - sourceData(iRow, 17)  'Reportables long
            sourceData(iRow, 16) = sourceData(iRow, 3) - sourceData(iRow, 18)  'Reportables Short
            'Calculate total reportable traders.
            sourceData(iRow, 63) = sourceData(iRow, 52) + sourceData(iRow, 54) + sourceData(iRow, 57) + sourceData(iRow, 60)
            sourceData(iRow, 64) = sourceData(iRow, 53) + sourceData(iRow, 55) + sourceData(iRow, 58) + sourceData(iRow, 61)
                        
            If iRow > 1 Then
                'Calculate change in Reportable positions.
                sourceData(iRow, 31) = sourceData(iRow, 15) - sourceData(iRow - 1, 15)
                sourceData(iRow, 32) = sourceData(iRow, 16) - sourceData(iRow - 1, 16)
            End If
                
        End If
        
    Next iRow
    ReDim sourceDates(LBound(sourceData, 1) To UBound(sourceData, 1))
    
    For iRow = LBound(sourceData, 1) To UBound(sourceData, 1)
        sourceDates(iRow) = sourceData(iRow, 1)
    Next iRow
        
    If UBound(sourceData, 1) > Time1 Then     'Year 3
    
        For iCount = 0 To 7
            INTE_B = Stochastic_Calculations(producerNetColumn + Array(0, 1, 2, 3, 4, 10, 11, 12)(iCount), Time1, sourceData, weeksToCalculateCount, sourceDates, Cap_Extremes:=True)  'Producer Net 3Y Array
            
            columnIndexByte = producerNetColumn + Array(5, 6, 7, 8, 9, 14, 15, 16)(iCount) ' used to calculate column number
            Z = 1
            For iRow = Start To Finish  'From First Missed to Most recent
                sourceData(iRow, columnIndexByte) = INTE_B(Z)
                Z = Z + 1
            Next iRow
        Next iCount
        
        Erase INTE_B
        
    End If
             
    If UBound(sourceData, 1) > Time2 Then   'Month6
        For iCount = 0 To 2
            INTE_B = Stochastic_Calculations(producerNetColumn + Array(10, 11, 12)(iCount), Time2, sourceData, weeksToCalculateCount, sourceDates, Cap_Extremes:=True) 'Commercial/Oi 6M Array
            
            columnIndexByte = producerNetColumn + Array(17, 18, 19)(iCount)
            Z = 1
            For iRow = Start To Finish          'From First Missed to Most recent
                sourceData(iRow, columnIndexByte) = INTE_B(Z)  '   xxx/oi
                Z = Z + 1
            Next iRow
            
            Erase INTE_B
        Next iCount
    End If
    
    columnIndexByte = producerNetColumn + 13   'Movement Index column
    iCount = columnIndexByte + 1
    
    For iRow = Start To Finish                  'First Missed to most recent or Last row if Weekl
        If iRow > Time1 + 6 Then                ' Movement Index Calculation
            sourceData(iRow, columnIndexByte) = sourceData(iRow, iCount) - sourceData(iRow - 6, iCount)
        End If
    Next iRow
    
    'The below code block is for adding only the missed data to  an Array called Intermediate_F
     Z = 1
    
    ReDim outputA(1 To weeksToCalculateCount, 1 To UBound(sourceData, 2))

    For iRow = Start To Finish
        For iCount = 1 To UBound(sourceData, 2)
            outputA(Z, iCount) = sourceData(iRow, iCount)
        Next iCount
        Z = Z + 1
    Next iRow

    Disaggregated_Multi_Calculations = outputA
    Exit Function
Propogate:
    Stop: Resume
    PropagateError Err, "Disaggregated_Multi_Calculations"
End Function

Public Function TFF_Multi_Calculations(ByRef sourceData() As Variant, weeksToCalculateCount As Long, Dealer_Column As Byte, Time1 As Long, Time2 As Long, Time3 As Long) As Variant()

    Dim iRow As Long, iCount As Byte, n As Long, Start As Long, Finish As Long, INTE_B() As Variant, _
    Z As Long, outputA() As Variant, sourceDates() As Date

    Start = UBound(sourceData, 1) - (weeksToCalculateCount - 1) 'First missing week in the case of 1 or more rows to be calculated
    Finish = UBound(sourceData, 1)
    
    For iRow = Start To Finish
    
        On Error GoTo Propogate
        
        For iCount = 0 To 4                      'Calculate Other,Non-Reportable and Leveraged Fund Net
                                                 'Dealers ,Asset Managers
            n = Array(4, 7, 10, 13, 18)(iCount)  'location of long column, short column = long column+1
            Z = Array(0, 1, 2, 3, 4)(iCount)
            sourceData(iRow, Dealer_Column + Z) = sourceData(iRow, n) - sourceData(iRow, n + 1)
            
        Next iCount
        
        If sourceData(iRow, 3) > 0 Then
            sourceData(iRow, Dealer_Column + 5) = sourceData(iRow, Dealer_Column) / sourceData(iRow, 3)  'Classification/OI
        End If
        
        If iRow >= 2 Then 'Calculate Change in Net positions for column 38 may Dealer or Asset Manger depending on if contract code is in the exceptions array
              sourceData(iRow, Dealer_Column + 12) = sourceData(iRow, Dealer_Column) - sourceData(iRow - 1, Dealer_Column)
        End If
        
        On Error GoTo NET_OI_Percentage_Unavailable
        
        For iCount = 0 To 2 'Net % OI per classificaion
            n = Array(38, 41, 44)(iCount) 'Long % OI column locations for Dealers, Asset Mangers and Leveraged Money
            sourceData(iRow, Dealer_Column + Array(13, 14, 15)(iCount)) = sourceData(iRow, n) - sourceData(iRow, n + 1)
NET_OI_Percentage_Unavailable:
            On Error GoTo -1
        Next iCount
    
    Next iRow
    
    On Error GoTo Propogate
    
    ReDim sourceDates(LBound(sourceData, 1) To UBound(sourceData, 1))
    
    For iRow = LBound(sourceData, 1) To UBound(sourceData, 1)
        sourceDates(iRow) = sourceData(iRow, 1)
    Next iRow
    
    If UBound(sourceData, 1) > Time1 Then     'Year 3 Indexes
    
        For iCount = 0 To 1
            INTE_B = Stochastic_Calculations(Dealer_Column + Array(0, 5)(iCount), Time1, sourceData, weeksToCalculateCount, sourceDates, Cap_Extremes:=True)
            n = Dealer_Column + Array(6, 8)(iCount)
            Z = 1
            For iRow = Start To Finish              'From First Missed to Most recent.
                sourceData(iRow, n) = INTE_B(Z)     'Dealer index 3Y 'Dealer/OI 3Y.
                Z = Z + 1
            Next iRow
            
            Erase INTE_B
        Next iCount
    
    End If
    
    If UBound(sourceData, 1) > Time2 Then   'Month6 willco
        'Dealer/Oi 6M Array
        INTE_B = Stochastic_Calculations(Dealer_Column + 5, Time2, sourceData, weeksToCalculateCount, sourceDates, Cap_Extremes:=True)
        
        n = Dealer_Column + 10
        Z = 1
        For iRow = Start To Finish              'From First Missed to Most recent.
            sourceData(iRow, n) = INTE_B(Z)     'Dealer/Oi 6M.
            Z = Z + 1
        Next iRow
        
        Erase INTE_B
        
    End If
                 
    If UBound(sourceData, 1) > Time3 Then   '1Y Indexes
    
        For iCount = 0 To 1
            'Dealer 1Y Array
            INTE_B = Stochastic_Calculations(Dealer_Column + Array(0, 5)(iCount), Time3, sourceData, weeksToCalculateCount, sourceDates, Cap_Extremes:=True)
          
            n = Dealer_Column + Array(7, 9)(iCount)
            Z = 1
            For iRow = Start To Finish
                sourceData(iRow, n) = INTE_B(Z)     'Dealer 1Y and DEALER/OI
                Z = Z + 1
            Next iRow
            
            Erase INTE_B
            
        Next iCount
    
    End If
    
    For iRow = Start To Finish                 'First Missed week to end of data set
        If iRow > Time1 + 6 Then               'Movement Index Calculation
            sourceData(iRow, Dealer_Column + 11) = sourceData(iRow, Dealer_Column + 8) - sourceData(iRow - 6, Dealer_Column + 8)
        End If
    Next iRow
            
    'The below code block is for adding only the missed data to  an Array called Intermediate_F
     n = 1

    ReDim outputA(1 To weeksToCalculateCount, 1 To UBound(sourceData, 2))

    For iRow = Start To Finish 'populate each row sequentially
        For iCount = 1 To UBound(sourceData, 2)
            outputA(n, iCount) = sourceData(iRow, iCount)
        Next iCount
        n = n + 1
    Next iRow

    TFF_Multi_Calculations = outputA
    Exit Function
Propogate:
    PropagateError Err, "TFF_Multi_Calculations"
End Function


