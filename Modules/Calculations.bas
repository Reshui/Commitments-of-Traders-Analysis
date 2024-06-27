Attribute VB_Name = "Calculations"
Option Explicit
Public Function Stochastic_Calculations(Column_Number_Reference As Long, Time_Period As Long, _
                                    inputA As Variant, Optional Missing_Weeks As Long = 1, Optional Cap_Extremes As Boolean = False) As Byte()
'===================================================================================================================
    'Purpose: Calculates Stochastic values for values found in inputA.
    'Inputs:    Column_NUmber_Reference - The column within inputA that stochastic values will be calculated for.
    '           Time_Period - The number of weeks used in each calculation.
    '           inputA - Input data used to generate calculations.
    '           Missing_Weeks - Maximum number of calculations that will be done.
    '           Cap_Extremes - If true then the values will be contained to region between 0 and 100.
    'Outputs:
'===================================================================================================================

    Dim Array_Column() As Double, iRow As Long, Array_Period() As Double, current_row As Long, _
    outputA() As Byte, totalRows As Long, initialRowToCalculate As Long, Minimum As Double, Maximum As Double
    
    On Error GoTo Propogate
    totalRows = UBound(inputA, 1) 'number of rows in the supplied array[Upper bound in 1st dimension]
    
    initialRowToCalculate = totalRows - (Missing_Weeks - 1) 'will be equal to totalRows if only 1 missing week
    
    ReDim Array_Period(1 To Time_Period)  'Temporary Array that will hold a certain range of data
    ReDim outputA(1 To Missing_Weeks) 'Array that will hold calculated values
    ReDim Array_Column(IIf(initialRowToCalculate > Time_Period, initialRowToCalculate - Time_Period, 1) To totalRows) 'Array composed of all data in a column
    
    If UBound(inputA, 2) = 1 Then Column_Number_Reference = 1 'for when a single column is supplied
    
    For iRow = IIf(initialRowToCalculate > Time_Period, initialRowToCalculate - Time_Period, 1) To totalRows
        'if starting row of data output is greater than the time period then offset the start of the queried array by the time period
        'otherwise start at 1...there should be checks to ensure there is enough data most of the time
        Array_Column(iRow) = inputA(iRow, Column_Number_Reference)
    Next iRow
    
    For current_row = initialRowToCalculate To totalRows
    
        If (current_row > Time_Period And Not Cap_Extremes) Or (current_row >= Time_Period And Cap_Extremes) Then   'Only calculate if there is enough data
        
            For iRow = 1 To Time_Period 'Fill array with the previous Time_Period number of values relative to the current row
                
                If Not Cap_Extremes Then
                    Array_Period(iRow) = Array_Column(current_row - iRow)
                Else
                    Array_Period(iRow) = Array_Column(current_row - (iRow - 1))
                End If
                
                If iRow = 1 Then
                    Minimum = Array_Period(iRow)
                    Maximum = Minimum
                ElseIf Array_Period(iRow) < Minimum Then
                    Minimum = Array_Period(iRow)
                ElseIf Array_Period(iRow) > Maximum Then
                    Maximum = Array_Period(iRow)
                End If
                                
            Next iRow
            'Stochastic calculation
            If Maximum <> Minimum Then
                outputA(Missing_Weeks - (totalRows - current_row)) = CByte(((Array_Column(current_row) - Minimum) / (Maximum - Minimum)) * 100)
            End If
            'ex for determining current location within array:    2 - ( 480 - 479 ) = 1
        End If
    
    Next current_row
    
    Stochastic_Calculations = outputA
    Exit Function
Propogate:
    PropagateError Err, "Stochastic_Calculations"
End Function
Public Function Legacy_Multi_Calculations(ByRef inputA() As Variant, weeksToCalculateCount As Long, commercialNetColumn As Byte, _
    Time1 As Long, Time2 As Long) As Variant()
    '======================================================================================================
    'Legacy Calculations for calculated columns
    '======================================================================================================
    Dim iRow As Long, iCount As Long, N As Long, Start As Long, Finish As Long, INTE_B() As Byte, Z As Long, outputA() As Variant
    
    On Error GoTo Propogate
    
    Start = UBound(inputA, 1) - (weeksToCalculateCount - 1)
    Finish = UBound(inputA, 1)
    
    'Time1 is Year3,Time2 is Month6
    
    For iRow = Start To Finish
        
        For iCount = 0 To 2 'Commercial Net,Non-Commercial Net,Non-Reportable
            N = Array(7, 4, 11)(iCount)
            inputA(iRow, commercialNetColumn + iCount) = inputA(iRow, N) - inputA(iRow, N + 1)
        Next iCount
        
        inputA(iRow, commercialNetColumn + 20) = inputA(iRow, 27) - inputA(iRow, 28) 'net %OI Commercial
        inputA(iRow, commercialNetColumn + 21) = inputA(iRow, 24) - inputA(iRow, 25) 'net %OI Non-Commercial
        'Commercial Net/OI
        If inputA(iRow, 3) <> 0 Then inputA(iRow, commercialNetColumn + 9) = inputA(iRow, commercialNetColumn) / inputA(iRow, 3)

        If inputA(iRow, 4) > 0 Or inputA(iRow, 5) > 0 Then
            'NC Long%
            inputA(iRow, commercialNetColumn + 13) = inputA(iRow, 4) / (inputA(iRow, 4) + inputA(iRow, 5))
            'NC Short%
            inputA(iRow, commercialNetColumn + 14) = 1 - inputA(iRow, commercialNetColumn + 13)
        End If

        If iRow >= 2 Then
            'Commercial Net Change
            inputA(iRow, commercialNetColumn + 15) = inputA(iRow, commercialNetColumn) - inputA(iRow - 1, commercialNetColumn)

            For iCount = 7 To 8  'Commercial Gross Long % Change, Commercial Gross Short % Change
                If inputA(iRow - 1, iCount) > 0 Then
                    inputA(iRow, commercialNetColumn + 9 + iCount) = (inputA(iRow, iCount) - inputA(iRow - 1, iCount)) / inputA(iRow - 1, iCount)
                End If
            Next iCount

        End If

        If inputA(iRow, 7) > 0 Or inputA(iRow, 8) > 0 Then
            inputA(iRow, commercialNetColumn + 18) = inputA(iRow, 7) / (inputA(iRow, 7) + inputA(iRow, 8))                  'Commercial Long %
            inputA(iRow, commercialNetColumn + 19) = 1 - inputA(iRow, commercialNetColumn + 18)                         'Commercial Short %
        End If

    Next iRow

    If UBound(inputA, 1) > Time1 Then
        'Calculate Three year index.
        For iCount = 0 To 3
        
            INTE_B = Stochastic_Calculations(commercialNetColumn + Array(0, 2, 9, 1)(iCount), Time1, inputA, weeksToCalculateCount, Cap_Extremes:=True)
            
            N = Array(3, 5, 11, 4)(iCount) + commercialNetColumn         'used to calculate column number
            Z = 1 'finish-iRow
            For iRow = Start To Finish
                inputA(iRow, N) = INTE_B(Z)                       '[0]Commercial index 3Y  [1]Non-Reportable 3Y   < values of iCount
                Z = Z + 1                                   '[2] Willco3Y            [3] Non-Commerical 3Y
            Next iRow
            
            Erase INTE_B
            
        Next iCount

    End If

    If UBound(inputA, 1) > Time2 Then '6M index
        
        For iCount = 0 To 3
            
            INTE_B = Stochastic_Calculations(commercialNetColumn + Array(0, 2, 9, 1)(iCount), Time2, inputA, weeksToCalculateCount, Cap_Extremes:=True)
            
            N = Array(6, 8, 10, 7)(iCount) + commercialNetColumn ' used to calculate column number
            Z = 1
            For iRow = Start To Finish
                inputA(iRow, N) = INTE_B(Z)   '[0]Commerical 6M [1]Non-Reportable 6M
                Z = Z + 1               '[2]WillCo6M      [3]Non Commercial 6M
            Next iRow
            
            Erase INTE_B
            
        Next iCount

    End If
    
    N = commercialNetColumn + 11 'Willco 3Y Column
    iCount = N + 1            'movement index column

    For iRow = Start To Finish 'First Missed to most recent do Movement Index Calculations

        If iRow > Time1 + 6 Then
            inputA(iRow, iCount) = inputA(iRow, N) - inputA(iRow - 6, N)
        End If

    Next iRow

    'The below code block is for adding only the missing data to the output array
    N = 1

    ReDim outputA(1 To weeksToCalculateCount, 1 To UBound(inputA, 2))

    For iRow = Start To Finish 'populate each row sequentially
        For iCount = 1 To UBound(inputA, 2)
            outputA(N, iCount) = inputA(iRow, iCount)
        Next iCount

        N = N + 1
    Next iRow
    
    Legacy_Multi_Calculations = outputA
    Exit Function
Propogate:
    PropagateError Err, "Legacy_Multi_Calculations"
End Function
Public Function Disaggregated_Multi_Calculations(ByRef inputA() As Variant, weeksToCalculateCount As Long, ByVal producerNetColumn As Byte, Time1 As Long, Time2 As Long) As Variant()

    Dim iceContractCodes$(), contractCodeColumn As Byte, iRow As Long, outputA() As Variant, _
    iCount As Byte, openInterest As Long, Start As Long, Finish As Long, INTE_B() As Byte, _
    Z As Long, columnIndexByte As Byte, isIceData As Boolean
    
    'Time1 is Year3,Time2 is Month6
    On Error GoTo Propogate
    iceContractCodes = Split("Wheat,B,RC,W,G,Cocoa", ",")
    
    contractCodeColumn = producerNetColumn - 3
    
    Start = UBound(inputA, 1) - (weeksToCalculateCount - 1) '-1 to incorpotate all missed weeks
    Finish = UBound(inputA, 1)
    isIceData = Not IsError(Application.Match(inputA(UBound(inputA, 1), contractCodeColumn), iceContractCodes, 0))
    
    For iRow = Start To Finish
    
        For iCount = 0 To 3 'Producer Net , Swap Net , Managed Net , Other Net
            columnIndexByte = Array(4, 6, 9, 12)(iCount)
            inputA(iRow, producerNetColumn + iCount) = inputA(iRow, columnIndexByte) - inputA(iRow, columnIndexByte + 1)
        Next iCount
        
        inputA(iRow, producerNetColumn + 4) = inputA(iRow, producerNetColumn) + inputA(iRow, producerNetColumn + 1) 'Commercial Net
        
        If iRow >= 2 Then
            For iCount = 0 To 2 'Producer Net Change,Swap Net Change,Commercial Net Change
                columnIndexByte = Array(0, 1, 4)(iCount)
                inputA(iRow, producerNetColumn + 20 + iCount) = inputA(iRow, producerNetColumn + columnIndexByte) - inputA(iRow - 1, producerNetColumn + columnIndexByte)
            Next iCount
        
            For iCount = 6 To 7 'Commercial Long/Short Change
                inputA(iRow, producerNetColumn + 17 + iCount) = (inputA(iRow, iCount - 2) + inputA(iRow, iCount)) - (inputA(iRow - 1, iCount - 2) + inputA(iRow - 1, iCount))
            Next iCount
        End If
        
        openInterest = inputA(iRow, 3)
        
        If inputA(iRow, 3) <> 0 Then
            'Producer/OI
            inputA(iRow, producerNetColumn + 11) = inputA(iRow, producerNetColumn) / openInterest
            'Swap/OI
            inputA(iRow, producerNetColumn + 12) = inputA(iRow, producerNetColumn + 1) / openInterest
        End If
        'Commercial/OI
        inputA(iRow, producerNetColumn + 10) = inputA(iRow, producerNetColumn + 11) + inputA(iRow, producerNetColumn + 12)
        
        If isIceData Then
            ' Calculate changes.
            If iRow > 1 Then
                For columnIndexByte = 3 To 18
                    Select Case columnIndexByte
                        Case 15, 16
                        Case Else
                            inputA(iRow, columnIndexByte + 16) = inputA(iRow, columnIndexByte) - inputA(iRow - 1, columnIndexByte)
                    End Select
                Next columnIndexByte
            End If
            'Below are the calculations for total reportable positions[OI-Non Reportable for position.]
            inputA(iRow, 15) = inputA(iRow, 3) - inputA(iRow, 17)  'Reportables long
            inputA(iRow, 16) = inputA(iRow, 3) - inputA(iRow, 18)  'Reportables Short
            'Calculate total reportable traders.
            inputA(iRow, 63) = inputA(iRow, 52) + inputA(iRow, 54) + inputA(iRow, 57) + inputA(iRow, 60)
            inputA(iRow, 64) = inputA(iRow, 53) + inputA(iRow, 55) + inputA(iRow, 58) + inputA(iRow, 61)
                        
            If iRow > 1 Then
                'Calculate change in Reportable positions.
                inputA(iRow, 31) = inputA(iRow, 15) - inputA(iRow - 1, 15)
                inputA(iRow, 32) = inputA(iRow, 16) - inputA(iRow - 1, 16)
            End If
                
        End If
        
    Next iRow
    
    If UBound(inputA, 1) > Time1 Then     'Year 3
    
        For iCount = 0 To 7
            INTE_B = Stochastic_Calculations(producerNetColumn + Array(0, 1, 2, 3, 4, 10, 11, 12)(iCount), Time1, inputA, weeksToCalculateCount, Cap_Extremes:=True)  'Producer Net 3Y Array
            
            columnIndexByte = producerNetColumn + Array(5, 6, 7, 8, 9, 14, 15, 16)(iCount) ' used to calculate column number
            Z = 1
            For iRow = Start To Finish  'From First Missed to Most recent
                inputA(iRow, columnIndexByte) = INTE_B(Z)
                Z = Z + 1
            Next iRow
        Next iCount
        
        Erase INTE_B
        
    End If
             
    If UBound(inputA, 1) > Time2 Then   'Month6
    
        For iCount = 0 To 2
            INTE_B = Stochastic_Calculations(producerNetColumn + Array(10, 11, 12)(iCount), Time2, inputA, weeksToCalculateCount, Cap_Extremes:=True) 'Commercial/Oi 6M Array
            
            columnIndexByte = producerNetColumn + Array(17, 18, 19)(iCount)
            Z = 1
            For iRow = Start To Finish          'From First Missed to Most recent
                inputA(iRow, columnIndexByte) = INTE_B(Z)  '   xxx/oi
                Z = Z + 1
            Next iRow
            
            Erase INTE_B
        Next iCount
       
    End If
    
    columnIndexByte = producerNetColumn + 13   'Movement Index column
    iCount = columnIndexByte + 1
    
    For iRow = Start To Finish                  'First Missed to most recent or Last row if Weekl
        If iRow > Time1 + 6 Then                ' Movement Index Calculation
            inputA(iRow, columnIndexByte) = inputA(iRow, iCount) - inputA(iRow - 6, iCount)
        End If
    Next iRow
    
    'The below code block is for adding only the missed data to  an Array called Intermediate_F
     Z = 1
    
    ReDim outputA(1 To weeksToCalculateCount, 1 To UBound(inputA, 2))

    For iRow = Start To Finish
        For iCount = 1 To UBound(inputA, 2)
            outputA(Z, iCount) = inputA(iRow, iCount)
        Next iCount
        Z = Z + 1
    Next iRow

    Disaggregated_Multi_Calculations = outputA
    Exit Function
Propogate:
    PropagateError Err, "Disaggregated_Multi_Calculations"
End Function

Public Function TFF_Multi_Calculations(ByRef inputA() As Variant, weeksToCalculateCount As Long, Dealer_Column As Byte, Time1 As Long, Time2 As Long, Time3 As Long) As Variant()

    Dim iRow As Long, iCount As Byte, N As Long, Start As Long, Finish As Long, INTE_B() As Byte, Z As Long, outputA() As Variant

    Start = UBound(inputA, 1) - (weeksToCalculateCount - 1) 'First missing week in the case of 1 or more rows to be calculated
    Finish = UBound(inputA, 1)
    
    For iRow = Start To Finish
    
        On Error GoTo Propogate
        
        For iCount = 0 To 4                      'Calculate Other,Non-Reportable and Leveraged Fund Net
                                            'Dealers ,Asset Managers
            N = Array(4, 7, 10, 13, 18)(iCount)  'location of long column, short column = long column+1
            Z = Array(0, 1, 2, 3, 4)(iCount)
            inputA(iRow, Dealer_Column + Z) = inputA(iRow, N) - inputA(iRow, N + 1)
            
        Next iCount
        
        If inputA(iRow, 3) > 0 Then
            inputA(iRow, Dealer_Column + 5) = inputA(iRow, Dealer_Column) / inputA(iRow, 3)  'Classification/OI
        End If
        
        If iRow >= 2 Then 'Calculate Change in Net positions for column 38 may Dealer or Asset Manger depending on if contract code is in the exceptions array
              inputA(iRow, Dealer_Column + 12) = inputA(iRow, Dealer_Column) - inputA(iRow - 1, Dealer_Column)
        End If
        
        On Error GoTo NET_OI_Percentage_Unavailable
        
        For iCount = 0 To 2 'Net % OI per classificaion
            N = Array(38, 41, 44)(iCount) 'Long % OI column locations for Dealers, Asset Mangers and Leveraged Money
            inputA(iRow, Dealer_Column + Array(13, 14, 15)(iCount)) = inputA(iRow, N) - inputA(iRow, N + 1)
NET_OI_Percentage_Unavailable:
            On Error GoTo -1
        Next iCount
    
    Next iRow
    
    On Error GoTo Propogate
    
    If UBound(inputA, 1) > Time1 Then     'Year 3 Indexes
    
        For iCount = 0 To 1
            INTE_B = Stochastic_Calculations(Dealer_Column + Array(0, 5)(iCount), Time1, inputA, weeksToCalculateCount, Cap_Extremes:=True)
            
            N = Dealer_Column + Array(6, 8)(iCount)
            Z = 1
            For iRow = Start To Finish          'From First Missed to Most recent.
                inputA(iRow, N) = INTE_B(Z)     'Dealer index 3Y 'Dealer/OI 3Y.
                Z = Z + 1
            Next iRow
            
            Erase INTE_B
        Next iCount
    
    End If
                        
    If UBound(inputA, 1) > Time2 Then   'Month6 willco
        
        INTE_B = Stochastic_Calculations(Dealer_Column + 5, Time2, inputA, weeksToCalculateCount, Cap_Extremes:=True)  'Dealer/Oi 6M Array
        
        N = Dealer_Column + 10
        Z = 1
        For iRow = Start To Finish          'From First Missed to Most recent.
            inputA(iRow, N) = INTE_B(Z)     'Dealer/Oi 6M.
            Z = Z + 1
        Next iRow
        
        Erase INTE_B
        
    End If
                 
    If UBound(inputA, 1) > Time3 Then   '1Y Indexes
    
        For iCount = 0 To 1
        
            INTE_B = Stochastic_Calculations(Dealer_Column + Array(0, 5)(iCount), Time3, inputA, weeksToCalculateCount, Cap_Extremes:=True) 'Dealer 1Y Array
          
            N = Dealer_Column + Array(7, 9)(iCount)
            Z = 1
            For iRow = Start To Finish          'From First Missed to Most recent
                inputA(iRow, N) = INTE_B(Z)     'Dealer 1Y and DEALER/OI
                Z = Z + 1
            Next iRow
            
            Erase INTE_B
            
        Next iCount
    
    End If
    
    For iRow = Start To Finish                 'First Missed week to end of data set
        If iRow > Time1 + 6 Then               'Movement Index Calculation
            inputA(iRow, Dealer_Column + 11) = inputA(iRow, Dealer_Column + 8) - inputA(iRow - 6, Dealer_Column + 8)
        End If
    Next iRow
            
    'The below code block is for adding only the missed data to  an Array called Intermediate_F
     N = 1

    ReDim outputA(1 To weeksToCalculateCount, 1 To UBound(inputA, 2))

    For iRow = Start To Finish 'populate each row sequentially
        For iCount = 1 To UBound(inputA, 2)
            outputA(N, iCount) = inputA(iRow, iCount)
        Next iCount
        N = N + 1
    Next iRow

    TFF_Multi_Calculations = outputA
    Exit Function
Propogate:
    PropagateError Err, "TFF_Multi_Calculations"
End Function


