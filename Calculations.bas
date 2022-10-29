Attribute VB_Name = "Calculations"
Public Function Legacy_Multi_Calculations(AR1 As Variant, Weeks_Missed As Integer, CommercialC As Byte, _
Time1 As Integer, Time2 As Integer) As Variant
'======================================================================================================
'Legacy Calculations for calculated columns
'======================================================================================================
Dim x As Integer, Y As Integer, N As Integer, Start As Integer, Finish As Integer, INTE_B() As Variant, Z As Integer

Start = UBound(AR1, 1) - (Weeks_Missed - 1)
Finish = UBound(AR1, 1)

    'Time1 is Year3,Time2 is Month6

On Error Resume Next

    For x = Start To Finish
        
        For Y = 0 To 2 'Commercial Net,Non-Commercial Net,Non-Reportable
            N = Array(7, 4, 11)(Y)
            AR1(x, CommercialC + Y) = AR1(x, N) - AR1(x, N + 1)
        Next Y
        
        AR1(x, CommercialC + 20) = AR1(x, 27) - AR1(x, 28) 'net %OI Commercial
        AR1(x, CommercialC + 21) = AR1(x, 24) - AR1(x, 25) 'net %OI Non-Commercial
        
        AR1(x, CommercialC + 9) = AR1(x, CommercialC) / (AR1(x, 3) - AR1(x, 6))     'Commercial/OI

        If AR1(x, 4) > 0 Or AR1(x, 5) > 0 Then
            AR1(x, CommercialC + 13) = AR1(x, 4) / (AR1(x, 4) + AR1(x, 5))      'NC Long%
            AR1(x, CommercialC + 14) = 1 - AR1(x, CommercialC + 13)             'NC Short%
        End If

        If x >= 2 Then

            AR1(x, CommercialC + 15) = AR1(x, CommercialC) - AR1(x - 1, CommercialC) 'Commercial Net Change

            For Y = 7 To 8  'Commercial Gross Long % Change & Commercial Gross Short % Change
                If AR1(x - 1, Y) > 0 Then
                    AR1(x, CommercialC + 9 + Y) = (AR1(x, Y) - AR1(x - 1, Y)) / AR1(x - 1, Y)
                End If
            Next Y

        End If

        If AR1(x, 7) > 0 Or AR1(x, 8) > 0 Then

            AR1(x, CommercialC + 18) = AR1(x, 7) / (AR1(x, 7) + AR1(x, 8))                  'Commercial Long %
            AR1(x, CommercialC + 19) = 1 - AR1(x, CommercialC + 18)                         'Commercial Short %

        End If

    Next x

On Error GoTo 0

    If UBound(AR1, 1) > Time1 Then '3Y index
        
        For Y = 0 To 3
        
            INTE_B = Stochastic_Calculations(CommercialC + Array(0, 2, 9, 1)(Y), Time1, AR1, Weeks_Missed, Cap_Extremes:=True)
            
            N = Array(3, 5, 11, 4)(Y) + CommercialC         'used to calculate column number
            Z = 1 'finish-x
            For x = Start To Finish
                AR1(x, N) = INTE_B(Z)                       '[0]Commercial index 3Y  [1]Non-Reportable 3Y   < values of Y
                Z = Z + 1                                   '[2] Willco3Y            [3] Non-Commerical 3Y
            Next x
            
            Erase INTE_B
            
        Next Y

    End If

    If UBound(AR1, 1) > Time2 Then '6M index
        
        For Y = 0 To 3
            
            INTE_B = Stochastic_Calculations(CommercialC + Array(0, 2, 9, 1)(Y), Time2, AR1, Weeks_Missed, Cap_Extremes:=True)
            
            N = Array(6, 8, 10, 7)(Y) + CommercialC ' used to calculate column number
            Z = 1
            For x = Start To Finish
                AR1(x, N) = INTE_B(Z)   '[0]Commerical 6M [1]Non-Reportable 6M
                Z = Z + 1               '[2]WillCo6M      [3]Non Commercial 6M
            Next x
            
            Erase INTE_B
            
        Next Y

    End If
    
    N = CommercialC + 11 'Willco 3Y Column
    Y = N + 1            'movement index column

    For x = Start To Finish 'First Missed to most recent do Movement Index Calculations

        If x > Time1 + 6 Then
            AR1(x, Y) = AR1(x, N) - AR1(x - 6, N)
        End If

    Next x

    'The below code block is for adding only the missing data to the output array
    N = 1

    ReDim INTE_B(1 To Weeks_Missed, 1 To UBound(AR1, 2))

    For x = Start To Finish 'populate each row sequentially

        For Y = 1 To UBound(AR1, 2)
            INTE_B(N, Y) = AR1(x, Y)
        Next Y

        N = N + 1

    Next x
    
    Legacy_Multi_Calculations = INTE_B

End Function
Public Function Disaggregated_Multi_Calculations(ByRef AR1 As Variant, Weeks_Missed As Integer, ByVal Producer_Column As Byte, Time1 As Integer, Time2 As Integer) As Variant

Dim Code_Exceptions() As String, Code_Column As Byte, rowIndex As Integer, _
Y As Integer, oiNoSpread As Long, Start As Integer, Finish As Integer, INTE_B() As Variant, Z As Integer, columnIndexByte As Byte, columnIndexInt As Integer

'Time1 is Year3,Time2 is Month6

Code_Exceptions = Split("Wheat,B,RC,W,G,Cocoa", ",")

Code_Column = Producer_Column - 3

Start = UBound(AR1, 1) - (Weeks_Missed - 1) '-1 to incorpotate all missed weeks

Finish = UBound(AR1, 1)

On Error Resume Next

    For rowIndex = Start To Finish
    
        For Y = 0 To 3 'Producer Net , Swap Net , Managed Net , Other Net
        
            columnIndexByte = Array(4, 6, 9, 12)(Y)
            
            AR1(rowIndex, Producer_Column + Y) = AR1(rowIndex, columnIndexByte) - AR1(rowIndex, columnIndexByte + 1)
            
        Next Y
        
        AR1(rowIndex, Producer_Column + 4) = AR1(rowIndex, Producer_Column) + AR1(rowIndex, Producer_Column + 1) 'Commercial Net
        
        If rowIndex >= 2 Then
        
            For Y = 0 To 2 'Producer Net Change,Swap Net Change,Commercial Net Change
            
                columnIndexByte = Array(0, 1, 4)(Y)
                
                AR1(rowIndex, Producer_Column + 20 + Y) = AR1(rowIndex, Producer_Column + columnIndexByte) - AR1(rowIndex - 1, Producer_Column + columnIndexByte)
                
            Next Y
        
            For Y = 6 To 7 'Commercial Long/Short Change
                AR1(rowIndex, Producer_Column + 17 + Y) = (AR1(rowIndex, Y - 2) + AR1(rowIndex, Y)) - (AR1(rowIndex - 1, Y - 2) + AR1(rowIndex - 1, Y))
            Next Y
        
        End If
        
        oiNoSpread = AR1(rowIndex, 3) - (AR1(rowIndex, 8) + AR1(rowIndex, 11) + AR1(rowIndex, 14)) 'OI without Spread Contracts
        
        AR1(rowIndex, Producer_Column + 11) = AR1(rowIndex, Producer_Column) / oiNoSpread      'Producer/OI
        AR1(rowIndex, Producer_Column + 12) = AR1(rowIndex, Producer_Column + 1) / oiNoSpread  'Swap/OI
        
        AR1(rowIndex, Producer_Column + 10) = AR1(rowIndex, Producer_Column + 11) + AR1(rowIndex, Producer_Column + 12)                        'Commercial/OI
        
    Next rowIndex

    If Not IsError(Application.Match(AR1(UBound(AR1, 1), Code_Column), Code_Exceptions, 0)) Then
    
        For rowIndex = Start To Finish
                
            For columnIndexByte = 3 To 18 'chaZges iZ iZdividual reportable positioZs
                
                Select Case columnIndexByte
                
                    Case 15, 16 'skip g= 15 aZd 16
                    Case Else
                        
                        If rowIndex > 1 Then AR1(rowIndex, columnIndexByte + 16) = AR1(rowIndex, columnIndexByte) - AR1(rowIndex - 1, columnIndexByte)
                
                End Select
            
            Next columnIndexByte
            
            'Below are the calculations for total reportable positions
            AR1(rowIndex, 15) = AR1(rowIndex, 3) - AR1(rowIndex, 17)  'Reportables long
            AR1(rowIndex, 16) = AR1(rowIndex, 3) - AR1(rowIndex, 18)  'Reportables Short
            
            AR1(rowIndex, 63) = AR1(rowIndex, 52) + AR1(rowIndex, 54) + AR1(rowIndex, 57) + AR1(rowIndex, 60)
            AR1(rowIndex, 64) = AR1(rowIndex, 53) + AR1(rowIndex, 55) + AR1(rowIndex, 58) + AR1(rowIndex, 61)
            
            If rowIndex > 1 Then
                AR1(rowIndex, 31) = AR1(rowIndex, 15) - AR1(rowIndex - 1, 15) 'Change in RPL
                AR1(rowIndex, 32) = AR1(rowIndex, 16) - AR1(rowIndex - 1, 16) 'Change in RPS
            End If
            
        Next rowIndex
               
    End If
    
    On Error GoTo 0
    
    If UBound(AR1, 1) > Time1 Then     'Year 3
    
        For Y = 0 To 7
        
            INTE_B = Stochastic_Calculations(Producer_Column + Array(0, 1, 2, 3, 4, 10, 11, 12)(Y), Time1, AR1, Weeks_Missed, Cap_Extremes:=True)  'Producer Net 3Y Array
            
            columnIndexInt = Producer_Column + Array(5, 6, 7, 8, 9, 14, 15, 16)(Y) ' used to calculate column number
            Z = 1
            For rowIndex = Start To Finish  'From First Missed to Most recent
                AR1(rowIndex, columnIndexInt) = INTE_B(Z)
                Z = Z + 1
            Next rowIndex
            
        Next Y
        
        Erase INTE_B
        
    End If
             
    If UBound(AR1, 1) > Time2 Then   'Month6
    
        For Y = 0 To 2
            
            INTE_B = Stochastic_Calculations(Producer_Column + Array(10, 11, 12)(Y), Time2, AR1, Weeks_Missed, Cap_Extremes:=True) 'Commercial/Oi 6M Array
            
            columnIndexInt = Producer_Column + Array(17, 18, 19)(Y)
            Z = 1
            For rowIndex = Start To Finish          'From First Missed to Most recent
                AR1(rowIndex, columnIndexInt) = INTE_B(Z)  '   xxx/oi
                Z = Z + 1
            Next rowIndex
            
            Erase INTE_B
            
        Next Y
       
    End If
    
    columnIndexInt = Producer_Column + 13   'Movement Index column
    Y = columnIndexInt + 1
    
    For rowIndex = Start To Finish                  'First Missed to most recent or Last row if Weekl
        If rowIndex > Time1 + 6 Then                ' Movement Index Calculation
            AR1(rowIndex, columnIndexInt) = AR1(rowIndex, Y) - AR1(rowIndex - 6, Y)
        End If
    Next rowIndex
    
    'The below code block is for adding only the missed data to  an Array called Intermediate_F
     Z = 1
    
    ReDim INTE_B(1 To Weeks_Missed, 1 To UBound(AR1, 2))

    For rowIndex = Start To Finish 'populate each row sequentially

        For Y = 1 To UBound(AR1, 2)
            INTE_B(Z, Y) = AR1(rowIndex, Y)
        Next Y
        Z = Z + 1

    Next rowIndex

    Disaggregated_Multi_Calculations = INTE_B
    
End Function

Public Function TFF_Multi_Calculations(AR1, Weeks_Missed As Integer, Dealer_Column As Byte, Time1 As Integer, Time2 As Integer, Time3 As Integer) As Variant

Dim x As Integer, Y As Integer, N As Integer, Start As Integer, _
Finish As Integer, INTE_B() As Variant, Z As Integer


Start = UBound(AR1, 1) - (Weeks_Missed - 1) 'First missing week in the case of 1 or more rows to be calculated
Finish = UBound(AR1, 1)

'Time1 is Year3,Time2 is Month6

For x = Start To Finish

    On Error Resume Next
    
    For Y = 0 To 4                      'Calculate Other,Non-Reportable and Leveraged Fund Net
                                        'Dealers ,Asset Managers
        N = Array(4, 7, 10, 13, 18)(Y) 'location of long column, short column = long column+1

        Z = Array(0, 1, 2, 3, 4)(Y)
        
        AR1(x, Dealer_Column + Z) = AR1(x, N) - AR1(x, N + 1)
        
    Next Y
    
    AR1(x, Dealer_Column + 5) = AR1(x, Dealer_Column) / AR1(x, 3)  'Classification/OI
  
    If x >= 2 Then 'Calculate Change in Net positions for column 38 may Dealer or Asset Manger depending on if contract code is in the exceptions array
          AR1(x, Dealer_Column + 12) = AR1(x, Dealer_Column) - AR1(x - 1, Dealer_Column)
    End If
    
    On Error GoTo NET_OI_Percentage_Unavailable
    
    For Y = 0 To 2 'Net % OI per classificaion
        N = Array(38, 41, 44)(Y) 'Long % OI column locations for Dealers, Asset Mangers and Leveraged Money
        AR1(x, Dealer_Column + Array(13, 14, 15)(Y)) = AR1(x, N) - AR1(x, N + 1)
        
NET_OI_Percentage_Unavailable: On Error GoTo -1
        
    Next Y

Next x

On Error Resume Next

If UBound(AR1, 1) > Time1 Then     'Year 3 Indexes

    For Y = 0 To 1
    
        INTE_B = Stochastic_Calculations(Dealer_Column + Array(0, 5)(Y), Time1, AR1, Weeks_Missed, Cap_Extremes:=True)
        
        N = Dealer_Column + Array(6, 8)(Y)
        Z = 1
        For x = Start To Finish 'From First Missed to Most recent
              
            AR1(x, N) = INTE_B(Z)                            'Dealer index 3Y 'Dealer/OI 3Y
            Z = Z + 1
        Next x
        
        Erase INTE_B
        
    Next Y

End If
                    
If UBound(AR1, 1) > Time2 Then   'Month6 willco
    
    INTE_B = Stochastic_Calculations(Dealer_Column + 5, Time2, AR1, Weeks_Missed, Cap_Extremes:=True)  'Dealer/Oi 6M Array
    
    N = Dealer_Column + 10
    Z = 1
    For x = Start To Finish          'From First Missed to Most recent
        
        AR1(x, N) = INTE_B(Z)   '                     Dealer/Oi 6M
        Z = Z + 1
    Next x
    
    Erase INTE_B
    
End If
             
If UBound(AR1, 1) > Time3 Then   '1Y Indexes

    For Y = 0 To 1
    
        INTE_B = Stochastic_Calculations(Dealer_Column + Array(0, 5)(Y), Time3, AR1, Weeks_Missed, Cap_Extremes:=True) 'Dealer 1Y Array
      
        N = Dealer_Column + Array(7, 9)(Y)
        Z = 1
        For x = Start To Finish          'From First Missed to Most recent
          
            AR1(x, N) = INTE_B(Z)    '          Dealer 1Y and DEALER/OI
            Z = Z + 1
        Next x
        
        Erase INTE_B
        
    Next Y

End If

For x = Start To Finish                 'First Missed week to end of data set

    If x > Time1 + 6 Then               'Movement Index Calculation
        
        AR1(x, Dealer_Column + 11) = AR1(x, Dealer_Column + 8) - AR1(x - 6, Dealer_Column + 8)
        
    End If

Next x
        
'The below code block is for adding only the missed data to  an Array called Intermediate_F
 N = 1

    ReDim INTE_B(1 To Weeks_Missed, 1 To UBound(AR1, 2))

    For x = Start To Finish 'populate each row sequentially

        For Y = 1 To UBound(AR1, 2)
        
            INTE_B(N, Y) = AR1(x, Y)
            
        Next Y

        N = N + 1

    Next x

    TFF_Multi_Calculations = INTE_B
    
End Function


