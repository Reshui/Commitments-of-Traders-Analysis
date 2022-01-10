Attribute VB_Name = "Calculations"
Public Function Legacy_Multi_Calculations(AR1 As Variant, Weeks_Missed As Long, CommercialC As Long, _
Time1 As Long, Time2 As Long) As Variant
'======================================================================================================
'Legacy Calculations for calculated columns
'======================================================================================================
Dim X As Long, Y As Long, N As Long, Start As Long, Finish As Long, INTE_B() As Variant, Z As Long

Start = UBound(AR1, 1) - (Weeks_Missed - 1)
Finish = UBound(AR1, 1)

    'Time1 is Year3,Time2 is Month6

On Error Resume Next

    For X = Start To Finish
        
        For Y = 0 To 2 'Commercial Net,Non-Commercial Net,Non-Reportable
            N = Array(7, 4, 11)(Y)
            AR1(X, CommercialC + Y) = AR1(X, N) - AR1(X, N + 1)
        Next Y
        
        AR1(X, CommercialC + 20) = AR1(X, 27) - AR1(X, 28) 'net %OI Commercial
        AR1(X, CommercialC + 21) = AR1(X, 24) - AR1(X, 25) 'net %OI Non-Commercial
        
        AR1(X, CommercialC + 9) = AR1(X, CommercialC) / (AR1(X, 3) - AR1(X, 6))     'Commercial/OI

        If AR1(X, 4) > 0 Or AR1(X, 5) > 0 Then
            AR1(X, CommercialC + 13) = AR1(X, 4) / (AR1(X, 4) + AR1(X, 5))      'NC Long%
            AR1(X, CommercialC + 14) = 1 - AR1(X, CommercialC + 13)             'NC Short%
        End If

        If X >= 2 Then

            AR1(X, CommercialC + 15) = AR1(X, CommercialC) - AR1(X - 1, CommercialC) 'Commercial Net Change

            For Y = 7 To 8  'Commercial Gross Long % Change & Commercial Gross Short % Change
                If AR1(X - 1, Y) > 0 Then
                    AR1(X, CommercialC + 9 + Y) = (AR1(X, Y) - AR1(X - 1, Y)) / AR1(X - 1, Y)
                End If
            Next Y

        End If

        If AR1(X, 7) > 0 Or AR1(X, 8) > 0 Then

            AR1(X, CommercialC + 18) = AR1(X, 7) / (AR1(X, 7) + AR1(X, 8))                  'Commercial Long %
            AR1(X, CommercialC + 19) = 1 - AR1(X, CommercialC + 18)                         'Commercial Short %

        End If

    Next X

On Error GoTo 0

    If UBound(AR1, 1) > Time1 Then '3Y index
        
        For Y = 0 To 3
        
            INTE_B = Stochastic_Calculations(CommercialC + Array(0, 2, 9, 1)(Y), Time1, AR1, Weeks_Missed)
            
            N = Array(3, 5, 11, 4)(Y) + CommercialC         'used to calculate column number
            Z = 1 'finish-x
            For X = Start To Finish
                AR1(X, N) = INTE_B(Z)                       '[0]Commercial index 3Y  [1]Non-Reportable 3Y   < values of Y
                Z = Z + 1                                   '[2] Willco3Y            [3] Non-Commerical 3Y
            Next X
            
            Erase INTE_B
            
        Next Y

    End If

    If UBound(AR1, 1) > Time2 Then '6M index
        
        For Y = 0 To 3
            
            INTE_B = Stochastic_Calculations(CommercialC + Array(0, 2, 9, 1)(Y), Time2, AR1, Weeks_Missed)
            
            N = Array(6, 8, 10, 7)(Y) + CommercialC ' used to calculate column number
            Z = 1
            For X = Start To Finish
                AR1(X, N) = INTE_B(Z)   '[0]Commerical 6M [1]Non-Reportable 6M
                Z = Z + 1               '[2]WillCo6M      [3]Non Commercial 6M
            Next X
            
            Erase INTE_B
            
        Next Y

    End If
    
    N = CommercialC + 11 'Willco 3Y Column
    Y = N + 1            'movement index column

    For X = Start To Finish 'First Missed to most recent do Movement Index Calculations

        If X > Time1 + 6 Then
            AR1(X, Y) = AR1(X, N) - AR1(X - 6, N)
        End If

    Next X

    'The below code block is for adding only the missing data to the output array
    N = 1

    ReDim INTE_B(1 To Weeks_Missed, 1 To UBound(AR1, 2))

    For X = Start To Finish 'populate each row sequentially

        For Y = 1 To UBound(AR1, 2)
            INTE_B(N, Y) = AR1(X, Y)
        Next Y

        N = N + 1

    Next X
    
    Legacy_Multi_Calculations = INTE_B

End Function
Public Function Disaggregated_Multi_Calculations(ByVal AR1, ByVal Weeks_Missed As Long, ByVal Producer_Column As Long, Time1 As Long, Time2 As Long) As Variant

Dim X As Long, Start As Long, Finish As Long, Y As Long, N As Long, G As Long, _
INTE_B() As Variant, Z As Long, Code_Exceptions() As String, Code_Column As Long

'Time1 is Year3,Time2 is Month6

Code_Exceptions = Split("Wheat,B,RC,W,G,Cocoa", ",")

Code_Column = Producer_Column - 3

Start = UBound(AR1, 1) - (Weeks_Missed - 1) '-1 to incorpotate all missed weeks

Finish = UBound(AR1, 1)

On Error Resume Next

For X = Start To Finish

    For Y = 0 To 3 'Producer Net , Swap Net , Managed Net , Other Net
    
        N = Array(4, 6, 9, 12)(Y)
        
        AR1(X, Producer_Column + Y) = AR1(X, N) - AR1(X, N + 1)
        
    Next Y
    
    AR1(X, Producer_Column + 4) = AR1(X, Producer_Column) + AR1(X, Producer_Column + 1) 'Commercial Net
    
    If X >= 2 Then
    
        For Y = 0 To 2 'Producer Net Change,Swap Net Change,Commercial Net Change
        
            N = Array(0, 1, 4)(Y)
            
            AR1(X, Producer_Column + 20 + Y) = AR1(X, Producer_Column + N) - AR1(X - 1, Producer_Column + N)
            
        Next Y
    
        For Y = 6 To 7 'Commercial Long/Short Change
        
            AR1(X, Producer_Column + 17 + Y) = (AR1(X, Y - 2) + AR1(X, Y)) - (AR1(X - 1, Y - 2) + AR1(X - 1, Y))
        
        Next Y
    
    End If
    
    N = AR1(X, 3) - (AR1(X, 8) + AR1(X, 11) + AR1(X, 14)) 'OI without Spread Contracts
    
    AR1(X, Producer_Column + 11) = AR1(X, Producer_Column) / N      'Producer/OI
    AR1(X, Producer_Column + 12) = AR1(X, Producer_Column + 1) / N  'Swap/OI
    AR1(X, Producer_Column + 10) = AR1(X, Producer_Column + 11) + AR1(X, Producer_Column + 12)                        'Commercial/OI
    
Next X

    If Not IsError(Application.Match(AR1(UBound(AR1, 1), Code_Column), Code_Exceptions, 0)) Then
    
        For X = Start To Finish
                
            For G = 3 To 18 'changes in individual reportable positions
                
                Select Case G
                
                    Case 15, 16 'skip g= 15 and 16
                    Case Else
                        
                        If X > 1 Then AR1(X, G + 16) = AR1(X, G) - AR1(X - 1, G)
                
                End Select
            
            Next G
            'Below are the calculations for total reportable positions
            AR1(X, 15) = AR1(X, 3) - AR1(X, 17)  'Reportables long
            AR1(X, 16) = AR1(X, 3) - AR1(X, 18)  'Reportables Short
            
            AR1(X, 63) = AR1(X, 52) + AR1(X, 54) + AR1(X, 57) + AR1(X, 60)
            AR1(X, 64) = AR1(X, 53) + AR1(X, 55) + AR1(X, 58) + AR1(X, 61)
            
            If X > 1 Then
                AR1(X, 31) = AR1(X, 15) - AR1(X - 1, 15) 'Change in RPL
                AR1(X, 32) = AR1(X, 16) - AR1(X - 1, 16) 'Change in RPS
            End If
            
        Next X
               
    End If
    
    On Error GoTo 0
    
    If UBound(AR1, 1) > Time1 Then     'Year 3
    
        For Y = 0 To 7
        
            INTE_B = Stochastic_Calculations(Producer_Column + Array(0, 1, 2, 3, 4, 10, 11, 12)(Y), Time1, AR1, Weeks_Missed)  'Producer Net 3Y Array
            
            N = Producer_Column + Array(5, 6, 7, 8, 9, 14, 15, 16)(Y) ' used to calculate column number
            Z = 1
            For X = Start To Finish  'From First Missed to Most recent
                
                AR1(X, N) = INTE_B(Z)
                Z = Z + 1
            Next X
            
        Next Y
        
        Erase INTE_B
        
    End If
             
    If UBound(AR1, 1) > Time2 Then   'Month6
    
        For Y = 0 To 2
            
            INTE_B = Stochastic_Calculations(Producer_Column + Array(10, 11, 12)(Y), Time2, AR1, Weeks_Missed) 'Commercial/Oi 6M Array
            
            N = Producer_Column + Array(17, 18, 19)(Y)
            Z = 1
            For X = Start To Finish          'From First Missed to Most recent
               
                AR1(X, N) = INTE_B(Z)  '   xxx/oi
                Z = Z + 1
            Next X
            
            Erase INTE_B
            
        Next Y
       
    End If
    
    N = Producer_Column + 13
    Y = N + 1
    For X = Start To Finish                  'First Missed to most recent or Last row if Weekl
    
        If X > Time1 + 6 Then                ' Movement Index Calculation
            
                AR1(X, N) = AR1(X, Y) - AR1(X - 6, Y)
            
        End If
    
    Next X
       'The below code block is for adding only the missed data to  an Array called Intermediate_F
     N = 1
    
    ReDim INTE_B(1 To Weeks_Missed, 1 To UBound(AR1, 2))

    For X = Start To Finish 'populate each row sequentially

        For Y = 1 To UBound(AR1, 2)
        
            INTE_B(N, Y) = AR1(X, Y)
            
        Next Y

        N = N + 1

    Next X

    Disaggregated_Multi_Calculations = INTE_B
    
End Function

Public Function TFF_Multi_Calculations(AR1, Weeks_Missed As Long, Dealer_Column As Long, Time1 As Long, Time2 As Long, Time3 As Long) As Variant

Dim X As Long, INTE_B() As Variant, Y As Long, N As Long, X1 As Long, Start As Long, _
Finish As Long, Z As Long

Start = UBound(AR1, 1) - (Weeks_Missed - 1) 'First missing week in the case of 1 or more rows to be calculated
Finish = UBound(AR1, 1)

'Time1 is Year3,Time2 is Month6

For X = Start To Finish

    On Error Resume Next
    
    For Y = 0 To 4                      'Calculate Other,Non-Reportable and Leveraged Fund Net
                                        'Dealers ,Asset Managers
        N = Array(4, 7, 10, 13, 18)(Y) 'location of long column, short column = long column+1

        Z = Array(0, 1, 2, 3, 4)(Y)
        
        AR1(X, Dealer_Column + Z) = AR1(X, N) - AR1(X, N + 1)
        
    Next Y
    
    AR1(X, Dealer_Column + 5) = AR1(X, Dealer_Column) / AR1(X, 3)  'Classification/OI
  
    If X >= 2 Then 'Calculate Change in Net positions for column 38 may Dealer or Asset Manger depending on if contract code is in the exceptions array
          AR1(X, Dealer_Column + 12) = AR1(X, Dealer_Column) - AR1(X - 1, Dealer_Column)
    End If
    
    On Error GoTo NET_OI_Percentage_Unavailable
    
    For Y = 0 To 2 'Net % OI per classificaion
        N = Array(38, 41, 44)(Y) 'Long % OI column locations for Dealers, Asset Mangers and Leveraged Money
        AR1(X, Dealer_Column + Array(13, 14, 15)(Y)) = AR1(X, N) - AR1(X, N + 1)
        
NET_OI_Percentage_Unavailable: On Error GoTo -1
        
    Next Y

Next X

On Error Resume Next

If UBound(AR1, 1) > Time1 Then     'Year 3 Indexes

    For Y = 0 To 1
    
        INTE_B = Stochastic_Calculations(Dealer_Column + Array(0, 5)(Y), Time1, AR1, Weeks_Missed)
        
        N = Dealer_Column + Array(6, 8)(Y)
        Z = 1
        For X = Start To Finish 'From First Missed to Most recent
              
            AR1(X, N) = INTE_B(Z)                            'Dealer index 3Y 'Dealer/OI 3Y
            Z = Z + 1
        Next X
        
        Erase INTE_B
        
    Next Y

End If
                    
If UBound(AR1, 1) > Time2 Then   'Month6 willco
    
    INTE_B = Stochastic_Calculations(Dealer_Column + 5, Time2, AR1, Weeks_Missed)  'Dealer/Oi 6M Array
    
    N = Dealer_Column + 10
    Z = 1
    For X = Start To Finish          'From First Missed to Most recent
        
        AR1(X, N) = INTE_B(Z)   '                     Dealer/Oi 6M
        Z = Z + 1
    Next X
    
    Erase INTE_B
    
End If
             
If UBound(AR1, 1) > Time3 Then   '1Y Indexes

    For Y = 0 To 1
    
        INTE_B = Stochastic_Calculations(Dealer_Column + Array(0, 5)(Y), Time3, AR1, Weeks_Missed) 'Dealer 1Y Array
      
        N = Dealer_Column + Array(7, 9)(Y)
        Z = 1
        For X = Start To Finish          'From First Missed to Most recent
          
            AR1(X, N) = INTE_B(Z)    '          Dealer 1Y and DEALER/OI
            Z = Z + 1
        Next X
        
        Erase INTE_B
        
    Next Y

End If

For X = Start To Finish                 'First Missed week to end of data set

    If X > Time1 + 6 Then               'Movement Index Calculation
        
        AR1(X, Dealer_Column + 11) = AR1(X, Dealer_Column + 8) - AR1(X - 6, Dealer_Column + 8)
        
    End If

Next X
        
'The below code block is for adding only the missed data to  an Array called Intermediate_F
 N = 1

    ReDim INTE_B(1 To Weeks_Missed, 1 To UBound(AR1, 2))

    For X = Start To Finish 'populate each row sequentially

        For Y = 1 To UBound(AR1, 2)
        
            INTE_B(N, Y) = AR1(X, Y)
            
        Next Y

        N = N + 1

    Next X

    TFF_Multi_Calculations = INTE_B
    
End Function


