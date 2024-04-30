VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadedData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public reportAbbreviation As Range
Public FullReportName As Range
Public LastCalculatedColumn As Range
Public CalculatedColumnCount As Range
Public RawDataCount As Range
Public CurrentContractName As Range
Public UsingCombined As Range
Public PendingUpdateInDatabase As Range
Public CurrentContractCode As Range
Public CurrentDatabasePath As Range

Option Explicit

Public Sub InitializeClass(reportType As String)

    Dim RR As Range, bb As Range

    With Variable_Sheet.ListObjects("Report_Abbreviation").Range

        Set RR = .columns(1).Find(reportType, , xlValues, xlWhole, , , False)
        Set bb = Range(RR, RR.Offset(, .columns.count - 1))
                
    End With
    
    With bb
        Set reportAbbreviation = .Cells(1)
        Set FullReportName = .Cells(2)
        Set LastCalculatedColumn = .Cells(3)
        Set CalculatedColumnCount = .Cells(4)
        Set RawDataCount = .Cells(5)
        Set CurrentContractName = .Cells(6)
        Set UsingCombined = .Cells(7)
        Set PendingUpdateInDatabase = .Cells(8)
        Set CurrentContractCode = .Cells(9)
        Set CurrentDatabasePath = .Cells(10)
    End With
        
    On Error Resume Next
    
    If Not IsEmpty(CurrentDatabasePath) Then
        If LenB(Dir(CurrentDatabasePath)) = 0 Then
            CurrentDatabasePath.ClearContents
        End If
    End If
    
End Sub
