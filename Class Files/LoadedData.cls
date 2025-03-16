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
Public lastCalculatedColumn As Range
Public CalculatedColumnCount As Range
Public RawDataCount As Range
Public CurrentContractName As Range
Public OpenInterestType As Range
Public PendingUpdateInDatabase As Range
Public CurrentContractCode As Range
Public CurrentDatabasePath As Range
Private pRowWithinTable As Range

Option Explicit

Public Property Get RowWithinTable() As Range
    Set RowWithinTable = pRowWithinTable
End Property

Public Sub InitializeClass(eReport As ReportEnum, Optional checkForSize As Boolean = False)

    Dim rr As Range, databasePath$

    With Variable_Sheet.ListObjects("Report_Abbreviation").DataBodyRange
        Set rr = .columns(1).Find(ConvertReportTypeEnum(eReport), , xlValues, xlWhole, , , False)
        Set pRowWithinTable = Range(rr, rr.Offset(, .columns.Count - 1))
    End With
    
    With pRowWithinTable
        Set reportAbbreviation = .Cells(1)
        Set FullReportName = .Cells(2)
        Set lastCalculatedColumn = .Cells(3)
        Set CalculatedColumnCount = .Cells(4)
        Set RawDataCount = .Cells(5)
        Set CurrentContractName = .Cells(6)
        Set OpenInterestType = .Cells(7)
        Set PendingUpdateInDatabase = .Cells(8)
        Set CurrentContractCode = .Cells(9)
        Set CurrentDatabasePath = .Cells(10)
    End With
        
    On Error GoTo Exit_Initializer
    
    If Not IsEmpty(CurrentDatabasePath) Then
        With CurrentDatabasePath
            databasePath = .Value2
            If LenB(Dir$(databasePath)) = 0 Or databasePath Like "*[[]*" Or databasePath Like "*[]]*" Then
                .ClearContents
            ElseIf checkForSize Then
                CheckMsAccessSize CurrentDatabasePath.Value2
            End If
        End With
    End If
Exit_Initializer:
End Sub
