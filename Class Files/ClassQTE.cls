VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassQTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents MyQueryTable As Excel.QueryTable
Attribute MyQueryTable.VB_VarHelpID = -1
Private LatestContractsQuery As Boolean
Private Procedure_To_Run$
Private Non_ListObject_QueryTable As Boolean
Public Refreshed As Boolean

Friend Sub HookUpQueryTable(Query_T As QueryTable, Procedure_2_Run$, QueryTableInListObject As QueryTable)

    Procedure_To_Run = Procedure_2_Run
    Non_ListObject_QueryTable = QueryTableInListObject
    Set MyQueryTable = QueryT
Exit_Hook_Up:
    LatestContractsQuery = False
End Sub
Private Sub MyQueryTable_BeforeRefresh(Cancel As Boolean)
    Refreshed = False
End Sub
Private Sub MyQueryTable_AfterRefresh(ByVal success As Boolean)
    
    Refreshed = True

    If LatestContractsQuery Then
        
        #If DatabaseFile Then
            Latest_Contracts_After_Refresh MyQueryTable, success
            Set MyQueryTable = Nothing
        #End If
        
    End If

End Sub
Sub HookUpLatestContracts(QT As QueryTable)
    Set MyQueryTable = QT
    LatestContractsQuery = True
End Sub


