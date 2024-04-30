VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContractInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pContractCode As String
Private pTableObject As ListObject
Private pPriceSymbol As String
Private pContractName As String
Private pUseYahooPrices As Boolean
Private pContractAvailability As String
Private pIsFavorited As Boolean

Option Explicit
Property Get HasSymbol() As Boolean
    HasSymbol = LenB(pPriceSymbol) > 0
End Property
Property Get IsFavorited() As Boolean
    IsFavorited = pIsFavorited
End Property

Property Get ContractName() As String
    ContractName = pContractName
End Property

Property Get ContractCode() As String
    ContractCode = pContractCode
End Property

Property Let ContractCode(newContractCode As String)
    pContractCode = newContractCode
End Property
Property Get priceSymbol() As String
    priceSymbol = pPriceSymbol
End Property

Property Let priceSymbol(newPriceSymbol As String)
    pPriceSymbol = newPriceSymbol
End Property

Property Get UseYahooPrices() As Boolean
    UseYahooPrices = pUseYahooPrices
End Property

Property Let UseYahooPrices(newUseYahooPrices As Boolean)
    pUseYahooPrices = newUseYahooPrices
End Property

Property Get DoesContractExistWithinReport(reportToTest As String) As Boolean
    DoesContractExistWithinReport = InStrB(1, pContractAvailability, reportToTest) > 0
End Property

Property Get ContractNameWithoutMarket() As String
    Dim fullName As String
    
    #If DatabaseFile Then
        fullName = pContractName
    #Else
        With pTableObject.DataBodyRange
            fullName = .Cells(.Rows.count, 2).Value2
        End With
    #End If
    
    ContractNameWithoutMarket = Left$(fullName, InStrRev(fullName, "-") - 1)
    
End Property

#If Not DatabaseFile Then

    Property Get TableSource() As ListObject
        Set TableSource = pTableObject
    End Property
    
    Property Set TableSource(newTableSource As ListObject)
        Set pTableObject = newTableSource
    End Property
    
    Property Get SourceWorksheet() As Worksheet
    
        If Not pTableObject Is Nothing Then
            Set SourceWorksheet = pTableObject.Parent
        Else
            Set SourceWorksheet = Nothing
        End If
        
    End Property
    
    Public Sub InitializeContract(newContractCode As String, newPriceSymbol As String, newUseYahooPrices As Boolean, Optional newTable As ListObject)
    
        pContractCode = newContractCode
        pPriceSymbol = newPriceSymbol
        pUseYahooPrices = newUseYahooPrices
        pIsFavorited = True
    
        If Not newTable Is Nothing Then Set pTableObject = newTable
        
    End Sub
#Else
    Public Sub InitializeBasicVersion(newContractCode As String, newContractName As String, newAvailability As String, newIsFavorited As Boolean, newPriceSymbol As String, newUseYahooPrices As Boolean)
    
        pContractCode = newContractCode
        pPriceSymbol = newPriceSymbol
        pUseYahooPrices = newUseYahooPrices
        pIsFavorited = newIsFavorited
        
        pContractName = newContractName
        pContractAvailability = newAvailability
    
    End Sub
#End If



