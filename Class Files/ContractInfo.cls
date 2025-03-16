VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContractInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pContractCode$
Private pTableObject As ListObject
Private pPriceSymbol$
Private pContractName$
Private pUseYahooPrices As Boolean
Private pContractAvailability$
Private pIsFavorited As Boolean

Option Explicit
Property Get AvailabilityDelimited$()
    AvailabilityDelimited = pContractAvailability
End Property
Property Get HasSymbol() As Boolean
    HasSymbol = LenB(pPriceSymbol) <> 0
End Property
Property Get IsFavorited() As Boolean
    IsFavorited = pIsFavorited
End Property

Property Get contractName$()
    contractName = pContractName
End Property

Property Get contractCode$()
    contractCode = pContractCode
End Property

Property Let contractCode(newContractCode$)
    pContractCode = newContractCode
End Property
Property Get PriceSymbol$()
    PriceSymbol = pPriceSymbol
End Property

Property Let PriceSymbol(newPriceSymbol$)
    pPriceSymbol = newPriceSymbol
End Property

Property Get UseYahooPrices() As Boolean
    UseYahooPrices = pUseYahooPrices
End Property

Property Let UseYahooPrices(newUseYahooPrices As Boolean)
    pUseYahooPrices = newUseYahooPrices
End Property

Property Get DoesContractExistWithinReport(reportToTest$) As Boolean
    DoesContractExistWithinReport = InStrB(1, pContractAvailability, reportToTest) <> 0
End Property

Property Get ContractNameWithoutMarket$()
    Dim fullName$
    
    #If DatabaseFile Then
        fullName = pContractName
    #Else
        With pTableObject.DataBodyRange
            fullName = .Cells(.Rows.Count, 2).Value2
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
    
    Public Sub InitializeContract(newContractCode$, newPriceSymbol$, newUseYahooPrices As Boolean, Optional newTable As ListObject)
    
        pContractCode = newContractCode
        pPriceSymbol = newPriceSymbol
        pUseYahooPrices = newUseYahooPrices
        pIsFavorited = True
    
        If Not newTable Is Nothing Then Set pTableObject = newTable
        
    End Sub
#Else
    Public Sub InitializeBasicVersion(newContractCode$, newContractName$, newAvailability$, newIsFavorited As Boolean, newPriceSymbol$, newUseYahooPrices As Boolean)
    
        pContractCode = newContractCode
        pPriceSymbol = newPriceSymbol
        pUseYahooPrices = newUseYahooPrices
        pIsFavorited = newIsFavorited
        
        pContractName = newContractName
        pContractAvailability = newAvailability
    
    End Sub
#End If



