VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StoredError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public pSource$
Public pDescription$
Public pNumber As Long
Public pHelpFile$
Public pHelpContext$
Public Sub Constructor(e As ErrObject)
    With e
        pSource = .Source
        pDescription = .Description
        pNumber = .Number
        pHelpFile = .HelpFile
        pHelpContext = .HelpContext
    End With
End Sub
Public Function HeldError() As ErrObject
    
    With Err
        .Description = pDescription
        .Source = pSource
        .Number = pNumber
        .HelpContext = pHelpContext
        .HelpFile = pHelpFile
    End With
    
    Set HeldError = Err
    
End Function
Public Function ErrorAvailable() As Boolean
    ErrorAvailable = pNumber <> 0
End Function
'Public Function Propogate(source$)
'    PropagateError HeldError, source
'End Function


