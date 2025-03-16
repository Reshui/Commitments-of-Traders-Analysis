VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
#If Not Mac And DatabaseFile Then

    Private pConnection As ADODB.Connection
    Private pStandardCMD As New Scripting.Dictionary
    
    Public Property Set Connection(value As ADODB.Connection)
        Set pConnection = value
    End Property
    Public Property Get Connection() As ADODB.Connection
        Set Connection = pConnection
    End Property
    
    Public Sub SetCommand(value As ADODB.command, key$)
        Set pStandardCMD.item(key) = value
    End Sub
    Public Property Get GetCommand(key$) As ADODB.command
        With pStandardCMD
            If .Exists(key) Then Set GetCommand = .item(key)
        End With
    End Property
    
    Public Function CloseConnection()
        If Not pConnection Is Nothing Then
            With pConnection
                If .State = adStateOpen Then .Close
            End With
            Set pConnection = Nothing
        End If
        Set pStandardCMD = Nothing
    End Function
#End If
