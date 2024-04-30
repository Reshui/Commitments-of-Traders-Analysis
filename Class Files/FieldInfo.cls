VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_editedFieldName As String
Private p_columnIndex As Byte
Private p_databaseFieldName As String
Private p_isMissing As Boolean

#If DatabaseFile Then

    Private p_enumType As DataTypeEnum
    
    Public Property Get DataType() As DataTypeEnum
        DataType = p_enumType
    End Property
    
    Private Function EnumType() As DataTypeEnum
        If InStrB(1, p_editedFieldName, "yyyy") > 0 Then
            EnumType = adDate
        ElseIf InStrB(1, p_editedFieldName, "pct") > 0 Then
            EnumType = adNumeric
        ElseIf InStrB(1, p_editedFieldName, "price") > 0 Then
            EnumType = adCurrency
        ElseIf DoesStringContainIntegerName() Then
            EnumType = adInteger
        Else
            EnumType = adBSTR
        End If
    End Function
    
#End If

Public Property Get ColumnIndex() As Byte
    
    If Not p_isMissing Then
        ColumnIndex = p_columnIndex
    Else
        Debug.Print "ColumnIndex access was denied because the IsMissing property is set to true."
        Err.Raise 17
    End If

End Property

Public Property Get IsMissing() As Boolean
    IsMissing = p_isMissing
End Property
Public Property Get EditedName() As String
    EditedName = p_editedFieldName
End Property
Public Property Get DatabaseName() As String
    DatabaseName = p_databaseFieldName
End Property
Public Property Get DatabaseNameForSQL() As String
    DatabaseNameForSQL = "[" + DatabaseName + "]"
End Property

Public Sub Constructor(fieldName As String, ColumnIndex As Byte, DatabaseName As String, Optional param_IsMissing As Boolean = False)
    p_editedFieldName = fieldName
    p_columnIndex = ColumnIndex
    p_databaseFieldName = DatabaseName
    p_isMissing = param_IsMissing
    #If DatabaseFile Then
        p_enumType = EnumType()
    #End If
End Sub
Public Sub EditDatabaseName(value As String)
    p_databaseFieldName = value
End Sub

Private Function DoesStringContainIntegerName() As Boolean
    
    Dim intNames() As String, Z As Byte
    intNames = Split("all,old,other,trader,yymmdd", ",")
    For Z = LBound(intNames) To UBound(intNames)
        If InStrB(1, p_editedFieldName, intNames(Z)) > 0 Then
            DoesStringContainIntegerName = True
            Exit For
        End If
    Next
    
End Function
Public Sub AdjustColumnIndex(value As Byte)
     p_columnIndex = value
End Sub


