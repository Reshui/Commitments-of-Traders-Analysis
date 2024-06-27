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

Private p_editedFieldName$
Private p_columnIndex As Byte
Private p_databaseFieldName$
Private p_isMissing As Boolean

#If DatabaseFile And Not Mac Then

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
        Err.Raise 17, "FieldInfo.ColumnIndex", "ColumnIndex access was denied because the IsMissing property is set to true."
    End If

End Property

Public Property Get isMissing() As Boolean
    isMissing = p_isMissing
End Property
Public Property Get EditedName$()
    EditedName = p_editedFieldName
End Property
Public Property Get DatabaseName$()
    DatabaseName = p_databaseFieldName
End Property
Public Property Get DatabaseNameForSQL$()
    DatabaseNameForSQL = "[" + DatabaseName + "]"
End Property

Public Sub Constructor(fieldName$, ColumnIndex As Byte, DatabaseName$, Optional param_IsMissing As Boolean = False)
    p_editedFieldName = fieldName
    p_columnIndex = ColumnIndex
    p_databaseFieldName = DatabaseName
    p_isMissing = param_IsMissing
    #If DatabaseFile And Not Mac Then
        p_enumType = EnumType()
    #End If
End Sub
Public Sub EditDatabaseName(value$)
    p_databaseFieldName = value
End Sub

Private Function DoesStringContainIntegerName() As Boolean
    
    Dim intNames$(), Z As Byte
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


