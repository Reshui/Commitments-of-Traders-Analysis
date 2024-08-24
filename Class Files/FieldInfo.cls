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
Private p_columnIndex As Integer
Private p_databaseFieldName$
Private p_isMissing As Boolean
Private p_enumType As FieldType
Private p_displayInTable As Boolean
Private p_isSocrataField  As Boolean

Public Enum FieldType
    DateTimeField = 7
    IntegerField = 3
    NumericField = 131
    CurrencyField = 6
    StringField = 8
End Enum

Public Property Get IsSocrataField() As Boolean
    IsSocrataField = p_isSocrataField
End Property
Public Property Get DataType() As FieldType
    DataType = p_enumType
End Property
Public Property Let DataType(value As FieldType)
    Select Case value
        Case DateTimeField, IntegerField, NumericField, CurrencyField, StringField
            p_enumType = value
        Case Else
            Err.Raise vbObjectError + 598, "FieldInfo.DataType", "Invalid value assignment. Must be a FieldType enum."
    End Select
End Property
Public Property Get ColumnIndex() As Byte
    
    If Not p_isMissing Then
        ColumnIndex = p_columnIndex
    Else
        Debug.Print "ColumnIndex access was denied because the IsMissing property is set to true."
        Err.Raise 17, "FieldInfo.ColumnIndex", "ColumnIndex access was denied because the IsMissing property is set to true."
    End If

End Property
Public Property Let ColumnIndex(value As Byte)
    p_columnIndex = value
End Property

Public Property Get IsMissing() As Boolean
    IsMissing = p_isMissing
End Property
Public Property Get EditedName$()
    EditedName = p_editedFieldName
End Property
Public Property Get DatabaseName$()
    DatabaseName = p_databaseFieldName
End Property
Public Property Get DatabaseNameForSQL$()
    DatabaseNameForSQL = "[" + p_databaseFieldName + "]"
End Property

Public Sub Constructor(fieldName$, ColumnIndex As Integer, DatabaseName$, Optional param_IsMissing As Boolean = False, Optional displayInTable As Boolean = False, Optional fromSocrataAPI As Boolean = False)
    
    p_editedFieldName = fieldName
    p_columnIndex = ColumnIndex
    p_databaseFieldName = DatabaseName
    p_isMissing = param_IsMissing
    p_displayInTable = displayInTable
    p_isSocrataField = fromSocrataAPI
    
    Call DetermineEnumType

End Sub
Public Sub EditDatabaseName(value$)
    p_databaseFieldName = value
End Sub

Private Function DoesStringContainIntegerName() As Boolean
    
    Dim intNames$(), Z As Byte
    intNames = Split("all,old,other,trader,yymmdd,cftc_commodity_code,cftc_region_code", ",")
    For Z = LBound(intNames) To UBound(intNames)
        If Not (Z = UBound(intNames) And p_isSocrataField) Then
            If InStrB(1, p_editedFieldName, intNames(Z)) <> 0 Then
                DoesStringContainIntegerName = True
                Exit For
            End If
        End If
    Next
    
End Function
Private Sub DetermineEnumType()
    If InStrB(1, p_editedFieldName, "yyyy") <> 0 Then
        p_enumType = FieldType.DateTimeField
    ElseIf InStrB(1, p_editedFieldName, "pct") <> 0 Then
        p_enumType = FieldType.NumericField
    ElseIf InStrB(1, p_editedFieldName, "price") <> 0 Then
        p_enumType = FieldType.CurrencyField
    ElseIf DoesStringContainIntegerName() Then
        p_enumType = FieldType.IntegerField
    Else
        p_enumType = FieldType.StringField
    End If
End Sub

'#If DatabaseFile Then
'    Public Function IsDatabaseTypeMismatch(recordSetField As Object)
'        IsDatabaseTypeMismatch = recordSetField.Type <> p_enumType
'    End Function
'#End If
