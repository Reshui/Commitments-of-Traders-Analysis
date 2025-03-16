VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents chx As MSForms.CheckBox
Attribute chx.VB_VarHelpID = -1
Private pActivatedColor&
Private pOffColor&

Public Sub Constructor(check As MSForms.CheckBox, Optional activatedColor& = 0, Optional offColor& = 0)
    
    Dim reddishColor&, blueishColor&
    Set chx = check
    
    reddishColor = RGB(236, 219, 219)
    blueishColor = RGB(208, 248, 255)
    
    pOffColor = IIf(offColor = 0, reddishColor, offColor)
    pActivatedColor = IIf(activatedColor = 0, blueishColor, activatedColor)
    
    ChangeColor
    
End Sub
Private Sub chx_Change()
    ChangeColor
End Sub
Sub ChangeColor()
    With chx
        .BackColor = IIf(.value = True, pActivatedColor, pOffColor)
    End With
End Sub
