VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LabelController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Label As MSForms.Label
Attribute Label.VB_VarHelpID = -1

Property Get setControlClass(ByRef con As MSForms.Control) As LabelController
    
    Set Label = con
    Set setControlClass = Me
    
End Property

Private Sub Label_Click()

    Call LabelClickEvent(Label)
    
End Sub

Private Sub Label_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Call LabelClickEvent(Label)
    
End Sub

Private Sub LabelClickEvent(con As MSForms.Control)
    
    MsgBox con.Name & "Ç≈Ç∑ÅI", vbOKOnly
    
End Sub
