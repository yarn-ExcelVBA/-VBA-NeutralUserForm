VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "CalcForm"
   ClientHeight    =   2400
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2964
   OleObjectBlob   =   "Form.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ControlCollection As New Collection

Private Sub UserForm_initialize()

''&UserForm起動時実行

    Call formInitialize
    Call setControls
    
End Sub

Private Sub formInitialize()
    
    Dim i As Long, j As Long
    Dim labelTop As Long, labelLeft As Long
    
    '/ FormSize Setting
    
    With Me
        
        .Height = 400: .Width = (.Height / 3) * 4
        
    End With

    For j = 1 To 4
    'j　縦、　i 横
        For i = 1 To 3
            
            labelTop = 40 + ((i - 1) * 80)
            labelLeft = 40 + ((j - 1) * 80)
            
            Call ControlAdd("Label", "", _
                             i, j, _
                             labelTop, labelLeft, _
                             40, 70, _
                             3, &HC5D8EB)
            
        Next i
    
    Next j
    
End Sub

Private Sub ControlAdd(conName As String, conCaption As String, _
                       conColumn As Long, conRow As Long, _
                       conTop As Long, conLeft As Long, _
                       conHeight As Long, conWidth As Long, _
                       conTextAlign As Long, conBackColor As Long)
    
    With Me.Controls.Add("Forms.Label.1")
        
        .Name = conName & conColumn & conRow: .Caption = conCaption
        
        .Top = conTop: .Left = conLeft
        
        .Height = conHeight: .Width = conWidth
        
        .Font.Name = "メイリオ": .FontSize = 14
        
        .TextAlign = conTextAlign: BackColor = conBackColor
        
        .SpecialEffect = 3: .BorderStyle = 1
        
    End With

End Sub

Private Sub setControls()

    Dim con As Control
    
    For Each con In Me.Controls
    
        If InStr(con.Name, "Label") <> 0 Then

            With New LabelController
            
               ControlCollection.Add .setControlClass(con)
                
            End With
            
        End If

    Next con
    
End Sub

Private Sub R_Click()

'Form再表示

    Dim i As Long
    For i = 1 To Me.Controls.Count - 1
        Me.Controls.Remove (1)
    Next
    UserForm_initialize
    
End Sub
