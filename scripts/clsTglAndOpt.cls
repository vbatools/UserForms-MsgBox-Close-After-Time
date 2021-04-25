VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTglAndOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents tg As MSForms.ToggleButton
Attribute tg.VB_VarHelpID = -1
Public WithEvents ob As MSForms.OptionButton
Attribute ob.VB_VarHelpID = -1

Private Sub ob_Click()
    If bTglBusy Then Exit Sub
    i = Mid$(ob.Name, 3)
    'if another optionButton is selected, select first toggleButton in row
    If i <> lRow Then
        lRow = i
        frm.Controls("tg" & i & "1") = True
    End If
End Sub

Private Sub tg_Change()
    If bTglBusy Then Exit Sub
    bTglBusy = True
    'if user clicks already selected toggleButton
    If tg Is oSelTgl Then tg = True: bTglBusy = False: Exit Sub
    oSelTgl = False
    Set oSelTgl = tg
    lRow = Mid$(tg.Name, 3, 1)
    lCol = Right$(tg.Name, 1)
    'switch on related optionButton
    frm.Controls("ob" & lRow) = True
    bTglBusy = False
End Sub

