VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMsgBox 
   Caption         =   "MsgBoxExt Demonstration VBATools.ru"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   OleObjectBlob   =   "frmMsgBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : frmMain
'* Created    : 13-01-2021 11:32
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Private Sub btQuit_Click()
    Unload Me
End Sub

Private Sub btRun_Click()
    Dim result

    i = Choose(lRow, vbOKOnly, vbOKCancel, vbYesNo, vbYesNoCancel, vbAbortRetryIgnore, vbRetryCancel)
    i = i + cbIcon
    i = i + Choose(lCol, vbDefaultButton1, vbDefaultButton2, vbDefaultButton3)

    result = MsgBoxEx(tbText, i, tbTitle, tbSec)
    Select Case result
        Case vbAbort: result = "Abort"
        Case vbCancel: result = "Cancel"
        Case vbIgnore: result = "Ignore"
        Case vbNo: result = "No"
        Case vbOK: result = "OK"
        Case vbRetry: result = "Retry"
        Case vbYes: result = "Yes"
        Case -1: result = "Timeout"
        Case Else: result = "Unknown: " & result
    End Select
    tbResult = result
End Sub


Private Sub UserForm_Initialize()

    Static clControls As New Collection
    Dim ctrl
    With clControls
        For Each ctrl In frButtons.Controls
            .Add New clsTglAndOpt
            If TypeOf ctrl Is MSForms.OptionButton Then
                Set .Item(.Count).ob = ctrl
            ElseIf TypeOf ctrl Is MSForms.ToggleButton Then
                Set .Item(.Count).tg = ctrl
            Else: .Remove (.Count)
            End If
        Next
    End With

    ReDim arr(0 To 4, 0 To 1)
    arr(0, 0) = "(none)"
    arr(1, 0) = "Exclamation": arr(1, 1) = vbExclamation
    arr(2, 0) = "Information": arr(2, 1) = vbInformation
    arr(3, 0) = "Question": arr(3, 1) = vbQuestion
    arr(4, 0) = "Critical": arr(4, 1) = vbCritical
    With cbIcon
        .List = arr
        .ListIndex = 2
    End With

    Set frm = Me
    Set oSelTgl = tg11
    lRow = 1
    lCol = 1

End Sub
