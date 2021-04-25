VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmForm 
   Caption         =   "VBATools.ru"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frmForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : frmForm
'* Created    : 13-01-2021 14:17
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Dim Time_when_me_close As Single
'

Private Sub CommandButton1_Click()
    Time_when_me_close = 0    'чтобы выйти из цикла досрочно
End Sub

Private Sub TextBox1_Change()
    Time_when_me_close = Time_when_me_close + VBA.CInt(TextBox1.Value)
End Sub

Private Sub UserForm_Activate()
    Time_when_me_close = Timer + 5    'спрячем через 5 сек
    Do
        DoEvents
        Label3.Caption = VBA.Round(Time_when_me_close - Timer, 1)
    Loop Until Timer > Time_when_me_close
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Time_when_me_close = 0    'чтобы выйти из цикла досрочно
End Sub
