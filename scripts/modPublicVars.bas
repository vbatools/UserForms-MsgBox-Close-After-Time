Attribute VB_Name = "modPublicVars"
Option Explicit

Public lRow&    'row of active toggleButton & optionButton
Public lCol&    'column of active toggleButton
Public i&    'temp
Public bTglBusy     As Boolean         'flag to skip event handling
Public oSelTgl      As MSForms.ToggleButton    'selected toggleButton
Public frm          As MSForms.UserForm
