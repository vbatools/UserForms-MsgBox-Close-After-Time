Attribute VB_Name = "modMsgBoxEx"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : modMsgBoxEx
'* Created    : 13-01-2021 11:32
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Public Function MsgBoxEx(Prompt, Optional Buttons As VbMsgBoxStyle = 0, Optional Title, Optional SecondsToWait = 0) As VbMsgBoxResult
    '---------------------------------------------------------------------------------------
    ' Procedure : MsgBoxEx
    ' Purpose   : MsgBox with timeout based on WScript.Shell Popup method. Creates .VBS file
    '             in temporary folder, runs it, returns result code, deletes the file.
    ' Arguments : First three are the same as for MsgBox, 4-th is timeout in seconds.
    '           : If 4-th arg. is omitted or <=0 then waits for user action infinitely.
    ' Ret.Value : The same as of Msgbox, -1 if timeout occured.
    ' Errors    : Raises error 735 if temporary folder can't be found.

    'Назначение  : MsgBox с таймаутом на основе WScript.Shell всплывающего окна оболочки. Создает Файл .VBS
    '              - во временной папке, запускает его, возвращает код результата, удаляет файл.
    'Аргументы   : первые три такие же, как и для MsgBox, 4-й-это тайм-аут в секундах.
    '            : Если 4-й арг. опущен или <=0, а затем бесконечно ждет действий пользователя.
    'Ret. Value  : то же самое, что и в Msgbox, -1, если произошел тайм-аут.
    'Ошибки      : вызывает ошибку 735, если временная папка не может быть найдена.
    '---------------------------------------------------------------------------------------

    Dim sTmp$, ff%, WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    sTmp = Environ("temp")
    If sTmp = "" Then
        sTmp = Environ("tmp")
        If sTmp = "" Then
            sTmp = WshShell.SpecialFolders("MyDocuments")
            If sTmp = "" Then Err.Raise 735    'Can't save file to TEMP directory
        End If
    End If
    sTmp = sTmp & Format$(Now, """\~MsgBoxEx""YYMMDDHHMMSS"".vbs""")
    ff = FreeFile
    Open sTmp For Output As ff

    If IsMissing(Title) Then Title = ""

    'Popup(<Text>,<SecondsToWait>,<Title>,<Type>)

    Print #ff, "WScript.Quit CreateObject(""WScript.Shell"").Popup (""" & Str2Code(Prompt) & _
            """, " & Int(SecondsToWait) & ", """ & Str2Code(Title) & """, " & Int(Buttons) & ")"
    Close ff
    MsgBoxEx = WshShell.Run(sTmp, 0, True)
    On Error Resume Next
    Kill sTmp
End Function

Private Function Str2Code$(s)
    '---------------------------------------------------------------------------------------
    ' Procedure : Str2Code
    ' Purpose   : Replaces combinations CR+LF, LF+CR, single chars CR, LF with " & vblf & "
    '             to be used in VBS code
    '---------------------------------------------------------------------------------------

    Str2Code = Replace$( _
            Replace$( _
            Replace$( _
            Replace$( _
            Replace$(s, """", """"""), _
            vbCrLf, vbLf), _
            vbLf & vbCr, vbLf), _
            vbCr, vbLf), _
            vbLf, """ & vblf & """)
End Function
