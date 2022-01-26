Dim objShell
Dim userIn

Set objShell = CreateObject("WScript.Shell")
userIn = InputBox("1. Set timer" & vbTab & "2. Abort shutdown", "Auto Shutdown")
If userIn = "" Then
   Wscript.Quit
ElseIf userIn = "1" Then
   userIn = InputBox("Set time(hour) for shutdown.", "Shutdown Timer", "1")
   objShell.Run "C:\WINDOWS\system32\shutdown.exe -s -t " & (userIn * 3600)
ElseIf userIn = "2" Then
   objShell.Run "C:\WINDOWS\system32\shutdown.exe -a"
End If

Set objShell = Nothing

Test1
