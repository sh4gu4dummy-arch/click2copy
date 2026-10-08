' Click2Copy.vbs - start Click2Copy with no console window.
' Double-click this file (or pin a shortcut to it). Closing the app is the only
' way to quit; there is no console window to keep open.
' If Electron is not installed yet, runs "Start Click2Copy.bat" visibly so
' setup can happen. Any arguments (e.g. a .c2copy file) are passed through.
Option Explicit

Dim fso, sh, appDir, exe, bat, q, args, i
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
q = Chr(34)

appDir = fso.GetParentFolderName(WScript.ScriptFullName)
exe = appDir & "\node_modules\electron\dist\electron.exe"
bat = appDir & "\Start Click2Copy.bat"

args = ""
For i = 0 To WScript.Arguments.Count - 1
  args = args & " " & q & WScript.Arguments(i) & q
Next

sh.CurrentDirectory = appDir

If fso.FileExists(exe) Then
  ' electron.exe is a GUI app: no console. Don't wait for it to exit.
  sh.Run q & exe & q & " " & q & appDir & q & args, 1, False
ElseIf fso.FileExists(bat) Then
  ' First run / missing Electron: visible console so npm setup is shown.
  sh.Run "cmd.exe /c " & q & q & bat & q & args & q, 1, False
Else
  MsgBox "Could not find Electron or ""Start Click2Copy.bat"" in:" & vbCrLf & appDir, vbExclamation, "Click2Copy"
End If
