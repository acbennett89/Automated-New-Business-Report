Option Explicit

Dim fso, shell, scriptDir
Dim pywPath, pyPath, uiScript, batPath, setupPath, cmd

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
pywPath = scriptDir & "\.venv\Scripts\pythonw.exe"
pyPath = scriptDir & "\.venv\Scripts\python.exe"
uiScript = scriptDir & "\Scripts\automation_ui.py"
batPath = scriptDir & "\run_reports.bat"
setupPath = scriptDir & "\Setup_UI_Launcher.bat"

If fso.FileExists(pywPath) And fso.FileExists(uiScript) Then
  cmd = """" & pywPath & """ """ & uiScript & """"
  shell.Run cmd, 0, False
ElseIf fso.FileExists(pyPath) And fso.FileExists(uiScript) Then
  cmd = """" & pyPath & """ """ & uiScript & """"
  shell.Run cmd, 1, False
ElseIf fso.FileExists(setupPath) Then
  cmd = "cmd /c """ & setupPath & """"
  shell.Run cmd, 1, False
ElseIf fso.FileExists(batPath) Then
  cmd = "cmd /c """ & batPath & """ ui"
  shell.Run cmd, 1, False
Else
  MsgBox "Could not find automation UI launcher files.", vbCritical, "New Biz Report Automation"
End If
