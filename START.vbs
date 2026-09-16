' IP Address Allocation Tool - Launcher
' This bypasses cmd.exe entirely, so no code-page / encoding issues.
' On error it shows a popup window with the reason.

Option Explicit
Dim sh, fso, here, py, cmd, rc, msg

Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)

' --- pick a Python interpreter ---
py = ""
If fso.FileExists(here & "\.venv\Scripts\pythonw.exe") Then
    py = here & "\.venv\Scripts\pythonw.exe"
ElseIf fso.FileExists(here & "\.venv\Scripts\python.exe") Then
    py = here & "\.venv\Scripts\python.exe"
End If

If py = "" Then
    ' try pythonw / python on PATH
    On Error Resume Next
    rc = sh.Run("pythonw -c ""import sys""", 0, True)
    If Err.Number = 0 And rc = 0 Then
        py = "pythonw"
    Else
        Err.Clear
        rc = sh.Run("python -c ""import sys""", 0, True)
        If Err.Number = 0 And rc = 0 Then py = "python"
    End If
    On Error GoTo 0
End If

If py = "" Then
    msg = "Python not found on this computer." & vbCrLf & vbCrLf & _
          "Please install Python 3.8 or newer from:" & vbCrLf & _
          "    https://www.python.org/downloads/" & vbCrLf & vbCrLf & _
          "IMPORTANT: check ""Add Python to PATH"" during install." & vbCrLf & vbCrLf & _
          "After that, double-click START.vbs again."
    MsgBox msg, 16, "IP Tool - Cannot Start"
    WScript.Quit 1
End If

' --- check dependencies, install if missing ---
cmd = """" & py & """ -c ""import PySide6, openpyxl, docx"""
On Error Resume Next
rc = sh.Run(cmd, 0, True)
On Error GoTo 0

If rc <> 0 Then
    ' Use install_deps.py: it upgrades pip and forces prebuilt wheels,
    ' so lxml will not try to compile and demand Visual C++.
    cmd = """" & py & """ """ & here & "\install_deps.py"""
    On Error Resume Next
    rc = sh.Run(cmd, 1, True)   ' window visible, user sees the report
    On Error GoTo 0
    If rc <> 0 Then
        msg = "Failed to install dependencies." & vbCrLf & vbCrLf & _
              "Most common cause: lxml has no prebuilt wheel for" & vbCrLf & _
              "this Python version, so pip tried to compile it and" & vbCrLf & _
              "no C++ compiler is installed." & vbCrLf & vbCrLf & _
              "Recommended fix: install Python 3.11 or 3.12 from" & vbCrLf & _
              "    https://www.python.org/downloads/" & vbCrLf & _
              "then delete the .venv folder and run START.vbs again." & vbCrLf & vbCrLf & _
              "Run CHECK.bat for a full environment report."
        MsgBox msg, 16, "IP Tool - Dependency Error"
        WScript.Quit 1
    End If
End If

' --- launch the app ---
sh.CurrentDirectory = here
cmd = """" & py & """ """ & here & "\app_gui.py"""
On Error Resume Next
rc = sh.Run(cmd, 0, False)
If Err.Number <> 0 Then
    msg = "Failed to launch the application." & vbCrLf & vbCrLf & _
          "Error: " & Err.Description & vbCrLf & vbCrLf & _
          "Command: " & cmd & vbCrLf & vbCrLf & _
          "Try running START.bat instead - it shows the full error."
    MsgBox msg, 16, "IP Tool - Launch Error"
    WScript.Quit 1
End If
On Error GoTo 0

WScript.Quit 0
