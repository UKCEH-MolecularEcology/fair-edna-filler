' FAIRe Metadata Filler — Windows silent launcher
' Double-click this file to start the tool with NO console window.
' A tray icon appears in the system taskbar; browser opens automatically.

Dim fso, shell, toolDir, python, cmd

Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Find the folder containing this script
toolDir = fso.GetParentFolderName(WScript.ScriptFullName)

' Locate Python (try py launcher first, then python, then python3)
Dim pyFound
pyFound = False
Dim pyCandidates(2)
pyCandidates(0) = "py"
pyCandidates(1) = "python"
pyCandidates(2) = "pythonw"

Dim i, result
For i = 0 To 2
    On Error Resume Next
    result = shell.Run("cmd /c " & pyCandidates(i) & " --version", 0, True)
    If Err.Number = 0 And result = 0 Then
        python = pyCandidates(i)
        pyFound = True
        Exit For
    End If
    On Error GoTo 0
Next

If Not pyFound Then
    MsgBox "Python not found." & vbCrLf & vbCrLf & _
           "Please install Python from https://www.python.org/downloads/" & vbCrLf & _
           "Make sure to tick 'Add Python to PATH' during installation.", _
           vbCritical, "FAIRe Metadata Filler"
    WScript.Quit 1
End If

' Install dependencies silently (first run only)
shell.Run python & " -m pip install --quiet pystray pillow pdfplumber anthropic", _
    0, True

' Launch the tray app without a console window (0 = hidden)
cmd = python & " """ & toolDir & "\tray_app.py"""
shell.Run cmd, 0, False
