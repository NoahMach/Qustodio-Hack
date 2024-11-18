' VBScript to prompt user for selection and run the chosen BAT file as administrator
Option Explicit

Dim objShell, strBATFile, strBATFile1, strBATFile2, userInput

' Define paths for your BAT files
strBATFile1 = "C:\qengine\qenginestart.bat"
strBATFile2 = "C:\qengine\qenginestop.bat"

' Prompt the user to select which BAT file to run
userInput = InputBox("Would you like to start or stop the 'qengine' process?")

' Check user input and select the appropriate BAT file
If userInput = "start" Then
    strBATFile = strBATFile1
ElseIf userInput = "stop" Then
    strBATFile = strBATFile2
Else
    MsgBox "Invalid input. Please enter start or stop.", vbExclamation, "Error"
    WScript.Quit
End If

' Create a shell object
Set objShell = CreateObject("Shell.Application")

' Run the selected BAT file as administrator
objShell.ShellExecute strBATFile, "", "", "runas", 1

' Cleanup
Set objShell = Nothing
