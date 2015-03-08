'Set objShell_wscript= CreateObject("WScript.Shell")
'Set fso = CreateObject("Scripting.FileSystemObject")

'Set SchedulesConfiguration = fso.CreateTextFile("C:\Tellstick\TestFile.txt", true)
'SchedulesConfiguration.WriteLine DeviceID & " Action: " & Action
'SchedulesConfiguration.Close
'Minutes = 1
'MsgBox "Before"
'Wscript.Sleep (1000*60)*Minutes
'MsgBox "After"
Dim args, DeviceID, Action, fso, objShell_wscript

Set fso = CreateObject("Scripting.FileSystemObject")
DocumentLocation = fso.GetParentFolderName(WScript.ScriptFullName)
Set objShell_wscript= CreateObject("WScript.Shell")

args = WScript.Arguments.Count
If args <> 2Then
  objShell_wscript.LogEvent 4,"Failed Call to vbs: ExecuteAction.vbs DeviceID Action"
  wscript.Quit
End If

DeviceID = Trim(WScript.Arguments.Item(0))
Action = Trim(WScript.Arguments.Item(1))

' This row should be appended with a call to tstools.exe or tdtools, not sure the name right now.
objShell_wscript.LogEvent 4, "Sent Command [" & Action & "] to device: [" & DeviceID & "]"