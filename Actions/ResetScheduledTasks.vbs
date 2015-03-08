' This is going to be run on Schedule: On startup / login
Option Explicit
Dim DocumentLocation, fso, objShell_wscript
Dim AllTasks, Task
Set fso = CreateObject("Scripting.FileSystemObject")
Set objShell_wscript= CreateObject("WScript.Shell")

DocumentLocation = fso.GetParentFolderName(WScript.ScriptFullName)
Sub Include (File)
			
		
			Dim moduleContents, moduleHandle
			
			' Open the script file and read the entire contents
			Set moduleHandle = fso.OpenTextFile(DocumentLocation & "\" & File, 1)
			moduleContents = moduleHandle.ReadAll()
			
			' Execute it to include it in the current running script.
			ExecuteGlobal moduleContents

End Sub

' Include the files we need.
Include "FetchScheduledTasks.vbs"
AllTasks = FetchScheduledTasks
AllTasks = HighlightActiveSchedule(AllTasks)

For Each Task in AllTasks
	' No need to send on/off commands to the applications schedules.
	If Not (Task.DeviceID = "Application") Then
		objShell_wscript.LogEvent 4, "ResetScheduledTasks.vbs Sent order to ExecuteAction.vbs to set device: [" & Task.DeviceID & "] to status [" & Task.CurrentStatus & "]"
		objShell_wscript.Run DocumentLocation & "\ExecuteAction.vbs " & chr(34) & Task.DeviceID & chr(34) & " " & chr(34) & Task.CurrentStatus & chr(34) , 0, True
	End If
Next

' Update the ApplicationConfiguration.txt and set "Reset