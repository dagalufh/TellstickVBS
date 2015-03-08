' This is going to be run on Schedule every x minutes
Option Explicit
Dim DocumentLocation, fso, objShell_wscript, Result
Dim AllTasks, Task, Schedule, CurrentWeather

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
Include "CheckWeather.vbs"
Include "ReadOptions.vbs"
Include "CreateScheduledTasks.vbs"

AllTasks = FetchScheduledTasks
AllTasks = HighlightActiveSchedule(AllTasks)

' ReadOptions
ReadOptions

' Refresh weather information
Set CurrentWeather = GetWeatherInfo(CityCountryID)

For Each Task in AllTasks
	' No need to send on/off commands to the applications schedules.
	If Not (Task.DeviceID = "Application") Then
			For each Schedule in Task.Action_On
				ReCreateSchedule Schedule
			Next
			
			For each Schedule in Task.Action_Off
				ReCreateSchedule Schedule
			Next
	End If
Next

Function ReCreateSchedule (Schedule)
	Dim Description, SplitDays, newDays, Day
	
	' Build the description, just like before creating a new schedule.
	Description = Schedule.Controller + ";" + Schedule.Randomizer + ";"+ Schedule.Weather_Good + ";"+ Schedule.Weather_Bad + ";" + Schedule.OriginalStart
				
	' We need to split up the days and just take out the three first letters.
	SplitDays = Split(Schedule.Days,",")
	newDays = ""
	For Each Day in SplitDays
		If (Len(Day)>0) Then
			newDays = newDays & Trim(Left(Day,4)) & ","
		End If 
	Next
	newDays = Left(newDays,Len(newDays)-1)
	
	'msgBox Task.DeviceID & "," & Schedule.Action & "," & Schedule.Number & "," &  newDays & "," &  Schedule.OriginalStart & "," &  Description
	If (CreateScheduledTask(Task.DeviceID,Schedule.Action,Schedule.Number, newDays, Schedule.OriginalStart, Description)) Then
		objShell_wscript.LogEvent 4, "UpdateScheduledTasks.vbs Schedule [" & Schedule.TaskName & "] has been updated with a new starting time."
	Else
		objShell_wscript.LogEvent 4, "UpdateScheduledTasks.vbs Schedule [" & Schedule.TaskName & "] failed to be updated with a new starting time."
	End If
End Function