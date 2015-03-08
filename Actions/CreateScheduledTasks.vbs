Option Explicit
Dim FolderAppend, IncludedByHTA, Foldername
Foldername = "VBS_Tellstick\"
FolderAppend = "\.."
'CreateScheduledTask "1-1","On","0","Sun,Sat","06:18:02","Sunrise;+,0;+,0;+,0;06:18:02"
Function CreateScheduledTask(DeviceID,Action,IterateNumber,Days,Time,Description)
	Dim ReturnValue, Value, SchedulesConfiguration, Found, SchedulesConfiguration_Contents, DeviceSchedule, DeviceSchedules, DocumentLocation, Result
	Dim objShell_wscript, fso, Parse_Description, Config_Randomize, Config_Randomize_Action, Original_Time
	Dim CurrentWeather, Config_Controller, Config_Good_Weather, Config_Bad_Weather

' Example for a task: Executes at 01.30, every day
' schtasks /Create /TN "VBS_Tellstick\DeviceID_Action_IterateNumber" /SC Weekly /D MON,TUE,WED,THU,FRI,SAT,SUN /ST 01:30 /TR "F:\Temp\vbs_tellstick\Action\ExecuteAction.vbs DeviceID Action"
' To be replaced with variables:
' DeviceID
' Action
' Days
' Time
' IterateNumber


' Define objects needed.
	Set objShell_wscript= CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	' Define the root folder.
	If (IncludedByHTA) Then
		FolderAppend = ""
		DocumentLocation = document.location
		
		DocumentLocation = fso.GetParentFolderName(Right(document.location,Len(document.location)-8))
		DocumentLocation = Replace(DocumentLocation,"/","\")	
		
	Else
		DocumentLocation = fso.GetParentFolderName(WScript.ScriptFullName)
	End If
	
	If (TypeName(CurrentWeather) = "Empty") Then	
		Set CurrentWeather = GetWeatherInfo(CityCountryID)
	End If
	'MsgBox Description
	Parse_Description = Split(Description,";")
	
	' If we didn't receive any weather information, don't calculate on it.
	If Not ( CurrentWeather.CityID = false) Then
		' Check what controller was selected
		Config_Controller = Split(Parse_Description(0),",")
		If (Config_Controller(0) = "Sunrise") Then
			Time = CurrentWeather.SunRise
		End If
		
		If (Config_Controller(0) = "Sundown") Then
			Time = CurrentWeather.SunSet
		End If
		
		' Good Weather
		If (inStr(Opt_GoodCodes, CurrentWeather.WeatherCode)>0) Then
			Config_Good_Weather = Split(Parse_Description(2),",")
			Time = DateAdd("h", Config_Good_Weather(0) & Config_Good_Weather(1),Time)
		Else
			' Bad Weather
			Config_Bad_Weather = Split(Parse_Description(2),",")
			Time = DateAdd("h", Config_Bad_Weather(0) & Config_Bad_Weather(1),Time)
		End If
	End If
	
	' Modify according to randomization
	Config_Randomize = Split(Parse_Description(1),",")
	Config_Randomize_Action = Config_Randomize(0)
	
	
	If (Config_Randomize_Action = "both") Then
	' Randomize the action to take
		Result = RandomizeBetween(0, 1)
		If (Result = 0) Then
			Config_Randomize_Action = "-"
		Else
			Config_Randomize_Action = "+"
		End If
	End If
	
	' Subtract or add to the time. (h = hour, n = minutes)
	Result = RandomizeBetween(0,Config_Randomize(1))
	Time = DateAdd("n", Config_Randomize_Action & Result,Time)
	' End of Randomization
	

	' First we create the scheduled task
	objShell_wscript.LogEvent 4,  "cmd /c schtasks /Create /TN " & Foldername & DeviceID & "_" & Action & "_" & IterateNumber & " /SC Weekly /D " & Days & " /ST " & Time & " /TR " & chr(34) & DocumentLocation & FolderAppend & "\Actions\ExecuteAction.vbs " & DeviceID & " " & Action & chr(34)
	ReturnValue = objShell_wscript.Run("cmd /c schtasks /Create /TN " & Foldername & DeviceID & "_" & Action & "_" & IterateNumber & " /SC Weekly /D " & Days & " /ST " & Time & " /TR " & chr(34) & DocumentLocation & FolderAppend & "\Actions\ExecuteAction.vbs " & DeviceID & " " & Action & chr(34) & " /F",0,True)
	If (ReturnValue <> 0) Then
		CreateScheduledTask = "Error Occured"
		MsgBox "Error Occured"
		Exit Function
	End If
	
	
	DeviceSchedules = array(" ")
	If fso.FileExists(DocumentLocation & FolderAppend & "\Configuration\SchedulesConfiguration.txt") Then
		Set SchedulesConfiguration = fso.OpenTextFile(DocumentLocation & FolderAppend & "\Configuration\SchedulesConfiguration.txt", 1)
		SchedulesConfiguration_Contents = SchedulesConfiguration.ReadAll
		SchedulesConfiguration.Close
		DeviceSchedules = split(SchedulesConfiguration_Contents, vbCrlf)
	End If
	
	Set SchedulesConfiguration = fso.CreateTextFile(DocumentLocation & FolderAppend & "\Configuration\SchedulesConfiguration.txt", true)
	
	Found = False
	
	For Each DeviceSchedule in DeviceSchedules
		If (inStr(DeviceSchedule, Foldername & DeviceID & "_" & Action & "_" & IterateNumber) > 0 ) Then
			SchedulesConfiguration.WriteLine Foldername & DeviceID & "_" & Action & "_" & IterateNumber & "{" & Description & "}"
			Found = True
		Else
			If (Len(DeviceSchedule)>2) Then
				SchedulesConfiguration.WriteLine DeviceSchedule
			End If
		End If	
	Next
	
	If Not Found Then
		SchedulesConfiguration.WriteLine Foldername & DeviceID & "_" & Action & "_" & IterateNumber & "{" & Description & "}"
	End If
	
	SchedulesConfiguration.Close
	CreateScheduledTask = True
'MsgBox DeviceID & "," & Action & "," & IterateNumber & "," & Days & "," & Time & "," & Description
End Function


Function RandomizeBetween (Min, Max)
	Randomize
	RandomizeBetween = Int((Max-Min+1)*Rnd+Min)
End Function



'CreateScheduledTask 1-1,"on",10,"mon","13:52","time;+,0;+,0;+,0"