' Create with folder /TN "Foldername\TaskName"
' Query with folder /TN Foldername\
Option Explicit
Dim Foldername, ReturnValue, Value, IncludedByHTA, FolderAppend
Foldername = "VBS_Tellstick\"
FolderAppend = "\.."

Function FetchScheduledTasks
	Dim objShell_wscript, fso, CommandResult, ReturnValue, objParser, TellstickTemp, Node, ScheduledTasksArray, Configuration, ConfigurationRow, SchedulesConfiguration_Contents, SchedulesConfiguration, ScheduledTaskObject, DeviceSchedules
	Dim Action_On, Action_Off
	Dim ConfigurationItems
	Dim DocumentLocation
	
	ScheduledTasksArray = array()	
	
	' Define objects needed.
	Set objShell_wscript= CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objParser = CreateObject( "Microsoft.XMLDOM" )
	
	' Define the root folder.
	If (IncludedByHTA) Then
		FolderAppend = ""
		'DocumentLocation = document.location
		
		'DocumentLocation = fso.GetParentFolderName(Right(document.location,Len(document.location)-8))
		'DocumentLocation = Replace(DocumentLocation,"/","\")
		
	Else
		DocumentLocation = fso.GetParentFolderName(WScript.ScriptFullName)
	End If	
	
	ReturnValue = objShell_wscript.Run("cmd /c schtasks /Query /xml one /TN " & Foldername & " > " & DocumentLocation & FolderAppend & "\Actions\TellstickTemp.xml",0,True)
	If (ReturnValue <> 0) Then
		FetchScheduledTasks = "Error Occured, no schedules found"
		Exit Function
	End If

	Set TellstickTemp = CreateObject( "Microsoft.XMLDOM" )
	
	' We don't want async loading.
	TellstickTemp.async = False	
	
	If (TellstickTemp.Load(DocumentLocation & FolderAppend & "\Actions\TellstickTemp.xml")) Then
		' Lets do some magic with the XML here.
		
		'For Each Node in TellstickTemp.childNodes
		'	MsgBox Node.nodeName
		'Next
		
		DisplayNode TellstickTemp.childNodes,ScheduledTasksArray, 0, 0
		
		If (UBound(ScheduledTasksArray) < 0) Then
			If (fso.FileExists(DocumentLocation & FolderAppend & "\Configuration\SchedulesConfiguration.txt")) Then
				fso.DeleteFile(DocumentLocation & FolderAppend & "\Configuration\SchedulesConfiguration.txt")
			End If
		Else
			' Fetch the configurationfile
			Set SchedulesConfiguration = fso.OpenTextFile(DocumentLocation & FolderAppend & "\Configuration\SchedulesConfiguration.txt", 1)
			SchedulesConfiguration_Contents = SchedulesConfiguration.ReadAll
			SchedulesConfiguration.Close
			
			
			DeviceSchedules = split(SchedulesConfiguration_Contents, vbCrlf)
			' For each device
			For Each ScheduledTaskObject in ScheduledTasksArray
				' For each row in the configurationfile
				For Each ConfigurationRow in DeviceSchedules
					 'For each on:action
					For Each Action_On in ScheduledTaskObject.Action_On
						'MsgBox ConfigurationRow
						If (Instr(1,ConfigurationRow, Action_On.TaskName,1) > 0) Then
							Configuration = Mid(ConfigurationRow, Instr(ConfigurationRow, "{")+1, (Instr(ConfigurationRow, "}")-Instr(ConfigurationRow, "{"))-1)
							'MsgBox Configuration
							
							ConfigurationItems = Split(Configuration,";")
							Action_On.Controller = ConfigurationItems(0)
							Action_On.Randomizer = ConfigurationItems(1)
							Action_On.Weather_Good = ConfigurationItems(2)
							Action_On.Weather_Bad = ConfigurationItems(3)
							Action_On.OriginalStart = ConfigurationItems(4)
							
							Exit For
						End If
					Next

					' For each off:action
					For Each Action_Off in ScheduledTaskObject.Action_Off
						If (Instr(1,ConfigurationRow, Action_Off.TaskName,1) > 0) Then
							Configuration = Mid(ConfigurationRow, Instr(ConfigurationRow, "{")+1, (Instr(ConfigurationRow, "}")-Instr(ConfigurationRow, "{"))-1)
							'MsgBox Configuration
							
							ConfigurationItems = Split(Configuration,";")
							Action_Off.Controller = ConfigurationItems(0)
							Action_Off.Randomizer = ConfigurationItems(1)
							Action_Off.Weather_Good = ConfigurationItems(2)
							Action_Off.Weather_Bad = ConfigurationItems(3)
							Action_Off.OriginalStart = ConfigurationItems(4)
							
							Exit For
						End If
					Next
				Next		
			Next
		End If
		
		CommandResult = ScheduledTasksArray
		
	Else
		CommandResult = "Failed to load temporary XML file. " & vbCrlf & TellstickTemp.parseError.reason
	End If
	
	
	
	' Return it to caller.
	FetchScheduledTasks = CommandResult
End Function

'Dim ArrayOfTasks
'ArrayOfTasks = FetchScheduledTasks
'ArrayOfTasks = HighlightActiveSchedule(ArrayOfTasks)
' https://msdn.microsoft.com/en-us/library/aa468547.aspx


' Debug output
'For Each Value in ReturnValue
'	MsgBox "The Task " & Value.TaskName & " will start " & Value.Start & " on the following days: " & Value.Days
'Next



Public Sub DisplayNode (Nodes, ScheduledTasksArray, CurrentObject, CurrentDevice)

   Dim xNode, DeviceID, CurrentAction, CurrentNumber,  Found, ScheduledTaskObject, ScheduleName
   For Each xNode In Nodes
	   
		' Debug output
		'MsgBox xNode.nodeName & " Type of Node: " & xNode.nodeType
		If xNode.nodeType = 3 and xNode.parentNode.nodeName = "StartBoundary" Then
			'MsgBox xNode.parentNode.nodeName & ":" & xNode.nodeValue
			CurrentObject.Start = xNode.nodeValue
		End If
		
		If xNode.nodeType = 1 and xNode.parentNode.nodeName = "DaysOfWeek" Then
		'MsgBox "NodeElement: " & xNode.nodeName & " Parent Node: " & xNode.parentNode.nodeName
		CurrentObject.Days = CurrentObject.Days & ", " & xNode.nodeName
		End If
		' Get the name
		If xNode.nodeType = 8 Then
			'MsgBox "Comment: " & xNode.nodeValue
			
			' Should we create an object for this DeviceID or does it already exist?
			'VBS_Tellstick\DeviceID_Action_IterateNumber
			ScheduleName = Split(xNode.nodeValue,"_")
			'DeviceID = Mid(xNode.nodeValue,inStr(xNode.nodeValue,"\")+1,(inStr(inStr(xNode.nodeValue,"\"),xNode.nodeValue,"_"))
			DeviceID = Mid(ScheduleName(1),InStr(ScheduleName(1),"\")+1)
			
			CurrentAction = ScheduleName(Ubound(ScheduleName)-1)
			CurrentNumber = ScheduleName(Ubound(ScheduleName))
			'MsgBox CurrentAction
			
			Found = False
			
			For Each ScheduledTaskObject in ScheduledTasksArray
				If (ScheduledTaskObject.DeviceID = DeviceID) Then
					Found = True
				End If
			Next
			
			If Not Found Then
				Set CurrentDevice = New Device
				CurrentDevice.DeviceID = DeviceID
				
				ReDim Preserve ScheduledTasksArray(UBound(ScheduledTasksArray)+1)
				Set ScheduledTasksArray(UBound(ScheduledTasksArray)) = CurrentDevice				
				
			End If
			
			' We always create one object atleast
			Set CurrentObject = New ScheduledTask
			CurrentObject.TaskName = trim(xNode.nodeValue)
			CurrentObject.Action = CurrentAction
			CurrentObject.Number = Trim(CurrentNumber)
			If (CurrentAction = "On") Then
				CurrentDevice.Add_Action_On CurrentObject
			End If
			
			If (CurrentAction = "Off") Then
				CurrentDevice.Add_Action_Off CurrentObject
			End If
		End If	
		
		If xNode.hasChildNodes Then
				 DisplayNode xNode.childNodes, ScheduledTasksArray, CurrentObject, CurrentDevice
		End If
		
   Next
End Sub


Function CountActions (DeviceID, Action)
	Dim Task, Found
	Found = False
	For Each Task in ArrayOfTasks
		If (Task.DeviceID = DeviceID) Then
			If (Action = "On") Then
				Found = True
				CountActions = Ubound(Task.Action_On)+1
			End If
			If (Action = "Off") Then
				Found = True
				CountActions = Ubound(Task.Action_Off)+1
			End If
		End If
	Next
	
	If Not Found Then
		CountActions = 0
	End If
End Function



Function HighlightActiveSchedule (ScheduledArray)
	Dim Mon, Tue, Wed, Thu, Fri, Sat, Sun, Action, Action_Days, Day, Temp, Device, Thing, DaysOfWeek, i , TodayReached, StartDayOfWeek, CurrentDayOfWeek, Time, DeviceStatus, Actions, CurrentTime
	' First, For each device in the already created array (FetchScheduledTasks) add each time for the correct day.
    If Not (TypeName(ScheduledArray) = "String") Then
        For Each Device in ScheduledArray

            ' Define one array for each day of the week
            Mon = array()
            Tue = array()
            Wed = array()
            Thu = array()
            Fri = array()
            Sat = array()
            Sun = array()

            ' For each action on and for each action off
            If (Ubound(Device.Action_On) >= 0) Then

                AddToDay Device.Action_On, Mon,Tue,Wed,Thu,Fri,Sat,Sun
            End If
            If (Ubound(Device.Action_Off) >= 0) Then
                AddToDay Device.Action_Off, Mon,Tue,Wed,Thu,Fri,Sat,Sun
            End If

            ' Now we need to bubblesort each of those days.
            Mon = BubbleSort_Days(Mon)
            Tue = BubbleSort_Days(Tue)
            Wed = BubbleSort_Days(Wed)
            Thu = BubbleSort_Days(Thu)
            Fri = BubbleSort_Days(Fri)
            Sat = BubbleSort_Days(Sat)
            Sun = BubbleSort_Days(Sun)

            DaysOfWeek = array(Sun,Mon,Tue,Wed,Thu,Fri,Sat)

            CurrentDayOfWeek = Weekday(date) -1 ' Zero-based day of week number
            CurrentTime = Right(String(1,"0") & DatePart("h",Now()),2) & ":" & Right(String(1,"0") & DatePart("n",Now()),2)
            TodayReached = False

            DeviceStatus = ""

            If (CurrentDayOfWeek=6) Then
                StartDayOfWeek = 0
            Else
                StartDayOfWeek = CurrentDayOfWeek+1
            End If

            Do

                If (StartDayOfWeek = CurrentDayOfWeek) Then
                    TodayReached = True
                End If

                Day = DaysOfWeek(StartDayOfWeek)
                For Each Time in Day
                    'MsgBox Time
                    Actions = Split(Time, ";")
                    If (TodayReached) Then
                        If (CurrentTime > Actions(0)) Then
                            DeviceStatus = Actions(1)
                        Else
                            Exit For
                        End If
                    Else	
                        DeviceStatus = Actions(1)
                    End If
                Next

                If (StartDayOfWeek = 6) Then
                    StartDayOfWeek = 0
                Else
                    StartDayOfWeek = StartDayOfWeek+1
                End If

            Loop While TodayReached = False
            Device.CurrentStatus = DeviceStatus
            'MsgBox "Device " & Device.DeviceID & " has status: "  & DeviceStatus
            'MsgBox Temp

        Next
	End If
	HighlightActiveSchedule = ScheduledArray
End Function

Function BubbleSort_Days (ArrayToSort)
	Dim Swap, CurrentValue, NextValue, TempValue, j
	' Foundations for bubblesort borrowed from https://helloacm.com/bubble-sort-in-vbscript/
	Do
		' Reset the swap variable to false each turn
	  Swap = False
	  
	  ' For each value in the array
	  For j = LBound(ArrayToSort) to UBound(ArrayToSort) - 1
		  ' If the current value is larger than the next value in the array
		  CurrentValue = Left(ArrayToSort(j),inStr(ArrayToSort(j),";")-1)
		  NextValue = Left(ArrayToSort(j + 1),inStr(ArrayToSort(j + 1),";")-1)
		  'MsgBox CurrentValue  & ">" & NextValue
		  If CurrentValue > NextValue Then
		  
			' Store next arrayvalue in a temporary variable
			 TempValue = ArrayToSort(j + 1)
			 ' Update next array value with the current position value
			 ArrayToSort(j + 1) = ArrayToSort(j)
			 ' Update current position value with that from the Temporary variable, that is, the value that was previously on the next position.
			 ArrayToSort(j) = TempValue
			 ' Make a note that we did a swap this time around.
			 Swap = True
		  End If
	  Next
	Loop Until Not Swap
	
	BubbleSort_Days = ArrayToSort
End Function

Sub AddToDay (ActionArray, Mon,Tue,Wed,Thu,Fri,Sat,Sun)
	Dim Action, Action_Days, Day
		For Each Action in ActionArray
			Action_Days = Split(Action.Days,",")
			For Each Day in Action_Days
				'MsgBox "Day: " & Day
				If (inStr(1,Day,"Monday")>0) Then
					ReDim Preserve Mon(UBound(Mon)+1)
					Mon(UBound(Mon)) = Mid(Action.Start,inStr(Action.Start, "T")+1) & ";" & Action.Action
				End If
				If (inStr(1,Day,"Tuesday")>0) Then
					ReDim Preserve Tue(UBound(Tue)+1)
					Tue(UBound(Tue)) = Mid(Action.Start,inStr(Action.Start, "T")+1) & ";" & Action.Action
				End If
				If (inStr(1,Day,"Wednesday",1)>0) Then
					ReDim Preserve Wed(UBound(Wed)+1)
					Wed(UBound(Wed)) = Mid(Action.Start,inStr(Action.Start, "T")+1) & ";" & Action.Action
				End If
				If (inStr(1,Day,"Thursday")>0) Then
					ReDim Preserve Thu(UBound(Thu)+1)
					Thu(UBound(Thu)) = Mid(Action.Start,inStr(Action.Start, "T")+1) & ";" & Action.Action
				End If
				If (inStr(1,Day,"Friday")>0) Then
					ReDim Preserve Fri(UBound(Fri)+1)
					Fri(UBound(Fri)) = Mid(Action.Start,inStr(Action.Start, "T")+1) & ";" & Action.Action
				End If
				If (inStr(1,Day,"Saturday")>0) Then
					ReDim Preserve Sat(UBound(Sat)+1)
					Sat(UBound(Sat)) = Mid(Action.Start,inStr(Action.Start, "T")+1) & ";" & Action.Action
				End If
				If (inStr(1,Day,"Sunday")>0) Then
					ReDim Preserve Sun(UBound(Sun)+1)
					Sun(UBound(Sun)) = Mid(Action.Start,inStr(Action.Start, "T")+1) & ";" & Action.Action
				End If
			Next
		Next
	End Sub

' Class
Class ScheduledTask
	Public TaskName
	Public Start
	Public OriginalStart
	Public Days
	Public Description
	Public Action
	Public Number
	Public Controller
	Public Randomizer
	Public Weather_Good
	Public Weather_Bad
End Class

Class Device
	Public DeviceID
	Public Action_On
	Public Action_Off
	Public CurrentStatus
	
	' Initialize the Class
	Private Sub Class_Initialize
		Action_On = array()
		Action_Off = array()
		CurrentStatus = "Unknown"
	End Sub
	
	Public Sub Add_Action_On(TaskObject)
		ReDim Preserve Action_On(UBound(Action_On)+1)
		Set Action_On(UBound(Action_On)) = TaskObject
	End Sub
	
	Public Sub Add_Action_Off(TaskObject)
		ReDim Preserve Action_Off(UBound(Action_Off)+1)
		Set Action_Off(UBound(Action_Off)) = TaskObject
	End Sub	
	
End Class
' Working with times or dates:
'CurrentTime = "10:42"
'DateAdd("h",-1,CurrentTime) ' Subtracts one hour from CurrentTime
'DateAdd("h",1,CurrentTime) ' Adds one hour to CurrentTime