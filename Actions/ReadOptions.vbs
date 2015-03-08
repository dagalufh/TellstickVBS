Dim IncludedByHTA
Dim Opt_Action_Repeat_Minutes, Opt_Action_Repeat_Times, Opt_Weather_Updater_Minutes,Opt_AutoRemote_Key, Opt_AutoRemote_Pass, Opt_CityCountry, CityCountryID, Opt_GoodCodes, Opt_PathToTDTool
Function ReadOptions
	Dim OptionsArray, Option_Split , FolderAppend, objShell_wscript, fso
	Dim AppConfiguration
	FolderAppend = "\.."
	
	' Define objects needed.
	Set objShell_wscript= CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	' Define the root folder.
	If (IncludedByHTA) Then
		FolderAppend = ""		
	Else
		DocumentLocation = fso.GetParentFolderName(WScript.ScriptFullName)
	End If	
	
	' Fetch the contents of the configuration file.
	Set AppConfiguration = fso.OpenTextFile(DocumentLocation & FolderAppend & "\Configuration\AppConfiguration.txt", 1)
	OptionsArray = Split(AppConfiguration.ReadAll,vbCrlf)	
	AppConfiguration.Close
		
	' Store the values in the globally defined variables.
	For i = 0 To UBound(OptionsArray)
		Option_Split = split(OptionsArray(i),";")
		If (UBound(Option_Split) > -1) Then 
			If Option_Split(0) = "Action_Repeat_Minutes" Then
				Opt_Action_Repeat_Minutes = Option_Split(1)
			End If
			
			If Option_Split(0) = "Action_Repeat_Times" Then
				Opt_Action_Repeat_Times = Option_Split(1)
			End If
			
			If Option_Split(0) = "Weather_Updater_Minutes" Then
				Opt_Weather_Updater_Minutes = Option_Split(1)
			End If
			
			If Option_Split(0) = "AutoRemote_Key" Then
				Opt_AutoRemote_Key = Option_Split(1)
			End If
			
			If Option_Split(0) = "AutoRemote_Pass" Then
				Opt_AutoRemote_Pass = Option_Split(1)
			End If
			
			If Option_Split(0) = "CityCountry" Then
				Opt_CityCountry = Option_Split(1)
			End If
			
			If Option_Split(0) = "CityCountryID" Then
				CityCountryID = Option_Split(1)
			End If
			
			If Option_Split(0) = "GoodCodes" Then
				Opt_GoodCodes = Option_Split(1)
			End If
			
			If Option_Split(0) = "PathToTDTool" Then
				Opt_PathToTDTool = Option_Split(1)
			End If
		
		End If
	Next
	ReadOptions = OptionsArray
End Function