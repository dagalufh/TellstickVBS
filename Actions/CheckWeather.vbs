' Fetches the contents of URL and returns it.
Function GetContentsFromURL_Weather (CityCountry)
	On Error Resume Next
	URL = "http://api.openweathermap.org/data/2.5/weather?mode=xml&" & CityCountry
	Set http = CreateObject("Microsoft.XmlHttp")
	http.open "GET", URL, FALSE	
	http.send ""		
	
	GetContentsFromURL_Weather = http.responseText
	Set http = Nothing
End Function

Function GetWeatherInfo (CityID)
	Set Result = CreateObject( "Microsoft.XMLDOM" )
	Dim CurrentWeather
	' We don't want async loading.
	Result.async = False	
	
	Result.LoadXml(GetContentsFromURL_Weather("id=" & CityID))
	Set CurrentWeather = new Weather
	
	if Not (Result.hasChildNodes) Then
		CurrentWeather.CityID = false
		Set GetWeatherInfo = CurrentWeather
	Else
		Set CurrentWeather = new Weather
		DisplayNode_Weather Result.childNodes, CurrentWeather
		Set GetWeatherInfo = CurrentWeather
	End If
		
End Function
'Set Test = GetWeatherInfo("2677234")
'MsgBox Test.Sunrise

' This is used to test the input from the user on the optionspanel.
Function TestWeather()
	If (Len(CityCountry.Value)>0) Then
	
	Set Result = CreateObject( "Microsoft.XMLDOM" )
	
	' We don't want async loading.
	Result.async = False	
	
	
		Result.LoadXml(GetContentsFromURL_Weather("q=" & CityCountry.Value))
		if Not (Result.hasChildNodes) Then
			MsgBox "Not found, recheck city and country code."
			CityCountryID = "0"
			
			'MsgBox "Failed to load temporary XML file. " & vbCrlf & Result.parseError.reason
			
		Else
			Set CurrentWeather = new Weather
			
			'MsgBox Result.hasChildNodes
			
			DisplayNode_Weather Result.childNodes, CurrentWeather
			CityCountryID = CurrentWeather.CityID
			
			MsgBox "Found a match, CityID: " & CurrentWeather.CityID
		End If
	Else
		CityCountryID = "0"
	End If
End Function

' Look for specific items in the returned xml code and store them in the objects properties.
Public Sub DisplayNode_Weather (Nodes, CurrentWeather)

   Dim xNode
   For Each xNode In Nodes
	   
		if (xNode.nodeName = "city") Then			
			CurrentWeather.CityID = xNode.getAttribute("id")	
			CurrentWeather.CityName = xNode.getAttribute("name")
		End If
		
		if (xNode.nodeName = "country") Then
			CurrentWeather.Country = xNode.nodeValue
		End If
		
		if (xNode.nodeName = "sun") Then		
			CurrentWeather.SunRise = Mid(xNode.getAttribute("rise"),inStr(xNode.getAttribute("rise"), "T")+1) 			
			CurrentWeather.SunSet = Mid(xNode.getAttribute("set"),inStr(xNode.getAttribute("set"), "T")+1) 	
		End If		
		
		if (xNode.nodeName = "weather") Then
			CurrentWeather.WeatherCode = xNode.getAttribute("number")					
		End If		
		
		if (xNode.nodeName = "CurrentTemperature") Then
			CurrentWeather.WeatherCode = xNode.getAttribute("value")					
		End If				
	
		If xNode.hasChildNodes Then
				 DisplayNode_Weather xNode.childNodes, CurrentWeather
		End If
		
   Next
End Sub

' Define the class for a Weather.
Class Weather
	Public CityID
	Public SunRise
	Public SunSet
	Public WeatherCode
	Public CityName
	Public Country
	Public CurrentTemperature
End Class