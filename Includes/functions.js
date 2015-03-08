function CreateSchedule() {
	// Verify that the values needed are selected.
	
	var DeviceID = $("#Select_Device").val();
	var Controller = $("#Select_Controller").val();
	var Action = $("#Select_Action").val();
	
	var IterateNumber = 0;
	var Days = "";
	var Time = $("#Time").val();
	var Description = "";
	
	var Time_Regex = new RegExp(/\d\d[":"]\d\d/);
	if (Time_Regex.test(Time) === false) {
		alert("The format of Time is not valid. Need to be in the format: 12:38");
		return false;
	}
	
	
	var Randomizer_Controller = $("#Select_Randomizer").val();
	var Randomizer_Value = $("#Select_Randomizer_Value").val();
	var Weather_Good_Controller = $("#Select_Weather_Good").val();
	var Weather_Good_Value = $("#Select_Weather_Good_Time").val();
	var Weather_Bad_Controller = $("#Select_Weather_Bad").val();
	var Weather_Bad_Value = $("#Select_Weather_Bad_Time").val();
	
	$("#DayOfWeek:checked").each(function() {
		Days += $(this).val() + ",";
	});
	
	Days = Days.substring(0,Days.length-1);
	
	// Check if the user has selected any days
	if (Days.length == 0) {
		alert("You must select atleast one day of the week.");
		return false;
	}
	// IterateNumber - Count number of scheduled events with this deviceID
	IterateNumber = CountActions(DeviceID, Action);

	// Description should be built up with configuration information
	Description = Controller + ";" + Randomizer_Controller + "," + Randomizer_Value + ";"+ Weather_Good_Controller + "," + Weather_Good_Value + ";"+ Weather_Bad_Controller + "," + Weather_Bad_Value + ";" + Time;
	
	
	// This should only be allowed if: user has selected atleast one day.. AND that the TDTOOL.exe is located! If we don't have that, no use in creating a schedule.
	CreateScheduledTask(DeviceID,Action,IterateNumber,Days,Time,Description);
	
	ArrayOfTasks = FetchScheduledTasks();
	if (CheckSystemSchedules(ArrayOfTasks)) {
				ArrayOfTasks = FetchScheduledTasks();
	}
	
	VBScript:ScheduleListBuilder(ArrayOfTasks);
}

function ChangeController() {
	if ($("#Select_Controller").val() == "Time") {
		//$(".CanHide").show();
		
	} else if  ($("#Select_Controller").val() == "Sundown")  {
		//$(".CanHide").hide();
		document.getElementById("Time").value = CurrentWeather.Sunset;
	} else if ($("#Select_Controller").val() == "Sunrise")  {
		document.getElementById("Time").value = CurrentWeather.Sunrise;
	}
}

function ShowTab(TabID) {
	
	Result = CheckConfigFileExists();
	
	if ( ( Result === false ) && (TabID !== "Tab2") ) {
		alert("Please save the options at least one time.");
	} else {
		$("#Tab1").hide();
		$("#Tab2").hide();
		$("#"+TabID).show();
	}

}
	
$(document).ready(function() {		
	
});
