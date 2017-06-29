'*******************************************************************************************************************************
'*******************************************************************************************************************************
'	Script Name							:				Flight Booking
'	Objective							:				End to End Scenario
'	UFT Version							:				12.00
'	QC Version							:				NA
'	Pre-requisites						:				NA  
'	Created By							:				Cigniti Technologies
'	Date Created						:				21-06-2017
'	Modification Date		       		:		   		NIL	 
'*******************************************************************************************************************************

DataTable.ImportSheet strProjectTestdataPath&Environment("TestName")&".xls", Environment("TestName"), "Global"

Dim intRow  
For intRow  = 1 To DataTable.GetRowCount
	On Error Resume Next
	Environment.Value("gErrorFlag") = False
	DataTable.SetCurrentRow intRow
	Environment.Value("gIteration") = intRow
	
	If UCase(Trim(DataTable.Value("Run"))) = "Y" Then		
		Call OpenApplication(DataTable.Value("ApplicationURL"),DataTable.Value("DirectoryPath"))
		Call LoginInToApplication(DataTable.Value("UserName"),DataTable.Value("Password"))
		Call fn_Select(WpfWindow("HP MyFlight Sample Application").WpfComboBox("fromCity"),"fromCity",DataTable.Value("FromCity")) @@ hightlight id_;_2101056624_;_script infofile_;_ZIP::ssf24.xml_;_
		Call fn_Select(WpfWindow("HP MyFlight Sample Application").WpfComboBox("toCity"),"ToCity",DataTable.Value("ToCity"))
		Call fn_Set(WpfWindow("HP MyFlight Sample Application").WpfCalendar("datePicker"),DataTable.Value("TravelDate"))
		Call fn_Select(WpfWindow("HP MyFlight Sample Application").WpfComboBox("Class"),"Classtype",DataTable.Value("ClassType"))
		Call fn_Select(WpfWindow("HP MyFlight Sample Application").WpfComboBox("numOfTickets"),"Tickets",DataTable.Value("Tickets"))
        call fn_Click(WpfWindow("HP MyFlight Sample Application").WpfButton("FIND FLIGHTS"),"FindFlight")
        call fn_ClickCell(WpfWindow("HP MyFlight Sample Application").WpfTable("flightsDataGrid"),1,1, "flightno")
        call fn_Click(WpfWindow("HP MyFlight Sample Application").WpfButton("SELECT FLIGHT"),"SelectFlight")
        Call fn_Set(WpfWindow("HP MyFlight Sample Application").WpfEdit("passengerName"),"PassengerName",DataTable.Value("Passengername"))
        call fn_Click(WpfWindow("HP MyFlight Sample Application").WpfButton("ORDER"),"Order")
        Call CheckApplicationForAndClose()
	End If	
Next





