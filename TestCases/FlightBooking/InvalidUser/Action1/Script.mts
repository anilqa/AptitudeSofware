'*******************************************************************************************************************************
 @@ hightlight id_;_1508932_;_script infofile_;_ZIP::ssf1.xml_;_
'*******************************************************************************************************************************
'	Script Name							:				Invalid User
'	Objective							:				End to End Scenario
'	UFT Version							:				12.00
'	QC Version							:				NA
'	Pre-requisites						:				NA  
'	Created By							:				Cigniti Technologies
'	Date Created						:				22-06-2017
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
 @@ hightlight id_;_65870_;_script infofile_;_ZIP::ssf3.xml_;_

		If WpfWindow("HP MyFlight Sample Application").Dialog("Login Failed").ExistWpfWindow("HP MyFlight Sample Application").Dialog("Login Failed").Exist Then
			Reporter.ReportEvent micPass,"Check Failed dialog exist","Successfully found failed login dialog box"				
			rptWriteReport "Pass","Check Failed dialog exist","Successfully found failed login dialog box"
            Window("HP MyFlight Sample Application").Dialog("Login Failed").WinButton("OK").DblClick	
		else
			Reporter.ReportEvent micFail,  "Check Failed dialog exist","Failed.login dialog was not found."				
			rptWriteReport "Fail","Check Failed dialog exist","Failed.login dialog was not found."				
			Environment.Value("gErrorFlag") = True

		End If
		
	'Call fn_Click(Window("HP MyFlight Sample Application").Dialog("Login Failed").WinButton("OK"),"DailogBox")   
	Call CheckApplicationForAndClose()

		
	

	End If	
Next


