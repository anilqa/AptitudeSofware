  Public Function fnGetRowCount(sSheetName)    
		sFile = strProjectTestdataPath&Environment("TestName")&".xls"
		sItemName = sSheetName

		Set DB_CONNECTION=CreateObject("ADODB.Connection")
		DB_CONNECTION.Open "DBQ="&sFile&";DefaultDir=C:\;Driver={Driver do Microsoft Excel(*.xls)};DriverId=790;FIL=excel 8.0;FILEDSN=C:\Program Files\Common Files\ODBC\Data Sources\matdsn2.dsn;MaxScanRows=8;PageTimeout=5;ReadOnly=0;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"
		iCheck = Instr(1,sItemName,"$")
		If iCheck = 0 Then
			sItemName = sItemName&"$"
		End If
		sQuery =  "SELECT Count(*) FROM ["&sItemName&"] WHERE Run = 'Y'"
		set Record_Set1=DB_CONNECTION.Execute(sQuery)
		iRowCountValue = 0

		Do While Not Record_Set1.EOF
				For Each Element In Record_Set1.Fields
						iRowCount = Record_Set1(iRowCountValue)
				Next
				Record_Set1.MoveNext
		Loop
		Record_Set1.Close
		Set Record_Set1=Nothing
		DB_CONNECTION.Close
		Set DB_CONNECTION=Nothing
        fnGetRowCount = iRowCount

End Function



Public Function fnGetTestData(sItemName)
	sFile = strProjectTestdataPath&Environment("TestName")&".xls"

	Set Data = CreateObject("Scripting.Dictionary")
	Data.RemoveAll
		
	iCheck = Instr(1,sItemName,"$")
	If iCheck = 0 Then
			sItemName = sItemName&"$"
	End If

	sQuery =  "SELECT * FROM ["&sItemName&"] Where Run = 'Y'"
	Set DB_CONNECTION=CreateObject("ADODB.Connection")
	
	DB_CONNECTION.Open "DBQ="&sFile&";DefaultDir=C:\;Driver={Driver do Microsoft Excel(*.xls)};DriverId=790;FIL=excel 8.0;FILEDSN=C:\Program Files\Common Files\ODBC\Data Sources\matdsn2.dsn;MaxScanRows=8;PageTimeout=5;ReadOnly=0;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"

	Set Record_Set1=DB_CONNECTION.Execute(sQuery)
	Set Record_Set2=DB_CONNECTION.Execute(sQuery)

	iRowCount = 0

	Do While Not Record_Set2.EOF
		iColumnCount = 0
		For Each Field In Record_Set1.Fields
			sColumnName = Field.Name& (iRowCount + 1)
			iRowValue = Record_Set2(iColumnCount)
			If IsNull(iRowValue) Then
				iRowValue = ""
			End If
			Data.Add sColumnName,iRowValue
		iColumnCount = iColumnCount + 1
		Next
		Record_Set2.MoveNext
		iRowCount = iRowCount + 1
	Loop

	Record_Set1.Close
	Set Record_Set1=Nothing
	Record_Set2.Close
	Set Record_Set2=Nothing
	DB_CONNECTION.Close
	Set DB_CONNECTION=Nothing
	Set fnGetTestData = Data	

End Function


' ---  mycode -----

function fn_Set(objObject,strObjName,strValue)
On error resume next
If Not Environment.Value("gErrorFlag") Then	
	If objObject.Exist(3) Then
		strObjTypeDescription = fn_GetObjectClass(objObject)
		Select Case strObjTypeDescription
			Case "calender"
				objObject.SetDate strValue
			Case else 
				objObject.set strValue
		End Select
		If err.number = 0 Then
			Reporter.ReportEvent micPass, strObjName + " " + strObjTypeDescription,strObjName + " entered as " + strValue				
			rptWriteReport "Pass",strObjName + " " + strObjTypeDescription,strObjName + " entered as " + strValue
		else
			Reporter.ReportEvent micFail,  strObjName + " " + strObjTypeDescription,"Unable to entered "&strValue + " in " & strObjName
			rptWriteReport "Fail", strObjName + " " + strObjTypeDescription,"Unable to entered "&strValue + " in " & strObjNames
			Environment.Value("gErrorFlag") = True
			err.clear
			Exit Function
		End If

	Else
		Reporter.ReportEvent micFail, "Enter " + strValue + " into " + strObjName + " field" , strObjName + " " + strObjTypeDescription + " not exist"
		rptWriteReport "Fail", "Enter " + strValue + " into " + strObjName + " field" , strObjName + " " + strObjTypeDescription + " not exist"
		Environment.Value("gErrorFlag") = True
		err.clear
	End If
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: fn_Set", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: fn_Set", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If
	err.clear			
End If
End function

function fn_Type(objObject,strObjName,strValue)
On error resume next
If Not Environment.Value("gErrorFlag") Then	
	
	If objObject.Exist(3) Then
		
		strObjTypeDescription = fn_GetObjectClass(objObject)
		If objObject.GetROProperty("visible") = 0 Then
			objObject.Type strValue
			Select Case strObjTypeDescription
				Case "Text field"					
					If err.number = 0 Then
						Reporter.ReportEvent micPass, "Enter " + strValue + " into " + strObjName + " field" , strValue + " entered into " + strObjName + " field"				
						rptWriteReport "Pass","Enter " + strValue + " into " + strObjName + " field" , strValue + " entered into " + strObjName + " field"				
					else
						Reporter.ReportEvent "Enter " + strValue + " into " + strObjName + " field" , "Failed to enter " + strValue + " into " + strObjName + " field"		
						rptWriteReport "Fail","Enter " + strValue + " into " + strObjName + " field" , "Failed to enter " + strValue + " into " + strObjName + " field"
						Environment.Value("gErrorFlag") = True
						err.clear
						Exit Function
					End If					
			End Select	
		Else
			Reporter.ReportEvent micFail, "Enter " + strValue + " into " + strObjName + " field" , strObjName + " field not enabled"
			rptWriteReport "Fail", "Enter " + strValue + " into " + strObjName + " field" , strObjName + " field not enabled"
			Environment.Value("gErrorFlag") = True
			err.clear
			Exit Function
		End If
	Else
		Reporter.ReportEvent micFail, "Enter " + strValue + " into " + strObjName + " field" , strObjName + " field not exist"
		rptWriteReport "Fail", "Enter " + strValue + " into " + strObjName + " field" , strObjName + " field not exist"
		Environment.Value("gErrorFlag") = True
		err.clear
		Exit Function
	End If
	
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: fn_Set", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: fn_Set", "Error: " + err.description
		Environment.Value("gErrorFlag") = False
	End If
	err.clear			
End If
End function

function fn_SelectDropDown(objObject,strObjName,strValue)
On error resume next
If Not Environment.Value("gErrorFlag") Then
	If objObject.Exist(3) Then
		strObjTypeDescription = fn_GetObjectClass(objObject)
		If objObject.GetROProperty("enabled") = true Then
			objObject.Select strValue
			If err.number = 0 Then
				Reporter.ReportEvent micPass, "Select " + strValue + " from " + strObjName + " " +strObjTypeDescription, strValue + " selected from " + strObjName + " " +strObjTypeDescription
				rptWriteReport "Pass", "Select " + strValue + " from " + strObjName + " " +strObjTypeDescription, strValue + " selected from " + strObjName + " " +strObjTypeDescription
			else
				Reporter.ReportEvent micFail, "Select " + strValue + " from " + strObjName + " " +strObjTypeDescription, "Failed to select " + strValue + " from "+ strObjName + " " +strObjTypeDescription
				rptWriteReport "Fail", "Select " + strValue + " from " + strObjName + " " +strObjTypeDescription, "Failed to select " + strValue + " from "+ strObjName + " " +strObjTypeDescription
				Environment.Value("gErrorFlag") = True
				err.clear
				Exit Function
			End If
		Else
			Reporter.ReportEvent micFail, "Select " + strValue + " from " + strObjName + " " +strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not enabled"
			rptWriteReport "Fail", "Select " + strValue + " from " + strObjName + " " +strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not enabled"
			Environment.Value("gErrorFlag") = True
			err.clear
			Exit Function
		End If
	Else
		Reporter.ReportEvent micFail, "Select " + strValue + " from " + strObjName + " " +strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not exist"
		rptWriteReport "Fail", "Select " + strValue + " from " + strObjName + " " +strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not exist"
		Environment.Value("gErrorFlag") = True
		err.clear
		Exit Function
	End If
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: fn_Select", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: fn_Select", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If
	err.clear	
End If

End function

function fn_Click(objObject,strObjName)
On error resume next
If Not Environment.Value("gErrorFlag") Then
	If objObject.Exist(3) Then
		strObjTypeDescription = fn_GetObjectClass(objObject)
		objObject.Click
		If err.number = 0 Then
				Reporter.ReportEvent micPass, "Click " + strObjName + " " + strObjTypeDescription , "successfully clicked " + strObjName  + " " + strObjTypeDescription
				rptWriteReport "Pass", "Click " + strObjName + " " + strObjTypeDescription , "successfully clicked " + strObjName + " " + strObjTypeDescription
		else
				Reporter.ReportEvent micFail, "Click " + strObjName + " " + strObjTypeDescription , "Failed to click " + strObjName + " " + strObjTypeDescription
				rptWriteReport "Fail", "Click " + strObjName + " " + strObjTypeDescription , "Failed to click " + strObjName + " " + strObjTypeDescription
				Environment.Value("gErrorFlag") = True
				err.clear
				Exit Function
		End If
	Else
		Reporter.ReportEvent micFail, "Click " + strObjName + " " + strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not exist"
		rptWriteReport "Fail", "Click " + strObjName + " " + strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not exist"
		Environment.Value("gErrorFlag") = True
		err.clear
		Exit Function
	End If
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: fn_Click", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: fn_Click", "Error: " + err.description
		Environment.Value("gErrorFlag") = False
	End If
	err.clear
	
End If
End function

function fn_ClickCell(objObject,row,col,strObjName)
On error resume next
If Not Environment.Value("gErrorFlag") Then
	If objObject.Exist(3) Then
		strObjTypeDescription = fn_GetObjectClass(objObject)
		If objObject.GetROProperty("enabled") = true Then
			objObject.selectcell row,col
			If err.number = 0 Then
				Reporter.ReportEvent micPass, "ClickCell on " + strObjName + " " + strObjTypeDescription , "In "&strObjName&" grid, successfully clicked on row:" & row  & " & col:" & col
				rptWriteReport "Pass", "ClickCell on " + strObjName + " " + strObjTypeDescription , "In "&strObjName&" grid, successfully clicked on row:" & row  & " & col:" & col
			else
				Reporter.ReportEvent micFail, "Click " + strObjName + " " + strObjTypeDescription , "Failed to click on row:" & row  & " & col:" & col
				rptWriteReport "Fail", "Click " + strObjName + " " + strObjTypeDescription , "Failed to click on row:" & row  & " & col:" & col
				Environment.Value("gErrorFlag") = True
				err.clear
				Exit Function
			End If
		Else
			Reporter.ReportEvent micFail, "Click " + strObjName + " " + strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not enabled"
			rptWriteReport "Fail", "Click " + strObjName + " " + strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not enabled"
			Environment.Value("gErrorFlag") = True
			err.clear
			Exit Function
		End If
	Else
		Reporter.ReportEvent micFail, "Click " + strObjName + " " + strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not exist"
		rptWriteReport "Fail", "Click " + strObjName + " " + strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not exist"
		Environment.Value("gErrorFlag") = True
		err.clear
		Exit Function
	End If
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: fn_Click", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: fn_Click", "Error: " + err.description
		Environment.Value("gErrorFlag") = False
	End If
	err.clear
	
End If

End function

function fn_DoubleClick(objObject,strObjName)
On error resume next
If Not Environment.Value("gErrorFlag") Then
	If objObject.Exist(3) Then
		strObjTypeDescription = fn_GetObjectClass(objObject)
		If objObject.GetROProperty("disabled") = 0 Then
			objObject.DblClick 0,0
			If err.number = 0 Then
				Reporter.ReportEvent micPass, "Double click " + strObjName + " " + strObjTypeDescription , "successfully clicked " + strObjName  + " " + strObjTypeDescription
				rptWriteReport "Pass", "Double click " + strObjName + " " + strObjTypeDescription , "successfully clicked " + strObjName + " " + strObjTypeDescription
			else
				Reporter.ReportEvent micFail, "Double Click " + strObjName + " " + strObjTypeDescription , "Failed to click " + strObjName + " " + strObjTypeDescription
				rptWriteReport "Fail", "Double Click " + strObjName + " " + strObjTypeDescription , "Failed to click " + strObjName + " " + strObjTypeDescription
				Environment.Value("gErrorFlag") = True
				err.clear
				Exit Function
			End If
		Else
			Reporter.ReportEvent micFail, "Double Click " + strObjName + " " + strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not enabled"
			rptWriteReport "Fail", "Double Click " + strObjName + " " + strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not enabled"
			Environment.Value("gErrorFlag") = True
			err.clear
			Exit Function
		End If
	Else
		Reporter.ReportEvent micFail, "Double Click " + strObjName + " " + strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not exist"
		rptWriteReport "Fail", "Double Click " + strObjName + " " + strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not exist"
		Environment.Value("gErrorFlag") = True
		err.clear
		Exit Function
	End If
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: fn_DoubleClick", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: fn_DoubleClick", "Error: " + err.description
		Environment.Value("gErrorFlag") = False
	End If
	err.clear
	
End If

End function

function fn_CheckForObjecct(objObject,strObjName,intTime)

On error resume next
If Not Environment.Value("gErrorFlag") Then
	If objObject.Exist(intTime) Then
		fn_CheckForObjecct = True
		Else
		fn_CheckForObjecct = False
	End If
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: fn_Set", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: fn_Set", "Error: " + err.description
		Environment.Value("gErrorFlag") = False
	End If
	err.clear
End If
End function

function fn_CheckForObjecctAndReport(objObject,strObjName,blnPresence)

On error resume next

If Not Environment.Value("gErrorFlag") Then
	
	If objObject.Exist(3) Then

		strObjTypeDescription = fn_GetObjectClass(objObject)
		If blnPresence Then
			Reporter.ReportEvent micPass, strObjTypeDescription + " existence", strObjName + " " + strObjTypeDescription + " exist"
			rptWriteReport "Pass",strObjTypeDescription + " existence", strObjName + " " + strObjTypeDescription + " exist"
			Else
			Reporter.ReportEvent micFail, strObjTypeDescription + " existence", strObjName + " " + strObjTypeDescription + " not exist"
			rptWriteReport "Fail",strObjTypeDescription + " existence", strObjName + " " + strObjTypeDescription + " not exist"			
		End If
		Else
		If blnPresence Then
			Reporter.ReportEvent micFail, strObjTypeDescription + " existence", strObjName + " " + strObjTypeDescription + " exist"
			rptWriteReport "Fail",strObjTypeDescription + " existence", strObjName + " " + strObjTypeDescription + " exist"
			Else
			Reporter.ReportEvent micPass, strObjTypeDescription + " existence", strObjName + " " + strObjTypeDescription + " not exist"
			rptWriteReport "Pass",strObjTypeDescription + " existence", strObjName + " " + strObjTypeDescription + " not exist"			
		End If		
	End If
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: fn_CheckForObjecctAndReport", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: fn_CheckForObjecctAndReport", "Error: " + err.description
		Environment.Value("gErrorFlag") = False
	End If
	err.clear
End If

End function

Function fn_GetObjectClass(objObject)
	On error resume next
	strObjClass = objObject.GetROProperty("wpftypename")	
	Select Case strObjClass
		Case "JavaEdit"
			fn_GetObjectClass = "Text field"
		Case else 
			fn_GetObjectClass = strObjClass
	End Select
	
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: fn_GetObjectClass", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: fn_GetObjectClass", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If
	err.clear
		
End Function

Function closeApplication
	Call CheckApplicationForAndClose
End Function

Function fn_Select(objObject,strObjectName,strValue)
On error resume next
If Not Environment.Value("gErrorFlag") Then
	If objObject.Exist(3) Then
		strObjTypeDescription = fn_GetObjectClass(objObject)
		If objObject.GetROProperty("enabled") = true Then
			objObject.Select strValue
			If err.number = 0 Then
				Reporter.ReportEvent micPass, "Menu selection",strValue&" is selected for " & strObjectName&" dropdown"
				rptWriteReport "Pass", "Menu selection",strValue&" is selected for " & strObjectName&" dropdown"
			else
				Reporter.ReportEvent micFail, "Menu selection","An error occurred while selecting menu " + strObjectName+" err.description"+err.description
				rptWriteReport "Fail", "Menu selection","An error occurred while selecting menu " + strObjectName+" err.description"+err.description
				Environment.Value("gErrorFlag") = True
				err.clear
				Exit Function
			End If
		Else
			Reporter.ReportEvent micFail, "Menu selection", strObjectName + " is not enabled"
			rptWriteReport "Fail", "Menu selection", strObjectName + " is not enabled"
			Environment.Value("gErrorFlag") = True
			err.clear
			Exit Function
		End If
	Else
		Reporter.ReportEvent micFail, "Menu selection", strObjectName + " not exist"
		rptWriteReport "Fail", "Menu selection", strObjectName + " not exist"
		Environment.Value("gErrorFlag") = True
		err.clear
		Exit Function
	End If
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: fn_Select", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: fn_Select", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If
	err.clear	
End If	
End Function
Function fn_getText(objObject,strObjectName)
On error resume next
If Not Environment.Value("gErrorFlag") Then
	If objObject.Exist(3) Then
		strObjTypeDescription = fn_GetObjectClass(objObject)
		temp = objObject.GetROProperty("text")
		If err.number = 0 Then
				Reporter.ReportEvent micPass, "Get Text","Text property for " + strObjectName+" is "+temp
				rptWriteReport "Pass", "Get Text","Text property for " + strObjectName+" is "+temp
				fn_getText = temp 
			else
				Reporter.ReportEvent micFail, "Get Text","An error occurred while retriveing text property from "+ strObjectName+" err.description"+err.description
				rptWriteReport "Fail", "Get Text","An error occurred while retriveing text property from "+ strObjectName+" err.description"+err.description
				Environment.Value("gErrorFlag") = True
				fn_getText =""
				err.clear
				Exit Function
			End If
	Else
		Reporter.ReportEvent micFail, "Get Text", strObjectName + " not exist"
		rptWriteReport "Fail", "Get Text", strObjectName + " not exist"
		Environment.Value("gErrorFlag") = True
		err.clear
		Exit Function
	End If
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: fn_getText", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: fn_getText", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If
	err.clear	
End If	
End Function

Function fn_getCellData(objObject,row,col,strObjectName)
On error resume next
If Not Environment.Value("gErrorFlag") Then
	If objObject.Exist(3) Then
		strObjTypeDescription = fn_GetObjectClass(objObject)
'		If objObject.getROProperty("rows") >=row and objObject.getROProperty("cols") >=col Then
			temp = objObject.GetCellData(row,col)
			If err.number = 0 Then
				Reporter.ReportEvent micPass, "GetCellData","CellData for " + strObjectName+" is "&cstr(temp)
				rptWriteReport "Pass", "GetCellData","CellData for " + strObjectName+" is "&cstr(temp)
				fn_getCellData = temp 
			else
				Reporter.ReportEvent micFail, "GetCellData","An error occurred while retriveing text property from "+ strObjectName+" err.description"+err.description
				rptWriteReport "Fail", "Get Text","An error occurred while retriveing text property from "+ strObjectName+" err.description"+err.description
				Environment.Value("gErrorFlag") = True
				fn_getCellData =""
				err.clear
				Exit Function
			End If
'		else
'			Reporter.ReportEvent micFail, "GetCellData","Rows & Cols are out of bound"
'			rptWriteReport "Fail", "Get Text","Rows & Cols are out of bound"
'		End If
'		
	Else
		Reporter.ReportEvent micFail, "Get Text", strObjectName + " not exist"
		rptWriteReport "Fail", "Get Text", strObjectName + " not exist"
		Environment.Value("gErrorFlag") = True
		err.clear
		Exit Function
	End If
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: fn_getText", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: fn_getText", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If
	err.clear	
End If	
End Function

Function fn_RClick(objObject,objTarget,x,y,objstrObjectName)
'On error resume next
If Not Environment.Value("gErrorFlag") Then
	If objObject.Exist(3) Then
		strObjTypeDescription = fn_GetObjectClass(objObject)
		If objObject.GetROProperty("disabled") = 0 Then
			objObject.Click x,y,"RIGHT"
			wait 1
			objTarget.Select
			wait 1
			If err.number = 0 Then
				Reporter.ReportEvent micPass, "RClick " + strObjName + " " + strObjTypeDescription , "successfully clicked " + strObjName  + " " + strObjTypeDescription
				rptWriteReport "Pass", "RClick " + strObjName + " " + strObjTypeDescription , "successfully clicked " + strObjName + " " + strObjTypeDescription
			else
				Reporter.ReportEvent micFail, "RClick " + strObjName + " " + strObjTypeDescription , "Failed to click " + strObjName + " " + strObjTypeDescription
				rptWriteReport "Fail", "RClick " + strObjName + " " + strObjTypeDescription , "Failed to click " + strObjName + " " + strObjTypeDescription
				Environment.Value("gErrorFlag") = True
				err.clear
				Exit Function
			End If
		Else
			Reporter.ReportEvent micFail, "RClick " + strObjName + " " + strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not enabled"
			rptWriteReport "Fail", "RClick " + strObjName + " " + strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not enabled"
			Environment.Value("gErrorFlag") = True
			err.clear
			Exit Function
		End If
	Else
		Reporter.ReportEvent micFail, "RClick " + strObjName + " " + strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not exist"
		rptWriteReport "Fail", "RClick " + strObjName + " " + strObjTypeDescription, strObjName + " " + strObjTypeDescription + " not exist"
		Environment.Value("gErrorFlag") = True
		err.clear
		Exit Function
	End If
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: fn_RClick", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: fn_RClick", "Error: " + err.description
		Environment.Value("gErrorFlag") = False
	End If
	err.clear
End if 
End Function

Function fn_SetCellData(objObject,row,col,value,strObjectName)
On error resume next
If Not Environment.Value("gErrorFlag") Then
	If objObject.Exist(3) Then
		strObjTypeDescription = fn_GetObjectClass(objObject)
		If objObject.GetROProperty("disabled") = 0 Then
			objObject.SetCellData row,col,value 
			If err.number = 0 Then
				Reporter.ReportEvent micPass, "SetCellData","For row:"+row+" col:"+col+" and value"+value+" is set for " + strObjectName
				rptWriteReport "Pass", "SetCellData","For row:"+row+" col:"+col+" and value"+value+" is set for " + strObjectName
			else
				Reporter.ReportEvent micFail, "SetCellData","An error occurred while settting row:"+row+" col:"+col+" and value"+value+" is set for " + strObjectName+" err.description"+err.description
				rptWriteReport "Fail", "Menu selection","An error occurred while settting row:"+row+" col:"+col+" and value"+value+" is set for " + strObjectName+" err.description"+err.description
				Environment.Value("gErrorFlag") = True
				err.clear
				Exit Function
			End If
		Else
			Reporter.ReportEvent micFail, "SetCellData", strObjectName + " is not enabled"
			rptWriteReport "Fail", "SetCellData", strObjectName + " is not enabled"
			Environment.Value("gErrorFlag") = True
			err.clear
			Exit Function
		End If
	Else
		Reporter.ReportEvent micFail, "SetCellData", strObjectName + " not exist"
		rptWriteReport "Fail", "SetCellData", strObjectName + " not exist"
		Environment.Value("gErrorFlag") = True
		err.clear
		Exit Function
	End If
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: fn_SetCellData", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: fn_SetCellData", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If
	err.clear	
End If		
End Function

Function fn_SelectTab(objObject,strTabName)
On error resume next
If Not Environment.Value("gErrorFlag") Then
	If objObject.Exist(3) Then
		strObjTypeDescription = fn_GetObjectClass(objObject)
		If objObject.GetROProperty("disabled") = 0 Then
			objObject.Select strTabName
			If err.number = 0 Then
				Reporter.ReportEvent micPass, "Select Tab","Following tab:" & strObjectName &" is selected succesfully"
				rptWriteReport "Pass", "Select Tab","Following tab:" & strObjectName &" is selected succesfully"
			else
				Reporter.ReportEvent micFail, "Select Tab","An error occurred while selecting:"&strObjectName
				rptWriteReport "Fail", "Select Tab","An error occurred while selecting:"&strObjectName
				Environment.Value("gErrorFlag") = True
				err.clear
				Exit Function
			End If
		Else
			Reporter.ReportEvent micFail, "Select Tab", strObjectName + " is not enabled"
			rptWriteReport "Fail", "Select Tab", strObjectName + " is not enabled"
			Environment.Value("gErrorFlag") = True
			err.clear
			Exit Function
		End If
	Else
		Reporter.ReportEvent micFail, "Select Tab", strObjectName + " not exist"
		rptWriteReport "Fail", "Select Tab", strObjectName + " not exist"
		Environment.Value("gErrorFlag") = True
		err.clear
		Exit Function
	End If
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: f_SelectTab", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: f_SelectTab", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If
	err.clear	
End If		
End Function
