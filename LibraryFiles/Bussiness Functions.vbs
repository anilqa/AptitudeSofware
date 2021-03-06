'-------- MyCode-------

Function CheckApplicationForAndClose
	Dim  AllProcess
	Dim  Process
	Dim  strFoundProcess
	strFoundProcess = False
	Set AllProcess = getobject("winmgmts:") 'create object
    For Each Process In AllProcess.InstancesOf("Win32_process") 'Get all the processes running in your PC
  		If (Instr (Ucase(Process.Name),"FLIGHTSGUI.EXE") = 1) Then 'Made all uppercase to remove ambiguity. Replace TASKMGR.EXE with your application name in CAPS.
			strFoundProcess = True
			Call rptWriteReport("Pass", "FlightGui", "Found already opened FlightGui application, attempting to close it")			 
			Exit for
		End If
	Next
	If strFoundProcess = true Then
		SystemUtil.CloseProcessByName("FlightsGUI.exe")		
		wait 5
	End If 
	Set AllProcess = nothing
End Function

Function OpenApplication(appPath,directoryPath)
	Call CheckApplicationForAndClose()
	On error resume next
	If not Environment.Value("gErrorFlag") Then		
		SystemUtil.Run appPath,"",directoryPath,""
		If fn_CheckForObjecct(WpfWindow("HP MyFlight Sample Application"),"Flight Application",10000) Then
			rptWriteReport "Pass","Launch Flight Application", "Application successfully launched"
		else
			gFun_WriteReport "Fail","Launch Flight Application","Application fialed to launch"
			Environment.Value("gErrorFlag") = True
		End If
		If err.number <> 0 Then
			Reporter.ReportEvent micFail,"Error occurred in the function: OpenApplication","Error: "+err.description
			rptWriteReport "Fail","Error occured in the function: Open Application","Error: "+err.description
			Environment.Value("gErrorFlag") = True			
		End If
		err.clear
	End If	
End Function

Function LoginInToApplication(UserName,Password)
	On error resume next
	If not Environment.Value("gErrorFlag") Then
		call fn_Set(WpfWindow("HP MyFlight Sample Application").WpfEdit("agentName"),"UserName",UserName)
		call fn_Set(WpfWindow("HP MyFlight Sample Application").WpfEdit("password"),"Password",Password)
		call fn_Click(WpfWindow("HP MyFlight Sample Application").WpfButton("OK"),"Login")
'		If fn_CheckForObjecct(JavaWindow("SkyChainHomeWin"),"SkyChain reservation screen",30) Then
'			rptWriteReport "Pass","Launch SkyChain Application", "Succesfully logged into the application"
'		else			
'			rptWriteReport "Fail","Unable to login into SkyChain Application","Failed to Login"
'			Environment.Value("gErrorFlag") = True
'		End If
		If err.number <> 0 Then
			Reporter.ReportEvent micFail,"Error occurred in the function: LoginInToApplication","Error: "+err.description
			rptWriteReport "Fail","Error occurred in the function: LoginInToApplication","Error: "+err.description
			Environment.Value("gErrorFlag") = True	
		End If	
		err.clear		
	End If
End Function

Function MaximizeWindow
	On error resume next
	If not Environment.Value("gErrorFlag")  Then
		JavaWindow("SkyChainHomeWin").Maximize
		If err.number <> 0 Then
			Reporter.ReportEvent micFail,"Error occurred in the function: LoginInToApplication","Error: "+err.description
			rptWriteReport "Fail","Error occurred in the function: LoginInToApplication","Error: "+err.description
			Environment.Value("gErrorFlag") = True	
		End If	
		err.clear
	End If
End Function

Function NavigateTo(strScreenName)
	On error resume next
	If not Environment.Value("gErrorFlag") Then
		If strcomp(lcase(strScreenName),"reservation",1)=0 Then			
			call fn_Select(JavaWindow("SkyChainHomeWin").JavaMenu("Shipment").JavaMenu("Reservation").JavaMenu("Reservation"),"Reservation")
			call fn_CheckForObjecct(JavaWindow("SkyChainHomeWin").JavaButton("GetNextBtn"),"Reservation",10)
		ElseIf strcomp(lcase(strScreenName),"flightrecord",1)=0 Then			
			call fn_Select(JavaWindow("SkyChainHomeWin").JavaMenu("Flight").JavaMenu("Flight Record"),"Flight Record")
			Call fn_CheckForObjecct(JavaWindow("SkyChainFlightWin").JavaButton("SearchBtn"),"FlightRecord",10)
		ElseIf strcomp(lcase(strScreenName),"docnetcharges",1)=0 Then
			call fn_Select(JavaWindow("SkyChainHomeWin").JavaMenu("Shipment").JavaMenu("Rates & Charges").JavaMenu("Doc. Net Rating"),"Doc. Net Rating")
			Call fn_CheckForObjecct(JavaWindow("SkyChainDocNetRatingWin").JavaButton("GetCharges"),"Doc. Net Rating",10)
		ElseIf strcomp(lcase(strScreenName),"awbcapture",1)=0  Then
			Call fn_Select(JavaWindow("SkyChainHomeWin").JavaMenu("Shipment").JavaMenu("AWB").JavaMenu("AWBCapture"),"AWB Capture")
			Call fn_CheckForObjecct(JavaWindow("SkyChainDocNetRatingWin").JavaButton("GetCharges"),"Doc. Net Rating",10)
		End If
		If err.number <> 0 Then
			Reporter.ReportEvent micFail,"Error occurred in the function: NavigateTo","Error: "+err.description
			rptWriteReport "Fail","Error occurred in the function: NavigateTo","Error: "+err.description
			Environment.Value("gErrorFlag") = True
		else
			gFunc_rptWriteReport "Pass","Screen selection","Able to select "&strScreeName&" in the application"
		End If
		err.clear
	End If
End Function

Function UpdateBranchCode(strBranchCode,strCaller)
	On error resume next 
	If not Environment.Value("gErrorFlag") Then
		call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("AgentBranchCode"),"Agent Branch Code",strBranchCode)				
		If fn_CheckForObjecct(JavaWindow("SkyChainHomeWin").JavaDialog("CustomerSearchWin").JavaTable("CustomerCodeTbl"),"CustomerSearch",10) Then
			JavaWindow("SkyChainHomeWin").JavaDialog("CustomerSearchWin").JavaTable("CustomerCodeTbl").ClickCell 0,1
		    call fn_Click(JavaWindow("SkyChainHomeWin").JavaDialog("CustomerSearchWin").JavaButton("OkBtn"),"Ok")
		    While JavaWindow("SkyChainHomeWin").JavaDialog("CustomerSearchWin").JavaTable("CustomerCodeTbl").Exist(0)
		    	wait 0,500
		    Wend
		    call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("Caller"),"Caller",strCaller)
		    Call fn_Click(JavaWindow("SkyChainHomeWin").JavaButton("GetNextBtn"),"GetNext")
		    do While true
		    	wait 0,500
		    	strTemp = JavaWindow("SkyChainHomeWin").JavaEdit("DocumentPrefix_2").GetROProperty("text")
		    	If len(trim(strTemp))>0 Then
		    		Exit do
		    	End If
		    loop
		else
			rptWriteReport "Fail","Error occured in the function: UpdateBranchCode", "Error: " + err.description
			Environment.Value("gErrorFlag") = true
		End If		
		If err.number <> 0 Then		
			Reporter.ReportEvent micFail, "Error occured in the function: UpdateBranchCode", "Error: " + err.description
			rptWriteReport "Fail","Error occured in the function: UpdateBranchCode", "Error: " + err.description
			Environment.Value("gErrorFlag") = true
		End If
		err.clear
	End If
End Function

Function UpdateShipmentInfo(strPieces,strWtVol,strOrg,strDest,strCommodity,strCommodityDesc)
	On error resume next
	If not Environment.Value("gErrorFlag") Then
		call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("Pieces"),"Pieces",strPieces)
		Call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("Wt-Vol"),"WtVol",strWtVol)
		Call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("CargoNumericField"),"CargoNumber",strPieces)
		Call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("Org"),"Org",strOrg)
		Call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("Dest"),"Dest",strDest)
		Call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("ManifestDesc"),"ManifestDesc",strCommodityDesc)
		Call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("Commodity"),"Commodity",strCommodity)
		JavaWindow("SkyChainHomeWin").JavaEdit("Commodity").Type micTab
		do 
			wait 2
			strTempCommodityDesc = JavaWindow("SkyChainHomeWin").JavaEdit("CommodityDesc").GetROProperty("text")
			If len(trim(strTempCommodityDesc))>0 Then
				Exit do 
			End If
		loop While false			
		'Call fn_Click(JavaWindow("SkyChainHomeWin").JavaButton("RoutingBtn"),"Routing")
'		If err.number <>0 Then
'			Reporter.ReportEvent micFail, "Error occured in the function: UpdateShipmentInfo", "Error: " + err.description
'			rptWriteReport "Fail","Error occured in the function: UpdateShipmentInfo", "Error: " + err.description
'			Environment.Value("gErrorFlag") = true
'		else
'			If JavaWindow("SkyChainHomeWin").JavaDialog("RoutingDetailsWin").JavaTable("RoutingDetailsTbl").Exist(20) Then
'				set objRouting = JavaWindow("SkyChainHomeWin").JavaDialog("RoutingDetailsWin").JavaTable("RoutingDetailsTbl")
'				rows = objRouting.GetROProperty("rows")
'				cols = objRouting.GetROProperty("cols")
'				strItinearyValue = "" 
'				For row = 1 To rows
'					For col = 0 To cols Step 1
'						strTemp = objRouting.GetCellData(row,col)
'						If len(trim(strTemp))>0 Then
'							 strItinearyValue = strItinearyValue+";"+Cstr(objRouting.GetCellData(row,col))
'						End If						
'					Next
'					Exit for 
'				Next
'				objRouting.ClickCell 1,1
'				call fn_Click(JavaWindow("SkyChainHomeWin").JavaDialog("RoutingDetailsWin").JavaButton("SelectBtn"),"Select")
'				While JavaWindow("SkyChainHomeWin").JavaDialog("RoutingDetailsWin").Exist(0)
'					wait 0,500
'				Wend
'				If err.number <>0 Then
'					Reporter.ReportEvent micFail, "Error occured in the function: UpdateShipmentInfo", "Error: " + err.description
'					rptWriteReport "Fail","Error occured in the function: UpdateShipmentInfo", "Error: " + err.description
'					Environment.Value("gErrorFlag") = true
'				else 
'					UpdateShipmentInfo = strItinearyValue
'				End if 
'			else
'			 	Reporter.ReportEvent micFail, "Error occured in the function: UpdateShipmentInfo", "Error: " + err.description
'				rptWriteReport "Fail","Error occured in the function: UpdateShipmentInfo", "Error: " + err.description
'				Environment.Value("gErrorFlag") = true
'			End If
'		End If
		If err.number <>0 Then
			Reporter.ReportEvent micFail, "Error occured in the function: UpdateShipmentInfo", "Error: " + err.description
			rptWriteReport "Fail","Error occured in the function: UpdateShipmentInfo", "Error: " + err.description
			Environment.Value("gErrorFlag") = true
		else 
			UpdateShipmentInfo = strItinearyValue
		End if 
		err.clear
	End If
End Function

Function UpdateProductDetails(strProduct,strChargeCode)
	On error resume next
	If not Environment.Value("gErrorFlag") Then
		If strChargeCode<>"" Then
			call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("ChargeCode"),"ChargeCode",strChargeCode)			
		End If		
		call wait (2)
		If strProduct<>"" Then
			If fn_CheckForObjecct(JavaWindow("SkyChainHomeWin").JavaEdit("Product"),"Product",10) Then
				call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("Product"),"Product",strProduct)								
			End If	
		End If		
		If err.number <> 0 Then
			Reporter.ReportEvent micFail, "Error occured in the function: UpdateShipmentInfo", "Error: " + err.description
			rptWriteReport "Fail","Error occured in the function: UpdateShipmentInfo", "Error: " + err.description
			Environment.Value("gErrorFlag") = true	
		End If
		err.clear
	End If
End Function

Function getFlightRecord(strCarrier,strFlightNumber,strFlightSuffix,strFlightDate)
	On error resume next 
	If not Environment.Value("gErrorFlag") Then
		Call fn_Set(JavaWindow("SkyChainFlightWin").JavaEdit("Carrier"),"Carrier",strCarrier)
		Call fn_Set(JavaWindow("SkyChainFlightWin").JavaEdit("FltNo"),"FlightNumber",strFlightNumber)
		Call fn_Set(JavaWindow("SkyChainFlightWin").JavaEdit("Sufix"),"Suffix",strFlightSuffix)
		Call fn_Set(JavaWindow("SkyChainFlightWin").JavaEdit("Date"),"FlightDate",strFlightDate)		
		Call fn_Click(JavaWindow("SkyChainFlightWin").JavaButton("SearchBtn"),"SearchButton")
		flag = true
		count = 1
		Do
			wait 0,500
			rows = JavaWindow("SkyChainFlightWin").JavaTable("SegmentAdvTbl").GetROProperty("rows")
			If rows>0 Then
				flag = false
			End If
			If count >5 Then
				flag = false
			End If
			count = count+1
		Loop While flag
		If rows>0 Then
			Reporter.ReportEvent micPass, "FlightRecord", "Flight Record found"
			rptWriteReport "Pass","FlightRecord", "Flight Record found"
		else
			Reporter.ReportEvent micFail, "FlightRecord", "Flight Record not found"
			rptWriteReport "Fail","FlightRecord", "Flight Record not found"
			Environment.Value("gErrorFlag")=true
		End If
	End If	
End Function


Function ValidateFlightRecord(strPrevFSAvailableWt,strPrevFSAvailableVol,strPrevTotalAvailableWt,strPrevTotalAvailableVol)
	On error resume next 
	If not Environment.Value("gErrorFlag") Then
		strFSAvailableWt = fn_getCellData(JavaWindow("SkyChainFlightWin").JavaTable("SegmentAdvTbl"),0,8,"FS Available Wt") 'Fs Available WT
		strFSAvailableVol = fn_getCellData(JavaWindow("SkyChainFlightWin").JavaTable("SegmentAdvTbl"),0,9,"FS Available Wt") ' FS available Vol
		strTotalAvailableWt = fn_getCellData(JavaWindow("SkyChainFlightWin").JavaTable("SegmentAdvTbl"),0,10,"FS Available Wt") 'Total Available Wt
		strTotalAvailableVol = fn_getCellData(JavaWindow("SkyChainFlightWin").JavaTable("SegmentAdvTbl"),0,11,"FS Available Wt") 'Total Available Vol
		If cint(strFSAvailableWt+1) = Cint(strPrevFSAvailableWt) Then
			Reporter.ReportEvent micPass, "FlightRecord update FS Available WT", "After updation flight record FS Available Wt is updated and reduced to "+cstr(strFSAvailableWt)
			rptWriteReport "Pass","FlightRecord update FS Available WT", "After updation flight record FS Available Wt is updated and reduced to "+cstr(strFSAvailableWt)
		else
			Reporter.ReportEvent micFail, "FlightRecord update FS Available WT", "After updation flight record not update in the booking window Wt:"+strPreviousWt+" and flight record:"+cstr(strFSAvailableWt)
			rptWriteReport "Fail","FlightRecord update FS Available WT", "After updation flight record not update in the booking window Wt:"+strPreviousWt+" and flight record:"+cstr(strFSAvailableWt)
		End If
		If (strTotalAvailableWt+1) = Cint(strPrevTotalAvailableWt)  Then
			Reporter.ReportEvent micPass, "FlightRecord update Total Avilable Wt", "After updation flight record TotalAvailable Wt is updated and reduced to "+cstr(strTotalAvailableWt)
			rptWriteReport "Pass","FlightRecord update Total Avilable Wt", "After updation flight record TotalAvailable Wt is updated and reduced to "+cstr(strTotalAvailableWt)
		else	
			Reporter.ReportEvent micFail, "FlightRecord update Total Avilable Wt", "After updation flight record TotalAvailable Wt is not updated in the booking window Wt:"+strPreviousWt+" and flight record:"+cstr(strTotalAvailableWt)
			rptWriteReport "Fail","FlightRecord update Total Avilable Wt", "After updation flight record TotalAvailable Wt is not updated in the booking window Wt:"+strPreviousWt+" and flight record:"+cstr(strTotalAvailableWt)
		End If
		If cint(strFSAvailableVol+1) = cint(strPrevFSAvailableVol) Then
			Reporter.ReportEvent micPass, "FlightRecord update FS Avilable Vol", "After updation flight record FS Avilable Vol is updated and added as "+cstr(strFSAvailableVol)
			rptWriteReport "Pass","FlightRecord update FS Avilable Vol", "After updation flight record FS Avilable Vol is updated and added as "+cstr(strFSAvailableVol)
		else
			Reporter.ReportEvent micFail, "FlightRecord update FS Avilable Vol", "After updation flight record FS Avilable Vol is not updated and in the booking window 1 and in the flight record"+cstr(strFSAvailableVol)
			rptWriteReport "Fail","FlightRecord update FS Avilable Vol", "After updation flight record FS Avilable Vol is not updated and in the booking window 1 and in the flight record"+cstr(strFSAvailableVol)
		End If
		If cint(strTotalAvailableVol+1) =cint(strPrevTotalAvailableVol) Then
			Reporter.ReportEvent micPass, "FlightRecord update Total Avilable Vol", "After updation flight record Total Avilable Vol is updated and added as "+Cstr(strFSAvailableVol)
			rptWriteReport "Pass","FlightRecord update Total Avilable Vol", "After updation flight record FS Avilable Vol is updated and added as "+Cstr(strFSAvailableVol)
		else
			Reporter.ReportEvent micFail, "FlightRecord update FS Avilable Vol", "After updation flight record Total Avilable Vol is not updated and in the booking window 1 and in the flight record"+Cstr(strFSAvailableVol)
			rptWriteReport "Fail","FlightRecord update FS Avilable Vol", "After updation flight record Total Avilable Vol is not updated and in the booking window 1 and in the flight record"+Cstr(strFSAvailableVol)
		End If
		If err.number <> 0 Then
			Reporter.ReportEvent micFail, "Error occured in the function: UpdateShipmentInfo", "Error: " + err.description
			rptWriteReport "Fail","Error occured in the function: UpdateShipmentInfo", "Error: " + err.description
			Environment.Value("gErrorFlag") = true	
		End If
		err.clear
	End If
End Function

Function ValidateDocNetCharges(strDocumentNubmer)
	On error resume next
	If not Environment.Value("gErrorFlag") Then
		Call fn_Set(JavaWindow("SkyChainDocNetRatingWin").JavaEdit("DocNumber"),"DocumentNumber",strDocumentNubmer)
		call fn_Click(JavaWindow("SkyChainDocNetRatingWin").JavaButton("GetCharges"),"GetCharges")
		If fn_CheckForObjecct(JavaWindow("SkyChainDocNetRatingWin").JavaStaticText("Complete"),"Complete",10) Then
			Reporter.ReportEvent micPass, "Doc. Net Charges Status", "After updation Doc. Net Charges status is complete"
			rptWriteReport "Pass","Doc. Net Charges Status", "After updation Doc. Net Charges status is complete"
		else
			Reporter.ReportEvent micFail, "Doc. Net Charges Status", "After updation Doc. Net Charges status is not complete"
			rptWriteReport "Fail","Doc. Net Charges Status", "After updation Doc. Net Charges status is not complete"
		End If
 		set objRateLines = JavaWindow("SkyChainDocNetRatingWin").JavaTable("RateLinesTbl")
		strTemp = objRateLines.GetCellData(0,2)
		If strcomp(strTemp,"Complete",1)=0 Then
			Reporter.ReportEvent micPass, "Doc. Net Charges RateLines", "After updation Doc. Net Charges RateLines status is complete"
			rptWriteReport "Pass","Doc. Net Charges RateLines", "After updation Doc. Net Charges RateLines status is complete"
		else
			Reporter.ReportEvent micFail, "Doc. Net Charges RateLines", "After updation Doc. Net Charges RateLines status is not complete"
			rptWriteReport "Fail","Doc. Net Charges RateLines", "After updation Doc. Net Charges RateLines status is not complete"
		End If
	End If
End Function

Function getData(strItineraryDetails,strPattern)
	On error resume next
	Set re = new Regexp
	re.pattern =strPattern
	re.ignorecase = true
	Set matches = re.Execute(strItineraryDetails)
	If matches.count>0 Then
		For Each match In matches
			getData = match.value
			Exit for
		Next
	else
		getData = ""
	End If	
	If err.number <> 0 Then
		Reporter.ReportEvent micFail, "Error occured in the function: getData", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: getData", "Error: " + err.description
		Environment.Value("gErrorFlag") = true	
	End If
	err.clear
	
End Function


Function CreateItinerary(strCarrier,strFltNo,strOrg,strDest,strBkgSts)
	On error resume next 
	If not Environment.Value("gErrorFlag") Then
		call fn_RClick(JavaWindow("SkyChainHomeWin").JavaObject("ItinearyTablePane"),JavaWindow("SkyChainHomeWin").JavaMenu("AddItinearyMenu"),"Itinarary Table")
		row = "#0"
		strDate = cstr(constructDate())
		Call fn_SetCellData(JavaWindow("SkyChainHomeWin").JavaTable("ItineraryTbl"),row,"Carr.",strCarrier,"Carrier")
		Call fn_SetCellData(JavaWindow("SkyChainHomeWin").JavaTable("ItineraryTbl"),row,"Flt Num",strFltNo,"FlightNumber")
		Call fn_SetCellData(JavaWindow("SkyChainHomeWin").JavaTable("ItineraryTbl"),row,"Date",strDate,"Date")
		Call fn_SetCellData(JavaWindow("SkyChainHomeWin").JavaTable("ItineraryTbl"),row,"Brd Pt",strOrg,"Org")
		Call fn_SetCellData(JavaWindow("SkyChainHomeWin").JavaTable("ItineraryTbl"),row,"Off Pt",strDest,"Dest")
		Call fn_SetCellData(JavaWindow("SkyChainHomeWin").JavaTable("ItinearyAdvTbl"),row,"Bkg Sts",strBkgSts,"BookingStatus")
		If err.number<>0 Then
			Reporter.ReportEvent micFail, "Error occured in the function: CreateItinerary", "Error: " + err.description
			rptWriteReport "Fail","Error occured in the function: CreateItinerary", "Error: " + err.description
			Environment.Value("gErrorFlag") = true	
		End If
		err.clear 
	End If
End Function

Function constructDate
	'dd = day(date)
	dd = 28
	mmm = MonthName(month(date))
	yyyy = year(date)
	constructDate = dd&"-"&mmm&"-"&yyyy
End Function

Function ValidateDBValues(strMsg,strMsgType,strDocumentNumberPrefix1,strDocumentNumberPrefix2,strOrg,strDest,strCallerName)
	On error resume next 
	If not Environment.Value("gErrorFlag") Then
		If instr(1,strMsg,strMsgType,1)>0 Then
			Reporter.ReportEvent micPass, "DB validation", strMsgType&" is found in the data base"
			rptWriteReport "Pass", "DB validation", strMsgType&" is found in the data base"
		else
			Reporter.ReportEvent micFail, "DB validation failed", strMsgType&" not found in the data base"
			rptWriteReport "Fail","DB validation failed", strMsgType&" not found in the data base"
		End If
		If instr(1,strMsg,strDocumentNumberPrefix1,1)>0 Then
			Reporter.ReportEvent micPass, "DB validation", "Document Number1:"&strDocumentNumberPrefix1&"  is found in the data base"
			rptWriteReport "Pass","DB validation", "Document Number1:"&strDocumentNumberPrefix1&"  is found in the data base"
		else	
			Reporter.ReportEvent micFail, "DB validation failed", "Document Number1:"&strDocumentNumberPrefix1&"  not found in the data base"
			rptWriteReport "Fail","DB validation failed", "Document Number1:"&strDocumentNumberPrefix1&"  not found in the data base"
		End If
		If instr(1,strMsg,strDocumentNumberPrefix2,1)>0 Then
			Reporter.ReportEvent micPass, "DB validation", "Document Number2:"&strDocumentNumberPrefix2&"  is found in the data base"
			rptWriteReport "Pass","DB validation", "Document Number2:"&strDocumentNumberPrefix2&"  is found in the data base"
		else
			Reporter.ReportEvent micFail, "DB validation failed", "Document Number2:"&strDocumentNumberPrefix2&"  not found in the data base"
			rptWriteReport "Fail","DB validation failed", "Document Number2:"&strDocumentNumberPrefix2&"  not found in the data base"
		End If
		If instr(1,strMsg,strOrg,1)>0 Then
			Reporter.ReportEvent micPass, "DB validation", "Org:"&strOrg&"  is found in the data base"
			rptWriteReport "Pass", "DB validation", "Org:"&strOrg&"  is found in the data base"
		else
			Reporter.ReportEvent micFail, "DB validation", "Org:"&strOrg&"  is found in the data base"
			rptWriteReport "Fail", "DB validation", "Org:"&strOrg&"  is found in the data base"
		End If
		If instr(1,strMsg,strDest,1)>0 Then
			Reporter.ReportEvent micPass, "DB validation", "Dest:"&strDest&"  is found in the data base"
			rptWriteReport "Pass", "DB validation", "Org:"&strOrg&"  is found in the data base"
		else
			Reporter.ReportEvent micFail, "DB validation Failed", "Dest:"&strDest&"  not found in the data base"
			rptWriteReport "Fail", "DB validation Failed", "Dest:"&strDest&"  not found in the data base"
		End If
		If instr(1,strMsg,strCallerName,1)>0 Then
			Reporter.ReportEvent micPass, "DB validation", "CallerName:"&strCallerName&"  is found in the data base"
			rptWriteReport "Pass", "DB validation", "CallerName:"&strCallerName&"  is found in the data base"
		else
			Reporter.ReportEvent micFail, "DB validation Failed", "CallerName:"&strCallerName&"  not found in the data base"
			rptWriteReport "Fail","DB validation Failed", "CallerName:"&strCallerName&"  not found in the data base"
		End If
		
		
		If err.number <> 0 Then
			Reporter.ReportEvent micFail, "Error occured in the function:ValidateDBValues ", "Error: " + err.description
			rptWriteReport "Fail","Error occured in the function:ValidateDBValues ", "Error: " + err.description
			Environment.Value("gErrorFlag") = true	
		End If
		err.clear
	End If	
	
End Function



Function getValuesFromDB(strUsername,strPassword,strSQL,strDataSource)
	Dim oData
	Dim dbMyDBConnection
	Set oData = CreateObject("ADODB.Recordset")
	Set dbMyDBConnection = CreateObject("ADODB.Connection")
	
	dbMyDBConnection.ConnectionString = "Provider=OraOLEDB.Oracle;Data Source="&strDataSource&";User ID=" & strUsername & ";Password=" & strPassword & ";"
	dbMyDBConnection.Open
	
	oData.Open strSQL, dbMyDBConnection
	
	If Not(oData.EOF) Then
	    strMsg = oData("msg_txt")
	    Reporter.ReportEvent micPass, "DataBase Retrival Passed", "Record found and msg_txt:"&strMsg
		rptWriteReport "Pass","DataBase Retrival Passed", "Record found and msg_txt:"&strMsg

	Else
	    Reporter.ReportEvent micFail, "DataBase Retrival Failed", "No Records found for the following query:"+strSQL
		rptWriteReport "Fail","DataBase Retrival Failed", "No Records found for the following query:"+strSQL
		strMsg =" "
	End If
	
	oData.Close
	dbMyDBConnection.Close
	
	Set oData = Nothing
	Set dbMyDBConnection = Nothing
	getValuesFromDB = strMsg
End Function

SUB SendMail(mailusername,mailpassword,mailto,mailSubject,mailBody)
    If not Environment.Value("gErrorFlag") Then
    	Dim objEmail
	    Const cdoSendUsingPort = 2  ' Send the message using SMTP
	    Const cdoBasicAuth = 1      ' Clear-text authentication
	    Const cdoTimeout = 60       ' Timeout for SMTP in seconds
	
	     mailServer = "smtp.gmail.com"
	     SMTPport = 25     '25 'SMTPport = 465
	     mailusername = "mercator.cigniti@gmail.com"
	     mailpassword = "Ctl@1234"
	
	     'mailto = "raghu.sudam@cigniti.com" 
	     'mailSubject = "my test-deleteme" 
	     'mailBody = "This is the email body" 
	
	    Set objEmail = CreateObject("CDO.Message")
	    Set objConf = objEmail.Configuration
	    Set objFlds = objConf.Fields
	
	    With objFlds
	        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
	        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = mailServer
	    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPport
	    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	        .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = cdoTimeout
	    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasicAuth
	    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = mailusername
	    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = mailpassword
	        .Update
	    End With
	
	    objEmail.To = mailto
	    objEmail.From = mailusername
	    objEmail.Subject = mailSubject
	    objEmail.TextBody = mailBody
	    'objEmail.AddAttachment "C:\report.pdf"
	    objEmail.Send
	
	    Set objFlds = Nothing
	    Set objConf = Nothing
	    Set objEmail = Nothing
	else
		msgbox "Test Case failed, no mail sent"
    End If
    
END SUB

Function SetMaskCode(strMaskCode)
	On error resume next 
	If not Environment.Value("gErrorFlag") Then
		call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("Mask Code"),"Mask Code",strMaskCode)
		call fn_DoubleClick(JavaWindow("SkyChainHomeWin").JavaEdit("Mask Code"),strMaskCode)				
		If fn_CheckForObjecct(JavaWindow("SkyChainHomeWin").JavaDialog("MaskCodesWin").JavaTable("FindMaskCode"),"MaskCodeTable",10) Then
			JavaWindow("SkyChainHomeWin").JavaDialog("MaskCodesWin").JavaTable("FindMaskCode").ClickCell 0,1
		    call fn_Click(JavaWindow("SkyChainHomeWin").JavaDialog("MaskCodesWin").JavaButton("OK"),"Ok")
		    While JavaWindow("SkyChainHomeWin").JavaDialog("MaskCodesWin").JavaTable("FindMaskCode").Exist(0)
		    	wait 0,500
		    Wend		    
		else
			rptWriteReport "Fail","Error occured in the function: SetMaskCode", "Error: " + err.description
			Environment.Value("gErrorFlag") = true
		End If		
		If err.number <> 0 Then		
			Reporter.ReportEvent micFail, "Error occured in the function: SetMaskCode", "Error: " + err.description
			rptWriteReport "Fail","Error occured in the function: SetMaskCode", "Error: " + err.description
			Environment.Value("gErrorFlag") = true
		End If
		err.clear
	End If	
End Function

Function DeleteRoutingAndItinerary()
	'On error resume next 
	If not Environment.Value("gErrorFlag") Then
		If JavaWindow("SkyChainHomeWin").JavaTable("RoutingInfoTbl").GetROProperty("rows")>0 Then
			call fn_RClick(JavaWindow("SkyChainHomeWin").JavaTable("RoutingInfoTbl"),JavaWindow("SkyChainHomeWin").JavaMenu("DeleteRouting"),15,10,"DeleteRouting")
			call fn_CheckForObjecct(JavaWindow("SkyChainHomeWin").JavaDialog("ConfirmationWin").JavaButton("Yes"),"Ok",10)
			Call fn_Click(JavaWindow("SkyChainHomeWin").JavaDialog("ConfirmationWin").JavaButton("Yes"),"Ok")
		else
			Reporter.ReportEvent micPass, "Routing Info Table", "No row available"
			rptWriteReport "Pass","Routing Info Table", "No row available"
		End If
		If JavaWindow("SkyChainHomeWin").JavaTable("ItineraryTbl").GetROProperty("rows")>0 Then
			call fn_RClick(JavaWindow("SkyChainHomeWin").JavaTable("ItineraryTbl"),JavaWindow("SkyChainHomeWin").JavaMenu("DeleteItinerary"),15,10,"Delete Itinerary")
			call fn_CheckForObjecct(JavaWindow("SkyChainHomeWin").JavaDialog("ConfirmationWin").JavaButton("Yes"),"Ok",10)
			Call fn_Click(JavaWindow("SkyChainHomeWin").JavaDialog("ConfirmationWin").JavaButton("Yes"),"Ok")
		else 
			Reporter.ReportEvent micPass, "Itinerary Table", "No row available"
			rptWriteReport "Pass","Itinerary Table", "No row available"
		End If
		If err.number <> 0 Then		
			Reporter.ReportEvent micFail, "Error occured in the function: DeleteRoutingAndItinerary", "Error: " + err.description
			rptWriteReport "Fail","Error occured in the function: DeleteRoutingAndItinerary", "Error: " + err.description
			Environment.Value("gErrorFlag") = true
		End If
		err.clear
	End If		
End Function

Function clickRoutingAndSelectFlight()
If not Environment.Value("gErrorFlag") Then
	Call fn_Click(JavaWindow("SkyChainHomeWin").JavaButton("RoutingBtn"),"Routing")
	If err.number <>0 Then
		Reporter.ReportEvent micFail, "Error occured in the function: clickRoutingAndSelectFlight", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: clickRoutingAndSelectFlight", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	else
		If JavaWindow("SkyChainHomeWin").JavaDialog("RoutingDetailsWin").JavaTable("RoutingDetailsTbl").Exist(20) Then
			set objRouting = JavaWindow("SkyChainHomeWin").JavaDialog("RoutingDetailsWin").JavaTable("RoutingDetailsTbl")
			rows = objRouting.GetROProperty("rows")
			cols = objRouting.GetROProperty("cols")
			strItinearyValue = "" 
			For row = 1 To rows
				For col = 0 To cols-1 Step 1
					strTemp = objRouting.GetCellData(row,col)
					If len(trim(strTemp))>0 Then
						 strItinearyValue = strItinearyValue+";"+Cstr(objRouting.GetCellData(row,col))
					End If						
				Next
				Exit for 
			Next			
			objRouting.ClickCell 1,1
			call fn_Click(JavaWindow("SkyChainHomeWin").JavaDialog("RoutingDetailsWin").JavaButton("SelectBtn"),"Select")
			While JavaWindow("SkyChainHomeWin").JavaDialog("RoutingDetailsWin").Exist(0)
				wait 0,500
			Wend
			clickRoutingAndSelectFlight = strItinearyValue
		Else 
			Reporter.ReportEvent micFail, "Error occured in the function: clickRoutingAndSelectFlight", "Error: " + err.description
			rptWriteReport "Fail","Error occured in the function: clickRoutingAndSelectFlight", "Error: " + err.description
			Environment.Value("gErrorFlag") = true
		End if 
	End if 	
	If err.number <>0 Then
		Reporter.ReportEvent micFail, "Error occured in the function: clickRoutingAndSelectFlight", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: clickRoutingAndSelectFlight", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End if 
	err.clear
End If		
End Function

Function WaitForControllerToClose
On error resume next
If Not Environment.Value("gErrorFlag") Then
	count=0
	Do
		wait 0,500
		count = count+1
		If count>20 Then
			Exit do
		End If
	Loop While JavaWindow("SkyChainHomeWin").JavaDialog("BusyStatusController").Exist(0)
	wait 1
	If JavaWindow("SkyChainHomeWin").JavaDialog("BusyStatusController").Exist(0) Then
		Reporter.ReportEvent micFail, "WaitForControllerToClose", "Application is taking more time to load"
		rptWriteReport "Fail","WaitForControllerToClose", "Application is taking more time to load"
		Environment.Value("gErrorFlag") = true		
	End If
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: WaitForControllerToClose", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: WaitForControllerToClose", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If
	err.clear	
End If
End Function


Function getShipmentData(strPieces,strWeight,strVolue)
On error resume next
If not Environment.Value("gErrorFlag") Then
	strPieces = fn_getText(JavaWindow("SkyChainHomeWin").JavaEdit("Pieces"),"Pieces")
	strWeight = fn_getText(JavaWindow("SkyChainHomeWin").JavaEdit("Wt-Vol"),"Weight")
	strVolue =  fn_getText(JavaWindow("SkyChainHomeWin").JavaEdit("CargoNumericField"),"Volume")

	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: getShipmentData", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: getShipmentData", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If
	err.clear
End If
End Function 

Function AcceptConfirmationIfExists(strObjName)
On error resume next 
If not Environment.Value("gErrorFlag") Then
	call WaitForControllerToClose()
	wait 2
	If strcomp(ucase(strObjName),"YES",1)=0 Then
		If fn_CheckForObjecct(JavaWindow("SkyChainHomeWin").JavaDialog("ConfirmationWin").JavaButton("Yes"),"Yes",5) Then
			call fn_Click(JavaWindow("SkyChainHomeWin").JavaDialog("ConfirmationWin").JavaButton("Yes"),"Yes")
		else
		    Reporter.ReportEvent micPass, "AcceptConfirmationIfExists", "No confirmation message"
			rptWriteReport "Pass","AcceptConfirmationIfExists", "No confirmation message"
		End If		
	ElseIf strcomp(ucase(strObjName),"OK",1)=0 Then
		If fn_CheckForObjecct(JavaWindow("SkyChainHomeWin").JavaDialog("ConfirmationWin").JavaButton("Ok"),"Ok",5) Then
			call fn_Click(JavaWindow("SkyChainHomeWin").JavaDialog("ConfirmationWin").JavaButton("Ok"),"Ok")
		else
		    Reporter.ReportEvent micPass, "AcceptConfirmationIfExists", "No confirmation message"
			rptWriteReport "Pass","AcceptConfirmationIfExists", "No confirmation message"
		End If			
	End If

End If
End Function

Function AddLocationToLocationTable(strLocation)
On error resume next
If not Environment.Value("gErrorFlag") Then
	JavaWindow("SkyChainHomeWin").JavaDialog("Modify Suggested Location").JavaTable("LocationTbl").ClickCell "#0","Location"
	wait 1
	Set obj = CreateObject("WScript.Shell")
	obj.SendKeys "{F9}"
	If fn_CheckForObjecct(JavaWindow("SkyChainHomeWin").JavaDialog("Location Search").JavaTable("LocationSearchTbl"),"Location Search",5) Then
		Call fn_Set(JavaWindow("SkyChainHomeWin").JavaDialog("Location Search").JavaEdit("Location"),"Location",strLocation)
		Call fn_Click(JavaWindow("SkyChainHomeWin").JavaDialog("Location Search").JavaButton("FindBtn"),"Find")
		wait 1
		call fn_ClickCell(JavaWindow("SkyChainHomeWin").JavaDialog("Location Search").JavaTable("LocationSearchTbl"),"#0","Location",strObjName)
		Call fn_Click(JavaWindow("SkyChainHomeWin").JavaDialog("Location Search").JavaButton("OkBtn"),"Ok")
		Call fn_Click(JavaWindow("SkyChainHomeWin").JavaDialog("Modify Suggested Location").JavaButton("SaveBtn"),"Save")
		Call AcceptConfirmationIfExists("OK")
		Call AcceptConfirmationIfExists("OK")
	else
		Reporter.ReportEvent micFail, "LocationSearch", "Location search window didn't show up"
		rptWriteReport "Fail","LocationSearch", "Location search window didn't show up"
		Environment.Value("gErrorFlag") = true
	End If	
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: getShipmentData", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: getShipmentData", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If	
	err.clear
End If
End Function

Function SearchInWorkBench(strCarrier,strFlightNumber,StrSuffix,StrDate)
On error resume next 
If not Environment.Value("gErrorFlag") Then
	Call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("Carrier"),"Carrier",strCarrier)
	Call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("FlightNo"),"Flight Nubmer",strFlightNumber)
	Call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("Suffix"),"Sufix",StrSuffix)
	Call fn_Set(JavaWindow("SkyChainHomeWin").JavaEdit("Date"),"Date",StrDate)
	Call fn_Click(JavaWindow("SkyChainHomeWin").JavaButton("SearchFlight"),"SearchFlight")
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: getShipmentData", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: getShipmentData", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If	
	err.clear	
End If
End Function


Function ValidateWorkBench(strCarrier,strFlightNumber,StrSuffix,StrDate)
On error resume next 
If not Environment.Value("gErrorFlag") Then
	If JavaWindow("SkyChainHomeWin").JavaTable("FlightDetails").GetROProperty("rows")>0 Then
		strAppCarrier = fn_getCellData(JavaWindow("SkyChainHomeWin").JavaTable("FlightDetails"),"#0","Carrier Code","Carrier Code")
		strAppFlightNubmer = fn_getCellData(JavaWindow("SkyChainHomeWin").JavaTable("FlightDetails"),"#0","Flight","Flight")
		strAppSuffix = fn_getCellData(JavaWindow("SkyChainHomeWin").JavaTable("FlightDetails"),"#0","Flt. Sufx","Flt. Sufx")
		strAppDate = fn_getCellData(JavaWindow("SkyChainHomeWin").JavaTable("FlightDetails"),"#0","Sch. Date","Sch. Date")
		
		If strcomp(strAppCarrier,strCarrier,1) =0 Then
			Reporter.ReportEvent micPass, "FlightDetails", "Carrier details are maching into both AWB Capture & Flight preparation work bench:"&strCarrier
			rptWriteReport "Pass","FlightDetails", "Carrier details are maching into both AWB Capture & Flight preparation work bench:"&strCarrier
		else
			Reporter.ReportEvent micFail, "FlightDetails", "Carrier details are not maching, AWB Capture:"&strCarrier&" & Flight preparation work bench:"&strAppCarrier
			rptWriteReport "Fail","FlightDetails", "Carrier details are not maching, AWB Capture:"&strCarrier&" & Flight preparation work bench:"&strAppCarrier
		End If
		
		If strAppFlightNubmer = cint(strFlightNubmer) Then
			Reporter.ReportEvent micPass, "FlightDetails", "Flight Number is maching into both AWB Capture & Flight preparation work bench:"&strFlightNubmer
			rptWriteReport "Pass","FlightDetails", "Flight Number is maching into both AWB Capture & Flight preparation work bench:"&strFlightNubmer
		else
			Reporter.ReportEvent micFail, "FlightDetails", "Flight Number is not maching, AWB Capture:"&strFlightNubmer&" & Flight preparation work bench:"&strAppFlightNubmer
			rptWriteReport "Fail","FlightDetails",  "Flight Number is not maching, AWB Capture:"&strFlightNubmer&" & Flight preparation work bench:"&strAppFlightNubmer
		End If
		
		If strcomp(strAppSuffix,StrSuffix,1) =0 Then
			Reporter.ReportEvent micPass, "FlightDetails", "Flight Suffix details are maching into both AWB Capture & Flight preparation work bench:"&StrSuffix
			rptWriteReport "Pass","FlightDetails", "Flight Suffix details are maching into both AWB Capture & Flight preparation work bench:"&StrSuffix
		else
			Reporter.ReportEvent micFail, "FlightDetails", "Flight Suffix are not maching, AWB Capture:"&StrSuffix&" & Flight preparation work bench:"&strAppSuffix
			rptWriteReport "Fail","FlightDetails", "Flight Suffix are not maching, AWB Capture:"&StrSuffix&" & Flight preparation work bench:"&strAppSuffix
		End If
		
		If strcomp(cstr(strAppDate),StrDate,1) =0 Then
			Reporter.ReportEvent micPass, "FlightDetails", "FlightDate details are maching into both AWB Capture & Flight preparation work bench:"&StrDate
			rptWriteReport "Pass","FlightDetails", "FlightDate details are maching into both AWB Capture & Flight preparation work bench:"&StrDate
		else
			Reporter.ReportEvent micFail, "FlightDetails", "FlightDate details are not maching, AWB Capture:"&StrDate&" & Flight preparation work bench:"&strAppDate
			rptWriteReport "Fail","FlightDetails", "FlightDate details are not maching, AWB Capture:"&StrDate&" & Flight preparation work bench:"&strAppDate
		End If
		
	else
		Reporter.ReportEvent micFail, "FlightDetails", "No records found"
		rptWriteReport "Fail","FlightDetails", "No records found"
		Environment.Value("gErrorFlag") = true
	End If	
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: ValidateWorkBench FlightDetails", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: ValidateWorkBench FlightDetails", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If	
	err.clear
End If
End Function

Function generateRandomNumber
	Dim max,min,rand
	max=99999
	min=1
	Randomize
	rand = Int((max-min+1)*Rnd+min)
	generateRandomNumber= rand	
End Function

Function ValidateAWBNumber(strDocumentPrefix1,strDocumentPrefix2,strPieces,strWeight,StrVolume)
On error resume next 
If not Environment.Value("gErrorFlag") Then
	If JavaWindow("SkyChainHomeWin").JavaTable("BookedShipmentsTbl").GetROProperty("rows")>0 Then
		row = findRows(strDocumentPrefix1&"-"&strDocumentPrefix2)
		strAppDocumentNumber = fn_getCellData(JavaWindow("SkyChainHomeWin").JavaTable("BookedShipmentsTbl"),row,"Doc Nbr.","Doc Nbr.")
		strAppPieces = fn_getCellData(JavaWindow("SkyChainHomeWin").JavaTable("BookedShipmentsAdvTbl"),row,"Pieces","Pieces")
		strAppWeight = fn_getCellData(JavaWindow("SkyChainHomeWin").JavaTable("BookedShipmentsAdvTbl"),row,"Weight","Weight")
		strAppVol = fn_getCellData(JavaWindow("SkyChainHomeWin").JavaTable("BookedShipmentsAdvTbl"),row,"Volume","Volume")
		strDocNubr = (strDocumentNumber&"-"&strDocumentPrefix2)
		If strcomp(strAppDocumentNumber,strDocNubr,1) =0 Then
			Reporter.ReportEvent micPass, "BookedShipments", "Document Number details are maching into both AWB Capture & Flight preparation work bench:"&strDocNubr
			rptWriteReport "Pass","BookedShipments", "Document Number details are maching into both AWB Capture & Flight preparation work bench:"&strDocNubr
		else
			Reporter.ReportEvent micFail, "BookedShipments", "Document Number details are not maching into both AWB Capture:"&strDocNubr &" & Flight preparation work bench:"&strAppDocumentNumber
			rptWriteReport "Fail", "BookedShipments", "Document Number details are not maching into both AWB Capture:"&strDocNubr &" & Flight preparation work bench:"&strAppDocumentNumber
		End If
		
		If strcomp(strAppPieces,strPieces,1) =0 Then
			Reporter.ReportEvent micPass, "BookedShipments", "Pieces details are maching into both AWB Capture & Flight preparation work bench:"&strPieces
			rptWriteReport "Pass","BookedShipments", "Pieces details are maching into both AWB Capture & Flight preparation work bench:"&strPieces
		else
			Reporter.ReportEvent micFail, "BookedShipments", "Pieces Number details are not maching into both AWB Capture:"&strPieces &" & Flight preparation work bench:"&strAppPieces
			rptWriteReport "Fail","BookedShipments", "Pieces Number details are not maching into both AWB Capture:"&strPieces &" & Flight preparation work bench:"&strAppPieces
		End If
		
		If strcomp(strAppWeight,strWeight,1) =0 Then
			Reporter.ReportEvent micPass, "BookedShipments", "Weight details are maching into both AWB Capture & Flight preparation work bench:"&strWeight
			rptWriteReport "Pass","BookedShipments", "Weight details are maching into both AWB Capture & Flight preparation work bench:"&strWeight
		else
			Reporter.ReportEvent micFail, "BookedShipments", "Weight Number details are not maching into both AWB Capture:"&strWeight &" & Flight preparation work bench:"&strAppWeight
			rptWriteReport "Fail", "BookedShipments", "Weight Number details are not maching into both AWB Capture:"&strWeight &" & Flight preparation work bench:"&strAppWeight
		End If
		
		If strcomp(StrAppVolume,StrVolume,1) =0 Then
			Reporter.ReportEvent micPass, "BookedShipments", "Volume details are maching into both AWB Capture & Flight preparation work bench:"&StrVolume
			rptWriteReport "Pass","BookedShipments", "Volume details are maching into both AWB Capture & Flight preparation work bench:"&StrVolume
		else
			Reporter.ReportEvent micFail, "BookedShipments", "Volume Number details are not maching into both AWB Capture:"&StrVolume &" & Flight preparation work bench:"&strAppWeight
			rptWriteReport "Fail","BookedShipments", "Volume Number details are not maching into both AWB Capture:"&StrVolume &" & Flight preparation work bench:"&strAppWeight
		End If
		
	else
		Reporter.ReportEvent micFail, "BookedShipments", "No records found"
		rptWriteReport "Fail","BookedShipments", "No records found"
		Environment.Value("gErrorFlag") = true
	End If	
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: BookedShipments", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: BookedShipments", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If	
	err.clear
End If	
End Function

Function findRows(strAwbNumber)
On error resume next 
If not Environment.Value("gErrorFlag") Then
	temp = "#0"
	rows = JavaWindow("SkyChainHomeWin").JavaTable("BookedShipmentsTbl").GetROProperty("rows")
	For row = 0 To rows-1 Step 1
		awbNumber = JavaWindow("SkyChainHomeWin").JavaTable("BookedShipmentsTbl").GetCellData("#"&row,"Doc Nbr.")
		If strcomp(awbNumber,strAwbNumber,1) =0 Then
			temp = "#"&row
			Exit for 
		End If
	Next
	findRows = temp
End If
End Function


Function PreManifestGenerate()
On error resume next 
If not Environment.Value("gErrorFlag") Then
	Call fn_RClick(JavaWindow("SkyChainHomeWin").JavaTable("BookedShipmentsTbl"),JavaWindow("SkyChain (EJB3_Product").JavaMenu("Pre-Manifest").JavaMenu("Generate"),69,8,objstrObjectName)
	If err.number <> 0 Then		
		Reporter.ReportEvent micFail, "Error occured in the function: PreManifestGenerate", "Error: " + err.description
		rptWriteReport "Fail","Error occured in the function: PreManifestGenerate", "Error: " + err.description
		Environment.Value("gErrorFlag") = true
	End If	
	err.clear
End If
End Function

Function queryDBForJobId(strUsername,strPassword,strSQL,strDataSource)
	If not Environment.Value("gErrorFlag") Then
		Dim oData
		Dim dbMyDBConnection
		Set oData = CreateObject("ADODB.Recordset")
		Set dbMyDBConnection = CreateObject("ADODB.Connection")
		
		dbMyDBConnection.ConnectionString = "Provider=OraOLEDB.Oracle;Data Source="&strDataSource&";User ID=" & strUsername & ";Password=" & strPassword & ";"
		dbMyDBConnection.Open
		Do
			jobNumber = generateRandomNumber()			
			strTempQ= Replace(strSQL,"%jobNumber%",jobNumber)
			oData.Open strTempQ, dbMyDBConnection
			
			If (oData.EOF) Then
			    Reporter.ReportEvent micPass, "DataBase Retrival", "No Records found for the following query:"+strTempQ
				rptWriteReport "Pass","DataBase Retrival", "No Records found for the following query:"+strTempQ
				Exit do 
			End If
		Loop While true		
		oData.Close
		dbMyDBConnection.Close		
		Set oData = Nothing
		Set dbMyDBConnection = Nothing
		queryDBForJobId = jobNumber
	End If
	
End Function


Function UpdateExcelSheet(strJobid,awb,strCarrier,strFlightNumber,strSuffix)
	On error resume next 
	Dim objExcel, strExcelPath, objSheet
	
	strExcelPath =  strProjectTestdataPath&Environment("TestName")&".xls"
	
	' Open specified spreadsheet and select the first worksheet.
	Set objExcel = CreateObject("Excel.Application")
	objExcel.WorkBooks.Open strExcelPath
	Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
	cols = objSheet.UsedRange.Columns.count
	If  not Environment.Value("gErrorFlag") Then
		For col = cols To cols\2 Step -1
			strColumnName = objSheet.cells(1,col)
			If strcomp(strColumnName,"JobNumber",1)=0 Then
				objSheet.cells(2,col).value = strJobid
			ElseIf strcomp(strColumnName,"AWB",1)=0 Then
				objSheet.cells(2,col).value = awb
			ElseIf strcomp(strColumnName,"Carrier",1)=0 Then
				objSheet.cells(2,col).value = strCarrier
			ElseIf strcomp(strColumnName,"FlightNumber",1)=0  Then
				objSheet.cells(2,col).value = strFlightNumber
			ElseIf strcomp(strColumnName,"Suffix",1)=0 Then
				objSheet.cells(2,col).value = strSuffix
			End If
		Next
	Else 	
		objSheet.cells(2,cols).Value =  "Fail"
	End If
	
	
	' Save and quit.
	objExcel.ActiveWorkbook.Save
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit
End Function

Function constructMailBody(strFilePath,awb,strPieces,strFlightNubmer,strDate,strStartTime,strJobid)
	strUnformatedMsg =  readTextFile(strFilePath)
	strUnformatedMsg = Replace(strUnformatedMsg,"%awbNumber%",awb)
	strUnformatedMsg = replace(strUnformatedMsg,"%pieces%",strPieces)
	strUnformatedMsg = replace(strUnformatedMsg,"%flightNumber%",strFlightNubmer)
	strtemp = constructServerDateFormat(strDate)	
	strUnformatedMsg = replace(strUnformatedMsg,"%date%",strtemp)
	strUnformatedMsg = replace(strUnformatedMsg,"%startTime%",strStartTime)
	strUnformatedMsg = replace(strUnformatedMsg,"%%jobNo%%",strJobid)
	writeInTextFile(strUnformatedMsg)
	constructMailBody = strUnformatedMsg
End Function

Function writeInTextFile(strUnformatedMsg)
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	strPath = strProjectPath&"Results\SendMail"
	temp = replace(now(),"/","_")
	temp = replace(temp,":","_")
	temp = replace(temp," ","_")
	strfileName = "\mail"&temp&".txt"
	outFile=strPath&strfileName
	Set objFile = objFSO.CreateTextFile(outFile,True)
	objFile.WriteLine strUnformatedMsg	
	objFile.Close
End Function

Function constructServerDateFormat(strDate)
	yyyy = Year(cDate(strDate))
	If month(cDate(strDate))<10 Then
		mm = "0"&month(strDate)
	Else 
		mm = month(cDate(strDate))
	End If
	dd = day(cDate(strDate))
	constructServerDateFormat = yyyy&"-"&mm&"-"&dd
End Function

Function readTextFile(strFilePath)
	Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(strProjectTestdataPath&"\sendMail.txt",1)
	strFileText = objFileToRead.ReadAll()
	objFileToRead.Close
	Set objFileToRead = Nothing	
	readTextFile = strFileText
End Function

