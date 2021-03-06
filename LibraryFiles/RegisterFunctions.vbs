' List of Functions
'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fneditset
'	Objective							:					Used to set the value in any textbox, Check and uncheck checkboxes and set radio button in any environment 
'	Input Parameters					:					objEdit,strValue
'	Output Parameters					:					NIL
'	Date Created						:					NIL
'	QTP Version							:					NIL
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		            NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fneditset(objEdit,strValue)
	Wait (MIN_WAIT)	
    Dim sFlag
	sFlag = True
    sObjName=objEdit.GetToProperty("TestobjName")
    If  objEdit.Exist(MIN_WAIT) Then
        If  objEdit.GetROProperty("disabled") = 0 Then
            objEdit.Set strValue
            Call rptWriteReport("Pass", sObjName,strValue&" is enterd into " &sObjName)
        Else
        	Call rptWriteReport("Fail", sObjName,strValue&" not entered as the field " &sObjName &" is Exist but not enabled")
            sFlag = False			
        End If
    Else
    		Call rptWriteReport("Fail", sObjName,strValue&" not entered as the field " &sObjName &" doesn't Exist")
   			sFlag = False			
    End If
    If sFlag = False Then
		Call fnErrorChecking()
	End If 
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fnselectitem
'	Objective							:					Used to select the value in any listbox, ComboBox and select the value of RadioGroup in any environment 
'	Input Parameters					:					objList,strValue
'	Output Parameters					:					NIL
'	Output Parameters					:					NIL
'	Date Created						:					NIL
'	QTP Version							:					NIL
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		            NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fnselectitem(objList,strValue)
	Wait (MIN_WAIT)
	Dim sFlag
	sFlag = True
    sObjName=objList.GetToProperty("TestobjName")
    If Len(strValue)>0 Then
	    If  objList.Exist(MIN_WAIT) Then
	        If  objList.GetROProperty("disabled") = 0 Then
	            objList.Select strValue
	            Call fnReportDetailedSuccess(sObjName,strValue&" Item is selected in "&sObjName)                    
	        Else
	            Call fnReportDetailedFailure(sObjName,"The object is disabled")
				sFlag = False			
	        End If
	    Else
	        Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
			sFlag = False		
	    End If
	 Else
	 	If  objList.Exist(MIN_WAIT) Then
	        If  objList.GetROProperty("disabled") = 0 Then
	            objList.Select 
	            Call fnReportDetailedSuccess(sObjName,strValue&" Item is selected in "&sObjName)                    
	        Else
	            Call fnReportDetailedFailure(sObjName,"The object is disabled")
				sFlag = False			
	        End If
	    Else
	        Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
			sFlag = False		
	    End If
	  End If 
	If sFlag = False Then
		Call fnErrorChecking()
	End If
End Function

'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fnpress
'	Objective							:					Used to Click on any Button, Image, Link in any environment 
'	Input Parameters					:					objButton
'	Output Parameters					:					NIL
'	Output Parameters					:					NIL
'	Date Created						:					NIL
'	QTP Version							:					NIL
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		            NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fnpress(objButton)
  Wait (MIN_WAIT)
  Dim sFlag
  sFlag = True
  sObjName=objButton.GetToProperty("TestobjName")
    If  objButton.Exist(MIN_WAIT) Then
        If  objButton.GetROProperty("disabled") = 0 Then
            objButton.Click
            Call rptWriteReport("Pass", sObjName,"Click operation performed on "&sObjName)  
                           
        Else
        	Call rptWriteReport("Fail", sObjName,"The object is disabled")  
            
            sFlag = False			
        End If
    Else
    	Call rptWriteReport("Fail", sObjName,"The object doesn't Exist")  
        sFlag = False			
    End If
	If sFlag = False Then
		Call fnErrorChecking()
	End If
End Function


''******************************************************************************************************************************************************************************************************************************************
''	Function Name						:					fntbcelldata
''	Objective							:					Used to Click on any Button, Image, Link in any environment 
''	Input Parameters					:					objTable
''	Output Parameters					:					NIL
''	Output Parameters					:					NIL
''	Date Created						:					NIL
''	QTP Version							:					NIL
''	QC Version							:					NIL
''	Pre-requisites						:					NIL
''	Created By							:					NIL
''	Modification Date					:		            NIL
''******************************************************************************************************************************************************************************************************************************************
'Public Function fntbcelldata(objTable)
'    Dim sFlag
'	sFlag = True
'    sObjName=objTable.GetToProperty("TestobjName")
'    If  objTable.Exist(MIN_WAIT) Then
'        If  objTable.GetROProperty("disabled") = 0 Then
'            cData=objTable.GetCellData(2,1)'Its for example we have defined row as 1 and column as 2 if u want all data iterate a loop by finding RowCoumnt, CoumnCoumnt
'            Call fnReportDetailedSuccess(sObjName,"cell value is" &cData)                    
'        Else
'            Call fnReportDetailedFailure(sObjName,"The object is disabled")
'            sFlag = False			
'        End If
'    Else
'        Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
'        sFlag = False			
'    End If
'	If sFlag = False Then
'		Call fnErrorChecking()
'	End If
'End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fnintext
'	Objective							:					Used to get the innertext of  all Objects(Validation)
'	Input Parameters					:					objElement
'	Output Parameters					:					strInrText
'	Output Parameters					:					NIL
'	Date Created						:					NIL
'	QTP Version							:					NIL
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		   NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fnintext(objElement,sPropertyName,sVal,sMessage)
    Wait (MIN_WAIT)
	Dim sFlag
	sFlag = True
    sObjName=objElement.GetToProperty("TestobjName")
    If  objElement.Exist(MIN_WAIT) Then
            text=objElement.GetROProperty(sPropertyName) 
            If Trim(sVal)=Trim(text) Then
            	Call fnReportDetailedSuccess(sMessage,"The Validation is Success and Value is "&sVal)
           	Else
           		Call fnReportDetailedFailure(sMessage,"The Validation failed.")
				sFlag = False			
            End If
    Else
        Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
        sFlag = False			
    End If
	If sFlag = False Then
		Call fnErrorChecking()
	End If
End Function

'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fnitempress
'	Objective							:					Used to select any option in ToolBar Menu of any Environment
'	Input Parameters					:					objToolBar, strItem
'	Output Parameters					:					NIL
'	Output Parameters					:					NIL
'	Date Created						:					NIL
'	QTP Version							:					NIL
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		  			NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fnitempress(objToolBar, strItem)
	Wait (MIN_WAIT)
	Dim sFlag
	sFlag = True
    sObjName = objToolBar.GetToProperty("TestobjName")
    If  objToolBar.Exist(MIN_WAIT) Then
        If  objToolBar.GetROProperty("disabled") = 0 Then
            objToolBar.Press strItem 
            Call fnReportDetailedSuccess(sObjName,"Item pressed in context menu")                    
        Else
            Call fnReportDetailedFailure(sObjName,"The object is disabled")
            sFlag = False			
        End If
    Else
        Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
         sFlag = False			
    End If
    If sFlag = False Then
		Call fnErrorChecking()
	End If
End Function

'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fngettext
'	Objective							:					Retrieving the innertext all Objects(Verification)
'	Input Parameters					:					objElement,sPropertyName,sMessage
'	Output Parameters					:					NIL
'	Output Parameters					:					NIL
'	Date Created						:					NIL
'	QTP Version							:					NIL
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		  			NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fngettext(objElement,sPropertyName,sMessage)
	wait (MIN_WAIT)
    Dim sFlag
	sFlag = True
    sObjName=objElement.GetToProperty("TestobjName")
    If  objElement.Exist(MIN_WAIT) Then
            text=objElement.GetROProperty(sPropertyName) 
            Call fnReportDetailedSuccess(sMessage,"Value is Retrieved and the Value is " &text)
    Else
        Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
        sFlag = False			
    End If
    fngettext = text
    If sFlag = False Then
		Call fnErrorChecking()
	End If
End Function

'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fnpophandler
'	Objective							:					Handling the pop-up functions.
'	Input Parameters					:					objToolBar
'	Output Parameters					:					NIL
'	Output Parameters					:					NIL
'	Date Created						:					NIL
'	QTP Version							:					NIL
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		  			NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fnpopuphandler(objToolBar)
    Wait (MIN_WAIT)
	Dim sFlag
	sFlag = True
    sObjName = objToolBar.GetToProperty("TestobjName")
    If  objToolBar.Exist(MIN_WAIT) Then
        If  objToolBar.GetROProperty("disabled") = 0 Then
            objToolBar.Click
            Call fnReportDetailedSuccess(sObjName,"Pop Up Handled")                    
        Else
            Call fnReportDetailedFailure(sObjName,"The object is disabled")
            sFlag = False			
        End If
    Else
        Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
         sFlag = False			
    End If
    If sFlag = False Then
		Call fnErrorChecking()
	End If
End Function

'****************************************RegisterFunc for Activex Environment******************************
	RegisterUserFunc "AcxEdit","fneditset","fneditset"
	RegisterUserFunc "AcxCheckBox","fneditset","fneditset" 
	RegisterUserFunc "AcxRadioButton","fneditset","fneditset" 
	RegisterUserFunc "AcxCalender","fneditset","fneditset" 	
	RegisterUserFunc "AcxComboBox","fnselectitem","fnselectitem" 
	RegisterUserFunc "AcxButton","fnpress","fnpress" 
	RegisterUserFunc "AcxTable","fntbcelldata","fntbcelldata" 
	
'/****************************************End RegisterFunc for Activex Environment******************************

'****************************************RegisterFunc for Java Environment******************************
	RegisterUserFunc "JavaEdit","fneditset","fneditset"
	RegisterUserFunc "JavaCheckBox","fneditset","fneditset" 
	RegisterUserFunc "JavaRadioButton","fneditset","fneditset" 
	RegisterUserFunc "JavaCalender","fneditset","fneditset" 
	RegisterUserFunc "JavaList","fnselectitem","fnselectitem"
	RegisterUserFunc "JavaButton","fnpress","fnpress"
	RegisterUserFunc "JavaLink","fnpress","fnpress"
	RegisterUserFunc "JavaTable","fntbcelldata","fntbcelldata"
	RegisterUserFunc "JavaToolBar","fnitempress","fnitempress"
'/****************************************End RegisterFunc for Java Environment******************************

'****************************************RegisterFunc for SAP Wiondows Environment******************************
	RegisterUserFunc "SAPGuiEdit","fneditset","fneditset" 
	RegisterUserFunc "SAPGuiCheckBox","fneditset","fneditset"
	RegisterUserFunc "SAPGuiRadioButton","fneditset","fneditset" 
	RegisterUserFunc "SAPGuiCalender","fneditset","fneditset"
	RegisterUserFunc "SAPGuiComboBox","fnselectitem","fnselectitem"
	RegisterUserFunc "SAPGuiButton","fnpress","fnpress"
	RegisterUserFunc "SAPGuiElement","fnintext","fnintext"
	RegisterUserFunc "SAPGuiTable","fntbcelldata","fntbcelldata"
	RegisterUserFunc "SAPGuiToolBar","fnitempress","fnitempress" 
'/****************************************End RegisterFunc for SAP Wiondows  Environment******************************

'****************************************RegisterFunc for SAP Web Environment******************************
	RegisterUserFunc "SAPEdit","fneditset","fneditset" 
	RegisterUserFunc "SAPCheckBox","fneditset","fneditset" 
	RegisterUserFunc "SAPCalender","fneditset","fneditset"
	RegisterUserFunc "SAPDropDownMenu","fnselectitem","fnselectitem" 
	RegisterUserFunc "SAPRadioGroup","fnselectitem","fnselectitem" 
	RegisterUserFunc "SAPButton","fnpress","fnpress"
	RegisterUserFunc "SAPTable","fntbcelldata","fntbcelldata" 
	
'/****************************************End RegisterFunc for SAP Web Environment******************************

'****************************************RegisterFunc for Standard Windows Environment******************************
	RegisterUserFunc "WinEdit","fneditset","fneditset"
	RegisterUserFunc "WinCheckBox","fneditset","fneditset"
	RegisterUserFunc "WinRadioButton","fneditset","fneditset" 
	RegisterUserFunc "WinCalender","fneditset","fneditset" 
	RegisterUserFunc "WinComboBox","fnselectitem","fnselectitem" 
	RegisterUserFunc "WinList","fnselectitem","fnselectitem" 
	RegisterUserFunc "WinButton","fnpress","fnpress" 
	RegisterUserFunc "WinTable","fntbcelldata","fntbcelldata"
	RegisterUserFunc "WinToolBar","fnitempress","fnitempress"
'/****************************************End RegisterFunc for Standard Windows Environment******************************

'****************************************RegisterFunc for .NET Windows Forms Environment******************************
	RegisterUserFunc "SwfEdit","fneditset","fneditset"
	RegisterUserFunc "SwfCheckBox","fneditset","fneditset"
	RegisterUserFunc "SwfRadioButton","fneditset","fneditset"
	RegisterUserFunc "SwfCalender","fneditset","fneditset" 
	RegisterUserFunc "SwfComboBox","fnselectitem","fnselectitem" 
	RegisterUserFunc "SwfList","fnselectitem","fnselectitem" 
	RegisterUserFunc "SwfButton","fnpress","fnpress"
	RegisterUserFunc "SwfTable","fntbcelldata","fntbcelldata"
	RegisterUserFunc "SwfToolBar","fnitempress","fnitempress" 
'/****************************************End RegisterFunc for .NET Windows Forms Environment******************************

'****************************************RegisterFunc for Visual Basic Environment******************************
	RegisterUserFunc "VbEdit","fneditset","fneditset" 
	RegisterUserFunc "VbCheckBox","fneditset","fneditset" 
	RegisterUserFunc "VbRadioButton","fneditset","fneditset"
	RegisterUserFunc "VbComboBox","fnselectitem","fnselectitem"
	RegisterUserFunc "VbButton","fnpress","fnpress"
'/****************************************End RegisterFunc for Visual Basic Environment******************************

'****************************************RegisterFunc for .NET Web Forms Environment******************************
	RegisterUserFunc "WbfToolBar","fnitempress","fnitempress" 
'/****************************************End RegisterFunc for .NET Web Forms Environment******************************

'****************************************RegisterFunc for Web  Environment******************************
'	RegisterUserFunc "WebEdit","fneditset","fneditset"
	RegisterUserFunc "WebCheckBox","fneditset","fneditset"
	RegisterUserFunc "WebList","fnselectitem","fnselectitem" 
	RegisterUserFunc "WebRadioGroup","fnselectitem","fnselectitem"
	RegisterUserFunc "WebButton","fnpress","fnpress"
	RegisterUserFunc "Image","fnpress","fnpress"
	RegisterUserFunc "Link","fnpress","fnpress"
	RegisterUserFunc "WebElement","fnpress","fnpress"
	RegisterUserFunc "WebTable","fntbcelldata","fntbcelldata"
	RegisterUserFunc "WebElement","fnintext","fnintext"
	RegisterUserFunc "WebEdit","fnintext","fnintext"
	RegisterUserFunc "WebEdit","fngettext","fngettext"
	RegisterUserFunc "WebButton","fnpopuphandler","fnpopuphandler"
	RegisterUserFunc "WebTable","fngettext","fngettext"
	RegisterUserFunc "WebFile","fnpress","fnpress"
'/****************************************End RegisterFunc for Web  Environment******************************

'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fntbcelldata
'	Objective							:				Performing actions on Link, WebCheckBox, WebElement & RadioButton.
'	Input Parameters					:					objTable,obj,objName,RefCol,ActionCol,iPosition
'	Output Parameters					:					NIL
'	Output Parameters					:					NIL
'	Date Created						:					11-02-2014
'	QTP Version							:					UFT 11.50
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		  			NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fntbcelldata(objTable,obj,objName,RefCol,ActionCol,iPosition)
    objTable.RefreshObject
	sblnStatus = True
    sFlag = True
    sObjName=objTable.GetToProperty("TestobjName")
    If  objTable.Exist(MIN_WAIT) Then
        If objTable.GetROProperty("disabled") = 0 Then	
        	rc=objTable.GetROProperty("rows")
        	''cc=objTable.GetROProperty("cols")
            For i = 1 To rc Step 1
                RetrieveVal = objTable.GetCellData(i,RefCol)
                If Trim(objName) = Trim(RetrieveVal) Then            
					set a=objTable.Childitem(i,ActionCol,obj,iPosition)
					Select Case obj
						'********** Link ***************************
						Case "Link" 
							objval = a.GetROProperty("innertext")
							If a.Exist(MIN_WAIT) Then
								a.click	
								Call fnReportDetailedSuccess(objName,"Link is Clicked")
							Exit Function
							Else
								Call fnReportDetailedFailure(objName,"Operation was Failed")
                                sFlag = False		
							End If 
						 '********** Link ***************************
						Case "Image" 
							objval = a.GetROProperty("innertext")
							If a.Exist(MIN_WAIT) Then
								a.click	
								Call fnReportDetailedSuccess(objName,"Image is Clicked")
							Exit Function
							Else
								Call fnReportDetailedFailure(objName,"Operation was Failed")
                                sFlag = False		
							End If 
						''********** WebCheckBox ***************************
						Case "WebCheckBox"
							If a.Exist(MIN_WAIT) Then
								a.Set "ON"
								Call fnReportDetailedSuccess(objName,"WebCheckBox is ON")
							Exit Function
							Else
								Call fnReportDetailedFailure(objName,"Operation was Failed")
                                sFlag = False		
							End If
						'********** RadioButton ***************************
						Case "WebRadioGroup"
							If a.Exist(MIN_WAIT) Then
								a.Select ActionCol
								Call fnReportDetailedSuccess(objName,"RadioButton is Selected")
							Exit Function
							Else
								Call fnReportDetailedFailure(objName,"Operation was Failed")
                                sFlag = False		
							End If
						'*********** WebElement ***************************
						Case "WebElement"
							If a.Exist(MIN_WAIT) Then
								a.Click
								Call fnReportDetailedSuccess(objName,"WebElement is Clicked")
							Exit Function
							Else
								Call fnReportDetailedFailure(objName,"Operation was Failed")
                                sFlag = False		
							End If
                    End Select
                End If
            Next    				
	   Else
            Call fnReportDetailedFailure(sObjName,"The object is disabled")
            sFlag = False		
        End If
	  Else
		  Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
          sFlag = False		
	  End If
	If sblnStatus = True Then
		Call fnReportDetailedFailure(objName,"Object to be searched does not Exist")
		Call fnErrorChecking()
	End If
    If sFlag = False Then
		Call fnErrorChecking()
	End If
End Function

'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fntbintext
'	Objective							:				Retrieving the RunTime Property of the object from webtable.
'	Input Parameters					:					objTable,objName,j,sProperty,sMessage,iPosition
'	Output Parameters					:					NIL
'	Output Parameters					:					NIL
'	Date Created						:					11-02-2014
'	QTP Version							:					UFT 11.50
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		  			NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fntbintext(objTable,obj,objName,RefCol,ActionCol,sProperty,sVal,sMessage,iPosition)
	Wait (MIN_WAIT)
	Dim sFlag
	sFlag = True
	sblnStatus = True
	objTable.RefreshObject	
    sObjName=objTable.GetToProperty("TestobjName")
    If  objTable.Exist(MIN_WAIT) Then
        If  objTable.GetROProperty("disabled") =0 Then
        	rc=objTable.GetROProperty("Rows")        	
            For i = 1 To rc Step 1
            	RetVal = objTable.GetCellData(i,RefCol)
            	If Trim(objName) = Trim(RetVal) Then            
            	set a=objTable.Childitem(i,ActionCol,obj,iPosition)
            		If a.exist(MIN_WAIT) Then
            			text=a.GetROProperty(sProperty) 
            			If Trim(sVal)=Trim(text) Then
            				Call fnReportDetailedSuccess(sMessage,"Value is Retrieved & Value is "&text)
							sblnStatus = False
						End If
					End IF
'            	Else
'					 Call fnReportDetailedFailure(sMessage,"Value Retrieved not match with the value sent")
'                   	 sFlag = False			
	        	End If							  
			Next	     
	     Else
         	Call fnReportDetailedFailure(sObjName,"The object is disabled")
             sFlag = False			
         End If
    Else
        Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
         sFlag = False			
    End If
	If sblnStatus = True Then
		Call fnReportDetailedFailure(objName,"Object to be searched does not Exist")
		Call fnErrorChecking()
	End If
	If sFlag = False Then
			Call fnErrorChecking()
	End If
End Function

'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fnwebcelldata
'	Objective								:					Performing actions on WebList & WebEdit in WebTable.
'	Input Parameters					:					objTable,obj,RefName,ActionName,RefCol,ActionCol,iPosition
'	Output Parameters					:					NIL
'	Output Parameters					:					NIL
'	Date Created						:					14-02-2014
'	QTP Version							:					UFT 11.50
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		  			NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fnwebcelldata(objTable,obj,RefName,RefCol,ActionCol,ActionName,iPosition)
	Dim sFlag
	sFlag = True
	sbln = True
	objTable.RefreshObject
    sObjName=objTable.GetToProperty("TestobjName")
    If  objTable.Exist(2) Then
        If  objTable.GetROProperty("disabled") =0 Then	
        	rc=objTable.GetROProperty("rows")
        	''cc=objTable.GetROProperty("cols")
            For i = 1 To rc Step 1
				RetrievedName = Trim(objTable.GetCellData(i,RefCol))
				If Trim(RefName) = Trim(RetrievedName)  Then            
            	set a=objTable.Childitem(i,ActionCol,obj,iPosition)
            	Select Case obj
            			'********** WebList ***************************
						Case "WebList"
							If a.Exist(MIN_WAIT) Then
            					a.select ActionName
            					Call fnReportDetailedSuccess(ActionName,"Item is selected from Weblist")
            				Exit Function
           					Else
           						Call fnReportDetailedFailure(ActionName,"Operation was Failed")
                                sFlag = False		
           					End If	
           				'********** WebEdit ***************************
						Case "WebEdit"
							If a.Exist(MIN_WAIT) Then
            					a.set ActionName
            					Call fnReportDetailedSuccess(ActionName,"Value is set to the Edit Field")
            				Exit Function
           					Else
           						Call fnReportDetailedFailure(ActionName,"Operation was Failed")
                                sFlag = False		
           					End If	
           					End Select
				End If
            Next
         Else
            Call fnReportDetailedFailure(sObjName,"The object is disabled")
            sFlag = False		
        End If
    Else
        Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
        sFlag = False		
    End If
	If sblnStatus = True Then
		Call fnReportDetailedFailure(objName,"Object to be searched does not Exist")
		Call fnErrorChecking()
	End If
	If sFlag = False Then
		Call fnErrorChecking()
	End If
End Function

'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fnradiogroup
'	Objective								:					Used to select  the value in Radio Group.
'	Input Parameters					:					objEdit,strValue
'	Output Parameters					:					NIL
'	Date Created						:					NIL
'	QTP Version							:					NIL
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		            NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fnradiogroup(objRadio,strValue)
	Wait (MIN_WAIT)
    Dim sFlag
	sFlag = True
    sObjName=objRadio.GetToProperty("TestobjName")
    If  objRadio.Exist(MIN_WAIT) Then
        If  objRadio.GetROProperty("disabled") = 0 Then
            objRadio.select strValue
            Call fnReportDetailedSuccess(sObjName,strValue&" is set to "&sObjName)                    
        Else
            Call fnReportDetailedFailure(sObjName,"The object is disabled")
            sFlag = False		
        End If
    Else
        Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
        sFlag = False		
    End If
    If sFlag = False Then
		Call fnErrorChecking()
	End If
 End Function

'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fnuniqtable
'	Objective								:					Used to do operations on WebEdit, Webcheckbox and WebList
'	Input Parameters					:					objEdit,strValue
'	Output Parameters					:					NIL
'	Date Created						:					NIL
'	QTP Version							:					NIL
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		            NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fnuniqtable(objTable,obj,Row,Col,ActionName,iPosition)
	objTable.RefreshObject
    Wait (MIN_WAIT)
	Dim sFlag
	sblnStatus = True
	sFlag = True
    sObjName=objTable.GetToProperty("TestobjName")
    If  objTable.Exist(2) Then
        If  objTable.GetROProperty("disabled") =0 Then	
  					set a=objTable.Childitem(Row,Col,obj,iPosition)
					Select Case obj
							''********** WebCheckBox ***************************
							Case "WebCheckBox"
							If a.Exist(MIN_WAIT) Then
								a.Set ActionName
								Call fnReportDetailedSuccess(objName,"WebCheckBox is ON")
							Exit Function
							Else
								Call fnReportDetailedFailure(objName,"Operation was Failed")
                                sFlag = False		
							End If
							'********** WebList ***************************
							Case "WebList"
                                If a.Exist(MIN_WAIT) Then
									a.select ActionName
									Call fnReportDetailedSuccess(ActionName,"Item is selected from Weblist")
								Exit Function
								Else
									Call fnReportDetailedFailure(ActionName,"Operation was Failed")
                                    sFlag = False			
								End If
							'********** WebEdit ***************************
							Case "WebEdit"
								If a.Exist(MIN_WAIT) Then
									a.set ActionName
									Call fnReportDetailedSuccess(ActionName,"Value is set to the Edit Field")
								Exit Function
								Else
									Call fnReportDetailedFailure(ActionName,"Operation was Failed")
                                    sFlag = False			
								End If	
								End Select
         Else
            Call fnReportDetailedFailure(sObjName,"The object is disabled")
            sFlag = False			
        End If
    Else
        Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
        sFlag = False			
    End If
    If sFlag = False Then
		Call fnErrorChecking()
	End If
	If sblnStatus = True Then
		Call fnReportDetailedFailure(objName,"Object to be searched does not Exist")
		Call fnErrorChecking()
	End If
End Function

'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					   fnstatuscheck
'	Objective								:						 Used to check the status of Object.
'	Input Parameters				:						 sObj,sObjType,sProperty,sValue,sMessage	
'	Output Parameters			:							   NIL
'	Date Created					:								27-Mar-2014
'	QTP Version						:								12.0
'	QC Version						:								  QC 11 
'	Pre-requisites					:								NIL  
'	Created By						:								Gallop Solutions
'	Modification Date		:		   
'******************************************************************************************************************************************************************************************************************************************
Public Function fnStatusCheck(sObj,sObjType,sProperty,sValue,sMessage)
		On Error Resume Next
		bWaitFlag = True
		iWait = 0
        iMaxWait = 20000
		Do While (iMaxWait > iWait)
			Select Case sObjType
					Case "WebElement"
							If sObj.WaitProperty(sProperty,sValue, iMaxWait) Then
									Call fnReportDetailedSuccess(sMessage,"Specified Property is achieved and the value is "&sValue)                    
									bWaitFlag = False
									Exit Do
							Else
									iWait = iWait + 1
									Wait 1
							End If

					Case "WebEdit"
							If ObjControl.WaitProperty("Visible","True") Then
									bWaitFlag = False
									Exit Do
							Else
									iWait = iWait + 1
									Wait 1
							End If

					Case "WinButton"
							If ObjControl.WaitProperty("Visible","True") Then
									bWaitFlag = False
									Exit Do
							Else
									iWait = iWait + 1
									Wait 1
							End If
			End Select
		Loop

		Err.Clear
		If bWaitFlag = True Then
		End If
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:			fntbretrievetext
'	Objective									:					Retrieving particular column data from webtable.
'	Input Parameters					:					objTable,objName,j,sProperty,sMessage,iPosition
'	Output Parameters				 :					NIL
'	Output Parameters				:					NIL
'	Date Created						:					08-06-2014
'	QTP Version							:					UFT 11.50
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		  			NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fntbretrievetext(objTable,sCol)
	objTable.RefreshObject
    Wait (MIN_WAIT)
	Dim sFlag
	sFlag = True
    sObjName=objTable.GetToProperty("TestobjName")
    If  objTable.Exist(MIN_WAIT) Then
        If  objTable.GetROProperty("disabled") =0 Then
        	rc=objTable.GetROProperty("Rows")
			For i = 1 To rc Step 1
                     If objTable.Exist Then
						Info = objTable.GetCellData(i,sCol)
						strValue = objTable.GetCellData(i,sCol)
						Call fnReportDetailedSuccess(sObjName,"Value is Retrieved & Value is "&strValue)
					Else
						Call fnReportDetailedFailure("Value","Value is not available")
                        sFlag = False			
					End IF
                Next
          Else
         	Call fnReportDetailedFailure(sObjName,"The object is disabled")
            sFlag = False			
         End If
    Else
        Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
        sFlag = False			
    End If
    If sFlag = False Then
		Call fnErrorChecking()
	End If
End Function

'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					 
'	Objective									:					Handling the Warning Messages or Pop-Ups displayed in the Functionality.
'	Input Parameters					:					oSetObject
'	Output Parameters				 :					NIL
'	Output Parameters				:					NIL
'	Date Created						:					06-05-2014
'	QTP Version							:					QTP 11
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		  			NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fnWarningHandler(oSetObject)
    Wait (MIN_WAIT)
	sObjName=oSetObject.GetToProperty("TestobjName")
	If oSetObject.WebButton("class:=PSPUSHBUTTONTBOK","html id:=#ICOK").Exist Then
		If oSetObject.WebElement("html id:=PTPOPUP_TITLE","html tag:=SPAN").Exist Then
			sWarning = oSetObject.WebElement("class:=popupText","html tag:=SPAN").GetROProperty("innertext")
			Call fnReportDetailedFailure("Warning",sWarning)
			oSetObject.WebButton("class:=PSPUSHBUTTONTBOK","html id:=#ICOK").Click
'			ExitAction()
		End If
	End If	
End Function

'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fnSearchItemInWebList					 
'	Objective									:					To search for items in the web list for multiple times (The options not being repetetive)
'	Input Parameters					:					
'	Output Parameters				 :					NIL
'	Output Parameters				:					NIL
'	Date Created						:					09-06-2014
'	QTP Version							:					QTP 11
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		  			NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function SearchItemInWebList (ByVal sTestObject, ByVal sSearchItem, ByVal sExpectedResult)
   ObjectExist = sTestObject.Exist(MAX_WAIT)
   If ObjectExist Then
       MatchFound = False
       WebListItems = Trim(sTestObject.GetROProperty("all items"))
       SplitWebListItems = Split(WebListItems, ";")
       For iList = 0 to UBound(SplitWebListItems)
            If Trim(SplitWebListItems(iList)) = sSearchItem Then
                MatchFound = True
                Exit For
            End If
       Next
           If ExpectedResult Then            
               If MatchFound Then
					Call fnReportDetailedSuccess(sSearchItem,"Value is Retrieved & Value is "&sSearchItem)
                    SearchItemInWebList = TRUE
               ElseIf Not MatchFound Then
					Call fnReportDetailedFailure(sSearchItem,"Cannot find the specified item " &sSearchItem & " in the web list")
					ExitRun()
                    SearchItemInWebList = FALSE
               End If
           ElseIf Not ExpectedResult Then
               If MatchFound Then
					Call fnReportDetailedSuccess(sSearchItem,"Specified item " &sSearchItem & " is present in the web list contrary to what was expected")
                    SearchItemInWebList = FALSE
           ElseIf Not MatchFound Then
					 Call fnReportDetailedFailure(sSearchItem,"Cannot find the specified item " &sSearchItem & " in the web list as expected")
					 ExitRun()
                     SearchItemInWebList = TRUE
               End If                         
           End If
    ElseIf Not ObjectExist Then
			  Call fnReportDetailedFailure(sSearchItem,"Object specified in function call cannot be found")			  
   End If
End Function

''************ Function Error Check ************************************************************
Public Function fnErrorChecking()
   Err.Clear
   on Error resume next
   ErrorName=Err.Description
	If Err.Number <> 0 Then
'		Call fnReportDetailedFailure("Operation Failed","Due to"&ErrorName)		
	End If
	On Error goto 0 
	ExitRun()
End Function


Public Function fnWbCellData(objTable,obj,objName,sCheckbox,RefCol,ActionCol,iPosition)
    objTable.RefreshObject
	sblnStatus = True
    sFlag = True
    sObjName=objTable.GetToProperty("TestobjName")
    If  objTable.Exist(MIN_WAIT) Then
        If objTable.GetROProperty("disabled") =0 Then	
        	rc=objTable.GetROProperty("rows")
        	''cc=objTable.GetROProperty("cols")
            For i = 1 To rc Step 1
                RetrieveVal = objTable.GetCellData(i,RefCol)
                If Trim(objName) = Trim(RetrieveVal) Then            
					set a=objTable.Childitem(i,ActionCol,obj,iPosition)
					Select Case obj						
						''********** WebCheckBox ***************************
						Case "WebCheckBox"
							If a.Exist(MIN_WAIT) Then
								a.Set sCheckbox
								Call fnReportDetailedSuccess(objName,"WebCheckBox is ON")
							Exit Function
							Else
								Call fnReportDetailedFailure(objName,"Operation was Failed")
                                sFlag = False		
							End If						
                    End Select
                End If
            Next    				
	   Else
            Call fnReportDetailedFailure(sObjName,"The object is disabled")
            sFlag = False		
        End If
	  Else
		  Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
          sFlag = False		
	  End If
	If sblnStatus = True Then
		Call fnReportDetailedFailure(objName,"Object to be searched does not Exist")
		Call fnErrorChecking()
	End If
    If sFlag = False Then
		Call fnErrorChecking()
	End If
End Function


'***************************** Retrieving the Text From WebTable In Particular Column Wise ******************************
Public Function fntbgettext(objTable,obj,sProperty,sRow,sCol,iPosition)
	objTable.RefreshObject
    Wait (MIN_WAIT)
	Dim sFlag
	sblnStatus = True
	sFlag = True
    sObjName=objTable.GetToProperty("TestobjName")
    If  objTable.Exist(2) Then
        If  objTable.GetROProperty("disabled") =0 Then	
  					set a=objTable.Childitem(sRow,sCol,obj,iPosition)
                    	If a.Exist(MIN_WAIT) Then
							text = a.GetROProperty(sProperty)
							Call fnReportDetailedSuccess(sObjName,"Retrieved Value is"&text)
						Exit Function
						Else
							Call fnReportDetailedFailure(sObjName,"Operation Failed")
                               sFlag = False		
						End If  	
          Else
         	Call fnReportDetailedFailure(sObjName,"The object is disabled")
            sFlag = False			
         End If
    Else
        Call fnReportDetailedFailure(sObjName,"The object doesnt Exist")
        sFlag = False			
    End If
    If sFlag = False Then
		Call fnErrorChecking()
	End If
End Function


''******************************************************************************************************************************************************************************************************************************************
'	Function Name						:		  					fnCheckBox
'	Objective							:				   	     	Used to Check and uncheck checkboxes in any environment 
'	Input Parameters					:				  			objCheckBoxName  (Name of the checkbox object during object spy)
'	Output Parameters					:							NIL
'	Date Created						:							NIL
'	QTP Version							:							NIL
'	QC Version							:							NIL
'	Pre-requisites						:							NIL
'	Created By							:							NIL
'	Modification Date					:		      				NIL
'******************************************************************************************************************************************************************************************************************************************

Public Function fnCheckBox(objCheckBoxName)
	Set objDesc = Description.Create()	
		objDesc("micclass").Value = "WebCheckBox"
		objDesc("html tag").Value= "INPUT"
		objDesc("type").value = "checkbox"
	Set ChkBoxCount = Browser("FSCM").Page("FSCM").ChildObjects(objDesc)	
	For i = 0 to ChkBoxCount.Count -1
		appVal = Trim(ChkBoxCount(i).GetROProperty("name"))
		if instr(appVal,objCheckBoxName) then
			ChkBoxCount(i).set "ON"
		End If
	Next
End Function


''******************************************************************************************************************************************************************************************************************************************
'	Function Name						:		  			fnAutoAdd
'	Objective							:				   	      Used to Add WebEdits and Images more than one automatically
'  Code type                                :                    Descrptive
'	Input Parameters					:				  objImageName (Name of the Image object) , objSeqhtmlid (HtmlId of the Edit Box) , objPagehtmlid(HtmlId of the Edit Box) during object spy
'	Output Parameters					:				NIL
'	Date Created						:					NIL
'	QTP Version							:					NIL
'	QC Version							:					NIL
'	Pre-requisites						:					NIL
'	Created By							:					NIL
'	Modification Date					:		      NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fnAutoAdd(objImageName,objSeqhtmlid,objPagehtmlid,strPageValue)
		flag = 0
		For k=0 to 2
				   On error resume next
					Set oFSCMPageObj = Browser("name:=Page Series Definition").Page("title:=Page Series Definition")
					Set dec1=description.Create
					dec1("micclass").value="Image"
					Set Img=oFSCMPageObj.ChildObjects(dec1)
					ImageCount = Img.count
					For j=0 to ImageCount
						 innervalue1=Img(j).getRoProperty("name")
						If instr(innervalue1,objImageName&flag&"$$img$0") > 0 Then
								Img(j).click
								Call fnReportDetailedSuccess(innervalue1,"Click operation performed in Add field button")                    
        					    flag=flag+1
					   End If
					Next
					On error goto 0
		Next
		wait 3
		On error resume next
		flag1=0
		flag2=0
		Set dec=description.Create
			dec("micclass").value="WebEdit"
			set Edit = oFSCMPageObj.childobjects(dec)
			WebEditCount = Edit.count
			For i = 0 To WebEditCount
			Set dec=description.Create
			dec("micclass").value="WebEdit"
			set Edit = oFSCMPageObj.childobjects(dec)
			innervalue=Edit(i).getroproperty("html id")
			If instr(innervalue,objSeqhtmlid&flag1) > 0 Then
					flag1=flag1+1
					Edit(i).set flag1
					 Call fnReportDetailedSuccess(innervalue,flag1&" is set to PageSequence"&flag1)                    
				
			End If
				If instr(innervalue,objPagehtmlid&flag2) > 0 Then
					Edit(i).Set strPageValue
					 Call fnReportDetailedSuccess(innervalue,strPageValue&" is set to "&innervalue)  
					flag2=flag2+1
				End If
		  Next
		On error goto 0
End Function
'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:		  						fnCompWarning
'	Objective							:				   	      		Function to handle warnings
'	Input Parameters					:				 		   	    NIL
'	Output Parameters					:								NIL
'	Date Created						:								26-dec-2014
'	QTP Version							:								12.0
'	QC Version							:								NIL
'	Pre-requisites						:								NIL
'	Created By							:								NIL
'	Modification Date					:		      					NIL
'******************************************************************************************************************************************************************************************************************************************
Public Function fnCompWarning()
   set objbrow=Browser("name:=Job Data").Page("title:=Job Data").Frame("name:=TargetContent").WebButton("html id:=#ICOK","visible:=True")
		Do 
			If objbrow.exist(5) Then
				Browser("name:=Job Data").Page("title:=Job Data").Frame("name:=TargetContent").WebButton("html id:=#ICOK").Click
				Wait 5
			Else
				sFlag = "False"
			End If 
	Loop Until sFlag = "False"
End Function
''******************************************************************************************************************************************************************************************************************************************
''	Function Name									:						fnWaitForObject
''	Objective										:			 			This function is used for Synchronisation. It will wait till the object is visible.
''	Input Parameters								:						ObjectName
''	Output Parameters								:						Nil
''	Date Created									:						06-June-2014
''	QTP Version										:						12.0
''	QC Version										:			  			QC 11.5
''	Pre-requisites									:						NIL  
''	Created By										:						Gallop Solutions
''	Modification Date								:		   
''******************************************************************************************************************************************************************************************************************************************
Function fnWaitForObject(sObjectName)
	fnWaitForObject=False
	Set RefObject=sObjectName
	counter=0
	'' Waits for (5x 20 = 100 sec) i.e. 1.66 mins
 	Do
		If RefObject.Exist(MID_WAIT) Then
				fnWaitForObject=True
				Exit Do
		End If
		counter=counter+1
	Loop until counter >=20
End Function
''******************************************************************************************************************************************************************************************************************************************
''Function Name										:					fnCloseForm
''Objective											:		 			Used to CloseForm
''Input Parameters									:					ObjectName, PropertyName
''Output Parameters									:					Nil
''Date Created										:					05-June-2014
''QTP Version										:					12.0
''QC Version										:		  			QC 11.5
''Pre-requisites									:					NIL  
''Created By										:					Gallop Solutions
''Modification Date									:		   
''******************************************************************************************************************************************************************************************************************************************
Function fnCloseFor(sObjectName)
	fnCloseForm=False
		Set RefObject = sObjectName
	If  RefObject.Exist Then
		RefObject.CloseForm
		fnCloseForm = True
	End If
	Set RefObject=Nothing
End Function

''******************************************************************************************************************************************************************************************************************************************
''	Function Name									:						fnObjExistance
''	Objective										:						Used to Verify the Object Exist in 1 sec
''	Input Parameters								:						Object Name,InputValue
''	Output Parameters								:						Nil
''	Date Created									:						27-May-2014
''	QTP Version										:						12.0
''	QC Version										:						QC 11.5
''	Pre-requisites									:						NIL  
''	Created By										:						Gallop Solutions
''	Modification Date								:		   
''******************************************************************************************************************************************************************************************************************************************
Function fnObjExistance(sObjectName)
	fnObjExistance=False
	If Not IsObject(sObjectName) Then
		Set RefObject=sObjectName
		tempObjectName=sObjectName
	Else
		Set RefObject = sObjectName	
		tempObjectName=RefObject.ToString	 
	End If
	If  RefObject.Exist(1)Then
		fnObjExistance = True
	End If
End Function
''******************************************************************************************************************************************************************************************************************************************
''Function Name										:					fnSelectMenu
''Objective											:		 			Used to Select submenu from the Menu
''Input Parameters									:					ObjectName,strItem
''Output Parameters									:					Nil
''Date Created										:					05-June-2015
''QTP Version										:					12.0
''QC Version										:		  			QC 11.5
''Pre-requisites									:					NIL  
''Created By										:					Gallop Solutions
''Modification Date									:		   
''******************************************************************************************************************************************************************************************************************************************
Function fnSelectMenu (sObjectName,strItem)
	fnSelectMenu=False
	If Not IsObject(sObjectName) Then
		Set RefObject=Eval(fnGetObjectHierarchy(sObjectName))
	Else
		Set RefObject = sObjectName
	End If
	If RefObject.Exist(MID_WAIT) Then
			RefObject.SelectMenu strItem
			Call rptWriteReport("Pass", "Click On " &sObjectName,"Successfully clicked on '"&strItem&"' of '"& sObjectName &"'")
			'Call fnReportDetailedSuccess("Click On " &tempObjectName,"Successfully clicked on '"& tempObjectName &"'")
			fnSelectMenu = True
	Else
			Call rptWriteReport("Fail", "Click On " &sObjectName,"Click operation not performed on '"& sObjectName &"' as the object does not exit")
			
			Exit Function
	End If
End Function     
''******************************************************************************************************************************************************************************************************************************************
''Function Name									:					fnOpenDialog
''Objective										:		 			Used to Uncheck the Oracle Checkbox
''Input Parameters								:					ObjectName, PropertyName
''Output Parameters								:					Nil
''Date Created									:					05-June-2014
''QTP Version									:					12.0
''QC Version									:				  	QC 11.5
''Pre-requisites								:					NIL  
''Created By									:					Gallop Solutions
''Modification Date								:		   
''******************************************************************************************************************************************************************************************************************************************
Function fnOpenDialog(sObjectName)
	fnOpenDialog=False
	If Not IsObject(sObjectName) Then
		Set RefObject=sObjectName
		tempObjectName=sObjectName 
	Else
		Set RefObject = sObjectName
		tempObjectName=RefObject.ToString
	End If
	If RefObject.Exist(MID_WAIT) Then
			RefObject.OpenDialog
			Call fnReportDetailedSuccess("Click On " &tempObjectName,"Successfully clicked on '"& tempObjectName &"'")
			fnOpenDialog = True
	Else
			Call fnReportDetailedFailure("Click On " &tempObjectName,"Click operation not performed on '"& tempObjectName &"' as the object does not exit")
			Exit Function
	End If
End Function  
''******************************************************************************************************************************************************************************************************************************************
''Function Name										:					fnCloseWindow
''Objective											:		 			Used to CloseWindow
''Input Parameters									:					ObjectName, PropertyName
''Output Parameters									:					Nil
''Date Created										:					
''UFT Version										:					12.0
''QC Version										:		  			
''Pre-requisites									:					NIL  
''Created By										:					
''Modification Date									:		   
''******************************************************************************************************************************************************************************************************************************************
Function fnCloseWindow(sObjectName)
	fnCloseWindow=False
	If Not IsObject(sObjectName) Then
	   Set RefObject=Eval(fnGetObjectHierarchy(sObjectName))
	Else 
		Set RefObject = sObjectName
	End If
    If RefObject.Exist(MIN_WAIT) Then
    	RefObject.CloseWindow
		fnCloseWindow = True
	 Set RefObject=Nothing
    End If	
End Function