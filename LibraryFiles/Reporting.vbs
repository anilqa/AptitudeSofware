'Option Explicit

'@Function Name <rpt_WriteReport>
'@CreationDate <02-mar-2014>
'@Dependency  
'@Author 
'@ModifiedDate
'@ModifiedBy 
'@Description	this method write the test case results in html file   
'@Documentation <param> and <param> will do….  
''gTestDir = "E:\Framework Dev\ScriptLess Framework Demo - Flight\Script Less Demo\"
Dim iSNO, iTestCaseNumber

Public Function rptWriteReport(ByVal strResult, ByVal strStepName , ByVal strExpected)
'***********************************************Initial Setup******************************************
  Dim objFso,objFolder,strResultsFolder,strTestcasesPath,objFile,status, html, link,arrPath,strResources  
  
  arrResource = Split(Environment("TestDir"),"TestCases") 
  strResources = arrResource(0) & "Resources"
  
  Set objFso = CreateObject("Scripting.FileSystemObject")
 
	If not objFso.FolderExists(strProjectResultPath) Then
		Set objFolder=objFso.CreateFolder(strProjectResultPath)
	End If
	
	If gFolderName=Empty Then
		rptFoldername
	End If
	 
   strResultsFolder=strProjectResultPath&"\"& gFolderName 
'   strResultsFolder1 = strResultsFolder
   
   If not objFso.FolderExists(strResultsFolder) Then
	   Set objFolder=objFso.CreateFolder(strResultsFolder)
   End If
   
   strLogos=strResultsFolder&"\Logos"
   If not objFso.FolderExists(strLogos) Then
	   Set objFolder=objFso.CreateFolder(strLogos)
	   objFso.CopyFile strResources&"\Pass.png",strLogos&"\"
	   objFso.CopyFile strResources&"\Fail.png",strLogos&"\"
	   objFso.CopyFile strResources&"\Companylogo.png",strLogos&"\"
	   objFso.CopyFile strResources&"\Clientlogo.png",strLogos&"\"
	   objFso.CopyFile strResources&"\PassWithScr.png",strLogos&"\"
   End If
   
   strTestcasesPath=strResultsFolder&"\Testcases"
   If not objFso.FolderExists(strTestcasesPath) Then
	   Set objFolder=objFso.CreateFolder(strTestcasesPath)
   End If
	If not objFso.FileExists(strTestcasesPath&"\"&Environment.Value("TestName")&".html") Then
				iSNO = 1
				Set  objFile = objFso.CreateTextFile(strTestcasesPath&"\"&Environment.Value("TestName")&".html",true, false)  
		
				objFile.writeline  "<html>" & VBNewLine
				objFile.writeline  "<head> " & VBNewLine
				objFile.writeline  "<style type=""text/css"">.passed{display: table-row; background-color: #E1E1E1; border: 1px solid #4D7C7B; color: #000000; font-size: 0.75em; td, th { padding: 5px; border: 1px solid #4D7C7B; text-align: inherit /; } th.Logos { padding: 5px; border: 0px solid #4D7C7B; text-align: inherit /;} td.justified { text-align: Left; } td.pass { font-weight: bold; color: green; } </style>" & VBNewLine
				objFile.writeline  "<style type=""text/css"">.failed{display: table-row;background-color: #FFFFFF; color: #000000; font-size: 0.7em; display: table-row;} </style>" & VBNewLine&"<style type=""text/css"">.notvisible{display: None; </style>"& VBNewLine &"<meta charset='UTF-8'> "
				objFile.writeline  "<title>Detailed Results Report</title>"& VBNewLine
				objFile.writeline  "<style type='text/css'>"& VBNewLine
				objFile.writeline  "body { background-color: #FFFFFF; font-family: Verdana, Geneva, sans-serif; text-align: center; } small { font-size: 0.7em; } table { box-shadow: 9px 9px 10px 4px #BDBDBD;border: 0px solid #4D7C7B; border-collapse: collapse; border-spacing: 0px; width: 1000px; margin-left: auto; margin-right: auto; } tr.heading { background-color: #041944;color: #FFFFFF; font-size: 0.7em; font-weight: bold; background:-o-linear-gradient(bottom, #999999 5%, #000000 100%);background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #999999), color-stop(1, #000000));background:-moz-linear-gradient( center top, #999999 5%, #000000 100%);filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#999999, endColorstr=#000000); background: -o-linear-gradient(top,#999999,000000);} tr.subheading { background-color: #FFFFFF; color: #000000; font-weight: bold; font-size: 0.7em; text-align: justify; } tr.section { background-color: #A4A4A4; color: #333300; cursor: pointer; font-weight: bold; font-size: 0.7em; text-align: justify; background:-o-linear-gradient(bottom, #56aaff 5%, #e5e5e5 100%); background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #56aaff), color-stop(1, #e5e5e5));background:-moz-linear-gradient( center top, #56aaff 5%, #e5e5e5 100%);filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#56aaff, endColorstr=#e5e5e5); background: -o-linear-gradient(top,#56aaff,e5e5e5);} tr.subsection { cursor: pointer; } td, th { padding: 5px; border: 1px solid #4D7C7B; text-align: inherit /; } th.Logos { padding: 5px; border: 0px solid #4D7C7B; text-align: inherit /;} " & VBNewLine
				objFile.writeline  "</style>"& VBNewLine
				objFile.writeline  "</head>" & VBNewLine
				objFile.writeline  "<body>" & VBNewLine & "</br>"
				objFile.writeline  "<table id='Logos'> " & VBNewLine 
				objFile.writeline  "<colgroup>" & VBNewLine 
				objFile.writeline  "<col style='width: 25%' />" & VBNewLine 
				objFile.writeline  "<col style='width: 25%' />" & VBNewLine 
				objFile.writeline  "<col style='width: 25%' />" & VBNewLine 
				objFile.writeline  "<col style='width: 25%' />" & VBNewLine 
				objFile.writeline  "</colgroup>" & VBNewLine
				objFile.writeline  "<thead>"& VBNewLine
				objFile.writeline  "<tr class='content'>" & VBNewLine
				objFile.writeline  "<th class ='Logos' colspan='2' ><img align ='left' src= '..\Logos\Clientlogo.png'  height=60 width=140></img></th>" & VBNewLine
				objFile.writeline  "<th class = 'Logos' colspan='2' > <img align ='right' src= '..\Logos\Companylogo.png' height=60 width=140></img></th> </tr> " & VBNewLine
				objFile.writeline  "</thead>" & VBNewLine
				objFile.writeline  "</table><table id='header'> " & VBNewLine
				objFile.writeline  "<colgroup> <col style='width: 25%' /> " & VBNewLine
				objFile.writeline  "<col style='width: 25%' /> " & VBNewLine
				objFile.writeline  "<col style='width: 25%' /> " & VBNewLine
				objFile.writeline  "<col style='width: 25%' /> " & VBNewLine
				objFile.writeline  "</colgroup>" & VBNewLine
				objFile.writeline  "<thead>" & VBNewLine
				objFile.writeline  "<tr class='heading'> " & VBNewLine
				objFile.writeline  "<th colspan='4' style='font-family:Copperplate Gothic Bold; font-size:1.4em;'> **"& Environment.Value("TestName") & " **</th>" & VBNewLine
				objFile.writeline  "</tr> " & VBNewLine
				iCurrentTime = Now()
				objFile.writeline  "<tr class='subheading'>" & VBNewLine
				objFile.writeline  "<th>&nbsp;Date&nbsp;&&nbsp;Time&nbsp;:&nbsp;</th> " & VBNewLine
				objFile.writeline  "<th>"& DatePart("d", iCurrentTime) & "-" & MonthName(Month(iCurrentTime), True) & "-" & DatePart("yyyy", iCurrentTime) & Space(1) & Hour(iCurrentTime) & ":" & Minute(iCurrentTime) & "</th>" & VBNewLine
				'objFile.writeline  "<th>&nbsp;Operating&nbsp;System&nbsp;:&nbsp;</th>" & VBNewLine
				'objFile.writeline  "<th> "& Environment.Value("OS") & "</th> " & VBNewLine
				'objFile.writeline  "</tr> " & VBNewLine
				'objFile.writeline  "<tr class='subheading'>" & VBNewLine
				objFile.writeline  "<th>&nbsp;Oracle&nbsp;Version&nbsp;:&nbsp;</th>" & VBNewLine
				objFile.writeline  "<th> "& Environment.Value("ORACLEVERSION") & "</th> " & VBNewLine
				'objFile.writeline  " <th>&nbsp;Executed&nbsp;on&nbsp;:&nbsp;</th>" & VBNewLine
				'objFile.writeline  "<th>" & Environment.Value("LocalHostName") & "</th> " & VBNewLine
				objFile.writeline  "</tr> " & VBNewLine
				objFile.writeline  "</thead>" & VBNewLine
				objFile.writeline  "</table>" & VBNewLine
				objFile.writeline  "<table id='main' cellpadding=""0"" cellspacing=""0""> " & VBNewLine
				objFile.writeline  "<Head>" & VBNewLine
				objFile.writeline  "<Body>" & VBNewLine
				objFile.writeline  "<colgroup>" & VBNewLine
				objFile.writeline  "<col style='width: 5%' /> <col style='width: 26%' /> <col style='width: 51%' /> " & VBNewLine
				objFile.writeline  "<col style='width: 8%' /> <col style='width: 10%' />" & VBNewLine

				objFile.writeline  "</colgroup>"
				objFile.writeline  "<thead>"
				objFile.writeline  "<tr class='heading'>"
				objFile.writeline  "<th>S.No</th> "
				objFile.writeline  "<th>Step"
				objFile.writeline  "<INPUT id=""txtStepValue"" value = ""Search"" onclick=""BlankStatus()"" onchange=""filterStatus()"">"
				objFile.writeline  "</th> "
				objFile.writeline  "<th>Details"
				objFile.writeline  "<INPUT id =""txtDetailsValue"" value = ""Search"" onclick = ""BlankDetails()"" onchange=""filterDetails()"">"
				objFile.writeline  "</th> "
				objFile.writeline  "<th> Status"
				objFile.writeline  "<select id=""filter"" onchange=""filter()"">"
				objFile.writeline  "<option value=""all"">All</option> "
				objFile.writeline  "<option value=""passed"">Passed</option>"
				objFile.writeline  "<option value=""failed"">Failed</option>"
				objFile.writeline  "</select>"
				objFile.writeline  "</th>"
				objFile.writeline  "<th>Time</th>"
				objFile.writeline  "</tr> "


				objFile.WriteBlankLines(5)
				objFile.writeline  "<script type=""text/javascript"">" & VBNewLine
				objFile.writeline  "function filter()" & VBNewLine
				objFile.writeline  "{" & VBNewLine
					objFile.writeline  "if(document.getElementById(""filter"").value==""passed"")" & VBNewLine
					objFile.writeline  "{" & VBNewLine
						objFile.writeline  "document.getElementsByTagName(""style"")[0].textContent = "".passed{display: table-row;background-color: #E1E1E1; border: 1px solid #4D7C7B; color: #000000; font-size: 0.75em;}"";" & VBNewLine
						objFile.writeline  "document.getElementsByTagName(""style"")[1].textContent = "".failed{display: none;}"";" & VBNewLine
					objFile.writeline  "}" & VBNewLine
					objFile.writeline  "else if (document.getElementById(""filter"").value==""failed"")" & VBNewLine
					objFile.writeline  "{" & VBNewLine
						objFile.writeline  "document.getElementsByTagName(""style"")[1].textContent = "".failed{display: table-row;background-color: #FFFFFF;color: #000000; font-size: 0.7em; display: table-row;}"";" & VBNewLine
						objFile.writeline  "document.getElementsByTagName(""style"")[0].textContent = "".passed{display: none;}"";" & VBNewLine
					objFile.writeline  "}" & VBNewLine
					objFile.writeline  "else" & VBNewLine
					objFile.writeline  "{" & VBNewLine
						objFile.writeline  "document.getElementsByTagName(""style"")[0].textContent = "".passed{display: table-row;background-color: #E1E1E1; border: 1px solid #4D7C7B; color: #000000; font-size: 0.75em;}"";" & VBNewLine
						objFile.writeline  "document.getElementsByTagName(""style"")[1].textContent = "".failed{display: table-row;background-color: #FFFFFF;color: #000000; font-size: 0.7em; display: table-row;}"";" & VBNewLine
					objFile.writeline  "}" & VBNewLine
				objFile.writeline  "}" & VBNewLine
				objFile.writeline  "</script>" & VBNewLine


				objFile.writeline  "<script type=""text/javascript"">"
				objFile.writeline  "function filterStatus()"
				objFile.writeline  "{"
				objFile.writeline  "searchtext = (document.getElementById(""txtStepValue"").value).toLowerCase();"
				objFile.writeline  "if(searchtext!="""")"
				objFile.writeline  "{"
				objFile.writeline  "var rowIndex = 0; // rowindex, in this case the first row of your table"
				objFile.writeline  "var table = document.getElementById('main'); // table to perform search on"
				objFile.writeline  "var row = table.getElementsByTagName(""tr"");"
				objFile.writeline  "irowcount = row.length"
				objFile.writeline  "for (i = 1; i < row.length; i++) {"
				objFile.writeline  "status = (row[i].getElementsByTagName(""td"")[1].textContent).toLowerCase();"
				objFile.writeline  "if (status.indexOf(searchtext) == -1) "
				objFile.writeline  "{"
				objFile.writeline  "row[i].className = 'content notvisible'"
				objFile.writeline  "}}}"
				objFile.writeline  "else {"
				objFile.writeline  "window.location.reload()"
				objFile.writeline  "}}"
				objFile.writeline  "</script>"


				objFile.writeline  "<script type=""text/javascript"">"
				objFile.writeline  "function filterDetails()"
				objFile.writeline  "{"
				objFile.writeline  "searchtext = (document.getElementById(""txtDetailsValue"").value).toLowerCase();"
				objFile.writeline  "if(searchtext!="""")"
				objFile.writeline  "{"
				objFile.writeline  "var rowIndex = 0; // rowindex, in this case the first row of your table"
				objFile.writeline  "var table = document.getElementById('main'); // table to perform search on"
				objFile.writeline  "var row = table.getElementsByTagName(""tr"");"
				objFile.writeline  "for (i = 1; i < row.length; i++) {"
				objFile.writeline  "Details = (row[i].getElementsByTagName(""td"")[2].textContent).toLowerCase();;"
				objFile.writeline  "if (Details.indexOf(searchtext) == -1) "
				objFile.writeline  "{"
				objFile.writeline  "row[i].className = 'content notvisible'"
				objFile.writeline  "}}}"
				objFile.writeline  "else {"
				objFile.writeline  "window.location.reload()"
				objFile.writeline  "}}"
				objFile.writeline  "</script>"
				
				objFile.writeline "<script type=""text/javascript"">"
				objFile.writeline "function BlankStatus()"
				objFile.writeline "{"
				objFile.writeline "document.getElementById(""txtStepValue"").value = """";"
				objFile.writeline "}"
				objFile.writeline "</script>"
				
				objFile.writeline "<script type=""text/javascript"">"
				objFile.writeline "function BlankDetails()"
				objFile.writeline "{"
				objFile.writeline "document.getElementById(""txtDetailsValue"").value = """";"
				objFile.writeline "}"
				objFile.writeline "</script>"


	Else
		Set objFile=objFso.OpenTextFile(strTestcasesPath&"\"&Environment.Value("TestName")&".html", 8,TRUE)    
   End If
		
   ''on Error Resume Next
   Err.Clear
	
    'status = UCase(Left(strResult,1))
    Select Case ucase(strResult)
		Case "PASS" 
				Reporter.ReportEvent micPass , strStepName , strActual
				objFile.WriteLine "<tr class='content passed' ><td>" & iSNO & "</td> "
				objFile.WriteLine "<td class='justified'>" & strStepName &"</td>"
				objFile.WriteLine "<td class='justified'>" & strExpected & "</td>"
				objFile.WriteLine "<td class='Pass' align='center'><img  src='" & "..\Logos\Pass.png' width='25' height='25'/></td> "
				iCurrentTime = Now()
				objFile.WriteLine "<td><small>" & DatePart("d", iCurrentTime) & "-" & MonthName(Month(iCurrentTime), True) & "-" & DatePart("yyyy", iCurrentTime) & Space(1) & Hour(iCurrentTime) & ":" & Minute(iCurrentTime) & ":" & Second(iCurrentTime)& "</small></td> </tr>"
				rptReportLog strStepName,strExpected,"Pass"						
		Case "FAIL"	
				Reporter.ReportEvent micFail , strStepName , strActual
				objFile.WriteLine "<tr class='content failed' ><td>" & iSNO & "</td> "
				objFile.WriteLine "<td class='justified'>" & strStepName &"</td>"
				objFile.WriteLine "<td class='justified'>" & strExpected & "</td> "
				link = rptScreenCapture()
				objFile.WriteLine "<td class='Fail' align='center'><a href="& link &">"
				iCurrentTime = Now
				objFile.WriteLine "<img  src='" & "..\Logos\Fail.png' width='25' height='25'/></td> <td><small>" & DatePart("d", iCurrentTime) & "-" & MonthName(Month(iCurrentTime), True) & "-" & DatePart("yyyy", iCurrentTime) & Space(1) & Hour(iCurrentTime) & ":" & Minute(iCurrentTime) & ":" & Second(iCurrentTime)& "</small></td> </tr>"
				rptReportLog strStepName,strExpected,"Fail"
			Case "PASSWITHSCREENSHOT"
				Reporter.ReportEvent micPass , strStepName , strActual
				objFile.WriteLine "<tr class='content passed' ><td>" & iSNO & "</td> "
				objFile.WriteLine "<td class='justified'>" & strStepName &"</td>"
				objFile.WriteLine "<td class='justified'>" & strExpected & "</td> "
				link = rptScreenCapture()
				objFile.WriteLine "<td class='Pass' align='center'><a href="& link &">"
				iCurrentTime = Now
				objFile.WriteLine "<img  src='" & "..\Logos\PassWithScr.png' width='25' height='25'/></td> <td><small>" & DatePart("d", iCurrentTime) & "-" & MonthName(Month(iCurrentTime), True) & "-" & DatePart("yyyy", iCurrentTime) & Space(1) & Hour(iCurrentTime) & ":" & Minute(iCurrentTime) & ":" & Second(iCurrentTime)& "</small></td> </tr>"
				rptReportLog strStepName,strExpected,"Pass"
				
		Case "DONE" 
				Reporter.ReportEvent micPass , strStepName , strActual
				objFile.WriteLine "<tr class='content passed' ><td>" & iSNO & "</td> "
				objFile.WriteLine "<td class='justified'>" & strStepName &"</td>"
				objFile.WriteLine "<td class='justified'>" & strExpected & "</td>"
				objFile.WriteLine "<td class='Pass' align='center'> Done </td> "
				iCurrentTime = Now()
				objFile.WriteLine "<td><small>" & DatePart("d", iCurrentTime) & "-" & MonthName(Month(iCurrentTime), True) & "-" & DatePart("yyyy", iCurrentTime) & Space(1) & Hour(iCurrentTime) & ":" & Minute(iCurrentTime) & ":" & Second(iCurrentTime)& "</small></td> </tr>"
				rptReportLog strStepName,strExpected,"Pass"	


	End Select
  iSNO = iSNO+1
End Function

'*****************************************************************
'@CreationDate <02-mar-2014>
'@Dependency  
'@Author 
'@ModifiedDate
'@ModifiedBy 
'@Description	it capture screen and send path of image fie to called function     
'@Documentation <param> and <param> will do….
'******************************************************************
Function rptScreenCapture()
	 Dim objFso,strResultsPath,strScreenshotPath,objFolder,strImagePath,strFilePath,strImagelinkPath,objDesktop
	 Set objFso = CreateObject("Scripting.FileSystemObject")
	 'strResultsPath=Environment.Value("RESULTSPATH")&"\"&gFolderName
	 'strScreenshotPath=strResultsFolder1 &"\screenshot"
	 strScreenshotPath=strProjectResultPath&"\"& gFolderName &"\Screenshot"
	  If not  objFso.FolderExists(strScreenshotPath) Then
		   Set objFolder=objFso.CreateFolder(strScreenshotPath)
	  End If
	
	 strImagePath="\Screenshot"&Replace(Replace(Replace(now(),":","_"),"/","_")," ","_") &".png"
	 strFilePath=strScreenshotPath&strImagePath
	 strImagelinkPath="..\Screenshot"&strImagePath
	 Set objDesktop = Desktop
	' Capture the Desktop
	 objDesktop.capturebitmap strFilePath ,  true
	'Add the Captured Screen shot to the Results file
	
	 rptScreenCapture=strImagelinkPath
End Function 


'*********************************************************
'@CreationDate <02-mar-2014>
'@Dependency  
'@Author 
'@ModifiedDate
'@ModifiedBy 
'@Description	it creates summary report of executed  test cases.      
'@Documentation <param> and <param> will do….
  
'***********************************************************
Function rptWriteResultsSummary()

		Dim strResultsPath, objSummary,objFilesummary,strResultsFolder,objFSO,objFolder,objFiles,iCount,iFailCount,iPassCount,objFile,FailedScriptPercentage,PassedSrciptPercentage,SummaryChart,html,SummaryChart1,wshShell
		arrResource = Split(Environment("TestDir"),"ScenarioScripts") 
		strResources = arrResource(0) & "Resources"
		'strResources=gTestDir &"Resources"
		
		SummaryChart =strProjectResultPath&"\"&gFolderName &"\SummaryChart.html"
		
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFolder = objFSO.GetFolder(strProjectResultPath & "\" & gFolderName & "\Testcases")
		Set objFiles = objFolder.Files
		

		
		Set objFile = objFSO.CreateTextFile(SummaryChart, true, false)
		
		objFile.writeline "<table id='Logos'> <colgroup> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> </colgroup> "
		objFile.writeline "<thead>  <tr class='content'> <th class ='Logos' colspan='2' > <img align ='left' src='.\Logos\Clientlogo.png ' height=60 width=140></img> </th>"
		objFile.writeline "<th class = 'Logos' colspan='2' > <img align ='right' src='.\Logos\Companylogo.png ' height=60 width=140></img> </th> </tr> </thead> </table> "		
		objFile.writeline "<html> <head> <script src='http://ajax.googleapis.com/ajax/libs/jquery/1.7.1/jquery.min.js' type='text/javascript'></script>"
		objFile.writeline "<script src='/js/highcharts.js' type='text/javascript'></script><script src='http://code.highcharts.com/highcharts.js'></script>"
		objFile.writeline "<script src='http://code.highcharts.com/highcharts-3d.js'></script><script src='http://code.highcharts.com/modules/exporting.js'></script>"
		objFile.writeline "<meta charset='UTF-8'> <title> Execution Summary Report</title><style type='text/css'>body {background-color: #FFFFFF; "
		objFile.writeline "font-family: Verdana, Geneva, sans-serif; text-align: center; } small { font-size: 0.7em; } table { box-shadow: 9px 9px 10px 4px #BDBDBD;"
		objFile.writeline "border: 0px solid #4D7C7B;border-collapse: collapse; border-spacing: 0px; width: 1000px; margin-left: auto; margin-right: auto; } "
		objFile.writeline "tr.heading { background-color: #041944;color: #FFFFFF; font-size: 0.7em; font-weight: bold; "
		objFile.writeline "background:-o-linear-gradient(bottom, #999999 5%, #000000 100%); "
		objFile.writeline "background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #999999), color-stop(1, #000000) );"
		objFile.writeline "background:-moz-linear-gradient( center top, #999999 5%, #000000 100% );"
		objFile.writeline "filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#999999, endColorstr=#000000); "
		objFile.writeline "background:-o-linear-"
		objFile.writeline "gradient(top,#999999,000000);} tr.subheading { background-color: #6A90B6;color: #000000; font-weight: bold; font-size: 0.7em; "
		objFile.writeline "text-align:justify; } tr.section { background-color: #A4A4A4; color: #333300; cursor: pointer; font-weight: bold;font-size: 0.8em; "
		objFile.writeline "text-align: justify;"
		objFile.writeline "background:-o-linear-gradient(bottom, #56aaff 5%, #e5e5e5 100%); "
		objFile.writeline "background:-webkit-gradient( linear, left top, left bottom,color-stop(0.05, #56aaff), color-stop(1, #e5e5e5) );"
		objFile.writeline "background:-moz-linear-gradient( center top, #56aaff 5%, #e5e5e5 100% );"
		objFile.writeline "filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#56aaff, endColorstr=#e5e5e5);"
		objFile.writeline "background:-o-linear-gradient(top,#56aaff,e5e5e5);} tr.subsection { cursor: pointer; } "
		objFile.writeline "tr.content { background-color: #FFFFFF; color:#000000; font-size: 0.7em; display: table-row; } "
		objFile.writeline "tr.content2 { background-color:#;E1E1E1border: 1px solid #4D7C7B;color: #000000; "
		objFile.writeline "font-size: 0.7em; display: table-row; } td, th { padding: 5px; border: 1px solid #4D7C7B; text-align: inherit/; } th.Logos {" 
		objFile.writeline "padding: 5px; "
		objFile.writeline "border: 0px solid #4D7C7B; text-align: inherit /;} td.justified { text-align: justify; } td.pass {font-weight: bold; color: green;"
		objFile.writeline "}" 
		objFile.writeline "td.fail { font-weight: bold; color: red; } td.done, td.screenshot { font-weight: bold; color: black; } "
		objFile.writeline "td.debug { font-weight: bold;color: blue; } td.warning { font-weight: bold; color: orange; } </style> </head> "
		objFile.writeline "<body> </br><table id='header'> "
		objFile.writeline "<colgroup> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> " 
		objFile.writeline "<col style='width: 25%' /> </colgroup> <thead> <tr class='heading'> <th colspan='4' style='font-family:Copperplate Gothic Bold;" 
		objFile.writeline "font-size:1.4em;'> Test Execution Summary Report </th> </tr> <tr class='subheading'>   "
		objFile.writeline "<th>&nbsp;Date&nbsp;&&nbsp;Time&nbsp;(India Standard Time)</th> <th>&nbsp;&nbsp;"& Now &"</th>"

		objFile.writeline "<th>&nbsp;&nbsp;Oracle Version</th> <th>&nbsp;&nbsp;"& Environment.Value("ORACLEVERSION") &"</th></tr></thead></table>"
		'<tr class='subheading'> <th>&nbsp;Suite Executed</th> <th>&nbsp;&nbsp;Regression</th> <th>&nbsp;Host Name</th>"
		'objFile.writeline "<th>&nbsp;&nbsp;"& "Test System - "& Environment.Value("LocalHostName") & "</th></tr></thead></table>"
		
		objFile.writeline "<table id='main'> <colgroup> <col style='width: 10%' /> <col style='width: 40%' /> <col style='width: 20%' /> <col style='width:" 
		objFile.writeline "30%' /> </colgroup> "
		objFile.writeline "<thead> <tr class='heading'> <th>S.No</th> <th>Test Case Title</th> <th>Test Execution Status</th> <th>Execution Time</th> </tr> </thead> <tbody>"
		Set objFile = Nothing
			

		iCount=0
		iFailCount=0
		iPassCount=0
		iTestCaseNumber = 0
		iTotalExecutionTime = 0		
		For Each Item In objFiles
		   If LCase(Right(Item.Name, 5)) = ".html" Or LCase(Right(Item.Name, 4)) = ".htm" Then
			  Set objFileDetailedReport = objFSO.OpenTextFile(Item.Path, 1, False)
				 strText = objFileDetailedReport.readAll()
				 
				 Set objReg = New RegExp
				 objReg.Pattern = "[\d]+-[a-zA-Z]+-[\d]+ [\d]+:[\d]+:[\d]+"			 
				 objReg.Global = True
				 Set objMatches =  objReg.Execute(strText)
				 iStepCount = objMatches.Count
				 
				 iStartTime = objMatches(0).Value
				 iEndTime = objMatches(iStepCount-1).Value
				 iExecutionTime = Round((CDbl(DateDiff("s",CDate(iStartTime),CDate(iEndTime)))/60), 2) &" Minutes"				  
				 iTotalExecutionTime = iTotalExecutionTime + Round((CDbl(DateDiff("s",CDate(iStartTime),CDate(iEndTime)))/60), 2)
				 
				 If Instr(strText,"Fail.png") > 0 Then
					rptwritesum Item.Name, "FAIL", iExecutionTime
					iFailCount = iFailCount +1
				 Else
					rptwritesum Item.Name,"PASS" ,iExecutionTime
					iPassCount = iPassCount +1 
				 End If
		   End If
		   Set objFileDetailedReport = Nothing
		Next 
			
		Set objFile=objFSO.openTextFile(SummaryChart, 8, True)		
		objFile.writeline "</table> <table id='footer'> <colgroup> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> </colgroup> "
		objFile.writeline "<tfoot> <tr class='heading'>	<th colspan='4'>Total Execution Time (Including Report Creation) : "&iTotalExecutionTime&" Minutes </th> </tr> <tr class='content'>"
		objFile.writeline "<td class='pass'>&nbsp;Tests passed</td>	<td class='pass'>&nbsp;"&iPassCount&"</td> <td class='fail'>&nbsp;Tests failed</td>	<td class='fail'>&nbsp; "&iFailCount&"</td> </tr> </tfoot> </table>"
			
		  If UCase(Environment.Value("AUTOSUMMARYDISPLAY")) = "YES" Then SystemUtil.Run SummaryChart, ,1
		  
		  set objFile=nothing
		  set objFSO=nothing
		  set objFolder=nothing
		  set objFiles=nothing 
'		 rptRenameFolder
End Function






'******************************************

'@CreationDate <02-mar-2014>
'@Dependency  
'@Author 
'@ModifiedDate
'@ModifiedBy 
'@Description	this method writes testcase results to summary file       
'@Documentation <param> and <param> will do….
'*********************************************

Function rptWriteSum(tname, tresult, iExecutionTime)		
		Dim strResultsPath,strResultsFolder,objFileSummary,objSummary
		strResources=gTestDir &"Resources"
		SummaryChart =strProjectResultPath&"\"& gFolderName &"\SummaryChart.html"
		'strResultsFolder1 &"\SummaryChart.html"
		Set objFileSummary=CreateObject("scripting.filesystemobject")
		Set objSummary=objFileSummary.openTextFile(SummaryChart, 8, True)
		iTestCaseNumber = iTestCaseNumber+1	
		If  StrComp(tresult,0,1) = 0 or StrComp(tresult,"PASS",1) = 0  then		 
			
			objSummary.writeline "<tr class='content2'><td class='justified'><font color='#153e7e' size='1' face='arial'><b>"&iTestCaseNumber&"</b>"
			objSummary.writeline "</font></td><td class='justified'> <a href=.\TestCases\" & tname & ">"& Split(tName,".")(0) &"</a></td>"
			objSummary.writeline "<td class='justified'>Pass</td><td class='justified'>"&iExecutionTime&"</td></tr></tbody>" 
			
				
		Elseif  StrComp(tresult,1,1) = 0 or StrComp(tresult,"FAIL",1) = 0  then									
	
			objSummary.writeline "<tr class='content2'><td class='justified'><font color='#153e7e' size='1' face='arial'><b>"&iTestCaseNumber&"</b>"
			objSummary.writeline "</font></td><td class='justified'> <a href=.\TestCases\" & tname & ">"& Split(tName,".")(0) &"</a></td>"
			objSummary.writeline "<td class='justified'>Fail</td><td class='justified'>"&iExecutionTime&"</td></tr></tbody>" 
						
		End If
		
		Set objSummary = Nothing
		Set objFileSummary = Nothing 
End Function
'''******************************************************************************************************************************************************************************************************************************************######################										--    
'@Function Name <rpt_ReportLog>
'@CreationDate <02-mar-2014>
'@Dependency  
'@Author 
'@ModifiedDate
'@ModifiedBy 
'@Description	-- This Function  useds to log the step details in log file  
'@Documentation <param> and <param> will do….

''******************************************************************************************************************************************************************************************************************************************######################

Public Function rptReportLog (ByVal strStepName, ByVal strExpected,ByVal strStatus)
   ''''On Error Resume Next
   Dim objFilesys,StrLogFloder,objFile
   set objFilesys = CreateObject("Scripting.FileSystemObject")
	'StrLogFloder=Environment.Value("RESULTSPATH")&"\"&gFolderName&"\Logs\"
	'StrLogFloder=strResultsFolder1 & "\Logs\"

		
		'strResources=gTestDir &"Resources"
		
		StrLogFloder =strProjectResultPath&"\"&gFolderName &"\Logs"
	
	
	If objFilesys.FolderExists(StrLogFloder)= False Then
			objFilesys.CreateFolder(StrLogFloder)
	End If
    If objFilesys.FileExists(StrLogFloder&"\"&Environment.Value("TestName")&".txt")= false Then
   
	Set objFile=objFilesys.CreateTextFile(Trim(StrLogFloder)&"\"&Environment.Value("TestName")&".txt")
	objFile.WriteLine "Test Name"&vbtab & "Expected" & vbtab & "Status" & vbtab & "Time"
	Set objFile=Nothing
    End if
     Set objFile = objFilesys.OpenTextFile(StrLogFloder&"\"&Environment.Value("TestName")&".txt",8,True)
     objFile.WriteLine Environment.Value("TestName")& vbTab & strExpected &  vbTab & Ucase(strStatus)& vbtab& Now 
	Set objFile = Nothing
	Set objFilesys = Nothing
'	Err.Clear
End Function


'''******************************************************************************************************************************************************************************************************************************************######################										--    
'@Function Name <rptFilename>
'@CreationDate <15-mar-2014>
'@Dependency  
'@Author 
'@ModifiedDate
'@ModifiedBy 
'@Description	--   
'@Documentation <param> and <param> will do….

''******************************************************************************************************************************************************************************************************************************************######################

sub rptFoldername 
Dim dDate,strdate,Filename
dDate=Now()
sTestName=Environment.Value("TestName")
Foldername= sTestName&Day(dDate)&"-"&Month(dDate)&"-"&hour(dDate)&"-"&Minute(dDate)
Environment("FolderName") = Foldername
gFolderName=Foldername	
End sub

'''******************************************************************************************************************************************************************************************************************************************######################										--    
'@Function Name <pCopyTestData>
'@CreationDate <24-apr-2015>
'@Dependency  <None>
'@Author <Gallop Solutions>
'@ModifiedDate <None>
'@ModifiedBy <None>
'@Description	Copies Test Data file from Test Data to Results folder.
'@Documentation <param> and <param> will do….
''******************************************************************************************************************************************************************************************************************************************######################
Public Function pCopyTestData
		Set objFso = CreateObject("Scripting.FileSystemObject")
		strResultsFolder=strProjectResultPath&"\"& gFolderName 
		arrTestdata = Split(Environment("TestDir"),"ScenarioScripts") 
		strTestDataPath = arrTestdata(0) & "TestData"
		strResultsTestData=strResultsFolder & "\TestData"

		If Not objFso.FolderExists(strResultsTestData) Then
			Set objFolder=objFso.CreateFolder(strResultsTestData)
		End If
		
		If objFso.FileExists(strTestDataPath & "\" & gScenarioName &"_Testdata.xls") = True Then
			 objFso.CopyFile strTestDataPath & "\" & gScenarioName &"_Testdata.xls",strResultsTestData &"\"
		End if
		Set objFile=Nothing
End Function


'Public Function fnReportDetailedSuccess(strStepName, strStepDesc)
'
'Call rptWriteReport("Pass", strStepName , strStepDesc)
'	
'End Function
'
'
'Public Function fnReportDetailedFailure(strStepName, strStepDesc)
'
'Call rptWriteReport("Fail", strStepName , strStepDesc)
'	
'End Function
'
'Public Function fnReportDetailedPassWithScreenShot(strStepName, strStepDesc)
'
'Call rptWriteReport("Passwithscreenshot", strStepName , strStepDesc)
'	
'End Function
'
'Public Function fnReportDetailedHeader(strHeader)
'	
'End Function