strProjectPath = Split(Environment("TestDir"),"TestCases")(0)
strProjectTestdataPath = Split(Environment("TestDir"),"TestCases")(0) & "TestData\"
str_ORHRXMLPath = Split(Environment("TestDir"),"TestCases")(0) &"ObjectRepository\HR.xml"
str_ORFSXMLPath = Split(Environment("TestDir"),"TestCases")(0) &"ObjectRepository\Finance.xml"
Environment("ERRORFLAG") = True
'Set oOracleLogin = Browser("OracleApplicationsHome").Page("OracleApplicationsHome")
'Set oFSObj = Browser("OracleEBS").Page("OracleEBS")
'Public oPage
'Public oFrame
'Dim rKey
Const MIN_WAIT = 2
Const MID_WAIT = 5
Const MAX_WAIT = 10

Dim strProjectResultPath
strProjectResultPath = Split(Environment("TestDir"),"TestCases")(0) & "Results"

Dim gFolderName
On Error Resume Next
gFolderName = Environment("FolderName")
On Error Goto 0
