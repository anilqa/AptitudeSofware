'***************************Loading ObjectRepository************************************
 RepositoriesCollection.Add("..\..\..\ObjectRepository\Common.tsr")
 RepositoriesCollection.Add("..\..\..\ObjectRepository\HR.tsr")
 RepositoriesCollection.Add("..\..\..\ObjectRepository\Finance.tsr")
 
'***************************Loading FunctionLibraries***********************************
 LoadFunctionLibrary("..\..\..\LibraryFiles\Common Functions.vbs")
 LoadFunctionLibrary("..\..\..\LibraryFiles\Bussiness Functions.vbs")
 LoadFunctionLibrary("..\..\..\LibraryFiles\RegisterFunctions.vbs")
 LoadFunctionLibrary("..\..\..\LibraryFiles\Reporting.vbs")
 LoadFunctionLibrary("..\..\..\Config\Global.vbs")
 '**************************Loading EnvironmentVariables********************************
 Environment.LoadFromFile("..\..\..\Config\Config.xml")