Rem Assigning TestCases file path to a variable
  strFilePath = "C:\Users\kabhishek\Desktop\TestCases.xls"
  Rem Sheet name
  strModuleSheetName = "Module"
  strSheetName = "TestCase"
  Rem Creating ADODB connection object
  Set objConn = CreateObject("ADODB.Connection")
  Rem connecting to Excel db
  With objConn
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "Data Source=" & strFilePath & ";" & _
                        "Extended Properties=Excel 8.0;"
    .Open          
  End With
  
  if Err.Number <> 0 Then
    Call Sub_voidTestLog("Fun_arrReadTestCasesFile - failed to connect to 'TestCases' file. Error Description: "&Err.Description,False)
    Call Fun_EnvWrite("strGenericFunctionStatus","Fail")
   
  End if
  Rem Creating ADODB record set object
  Set objRS = CreateObject("ADODB.Recordset")
  
  strSQL="SELECT ModuleNames FROM [Modules$] where ToBeExecuted ='Yes'"
  'strSQL="SELECT TestCaseID FROM [TestCase$] where ModuleName IN('Module2','A')"
			
  objRS.open strSQL, objConn
   Do while not objRS.eof     
    for each field in objRS.fields
      if Trim(strtext)="" then
        strtext = field.value
      Else
        strtext=strtext &","& field.value 
      End if
      
      if Err.Number <> 0 Then 
       
      End if  
      
      Exit For
    next                   
    objRS.moveNext()
  loop
  MsgBox strtext
  objRS.close