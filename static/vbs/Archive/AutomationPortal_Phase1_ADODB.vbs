'Author: Automation QA Team
' Public Variable Declaration
'QC Connection variables
On Error Resume Next
StartTime = Timer

Public gbQCConnection, gbqcURL, gbqcID, qcDomain, gbQCProjectName, gbQCConnectionResult
Public  gbTestSetFound, gbCurrTestCaseName, gbRunsCt, gbTestScriptResultPath, gbExcelRowNumber, gbExcelUsedRowCount
Public gbExecuteScript, gbTestPlanFolderPath, gbTestScriptNameToExecute, gbTestSetFolderPath, gbTestSetName, gbRemoteMachineName, gb_rt_IsStart, gb_rt_IsStop, gbTrigger
'Excel object
Public gbobjExcel, gbobjWorkBook, gbobjsheet, Wscript, gbNewTestSetName, gbNewCreatedTestSetName, gbTestScriptNameToAdd, gbTestType, gbSanityName, gbEnv
'Adodob.connection object and variables
Public sFilenameWithPath, gbSheetName, gbWhereQuery, gbColumnName, gbobjExcelConnection, gbobjExcelRecordSet
Public gbIsTestSetCreated

'Test Script type in QC: QUICKTEST_TEST or MANUAL
gbTestType = "QUICKTEST_TEST"

'*******************************************
'ADODB Record set
gbFilenameWithPath = "C:\Users\automationqateam\Project\venv2\static_cdn\data\RMData.xls"
gbSheetName = "RM"
gbWhereQuery = ""
ReturnVal = ExecuteQCWorkFlow(gbFilenameWithPath, gbSheetName, gbWhereQuery)

Call CloseExcelADODBConnection()


Public Function ExecuteQCWorkFlow(ByVal gbFilenameWithPath, ByVal gbSheetName, ByVal gbWhereQuery)
	On Error Resume Next
	Call OpenExcelADODBConnection()
	gbExcelRowNumber = 2
	For i=0  to gbobjExcelRecordSet.recordcount -1
		j = i	
	    'Get Trigger Value Data
		gbTrigger = gbobjExcelRecordSet.fields("Trigger")
		gbExcelRowNumber = i+2		
		If (IsNull(gbTrigger) Or (gbTrigger = "Error")) Then
		
			'Get Excel Data
			gbExecuteScript = gbobjExcelRecordSet.fields("Execute")
			gbTestPlanFolderPath = gbobjExcelRecordSet.fields("TestPlanFolderPath")
			gbTestScriptNameToExecute = gbobjExcelRecordSet.fields("TestScriptNameToExecute")
			testSetPath = gbobjExcelRecordSet.fields("TestSetFolderPath")
			gbEnv = gbobjExcelRecordSet.fields("Environment")
			gbTestSetFolderPath = testSetPath & "\" & gbEnv
			gbSanityName = gbobjExcelRecordSet.fields("SanityName")
			gbRemoteMachineName = gbobjExcelRecordSet.fields("RemoteMachineName")
		
			'Set Trigger Status
			gbobjExcelRecordSet.fields("Trigger") = "YES"
			Call CloseExcelADODBConnection()				

			gbIsExecutionStarted = QCWorkFlow(gbExcelRowNumber)
		 	Call WaitTime(2)
			Exit Function	
		 	
		 End If
	
		Call OpenExcelADODBConnection()
	    j = j+1
	    'increase excel recordset count
	    gbobjExcelRecordSet.move j
	Next
	
	'Close Recordset and Excel Connection
	Call CloseExcelADODBConnection()
	ExecuteQCWorkFlow = gbIsExecutionStarted
End Function


Public Function QCWorkFlow(ByVal gbExcelRowNumber)
	On Error Resume Next
	'Create Global QC Connection
	Set gbQCConnection = CreateObject("TDApiOle80.TDConnection")

	' Establish QC Connection
	gbQCConnectionResult = ConnectToQC()
	
	If gbQCConnectionResult = "QCConnectionPass" Then
		Call OpenExcelADODBConnection()
		
		gbobjExcelRecordSet.MoveFirst	' get the value of second row from excel
		gbIsTestSetCreated = gbobjExcelRecordSet.fields("IsTestSetCreated")
		If IsNull(gbIsTestSetCreated) Then
			' First Time only: Create TestSet in ALM and Pull Required Test Cases				 
			gbNewTestSetName = ALM_CreateTestSetAndAddTestScripts(gbTestPlanFolderPath, gbTestSetFolderPath, gbTestType)
		Else
			' Get Test Set name
			gbNewTestSetName = gbobjExcelRecordSet.fields("ALM_NewTestSetName")
		End If
		Call CloseExcelADODBConnection()
		EndTimer = Timer
		ElapsedTime = EndTimer - StartTime
	'	msgbox ElapsedTime
		
		' Function call to Execute Script workflow
		gbExecuteScript = ExecuteScriptUsingUFTQCOTA(gbExecuteScript, gbTestPlanFolderPath, gbTestScriptNameToExecute, gbTestSetFolderPath, gbNewTestSetName, gbRemoteMachineName, gbExcelRowNumber, gb_rt_IsStart, gb_rt_IsStop)
	Else
		Exit Function		
	End If
	
'	' Release QC Objects
	Call ReleaseQCObject()
	
End Function
	
'Establish QC Connection
Public Function ConnectToQC()
	On Error Resume Next
	gbqcURL = "http://sv2wnecoqc01:8080/qcbin"
	gbqcID = "automationqateam"
	qcDomain = "DEFAULT"
	gbQCProjectName = "ECO"
	gbQCPassword = "Welcome2"
	
	' Connect to QC
	gbQCConnection.InitConnectionEx gbqcURL
	gbQCConnection.Login gbqcID, gbQCPassword 	' Password tmp
	gbQCConnection.Connect qcDomain, gbQCProjectName 
	Call WaitTime(2)	

	
	If Err.Number = 0 Then
		ConnectToQC = "QCConnectionPass"
	Else
		ConnectToQC = "QCConnectionFail"
		Call OpenExcelADODBConnection()
			gbobjExcelRecordSet.move gbExcelRowNumber-2
			gbobjExcelRecordSet.fields("RunningStatus") = "QCConnectionFail"
		Call CloseExcelADODBConnection()
		'msgbox "QCConnectionFail on Machine : " & gbRemoteMachineName & " : For Script name : " & gbTestScriptNameToExecute
	End If
	
End Function


Public Function ALM_CreateTestSetAndAddTestScripts(ByVal testPlanPath,ByVal testLabPath, ByVal testType)
	On Error Resume Next
	' Define Test Set Naming convention
	sTmp = Replace(Time, ":", "_")
	sAMPM = Split(sTmp, " ")
	sArr = Split(sTmp, "_")
	sTime = sArr(0) & "_" & sArr(1) & "_" & sAMPM(1)
	'Get Local Host Name
	Set wshNetwork = CreateObject("WScript.Network")
	strLocalHostName = wshNetwork.ComputerName

	' create New test Set name for Team and Env from First Row
	gbNewTestSetName = "AP_"& strLocalHostName & "_" & gbSanityName &"_"& Replace(Date,"/","_") & "_" & sTime & "_" & gbEnv	
	
'    Set QCTreeManager=gbQCConnection.TreeManager
'    Set TestNode=QCTreeManager.nodebypath(testPlanPath)
'    Set TestFact = TestNode.TestFactory
'    Set TestsList = TestFact.NewList("")    
    
    Set QCTSTreeManager = gbQCConnection.TestSetTreeManager     
    Set  TreeNode=QCTSTreeManager.NodebyPath(testLabPath)
    Set TestSetFact=TreeNode.TestSetFactory
    
    Set NewTestSet=TestSetFact.AddItem(Null)' Creates new testset
    NewTestSet.name=gbNewTestSetName
    NewTestSet.Field("CY_COMMENT")=gbNewTestSetName
    NewTestSet.status="Open"
    NewTestSet.post
    
    Call OpenExcelADODBConnection()
    
    Set TSTestFactory=NewTestSet.TSTestFactory
    For i=0  to gbobjExcelRecordSet.recordcount -1
    	' get current test plan path    	
    	testPlanPath = gbobjExcelRecordSet.fields("TestPlanFolderPath")
	    Set QCTreeManager=gbQCConnection.TreeManager
	    Set TestNode=QCTreeManager.nodebypath(testPlanPath)
	    Set TestFact = TestNode.TestFactory
	    Set TestsList = TestFact.NewList("") 		    
    
		gbTestScriptNameToAdd = gbobjExcelRecordSet.fields("TestScriptNameToExecute") 
	    For Each Tests in TestsList
	        If Tests.field("TS_TYPE")= testType Then
	        	If Tests.Name = gbTestScriptNameToAdd Then
	        		TSTestFactory.Additem(tests)
	        		Exit For
	        	End If
	            
	        End If
	    Next
	   	'increase excel recordset count
	    gbobjExcelRecordSet.movenext
    Next
    'Update the Excel Sheet that New Test set is Created as : YES
	gbobjExcelRecordSet.MoveFirst
	gbobjExcelRecordSet.fields("IsTestSetCreated") = "YES"
	gbobjExcelRecordSet.fields("ALM_NewTestSetName") = gbNewTestSetName
	Call CloseExcelADODBConnection()
  
    ALM_CreateTestSetAndAddTestScripts = gbNewTestSetName

End Function
	
Public Function ExecuteScriptUsingUFTQCOTA(ByVal gbExecuteScript, ByVal gbTestPlanFolderPath, ByVal gbTestScriptNameToExecute, ByVal gbTestSetFolderPath, ByVal gbTestSetName, ByVal gbRemoteMachineName, ByVal gbExcelRowNumber, ByRef rt_IsStart, ByRef rt_IsStop)
	On Error Resume Next
	'Start Execution
	Set TSetFact = gbQCConnection.TestSetFactory
	Set tsTreeMgr = gbQCConnection.TestSetTreeManager 
	Set tsFolder = tsTreeMgr.NodeByPath(gbTestSetFolderPath)
	Set tsList = tsFolder.FindTestSets(gbTestSetName)
	strTable = ""  
	Set theTestSet = tsList.Item(1)  
	For Testset_Count=1 to tsList.Count
		Set theTestSet = tsList.Item(Testset_Count)   
	    Set TSTestFact = theTestSet.TSTestFactory  
	    TSName = theTestSet.Name  
	    Set TestSetTestsList = TSTestFact.NewList("")
	    
	    For Each theTSTest In TestSetTestsList 
	    	TestName = theTSTest.Test.Name  
			If TestName = gbTestScriptNameToExecute Then
	  	
		     	TestScript = gbTestPlanFolderPath & "\" & TestName 
				gbTestScriptResultPath = gbTestSetFolderPath & "\" & gbTestSetName     
		      	TestStatus = theTSTest.Status 
		    	
		    	' Create Object for UFT
		    	On Error Goto 0
		    	Set qtApp = CreateObject("QuickTest.Application",gbRemoteMachineName)   
		         If qtApp.launched <> True then					       
		             qtApp.Launch   
		         End If  
		        qtApp.Visible = "true"  
		        'handle error, if UFT crash
		        If Err.Number <> 0 Then
		        	Call OpenExcelADODBConnection()
		        		gbobjExcelRecordSet.fields("Trigger") = "Error"
		        		gbobjExcelRecordSet.fields("ErrorMessage") = Err.Description		        			        		
		        	Call CloseExcelADODBConnection()
		        	Exit Function ' here server will handle the Error and re-trigger the vbs to start the script again
		        End If 
		        On Error Goto 0
		        
				' Connect UFT to QC in IDE, if not connected for Automationqateam user
	'        	 If Not qtApp.gbQCConnection.IsConnected Then  
	'    	       qtApp.gbQCConnection.Connect gbqcURL,qcDomain,gbQCProjectName,gbqcID,gbQCPassword,False  
	'      		 End If 
	      		 
        		' Open the Script        	
			      qtApp.Open TestScript, True  ' Open the test in read-only mode  
			      
			     'handle error, if UFT not able to open script from QC
			        If Err.Number <> 0 Then
			        	Call OpenExcelADODBConnection()
			        		gbobjExcelRecordSet.fields("Trigger") = "Error"
			        		gbobjExcelRecordSet.fields("ErrorMessage") = Err.Description		        			        		
			        	Call CloseExcelADODBConnection()
			        	Exit Function ' here server will handle the Error and re-trigger the vbs to start the script again
			        End If 
			        On Error Goto 0
		        
			      Set qtTest = qtApp.Test 
	        	
	        	' Select Run Test Instance from QC based on created Test set name
	        	  Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions")  
	        	  qtResultsOpt.TDTestInstance = 1 
	        	  qtResultsOpt.TDRunName= "Run_" & Month(Now) & "-" & Day(Now) & "_" & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now)
	        	  qtResultsOpt.TDTestSet = gbTestScriptResultPath ' Path to the Test Set where we should save the results 
	        	 
	        	   Set fso=createobject("Scripting.FileSystemObject")   
				    If fso.FolderExists("C:\Res1") Then   
				         fso.DeleteFolder("C:\Res1")  
				    End If  
					    qtResultsOpt.ResultsLocation = "C:\Res1"
						
						'update status in Excel
						Call OpenExcelADODBConnection()
							gbobjExcelRecordSet.move gbExcelRowNumber-2
							gbobjExcelRecordSet.fields("RunningStatus") = "Running"
							'Msgbox "RunStatus for Script : " & gbTestScriptNameToExecute & " :Updated as: Running-->"
						Call CloseExcelADODBConnection()
						
						'Run Script
					    qtTest.Run qtResultsOpt,True 
						
		   				'handle error, if UFT not able to execute script from QC
				        If Err.Number <> 0 Then
				        	Call OpenExcelADODBConnection()
				        		gbobjExcelRecordSet.fields("Trigger") = "Error"
				        		gbobjExcelRecordSet.fields("ErrorMessage") = Err.Description		        			        		
				        	Call CloseExcelADODBConnection()
				        	Exit Function ' here server will handle the Error and re-trigger the vbs to start the script again
				        End If 
			       		On Error Goto 0
						' Get status of Execution after execution completed
					    TestStatus = qtTest.LastRunResults.Status	
						
						Call OpenExcelADODBConnection()
							gbobjExcelRecordSet.move gbExcelRowNumber-2
							If TestStatus = "Warning" Then
								TestStatus = "Passed"								
							End If
							gbobjExcelRecordSet.fields("RunningStatus") = TestStatus
						'	Msgbox "RunStatus for Script : " & gbTestScriptNameToExecute & " :Updated as: " & TestStatus
						Call CloseExcelADODBConnection()

						ExecuteScriptUsingUFTQCOTA = TestStatus
					
					    qtTest.Close  
					    qtApp.quit  
					    
					    Set qtApp = Nothing  	
					    Call WaitTime()
					    Exit Function
				End If
	    Next
	Next
	
End Function




Public Function OpenExcelADODBConnection()
	On Error Resume Next
	Set gbobjExcelConnection = CreateObject("ADODB.Connection")
	gbobjExcelConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & gbFilenameWithPath & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;"";"
	Call WaitTime(1)
	
	'Custor Type for VbScript to handle Excel
	Const adOpenStatic = 3			''---- CursorTypeEnum Values ----
	Const adLockPessimistic = 2		''---- LockTypeEnum Values ----
	Const adCmdText = "&H0001"		''---- CommandTypeEnum Values ----
	
	sql_text="Select * FROM [" & gbSheetName & "$]" & gbWhereQuery
	
	' create RecordSet
	Set gbobjExcelRecordSet = CreateObject("ADODB.Recordset")
	
	'Execute SQL and store results in reocrdset'
	gbobjExcelRecordSet.Open sql_text , gbobjExcelConnection, adOpenStatic, adLockPessimistic, adCmdText
	Call WaitTime(1)
End Function

Public Function CloseExcelADODBConnection()
	On Error Resume Next
	'Close and Discard all variables '
	gbobjExcelRecordSet.save
	gbobjExcelRecordSet.Close
	gbobjExcelConnection.Close
	Set gbobjExcelRecordSet = Nothing
	Set gbobjExcelConnection = Nothing
	If Err.Number <> 0 Then
		Err.Number = 0
		Err.Description = ""
	End If 
		
	Call WaitTime(2)
End Function

	
'Get QC Script Run Status
Public Function QC_GetExecutionStatusOfScript(ByVal gbTestSetFolderPath, ByVal gbTestSetName, ByVal gbTestScriptNameToExecute)
	On Error Resume Next
	Set TestSets = gbQCConnection.TestSetTreeManager.NodeByPath(gbTestSetFolderPath).TestSetFactory.NewList("")
	
	If gbTestSetName <> "" Then
		For Each testSet In TestSets
			If testSet.Name = gbTestSetName Then
				gbTestSetFound = True
				Set oTestSet  = testSet.TsTestFactory
				Exit For
			End If
		Next
		If Not gbTestSetFound Then 
			Exit Function
		End If
		
		Set oTestSetFilter = oTestSet.Filter
		
		If gbTestSetFound Then
			For Each TestRunInstance in oTestSet.NewList(oTestSetFilter.Text)
				gbCurrTestCaseName = Split(TestRunInstance.Name,"]")(1)
				If gbCurrTestCaseName = gbTestScriptNameToExecute Then
					gbCurrTestCaseRunStatus = Split(TestRunInstance.Status,"]")(1)
					'Return Script Run status
					QC_GetExecutionStatusOfScript = gbCurrTestCaseRunStatus
				End If
				
			Next
		End If
	Else
		Exit Function
	End If
End Function

Public Function ReleaseQCObject()
	On Error Resume Next
	' Release Global Varaible 
	 gbQCProjectName = Null
	 qcDomain = NULL
	 gbqcID = NULL
	 gbqcURL = NULL
 
	 'Release Global Object
	 Set gbQCConnection = Nothing

	If Err.Number <> 0 Then
		Err.Number = 0
		Err.Description = ""
	End If 
		
	Call WaitTime(2)
End Function
	
	
Public Function WaitTime()
	StartTime = Timer
	While Timer - StartTime < 5
	Wend

End Function

Public Function WaitTime(ByVal Seconds)
	StartTime = Timer
	While Timer - StartTime < Seconds
	Wend

End Function	

'************************************************************************************

'For Example
'Call KillProcess("notepad++.exe")
Public Function KillProcess(ByVal ProcessName)
	On Error Resume Next
	Const intTerminationCode=0
	Dim  ObjService,ObjInstance,Process,AppOpenProcess,strProPath,intStatus
    For i = 1 to 5
				KillProcess = True
				Set objService = getobject("winmgmts:")
				For Each Process In objService.InstancesOf("Win32_process")
							AppOpenProcess = True
							If Ucase(Process.Name)=Ucase(ProcessName) Then
										AppOpenProcess = False
										strProPath = "Win32_Process.Handle=" & Process.processid
										Set objInstance = objService.Get(strProPath)
										intStatus = objInstance.Terminate(intTerminationCode)
										If intStatus=0 Then
												KillProcess = True
										Else
												KillProcess = False
										End If
										If KillProcess = True Then
												On Error Goto 0
												Exit For
										End If
							End If
				Next
				Set objInstance = nothing
				On Error Goto 0
			
	Next
	On Error Goto 0
End Function
