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
Set dict_ScriptWithHostName = CreateObject("Scripting.Dictionary")
'Set dict_HostName = CreateObject("Scripting.Dictionary")

'**************************************************
'PostGre Database Variables
Public objDBConnection, objDBRecordset, DBTableName, QCConnectTableName, gbID
'******************************************************

'Test Script type in QC: QUICKTEST_TEST or MANUAL
gbTestType = "QUICKTEST_TEST"

QCConnectTableName = "automationUI_ate_connecttoqc"
DBTableName = "automationUI_ate_rmdata"
gbDBRowNum = "1"
ReturnVal = ExecuteQCWorkFlow() 



Public Function ExecuteQCWorkFlow()
	On Error Resume Next
	Call ConnectToPostGresDB()
	Call SelectRecordfromDB(DBTableName)
	gbID = 1
	Do Until objDBRecordset.Eof
		ID = objDBRecordset.Fields("RowID").Value
		gbID = ID
	    'Get Trigger Value Data
	     gbTrigger = objDBRecordset.Fields("Trigger").Value
		If ((gbTrigger = "") Or (gbTrigger = "Error")) Then
			'Get DB Data
			gbExecuteScript = objDBRecordset.Fields("Execute").Value
			gbTestPlanFolderPath = objDBRecordset.Fields("TestPlanFolderPath").Value
			gbTestScriptNameToExecute = objDBRecordset.Fields("TestScriptNameToExecute").Value
			testSetPath = objDBRecordset.Fields("TestSetFolderPath").Value
			gbEnv = objDBRecordset.Fields("Environment").Value
			gbTestSetFolderPath = testSetPath & "\" & gbEnv
			gbSanityName = objDBRecordset.Fields("SanityName").Value
			gbRemoteMachineName = objDBRecordset.Fields("RemoteMachineName").Value

			Call ReleaseDBobjects()

			gbIsExecutionStarted = QCWorkFlow(gbExcelRowNumber)
		 	Call WaitTime(2)
			Exit Function	
		 	
		 End If
			
		Call ConnectToPostGresDB()
		Call SelectRecordfromDB(DBTableName)

		objDBRecordset.MoveNext
	Loop
'	Next
	
	'Close Recordset and Excel Connection
	Call ReleaseDBobjects()
	ExecuteQCWorkFlow = gbIsExecutionStarted
End Function


Public Function QCWorkFlow(ByVal gbExcelRowNumber)
	On Error Resume Next
	'Create Global QC Connection
	Set gbQCConnection = CreateObject("TDApiOle80.TDConnection")

	' Establish QC Connection
	gbQCConnectionResult = ConnectToQC()
	
	If gbQCConnectionResult = "QCConnectionPass" Then
		Call ConnectToPostGresDB()
		Call SelectRecordfromDB(DBTableName)
		objDBRecordset.MoveFirst
		gbIsTestSetCreated = objDBRecordset.Fields("IsTestSetCreated").Value
	
		If gbIsTestSetCreated= "" Then
			gbNewTestSetName = ALM_CreateTestSetAndAddTestScripts(gbTestPlanFolderPath, gbTestSetFolderPath, gbTestType)
		Else
			' Get Test Set name
			gbNewTestSetName = objDBRecordset.Fields("ALM_NewTestSetName").Value
		End If
	
		Call ReleaseDBobjects()
		'updated code to run from QC - start
		'Get Host name in Dictionary
		gbRemoteMachineName = ""
		Call ConnectToPostGresDB()
		Call SelectRecordfromDB(DBTableName)
		Do Until objDBRecordset.Eof
		
			ScriptName = objDBRecordset.Fields("TestScriptNameToExecute").Value
			gbRemoteMachineName = objDBRecordset.Fields("RemoteMachineName").Value
			' store values into Dictionary
			dict_ScriptWithHostName.Add ScriptName, gbRemoteMachineName
			
			objDBRecordset.MoveNext
		Loop
		'updated code to run from QC - end
		
		Call ReleaseDBobjects()
		
		'update Trigger as Yes for all records in DB
		Call ConnectToPostGresDB()
		RecordCount =  SelectCountFromDB(DBTableName)
		For RowID = 1 To CInt(RecordCount)
    		return = UpdateRecordIntoDB(DBTableName, "Trigger", "YES", RowID)
		
		Next
		Call ReleaseDBobjects()
	
		EndTimer = Timer
		ElapsedTime = EndTimer - StartTime
		gbExecuteScript = ExecuteScriptsFromQC(gbTestSetFolderPath, gbNewTestSetName, dict_ScriptWithHostName)
	Else
		Exit Function		
	End If
	
	Call ReleaseQCObject()
	
End Function
	
'Establish QC Connection
Public Function ConnectToQC()
	On Error Resume Next
	
	Call ConnectToPostGresDB()
	Call SelectRecordfromDB(QCConnectTableName)

	gbqcURL = objDBRecordset.Fields("qcServer").Value
	gbqcID = objDBRecordset.Fields("qcUser").Value
	gbQCPassword = objDBRecordset.Fields("qcPassword").Value
	qcDomain = objDBRecordset.Fields("qcDomain").Value
	gbQCProjectName = objDBRecordset.Fields("qcProject").Value
		
	' Connect to QC
	gbQCConnection.InitConnectionEx gbqcURL
	gbQCConnection.Login gbqcID, gbQCPassword 	' Password tmp
	gbQCConnection.Connect qcDomain, gbQCProjectName 
	Call WaitTime(2)	
	
	If Err.Number = 0 Then
		ConnectToQC = "QCConnectionPass"
	Else
		ConnectToQC = "QCConnectionFail"
	End If
	Call ReleaseDBobjects()
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
	
    Set QCTSTreeManager = gbQCConnection.TestSetTreeManager     
    Set  TreeNode=QCTSTreeManager.NodebyPath(testLabPath)
    Set TestSetFact=TreeNode.TestSetFactory
    
    Set NewTestSet=TestSetFact.AddItem(Null)' Creates new testset
    NewTestSet.name=gbNewTestSetName
    NewTestSet.Field("CY_COMMENT")=gbNewTestSetName
    NewTestSet.status="Open"
    NewTestSet.post
    
    Call ConnectToPostGresDB()
	Call SelectRecordfromDB(DBTableName)
	
	Call ConnectToPostGresDB()
	RecordCount =  SelectCountFromDB(DBTableName)

    Set TSTestFactory=NewTestSet.TSTestFactory
    For RowID = 1 To CInt(RecordCount)
    	gbDBRowNum = RowID
    	Call SelectRecordfromDB(DBTableName)
    	' get current test plan path    
		testPlanPath = objDBRecordset.Fields("TestPlanFolderPath").Value    	
	    Set QCTreeManager=gbQCConnection.TreeManager
	    Set TestNode=QCTreeManager.nodebypath(testPlanPath)
	    Set TestFact = TestNode.TestFactory
	    Set TestsList = TestFact.NewList("") 	
	    
    	gbTestScriptNameToAdd = objDBRecordset.Fields("TestScriptNameToExecute").Value
	    For Each Tests in TestsList
	        If Tests.field("TS_TYPE")= testType Then
	        	If Tests.Name = gbTestScriptNameToAdd Then
	        		TSTestFactory.Additem(tests)
	        		Exit For
	        	End If
	            
	        End If
	    Next
	   	
	Next
	Call ReleaseDBobjects()
    'update db
    Call ConnectToPostGresDB()
    	return = UpdateRecordIntoDB(DBTableName, "IsTestSetCreated", "YES", 1)
		return = UpdateRecordIntoDB(DBTableName, "ALM_NewTestSetName", gbNewTestSetName, 1)
	Call ReleaseDBobjects()
    ALM_CreateTestSetAndAddTestScripts = gbNewTestSetName

End Function

Public Function ExecuteScriptsFromQC(ByVal gbTestSetFolderPath, ByVal gbTestSetName, ByVal dict_ScriptWithHostName)
	
	Set TSetFact = gbQCConnection.TestSetFactory
	Set tsTreeMgr = gbQCConnection.TestSetTreeManager
	
	Set TestFolderNode = tsTreeMgr.NodeByPath(gbTestSetFolderPath)
	Set SubFolderNode = TestFolderNode.NewList
	Set TestSetNode = TestFolderNode.TestSetFactory
	Set TestSetList = TestSetNode.NewList("")
	
	arrScriptName = dict_ScriptWithHostName.Keys
	arrHostName = dict_ScriptWithHostName.Items
	
	For Each TestSet In TestSetList
		TestSet.AutoPost = True
		Set Scheduler = TestSet.StartExecution("")    
		
		If TestSet.Name = gbTestSetName Then		
			Set TestFactory=TestSet.tsTestFactory
			Set tsFilter = TestFactory.Filter
			tsFilter.Filter("TC_CYCLE_ID") = TestSet.ID
			Set testList = TestFactory.NewList(tsFilter.Text)		
			
		    Dim tsFilter 'As TDFilter
		    Dim TSTst 'As TSTest
			
		    For Each TSTst In testList        
		        TSTst.AutoPost = True
		        'Assign Host name to scripts in order as per test data sheet
		        For i = LBound(arrScriptName) to UBound(arrScriptName)
		        	If TSTst.Name = "[1]" & arrScriptName(i) Then
		        		TSTst.HostName = arrHostName(i)
	        		    Scheduler.RunOnHost(TSTst.ID) = TSTst.HostName
	        		    Exit For
		       		End If 
		        Next
		        
		        Scheduler.RunAllLocally = False
		        
		    Next
		    
			TestSet.Refresh
			TestFolderNode.Refresh
		End If
	Next
	
	Scheduler.Run
		
	' Get the execution status object.
	Set execStatus = Scheduler.ExecutionStatus
	' Track the events and statuses.
	Dim RunFinished, i
	Dim ExecEventInfoObj, EventsList
	Dim TestExecStatusObj
	
	Do While RunFinished = False	    
	    execStatus.RefreshExecStatusInfo "all", True
	    RunFinished = execStatus.Finished
	    Set EventsList = execStatus.EventsList
	
	    For i = 1 To execStatus.Count
	        Set TestExecStatusObj = execStatus.Item(i)
	    Next
	    
	Loop 'Loop While execStatus.Finished = False
	
	Set ExecEventInfoObj=Nothing
	Set execStatus=Nothing
	Set EventsList=Nothing	
	'Set strNowEvent=Nothing
	Set Scheduler = Nothing
	RunFinished = False
		
	Set TestFactory=Nothing
	Set testList=Nothing
	
	gbQCConnection.Disconnect
	Set gbQCConnection = Nothing

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
						        	
					Call ConnectToPostGresDB()
					
					return = UpdateRecordIntoDB(DBTableName, "Trigger", "Error", 1)
					return = UpdateRecordIntoDB(DBTableName, "ErrorMessage", Err.Description, 1)
				
					Call ReleaseDBobjects()
		        	Exit Function ' here server will handle the Error and re-trigger the vbs to start the script again
		        End If 
		        On Error Goto 0
	      		 
        		' Open the Script        	
			      qtApp.Open TestScript, True  ' Open the test in read-only mode  
			      
			     'handle error, if UFT not able to open script from QC
			        If Err.Number <> 0 Then
			        	Call ConnectToPostGresDB()
					
						return = UpdateRecordIntoDB(DBTableName, "Trigger", "Error", 1)
						return = UpdateRecordIntoDB(DBTableName, "ErrorMessage", Err.Description, 1)
					
						Call ReleaseDBobjects()
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
						
						Call ConnectToPostGresDB()
						Do Until objDBRecordset.Eof
							return = UpdateRecordIntoDB(DBTableName, "RunningStatus", "Running", gbID)						
							objDBRecordset.MoveNext
						Loop
						Call ReleaseDBobjects()
						
					'Run Script
					    qtTest.Run qtResultsOpt,True 
						
		   				'handle error, if UFT not able to execute script from QC
				        If Err.Number <> 0 Then
				        	Call ConnectToPostGresDB()
						
								return = UpdateRecordIntoDB(DBTableName, "Trigger", "Error", 1)
								return = UpdateRecordIntoDB(DBTableName, "ErrorMessage", Err.Description, 1)
							Call ReleaseDBobjects()
				        	Exit Function ' here server will handle the Error and re-trigger the vbs to start the script again
				        End If 
			       		On Error Goto 0
						' Get status of Execution after execution completed
					    TestStatus = qtTest.LastRunResults.Status	
						
						Call ConnectToPostGresDB()
					
							If TestStatus = "Warning" Then
								TestStatus = "Passed"								
							End If
						Do Until objDBRecordset.Eof
							return = UpdateRecordIntoDB(DBTableName, "RunningStatus", TestStatus, gbID)
							objDBRecordset.MoveNext
						Loop
						Call ReleaseDBobjects()

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


'*******************************************
'PostGre Database Functions

Public Function ConnectToPostGresDB()
	On Error Resume Next
	Set objDBConnection = CreateObject("ADODB.Connection")
	objDBConnection.Open "Provider=PostgreSQL OLE DB Provider;Data Source=localhost;location=ate;User ID=postgres;password=postgres"
	If err.number = 0 Then
		'msgbox "Database Connection is successfull"
		ConnectToPostGresDB = objDBConnection
	else
		ConnectToPostGresDB = Err.Description
	End If	
End Function

Public Function UpdateRecordIntoDB(ByVal TableName, ByVal ColumnName, ByVal ColumnValue, ByVal RowID)
	On Error Resume Next
	Set objDBRecordset = CreateObject("ADODB.Recordset")
	Set objDBRecordset =  objDBConnection.Execute("UPDATE "& """"& TableName &""""&" SET "&""""& ColumnName &""""&"='"& ColumnValue &"' where id='"& RowID &"'")
	If err.number = 0 Then
		UpdateRecordIntoDB = True
	else
		UpdateRecordIntoDB = False
	End If
End Function

Public Function SelectRecordfromDB(ByVal TableName)
	On Error Resume Next
	Set objDBRecordset = CreateObject("ADODB.Recordset")
	If gbDBRowNum <> "" Then
		rowID = gbDBRowNum
	Else
		'default first row
		rowID = 1
	End If
	
	Set objDBRecordset =  objDBConnection.Execute("SELECT * FROM "& """"  &TableName & """"&" where id = " & rowID)
	If err.number = 0 Then
		SelectRecordfromDB = objDBRecordset
	else
		SelectRecordfromDB = Err.Description
	End If
	
	'msgbox objDBRecordset.Fields(0).Value

End Function

Public Function SelectAllRecordfromDB(ByVal TableName)
	On Error Resume Next
	Set objDBRecordset = CreateObject("ADODB.Recordset")
	Set objDBRecordset =  objDBConnection.Execute("SELECT * FROM "& """"  &TableName & """")
	If err.number = 0 Then
		SelectRecordfromDB = objDBRecordset
	else
		SelectRecordfromDB = Err.Description
	End If
End Function

Public Function SelectCountFromDB(ByVal TableName)
	On Error Resume Next
	Set objDBRecordset = CreateObject("ADODB.Recordset")
	Set objDBRecordset =  objDBConnection.Execute("SELECT Count(*) FROM "& """"  &TableName & """")
	RecordCount = objDBRecordset.Fields(0).Value
	If err.number = 0 Then
		SelectCountFromDB = RecordCount
	else
		SelectCountFromDB = Err.Description
	End If

End Function

Public Function ReleaseDBobjects()
	On Error Resume Next
	objDBConnection.Close
	Set objDBRecordset = Nothing
	Set objDBConnection = Nothing
	If Err.Number <> 0 Then
		Err.Number = 0
		Err.Description = ""
	End If 
	
End Function


'******************************************


