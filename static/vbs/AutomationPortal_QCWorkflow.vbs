' Public Variable Declaration
'QC Connection variables
On Error Resume Next
StartTime = Timer
Public gbQCConnection, gbqcURL, gbqcID, qcDomain, gbQCProjectName, gbQCConnectionResult
Public  gbTestSetFound, gbCurrTestCaseName, gbRunsCt, gbTestScriptResultPath, gbExcelRowNumber, gbExcelUsedRowCount
Public gbExecuteScript, gbTestPlanFolderPath, gbTestScriptNameToExecute, gbTestSetFolderPath, gbTestSetName, gbRemoteMachineName, gb_rt_IsStart, gb_rt_IsStop, gbTrigger
Public gbobjExcel, gbobjWorkBook, gbobjsheet, Wscript, gbNewTestSetName, gbNewCreatedTestSetName, gbTestScriptNameToAdd, gbTestType
'Test Script type in QC: QUICKTEST_TEST or MANUAL
gbTestType = "QUICKTEST_TEST"

'Create Global QC Connection
Set gbQCConnection = CreateObject("TDApiOle80.TDConnection")

'Declare Dynamic Public Array
Call KillProcess("EXCEL.exe")
If Err.Number <> 0 Then
	Err.Number = 0
	Err.Description = ""
End If

Call RetrieveExcelData()


Public Function RetrieveExcelData()
	On Error Resume Next
	Call OpenWorkBook()
	
	'Read Excel Record to Execute Script
	gbExcelUsedRowCount = gbobjsheet.usedrange.rows.count
	For Row = 2 To gbExcelUsedRowCount
		gbExecuteScript = gbobjsheet.cells(Row,1).value
		gbTestPlanFolderPath = gbobjsheet.cells(Row,2).value
		gbTestScriptNameToExecute = gbobjsheet.cells(Row,3).value
		testSetPath = gbobjsheet.cells(Row,4).value
		Env = gbobjsheet.cells(2,11).value
		gbTestSetFolderPath = testSetPath & "\" & Env
		'SanityName = gbobjsheet.cells(Row,5).value
		gbRemoteMachineName = gbobjsheet.cells(Row,6).value
		gbTrigger = gbobjsheet.cells(Row,10).value
		gbExcelRowNumber = Row		
		If Ucase(gbExecuteScript) = "YES" Then
			If Ucase(gbTrigger) = "" Then
				'Set Trigger Status
				gbobjsheet.cells(gbExcelRowNumber,10).value = "YES"					
				Call SaveAndCloseWorkBook()
				
				gbIsExecutionStarted = QCWorkFlow(gbExcelRowNumber)
			 	Call WaitTime()	
			 	
	'		 	' get QC Runninsg Status
	'		 	gbExecutionStatus = QC_GetExecutionStatusOfScript(gbTestSetFolderPath,gbTestSetName, gbTestScriptNameToExecute)
	'			'Set Final Run Status
	'			Call OpenWorkBook()
	'			gbobjsheet.cells(gbExcelRowNumber,9).value = gbExecutionStatus			
	'			Call SaveAndCloseWorkBook()
	'			'ReleaseExcelObject()		
	''			Set ClassObj = Nothing	
		 	End If
		End If
		If Row = gbExcelUsedRowCount Then			
			Exit Function	
		End If	
	Next		
End Function

Public Function QCWorkFlow(ByVal gbExcelRowNumber)
	On Error Resume Next
	'msgbox "Excel Row: " & gbExcelRowNumber
	' Establish QC Connection
	gbQCConnectionResult = ConnectToQC()
	
	If gbQCConnectionResult = "QCConnectionPass" Then
		Call OpenWorkBook()
		If gbobjsheet.cells(2,12).value <> "YES" Then
			' First Time only: Create TestSet in ALM and Pull Required Test Cases				 
			gbNewTestSetName = ALM_CreateTestSetAndAddTestScripts(gbTestPlanFolderPath, gbTestSetFolderPath, gbTestType )
		Else
			' Get Test Set name
			gbNewTestSetName = gbobjsheet.cells(2,13).value
		End If
		Call SaveAndCloseWorkBook()
				
'		' Set flag in Excel for Start and stop combination
'		'Start as YES and Stop as NO
'		Call OpenWorkBook()
'		gbobjsheet.cells(gbExcelRowNumber,7).value = "YES"
'		gbobjsheet.cells(gbExcelRowNumber,8).value = "NO"	
'		'Set Running Status
'		gbobjsheet.cells(gbExcelRowNumber,9).value = "Launching UFT"			
'		Call SaveAndCloseWorkBook()
		gbExecuteScript = ExecuteScriptUsingUFTQCOTA(gbExecuteScript, gbTestPlanFolderPath, gbTestScriptNameToExecute, gbTestSetFolderPath, gbNewTestSetName, gbRemoteMachineName, gbExcelRowNumber, gb_rt_IsStart, gb_rt_IsStop)
	Else
		Exit Function		
	End If
	
	' Release QC Objects
	ReleaseQCObject = ReleaseQCObject()
	' Release Excel Objects
	ReleaseExcelObject = ReleaseExcelObject()
	
End Function
	
'Establish QC Connection
Public Function ConnectToQC()
	On Error Resume Next
'	Set gbQCConnection = CreateObject("TDApiOle80.TDConnection")
	gbqcURL = "http://sv2wnecoqc01:8080/qcbin"
	gbqcID = "automationqateam"
	qcDomain = "DEFAULT"
	gbQCProjectName = "ECO"
	gbQCPassword = "Welcome1"
	
	' Excel Msg
'	'Set Running Status
'	Call OpenWorkBook()
'	gbobjsheet.cells(gbExcelRowNumber,9).value = "Connecting QC"	
'	'mgbox "Connecting QC"	
'	Call SaveAndCloseWorkBook()
	
	' Connect to QC
	gbQCConnection.InitConnectionEx gbqcURL
	gbQCConnection.Login gbqcID, gbQCPassword 	' Password tmp
	gbQCConnection.Connect qcDomain, gbQCProjectName 
	Call WaitTime()	

	
	If Err.Number = 0 Then
		ConnectToQC = "QCConnectionPass"
'		'Set Running Status
'		Call OpenWorkBook()
'		gbobjsheet.cells(gbExcelRowNumber,9).value = "QCConnectionPass"	
		msgbox "QCConnectionPass on Machine : " & gbRemoteMachineName		
'		Call SaveAndCloseWorkBook()
		
	Else
		ConnectToQC = "QCConnectionFail"
'		'Set Running Status
'		Call OpenWorkBook()
'		gbobjsheet.cells(gbExcelRowNumber,9).value = "QCConnectionFail"	
		msgbox "QCConnectionFail on Machine : " & gbRemoteMachineName		
'		Call SaveAndCloseWorkBook()
	End If
	 EndTime = Timer
 	ElapsedTime = EndTime - StartTime
' 	msgbox ElapsedTime
	
End Function


'' Create Test Set at Run Time in ALM
''Call ConnectToQC()
'
'gbNewTestSetName = "RelMgmt"
'gbTestSetFolderPath = "Root\ECO Automation\Product-Execution\AutomationPortal\UAT4" & "\"
'''Function to create New Test Set in ALM
''gbNewCreatedTestSetName = ALM_CreateNewTestSet(gbTestSetFolderPath, gbNewTestSetName)
'
''Function to Add Test Scripts under New Created Test Set
'gbTestPlanFolderPath = "[QualityCenter]Subject\Automation\ECO\Products\AutomationPortal" & "\"
'gbTestScriptNameToAdd = "AP_Product_ConfigurableAccessories"
'gbNewCreatedTestSetName = "AP_RelMgmt_5_30_2017_11_08_UAT4"
''gbAddTestScript = ALM_AddTestScriptsUnderNewTestSet(gbTestSetFolderPath, gbNewCreatedTestSetName, gbTestPlanFolderPath, gbTestScriptNameToAdd, gbTestType)
'
'NewCreatedTestSetName = ALM_CreateTestSetAndAddTestScripts(gbTestPlanFolderPath, gbTestSetFolderPath, gbTestType )

Public Function ALM_CreateTestSetAndAddTestScripts(ByVal testPlanPath,ByVal testLabPath, ByVal testType)
	On Error Resume Next
	' Define Test Set Naming convention
	sTmp = Replace(Time, ":", "_")
	sArr = Split(sTmp, "_")
	sTime = sArr(0) & "_" & sArr(1)
	'Get Local Host Name
	Set wshNetwork = CreateObject("WScript.Network")
	strLocalHostName = wshNetwork.ComputerName

	' Get Test Set name for Team and Env from First Row
	SanityName = gbobjsheet.cells(2,5).value
	Env = gbobjsheet.cells(2,11).value
	gbNewTestSetName = "AP_"& strLocalHostName & "_" & SanityName &"_"& Replace(Date,"/","_") & "_" & sTime & "_" & Env
	
	
    Set QCTreeManager=gbQCConnection.TreeManager
    Set TestNode=QCTreeManager.nodebypath(testPlanPath)
    Set TestFact = TestNode.TestFactory
    Set TestsList = TestFact.NewList("")
    
    
    Set QCTSTreeManager = gbQCConnection.TestSetTreeManager     
    Set  TreeNode=QCTSTreeManager.NodebyPath(testLabPath)
    Set TestSetFact=TreeNode.TestSetFactory
    
    Set NewTestSet=TestSetFact.AddItem(Null)' Creates new testset
    NewTestSet.name=gbNewTestSetName
    NewTestSet.Field("CY_COMMENT")=gbNewTestSetName
    NewTestSet.status="Open"
    NewTestSet.post
    
    Set TSTestFactory=NewTestSet.TSTestFactory
    For ExcelRowNumber = 2 to gbExcelUsedRowCount  
		gbTestScriptNameToAdd = gbobjsheet.cells(ExcelRowNumber,3).value 
	    For Each Tests in TestsList
	        If Tests.field("TS_TYPE")= testType Then
	        	If Tests.Name = gbTestScriptNameToAdd Then
	        		TSTestFactory.Additem(tests)
	        		Exit For
	        	End If
	            
	        End If
	    Next
    Next
    'Update the Excel Sheet that New Test set is Created as : YES
    gbobjsheet.cells(2,12).value = "YES"
    gbobjsheet.cells(2,13).value = gbNewTestSetName
    ALM_CreateTestSetAndAddTestScripts = gbNewTestSetName

End Function


	
Public Function ExecuteScriptUsingUFTQCOTA(ByVal gbExecuteScript, ByVal gbTestPlanFolderPath, ByVal gbTestScriptNameToExecute, ByVal gbTestSetFolderPath, ByVal gbTestSetName, ByVal gbRemoteMachineName, ByVal gbExcelRowNumber, ByRef rt_IsStart, ByRef rt_IsStop)
	On Error Resume Next
	'excel flag
	gb_rt_IsStart = "YES"
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
		    	
		    	' Create Object for QTP
		    	Set qtApp = CreateObject("QuickTest.Application",gbRemoteMachineName)   
		         If qtApp.launched <> True then					       
		             qtApp.Launch   
		         End If  
		        qtApp.Visible = "true"  
		        
'		        'Set Running Status
'		        Call OpenWorkBook()
'				gbobjsheet.cells(gbExcelRowNumber,9).value = "Loading Script"	
'				'mgbox "Loading Script"			
'				Call SaveAndCloseWorkBook()
	        	
	'        	 If Not qtApp.gbQCConnection.IsConnected Then  
	'    	       qtApp.gbQCConnection.Connect gbqcURL,qcDomain,gbQCProjectName,gbqcID,gbQCPassword,False  
	'      		 End If 
	      		 
	          	' Create the Application object  
	          	
	'          	' tmp code
	'          		If gbExcelRowNumber = 2 Then
	'          			qtApp.Open gbTestPlanFolderPath & "\" & "AP_PrintConfiguredQuote_PCKBP", True  ' Open the test in read-only mode 
	'          			'msgbox "Script Loading AP_PrintConfiguredQuote_PCKBP on Machine " & gbRemoteMachineName
	'          		ElseIf gbExcelRowNumber = 3 Then
	'          			qtApp.Open gbTestPlanFolderPath & "\" & "AP_Product_Private Cage with kVA Based Power", True  ' Open the test in read-only mode 
	'          			'msgbox "Script Loading AP_Product_Private Cage with kVA Based Power on Machine " & gbRemoteMachineName
	'          		ElseIf gbExcelRowNumber = 4 Then
	'          			qtApp.Open gbTestPlanFolderPath & "\" & "AP_Product_ConfigurableAccessories", True  ' Open the test in read-only mode 
	'          			'msgbox "Script Loading AP_Product_ConfigurableAccessories on Machine " & gbRemoteMachineName
	'          		ElseIf gbExcelRowNumber = 5 Then
	'          			qtApp.Open gbTestPlanFolderPath & "\" & "AP_Product_EQX_IEPP_Equinix Internet Exchange Port", True  ' Open the test in read-only mode
	'					'msgbox "Script Loading AP_Product_EQX_IEPP_Equinix Internet Exchange Port on Machine " & gbRemoteMachineName	          			
	'          		End If
	'          	' tmp code end
	          	
			      qtApp.Open TestScript, True  ' Open the test in read-only mode  
			      
			      Set qtTest = qtApp.Test 
	        	
	        	  Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions")  
	        	  qtResultsOpt.TDTestInstance = 1 
	        	 qtResultsOpt.TDRunName= "Run_" & Month(Now) & "-" & Day(Now) & "_" & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now)
	        	 qtResultsOpt.TDTestSet = gbTestScriptResultPath ' Path to the Test Set where we should save the results 
	        	 
	        	   Set fso=createobject("Scripting.FileSystemObject")   
				    If fso.FolderExists("C:\Res1") Then   
				         fso.DeleteFolder("C:\Res1")  
				    End If  
					    qtResultsOpt.ResultsLocation = "C:\Res1"  
					    qtTest.Run qtResultsOpt,True  
'					   	'Set Running Status
'					   	Call OpenWorkBook()
'						gbobjsheet.cells(gbExcelRowNumber,9).value = "Script Running"
'						'mgbox "Script Running"					
'						Call SaveAndCloseWorkBook()
							    
					    TestStatus = qtTest.LastRunResults.Status				    
'					   	'Set Running Status
'					   	Call OpenWorkBook()
'						gbobjsheet.cells(gbExcelRowNumber,9).value = TestStatus
						ExecuteScriptUsingUFTQCOTA = TestStatus
'						'mgbox TestStatus					
'						Call SaveAndCloseWorkBook()
						
					    qtTest.Close  
					    qtApp.quit  
					    
					    Set qtApp = Nothing  	
					    Call WaitTime()
					    Exit Function
				End If
	    Next
	Next
	
End Function

Public Function OpenWorkBook()
	Set gbobjExcel = CreateObject("Excel.Application")
	Set gbobjWorkBook = gbobjExcel.Workbooks.Open("C:\Users\automationqateam\Project\venv2\static_cdn\data\AutomationPortal_Data.xls")
'	Set gbobjWorkBook = gbobjExcel.Workbooks.Open("\\sgw-filesvr3\GDC\GDC_Team\QA\ECO-Product-Docs\Automation\AutomationPortal\QCWorkFlow\AutomationPortal_Data_V1.xls")
'	Set gbobjWorkBook = gbobjExcel.Workbooks.Open("\\sgw-filesvr3\GDC\GDC_Team\QA\ECO-Product-Docs\Automation\AutomationPortal\QCWorkFlow\AutomationPortal_Data.xlsx")
	set gbobjsheet=gbobjWorkBook.sheets(1)
	Call WaitTime()	
End Function

Public Function SaveAndCloseWorkBook()
	gbobjWorkBook.Save
	Call WaitTime()	
	gbobjWorkBook.Close
	Call ReleaseExcelObject()
	
	Call KillProcess("EXCEL.exe")
	If Err.Number <> 0 Then
		Err.Number = 0
		Err.Description = ""
	End If
	Call WaitTime()	
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

End Function
	
Public Function ReleaseExcelObject()
	On Error Resume Next	
	' Release Excel Object
	 gbobjWorkBook.Close
	 gbobjExcel.quit
End Function
	
Public Function WaitTime()
	StartTime = Timer
	While Timer - StartTime < 5
	Wend

End Function
	

'************************************************************************************

'For Example
'Call KillProcess("wscript.exe")
'Call KillProcess("UFTRemoteAgent.exe")
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
'	
'	
'		'*********************************************************************
'	' Get LAN IP Address of HostName
'	'e.g. My laptop ip address of 
''	IPAddress_LAN = GetLAN_IPAddress("SGLP8DQSK72")
''	print IPAddress_LAN
'	Function GetLAN_IPAddress(ByVal HostName)
'		On Error Resume Next
'	  Dim Wscript, objExec, StrOutput, RegEx 
'	
'	  Set Wscript = CreateObject("WScript.Shell")
'	
'	'  Set objExec = Wscript.Exec("ping " & ComputerName & " -n 1")
'	  Set objExec = Wscript.Exec("ipconfig")
'	  StrOutput = objExec.StdOut.ReadAll
'	  
'	  'Store the Result of IPConfig Command into Text file.
'	  Set objFSO=CreateObject("Scripting.FileSystemObject")
'	  outFile="C:\Temp\IPAddress_"& HostName & ".txt"
'	  Set objFile = objFSO.CreateTextFile(outFile,True)
'	  objFile.Write StrOutput & vbCrLf
'	  objFile.Close
'	  
'	  'Read File
'	  isLANIPAddressSection_Found = False
'	  Set objFile = objFSO.OpenTextFile(outFile)
'	  Do Until objFile.AtEndOfStream
'	    strLine= objFile.ReadLine
'	    If Instr(strLine, "Ethernet adapter Local Area Connection:") Then
'	    	isLANIPAddressSection_Found = True    	
'	    End If
'	    ' Find IP Address under Ethernet Section
'	    If isLANIPAddressSection_Found = True Then
'	    	If Instr(strLine, "IPv4 Address. . . . . . . . . . . :") Then
'	    		LAN_IP_Address = Split(strLine, ":")
'	    		GetLAN_IPAddress = Trim(LAN_IP_Address(1))
'	    		objFile.Close
'	    		Exit Do
'	    	End If
'	    
'	    End If
'	  Loop
'	  
'	  return = objFSO.DeleteFile("C:\Temp\IPAddress_"& HostName & ".txt", True)
'	  
'	End Function
'	
'	'bVMFlag =  IsMachineVirtual()
'	Public Function IsMachineVirtual()
'		On Error Resume Next
'		SystemName = "localhost"
'		
'		Set tmpObj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
'		SystemName & "\root\cimv2").InstancesOf ("Win32_ComputerSystem")
'		for each tmpItem in tmpObj
'		  'MakeModel = trim(tmpItem.Manufacturer) & " " & trim(tmpItem.Model)
'		  MakeModel = trim(tmpItem.Model)
'		next
'		Set tmpObj = Nothing
'		Set tmpItem = Nothing
'		
'		If InStr(1,MakeModel,"VMware",1)>0 Then
'			IsMachineVirtual = True
'		Else
'			IsMachineVirtual = False
'		End If
'	End Function
'	
'	Public Function QCUnLockUser()
'		
'		Set QCConnection=QCUtil.QCConnection 
'		Set con=QCConnection.command 
'		con.CommandText="DELETE FROM LOCKS WHERE LK_USER ='automationqateam'" 
'		Set recset=con.execute
'	
'	End Function
''End Class
'
''**********************************************************************
'Public Function ALM_CreateNewTestSet(ByVal gbTestSetFolderPath, ByVal gbNewTestSetName)
'	On Error Resume Next
'	
'	Set tsTreeMgr_1 = gbQCConnection.TestSetTreeManager 
'	Set tsFolder_1 = tsTreeMgr_1.NodeByPath(gbTestSetFolderPath)
'	Set tsTestSet_1 = tsFolder_1.TestSetFactory
'	Set aFilter = tsTestSet_1.Filter
'	aFilter.Filter("CY_CYCLE") = gbNewTestSetName
'	Set lst = tsTestSet_1.NewList(aFilter.Text)
'	If lst.Count = 0 Then
'        Set TstSet = tsTestSet_1.AddItem(Null)
'        TstSet.Field("CY_CYCLE") = gbNewTestSetName
'        TstSet.Post
'        msgbox "New Test Set '"& gbNewTestSetName & "' is Created under Test Lab Path : '" & gbTestSetFolderPath & "."
'    Else
'        Set TstSet = lst.Item(1)
'        msgbox "New Test set name is not created due to error : " & err.description
'    End If
' 	ALM_CreateNewTestSet = gbNewTestSetName
'	
'End Function
'
'Public Function ALM_AddTestScriptsUnderNewTestSet(ByVal gbTestSetFolderPath, ByVal gbNewCreatedTestSetName, ByVal gbTestPlanFolderPath, ByVal gbTestScriptNameToAdd, ByVal gbTestType)
'	On Error Resume Next	 
'	
'	'Test Plan Factory
'	 Set tsPlanTreeMgr_2 = gbQCConnection.TreeManager
'	Set TestPlanPathNode_2 = tsPlanTreeMgr_2.nodebypath(gbTestPlanFolderPath)
'	Set TestPlanFact = TestPlanPathNode_2.TestFactory
'	Set TestsList = TestPlanFact.NewList("")
'	Counter = 0   
'	For each TestScript in TestsList
'		Set objTest = TestScript
'		Counter = Counter+1
'
'		If TestScript.Name = gbTestScriptNameToAdd Then	
'			TestScriptID = TestScript.ID
'			TestScriptFound = True
'			Exit For
'		End If		
'	Next
'	
'	If TestScriptFound = True Then	
'		
'		
'		' Test Lab Factory
'		Set tsLabTreeMgr_2 = gbQCConnection.TestSetTreeManager
'		Set tsLabFolder_2 = tsLabTreeMgr_2.NodeByPath(gbTestSetFolderPath)
'		Set tsList = tsLabFolder_2.FindTestSets(gbNewCreatedTestSetName)
'		
''		tsList.Additem(TestScriptID)
'		
'	'	Set NewTestSet=tsTestSetFact_2.AddItem(Null)
'	'	NewTestSet.name=gbNewCreatedTestSetName
'	'	NewTestSet.Field("CY_COMMENT")=gbNewCreatedTestSetName
'	'	NewTestSet.status="Open"
'	'	NewTestSet.post
'		
'		Set TSTestFactory=tsLabFolder_2.TSTestFactory
'		Set newTSTest = TSTestFactory.AddItem(TestScriptID)
'		' Post the change to the DB
'		newTSTest.Post
'		
''		For Each Tests in TestsList
''	        If Tests.field("TS_TYPE")= gbTestType Then
''	            TSTestFactory.Additem("AP_PrintConfiguredQuote_PCKBP")
''	        End If
''	    Next
'	End If
'
''	Set tsList = tsTestSet_2.FindTestSets(gbTestSetName)
''	Set TstSet = tsList.AddItem(Null)
''	
''	'Set tsTestSet_2 = tsList.TestSetFactory
''	set tsTempTest = tsList.AddItem("154784") 
''	tsTempTest.Post
'End Function
'
'