'Author: Automation QA Team
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

'2 is default Excel Row number
gbExcelRowNumber = "2"
Call QCWorkFlow(gbExcelRowNumber)

Public Function QCWorkFlow(ByVal gbExcelRowNumber)
	On Error Resume Next
	' Establish QC Connection
	gbQCConnectionResult = ConnectToQC()
	
	If gbQCConnectionResult = "QCConnectionPass" Then
		
		gbExecuteScript = "YES"
		gbTestPlanFolderPath = "[QualityCenter]Subject\Automation\ECO\AutomationPortalScript"
		gbTestScriptNameToExecute = "AP_AutoEmailScript"
		gbTestSetFolderPath = "Root\ECO Automation\Product-Execution\AutomationPortal"
		gbNewTestSetName = "AP_AutoEmail"
		gbRemoteMachineName = ""
		
		gbExecuteScript = ExecuteAutoEmail(gbExecuteScript, gbTestPlanFolderPath, gbTestScriptNameToExecute, gbTestSetFolderPath, gbNewTestSetName, gbRemoteMachineName, gbExcelRowNumber, "", "")
	Else
		Exit Function		
	End If
	
	' Release QC Objects
	ReleaseQCObject = ReleaseQCObject()

	
End Function
	
'Establish QC Connection
Public Function ConnectToQC()
	On Error Resume Next
'	Set gbQCConnection = CreateObject("TDApiOle80.TDConnection")
	gbqcURL = "http://sv2wnecoqc01:8080/qcbin"
	gbqcID = "automationqateam"
	qcDomain = "DEFAULT"
	gbQCProjectName = "ECO"
	gbQCPassword = "Welcome2"
	
	' Connect to QC
	gbQCConnection.InitConnectionEx gbqcURL
	gbQCConnection.Login gbqcID, gbQCPassword 	' Password tmp
	gbQCConnection.Connect qcDomain, gbQCProjectName 
	Call WaitTime()	
	
	If Err.Number = 0 Then
		ConnectToQC = "QCConnectionPass"	
	Else
		ConnectToQC = "QCConnectionFail"
	End If
	On Error Goto 0
	On Error Resume Next

	
End Function


	
Public Function ExecuteAutoEmail(ByVal gbExecuteScript, ByVal gbTestPlanFolderPath, ByVal gbTestScriptNameToExecute, ByVal gbTestSetFolderPath, ByVal gbTestSetName, ByVal gbRemoteMachineName, ByVal gbExcelRowNumber, ByRef rt_IsStart, ByRef rt_IsStop)
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

		    	' Create Object for QTP for Remote machine
'		    	Set qtApp = CreateObject("QuickTest.Application",gbRemoteMachineName)   
				
				'create object of QTP for local machinee
		         Set qtApp = CreateObject("QuickTest.Application") 
		         
		         If qtApp.launched <> True Then      
		         	qtApp.Launch   
		         End If  
		        qtApp.Visible = "False"  
		        
	
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
							    
					    TestStatus = qtTest.LastRunResults.Status				    
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

End Function
	
Public Function WaitTime()
	StartTime = Timer
	While Timer - StartTime < 5
	Wend

End Function
	
'************************************************************************************
