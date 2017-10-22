'Get Current status from Test Lab for given Test Set and update in RM Data Sheet of Portal excel
On Error Resume Next
Public gbFilenameWithPath, gbSheetName, gbWhereQuery
Public gbTestPlanFolderPath, gbNewTestSetName, gbobjExcelConnection, gbobjExcelRecordSet, gbQCConnection

Set dict_ScriptNameFromRMData = CreateObject("Scripting.Dictionary")
Set dict_ScriptNameFromALM = CreateObject("Scripting.Dictionary")
Set dict_ScriptStatusFromALM = CreateObject("Scripting.Dictionary")

Set gbQCConnection = CreateObject("TDApiOle80.TDConnection")

'ADODB Record set
gbFilenameWithPath = "C:\Users\automationqateam\Project\venv2\static_cdn\data\RMData.xls"
gbSheetName = "RM"
gbWhereQuery = ""
gbNewTestSetName = ""
gbTestScriptNameToExecute = ""


'Set dict_FailTCName = CreateObject("Scripting.Dictionary")
'For k = LBound(gbArrFailedTestName) To UBound(gbArrFailedTestName)
'	dict_FailTCName.Add "TCName_"& k, gbArrFailedTestName(k)
'Next

'Get Scripts name from RM Data into Dictionary
Call OpenExcelADODBConnection()
	testSetPath = gbobjExcelRecordSet.fields("TestSetFolderPath")
	gbEnv = gbobjExcelRecordSet.fields("Environment")
	gbTestPlanFolderPath = testSetPath & "\" & gbEnv
			
	gbNewTestSetName = gbobjExcelRecordSet.fields("ALM_NewTestSetName")
	For i=0  to gbobjExcelRecordSet.recordcount -1
		QCRunStatus = ""
		gbTestScriptNameToGetStatus = ""
		gbTestScriptNameToGetStatus = gbobjExcelRecordSet.fields("TestScriptNameToExecute")		
		'Store script name in dictionary
		dict_ScriptNameFromRMData.Add "ScriptName_"& i, gbTestScriptNameToGetStatus	
		QCRunStatus = QC_GetExecutionStatusOfScript(gbTestPlanFolderPath, gbNewTestSetName, gbTestScriptNameToGetStatus)
		If QCRunStatus = "No Run" Then
			'if status is no run after 10 mins of first call from vbs, then re-start the script with this flag
			gbobjExcelRecordSet.fields("Trigger") = "Error"	
			gbobjExcelRecordSet.fields("RunningStatus") = QCRunStatus			
			gbobjExcelRecordSet.fields("ErrorMessage") = "Script Loading Error on VM ; Still QC Status is: " & QCRunStatus
		Else
			gbobjExcelRecordSet.fields("RunningStatus") = QCRunStatus
			
		End If	
		gbobjExcelRecordSet.save
		If Err.Number <> 0 Then
			Err.Number = 0
			Err.Description = ""
		End If		
		
		gbobjExcelRecordSet.movenext
	Next
Call CloseExcelADODBConnection()



'Get QC Script Run Status
Public Function QC_GetExecutionStatusOfScript(ByVal gbTestPlanFolderPath, ByVal gbNewTestSetName, ByVal gbTestScriptNameToGetStatus)
	On Error Resume Next
		' Establish QC Connection
	gbQCConnectionResult = ConnectToQC()
	
	If gbQCConnectionResult = "QCConnectionPass" Then
	
		Set TestSets = gbQCConnection.TestSetTreeManager.NodeByPath(gbTestPlanFolderPath).TestSetFactory.NewList("")
		
		If gbNewTestSetName <> "" Then
			For Each testSet In TestSets
				If testSet.Name = gbNewTestSetName Then
					gbTestSetFound = True
					Set oTestSet  = testSet.TsTestFactory
					Exit For
				End If
			Next
			If Not gbTestSetFound Then 
				Exit Function
			End If
			
			Set oTestSetFilter = oTestSet.Filter
			
			If gbNewTestSetName Then
				For Each TestRunInstance in oTestSet.NewList(oTestSetFilter.Text)
					gbCurrTestCaseName = Split(TestRunInstance.Name,"]")(1)
					If gbCurrTestCaseName = gbTestScriptNameToGetStatus Then
						gbCurrTestCaseRunStatus = TestRunInstance.status
						'Return Script Run status
						QC_GetExecutionStatusOfScript = gbCurrTestCaseRunStatus
						Exit Function
					End If
					
				Next
			End If
		Else
			Exit Function
		End If
	Else
		Exit Function
	End If
End Function





'*****************************************************
' PUBLIC FUNCTION 
'*****************************************************


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

'*************************************************************