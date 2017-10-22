'Get Current status from Test Lab for given Test Set and update in RM Data Sheet of Portal excel
On Error Resume Next
Public gbFilenameWithPath, gbSheetName, gbWhereQuery
Public gbTestPlanFolderPath, gbNewTestSetName, gbobjExcelConnection, gbobjExcelRecordSet, gbQCConnection

Set dict_ScriptNameFromRMData = CreateObject("Scripting.Dictionary")
Set dict_ScriptNameFromALM = CreateObject("Scripting.Dictionary")
Set dict_ScriptStatusFromALM = CreateObject("Scripting.Dictionary")

Set gbQCConnection = CreateObject("TDApiOle80.TDConnection")

'**************************************************
'PostGre Database Variables
Public objDBConnection, objDBRecordset, DBTableName, QCConnectTableName, gbID, gbDBRowNum
'******************************************************

QCConnectTableName = "automationUI_ate_connecttoqc"
DBTableName = "automationUI_ate_rmdata"

gbNewTestSetName = ""
gbTestScriptNameToExecute = ""
gbDBRowNum = "1"
Call ConnectToPostGresDB()
Call SelectRecordfromDB(DBTableName)

testSetPath = objDBRecordset.Fields("TestSetFolderPath").Value
gbEnv = objDBRecordset.Fields("Environment").Value
gbSetFolderPath = testSetPath & "\" & gbEnv
gbNewTestSetName = objDBRecordset.Fields("ALM_NewTestSetName").Value
Call ReleaseDBobjects()

'Get Scripts name from RM Data into Dictionary
Call ConnectToPostGresDB()
RecordCount =  SelectCountFromDB(DBTableName)
Call ReleaseDBobjects()

If gbNewTestSetName <> "" Then
	
	gbQCConnectionResult = ConnectToQC()

	For Iteration = 1 To CInt(RecordCount)
		gbDBRowNum = Iteration
		Call ConnectToPostGresDB()
		Call SelectRecordfromDB(DBTableName)	
		QCRunStatus = ""
		gbTestScriptNameToGetStatus = ""		
		gbTestScriptNameToGetStatus = objDBRecordset.Fields("TestScriptNameToExecute").Value

		QCRunStatus = QC_GetExecutionStatusOfScript(gbSetFolderPath, gbNewTestSetName, gbTestScriptNameToGetStatus)

		Call ReleaseDBobjects()
		Call ConnectToPostGresDB()
		If QCRunStatus = "No Run" Then
			'if status is no run after 10 mins of first call from vbs, then re-start the script with this flag
			return = UpdateRecordIntoDB(DBTableName, "Trigger", "Error", gbDBRowNum)
			return = UpdateRecordIntoDB(DBTableName, "RunningStatus", QCRunStatus, gbDBRowNum)
			return = UpdateRecordIntoDB(DBTableName, "ErrorMessage", "Script Loading Error on VM ; Still QC Status is: " & QCRunStatus, gbDBRowNum)

		Else
			return = UpdateRecordIntoDB(DBTableName, "RunningStatus", QCRunStatus, gbDBRowNum)
		
		End If	
		Call ReleaseDBobjects()
	Next
	Call ReleaseQCObject()
End If
'Get QC Script Run Status
Public Function QC_GetExecutionStatusOfScript(ByVal gbTestPlanFolderPath, ByVal gbNewTestSetName, ByVal gbTestScriptNameToGetStatus)
	On Error Resume Next
		' Establish QC Connection	
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
			
			If gbTestSetFound Then
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

'*******************************************
'PostGre Database Functions

Public Function ConnectToPostGresDB()
	On Error Resume Next
	Set objDBConnection = CreateObject("ADODB.Connection")	 
	objDBConnection.Open "Provider=PostgreSQL OLE DB Provider;Data Source=localhost;location=ate;User ID=postgres;password=postgres"
	If err.number = 0 Then	
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