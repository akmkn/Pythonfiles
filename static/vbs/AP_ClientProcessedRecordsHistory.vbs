' Collect Processed Client Requested Records in Incremental way into Excel.
On Error Resume Next
Public gbobjExcelConnection, gbobjExcelRecordSet, gbFilenameWithPath, gbSheetName
Public gbobjExcel, gbobjWorkBook, gbobjsheet
Public gbClientIPAddress, gbDateTime
Public gbTestSetFolderPath, gbALM_NewTestSetName, gbMailTo, sDateTime
Set dict_PortalRequestExecutionHistory = CreateObject("Scripting.Dictionary")


'ADODB Record set

'FilePath of Client IP and DateTime of Request from Proceed to Submit File
ProceedToSubmitFilePath = "C:\Users\automationqateam\Project\venv2\static_cdn\data\ProceedToSubmit.xls"
ProceedToSubitSheetName = "Sheet1"

'FilePath of TestSet Path, Testsetname, WatchList Email ID of Request from Proceed to Submit File
RMDataFilePath = "C:\Users\automationqateam\Project\venv2\static_cdn\data\RMData.xls"
RMDataSheetName = "RM"


'FilePath Processed Client Requests Records into File in Incremental way.
ClientRequestRecordFilePath = "C:\Users\automationqateam\Project\venv2\static_cdn\data\PortalProcessedRequests_Records.xls"
ClientRequestSheetName = "ExecutionRecords"

'Get Client IP and DateTime of Request from Proceed to Submit File
Call OpenExcelADODBConnection(ProceedToSubmitFilePath, ProceedToSubitSheetName)
	'goto first record
	gbobjExcelRecordSet.moveFirst		' Operation is not allowed when object is closed error. in vbs
	gbClientIPAddress = gbobjExcelRecordSet.fields("ClientRequestIP")
	gbDateTime = gbobjExcelRecordSet.fields("RequestTime")
	
'	' using dictionary to read, if required
'	dict_PortalRequestExecutionHistory.Add "gbClientIPAddress1", gbClientIPAddress
'	msgbox dict_PortalRequestExecutionHistory.Item("gbClientIPAddress1")	
Call CloseExcelADODBConnection()

'Get TestSet Path, Testsetname, WatchList Email ID of Request from Proceed to Submit File
Call OpenExcelADODBConnection(RMDataFilePath, RMDataSheetName)
	'goto first record
	gbobjExcelRecordSet.moveFirst
	gbTestSetFolderPath = gbobjExcelRecordSet.fields("TestSetFolderPath")
	gbALM_NewTestSetName = gbobjExcelRecordSet.fields("ALM_NewTestSetName")
	gbMailTo = gbobjExcelRecordSet.fields("MailTo")	
	
Call CloseExcelADODBConnection()

'Store Processed Client Requests Records into File in Incremental way.
Call OpenWorkBook(ClientRequestRecordFilePath)
	'goto first record
	sDateTime = Now()
	RequestCounter = gbobjsheet.UsedRange.Rows.Count

	gbobjsheet.cells(RequestCounter+1,1).value = RequestCounter
	gbobjsheet.cells(RequestCounter+1,2).value = gbClientIPAddress
	gbobjsheet.cells(RequestCounter+1,3).value = gbDateTime
	gbobjsheet.cells(RequestCounter+1,4).value = gbTestSetFolderPath
	gbobjsheet.cells(RequestCounter+1,5).value = gbALM_NewTestSetName
	gbobjsheet.cells(RequestCounter+1,6).value = gbMailTo
	gbobjsheet.cells(RequestCounter+1,7).value = sDateTime
	If Err.Number <> 0 Then
		Err.Number = 0
		Err.Description = ""
	End If
	
'		'using dictionary to write, if required
'	gbobjsheet.cells(RequestCounter,2).value = dict_PortalRequestExecutionHistory.Item("gbClientIPAddress")
'	msgbox dict_PortalRequestExecutionHistory.Item("gbClientIPAddress")
	
Call SaveAndCloseWorkBook()


' showing error on storing value 3021 : "Either BOF or EOF is True, or the current record has been deleted. Requested operation requires a current record."
''Store Processed Client Requests Records into File in Incremental way.
'Call OpenExcelADODBConnection(ClientRequestRecordFilePath, ClientRequestSheetName)
'	'goto first record
'	RequestCounter = gbobjExcelRecordSet.recordcount
'	RequestCounter = RequestCounter + 1
'	gbobjExcelRecordSet.moveNext
'	
'	gbobjExcelRecordSet.fields("RequestCounter") = RequestCounter
'	gbobjExcelRecordSet.fields("Client_IP_Address") = gbClientIPAddress
'	gbobjExcelRecordSet.fields("DateTime") = gbDateTime
'	gbobjExcelRecordSet.fields("TestSetFolderPath") = gbTestSetFolderPath
'	gbobjExcelRecordSet.fields("ALM_NewTestSetName") = gbALM_NewTestSetName
'	gbobjExcelRecordSet.fields("MailTo") = gbMailTo
'	If Err.Number <> 0 Then
'		Err.Number = 0
'		Err.Description = ""
'	End If
'Call CloseExcelADODBConnection()

Call ReleasePublicVariables()

'***************************************************
'*               PUBLIC FUNCTIONS	*
'***************************************************
'_cr stands for Client's Request on portal
'Public Function OpenExcelADODBConnection(ByVal gbFilenameWithPath, gbSheetName)
'	On Error Resume Next
'	Set gbobjExcelConnection = CreateObject("ADODB.Connection")
'	gbobjExcelConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & gbFilenameWithPath & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;"";"
'	Call WaitTime(1)
'	
'	'Custor Type for VbScript to handle Excel
'	Const adOpenStatic = 3			''---- CursorTypeEnum Values ----
'	Const adLockOptimistic = 3		''---- LockTypeEnum Values ----
'	Const adCmdText = "&H0001"		''---- CommandTypeEnum Values ----
'	
'	sql_text="Select * FROM [" & gbSheetName & "$]"
'	
'	' create RecordSet
'	Set gbobjExcelRecordSet = CreateObject("ADODB.Recordset")
'	
'	'Execute SQL and store results in reocrdset'
'	gbobjExcelRecordSet.Open sql_text , gbobjExcelConnection, adOpenStatic, adLockOptimistic, adCmdText
'	Call WaitTime(1)
'End Function

Public Function OpenExcelADODBConnection(ByVal gbFilenameWithPath, ByVal gbSheetName)
	On Error Resume Next
	Set gbobjExcelConnection = CreateObject("ADODB.Connection")
	gbobjExcelConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & gbFilenameWithPath & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;"";"
	Call WaitTime(1)
	
	'Custor Type for VbScript to handle Excel
	Const adOpenStatic = 3			''---- CursorTypeEnum Values ----
	Const adLockPessimistic = 2		''---- LockTypeEnum Values ----
	Const adCmdText = "&H0001"		''---- CommandTypeEnum Values ----
	
	sql_text="Select * FROM [" & gbSheetName & "$]" 
	
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

Public Function OpenWorkBook(Byval ExcelFilePath)
	On Error Resume Next
	Set gbobjExcel = CreateObject("Excel.Application")
'	Set gbobjWorkBook = gbobjExcel.Workbooks.Open("C:\Users\automationqateam\Project\venv2\static_cdn\data\RMData.xls")
	Set gbobjWorkBook = gbobjExcel.Workbooks.Open(ExcelFilePath)

'	Set gbobjWorkBook = gbobjExcel.Workbooks.Open("\\sgw-filesvr3\GDC\GDC_Team\QA\ECO-Product-Docs\Automation\AutomationPortal\QCWorkFlow\AutomationPortal_Data.xlsx")
	set gbobjsheet=gbobjWorkBook.sheets(1)
	Call WaitTime(2)	
End Function

Public Function SaveAndCloseWorkBook()
	On Error Resume Next
	gbobjWorkBook.Save
	Call WaitTime(2)	
	gbobjWorkBook.Close
	Call ReleaseExcelObject()
	
	If Err.Number <> 0 Then
		Err.Number = 0
		Err.Description = ""
	End If
	Call WaitTime(2)	
End Function

Public Function ReleaseExcelObject()
	On Error Resume Next	
	' Release Excel Object
	 gbobjWorkBook.Close
	 gbobjExcel.quit
End Function

Public Function WaitTime(ByVal Seconds)
	 On Error Resume Next
	StartTime = Timer
	While Timer - StartTime < Seconds
	Wend

End Function	

Public Function ReleasePublicVariables()
	On Error Resume Next
	gbClientIPAddress = Null
	gbDateTime = Null
	gbTestSetFolderPath = Null
	gbALM_NewTestSetName = Null
	gbMailTo = Null
End Function
'***************************************************