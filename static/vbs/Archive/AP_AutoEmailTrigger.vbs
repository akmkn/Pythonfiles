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

'2 is default Excel Row number
gbExcelRowNumber = "2"
Call QCWorkFlow(gbExcelRowNumber)


Public Function OpenWorkBook()
	Set gbobjExcel = CreateObject("Excel.Application")
'	Set gbobjWorkBook = gbobjExcel.Workbooks.Open("C:\Users\automationqateam\Project\venv2\static_cdn\data\RMData.xls")
	Set gbobjWorkBook = gbobjExcel.Workbooks.Open(""C:\Users\automationqateam\Project\venv2\static_cdn\data\RMData.xls")

'	Set gbobjWorkBook = gbobjExcel.Workbooks.Open("\\sgw-filesvr3\GDC\GDC_Team\QA\ECO-Product-Docs\Automation\AutomationPortal\QCWorkFlow\AutomationPortal_Data.xlsx")
	set gbobjsheet=gbobjWorkBook.sheets(1)
	Call WaitTime()	
End Function


Public Function QCWorkFlow(ByVal gbExcelRowNumber)
	On Error Resume Next
	'msgbox "Excel Row: " & gbExcelRowNumber
	' Establish QC Connection
	gbQCConnectionResult = ConnectToQC()
	
	If gbQCConnectionResult = "QCConnectionPass" Then
		
		gbExecuteScript = "YES"
		gbTestPlanFolderPath = "[QualityCenter]Subject\Automation\ECO\AutomationPortalScript"
		gbTestScriptNameToExecute = "AP_AutoEmailScript"
		gbTestSetFolderPath = "Root\ECO Automation\Product-Execution\AutomationPortal"
		gbNewTestSetName = "AP_AutoEmail"
		' get machine name from workbook
		Call OpenWorkBook()
		gbRemoteMachineName = gbobjsheet.cells(2,6).value
		Call SaveAndCloseWorkBook()
		
		gbExecuteScript = ExecuteAutoEmail(gbExecuteScript, gbTestPlanFolderPath, gbTestScriptNameToExecute, gbTestSetFolderPath, gbNewTestSetName, gbRemoteMachineName, gbExcelRowNumber, "", "")
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
		msgbox "QCConnectionPass for Auto Email"		
'		Call SaveAndCloseWorkBook()
		
	Else
		ConnectToQC = "QCConnectionFail"
'		'Set Running Status
'		Call OpenWorkBook()
'		gbobjsheet.cells(gbExcelRowNumber,9).value = "QCConnectionFail"	
		msgbox "QCConnectionFail for Auto Email"		
'		Call SaveAndCloseWorkBook()
	End If
	On Error Goto 0
	On Error Resume Next
' 	msgbox ElapsedTime
	
End Function


	
Public Function ExecuteAutoEmail(ByVal gbExecuteScript, ByVal gbTestPlanFolderPath, ByVal gbTestScriptNameToExecute, ByVal gbTestSetFolderPath, ByVal gbTestSetName, ByVal gbRemoteMachineName, ByVal gbExcelRowNumber, ByRef rt_IsStart, ByRef rt_IsStop)
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
		    	msgbox "Launching QTP on Local Machine "
		    	' Create Object for QTP for Remote machine
'		    	Set qtApp = CreateObject("QuickTest.Application",gbRemoteMachineName)   
				
				'create object of QTP for local machinee
		         Set qtApp = CreateObject("QuickTest.Application") 
		         
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

'
''tmp function ReadExcel
''***************************
'Dim ReturnValue
'myXlsFile = "C:\Ashish\Equinix\Automation Portal- Interface - vbs uft\QC Workflow Script vbs\QC Workflow VBS\AutomationPortal_Data_V1.xls"
'mySheet = "RM"
'my1stCell = "A1"
'myLastCell = "O100"
'blnHeader = true
'ReadExcelRowNumber = 2
'ColumnName = "MailTo"
'
'ReturnValue = ReadExcel_SpecificRowColmnData(myXlsFile, mySheet, my1stCell, myLastCell, blnHeader, ReadExcelRowNumber, ColumnName)
'msgbox ReturnValue
'a=1
'
'Public Function ReadExcel_SpecificRowColmnData(ByVal myXlsFile, ByVal mySheet, ByVal my1stCell, ByVal myLastCell, ByVal blnHeader, ByVal ReadExcelRowNumber, ByVal ColumnName)
'' Function :  ReadExcel
'' This function reads data from an Excel sheet without using MS-Office
''
'' Arguments:
'' myXlsFile   [string]   The path and file name of the Excel file
'' mySheet     [string]   The name of the worksheet used (e.g. "Sheet1")
'' my1stCell   [string]   The index of the first cell to be read (e.g. "A1")
'' myLastCell  [string]   The index of the last cell to be read (e.g. "D100")
'' blnHeader   [boolean]  True if the first row in the sheet is a header
''
'' Returns:
'' The values read from the Excel sheet are returned in a two-dimensional
'' array; the first dimension holds the columns, the second dimension holds
'' the rows read from the Excel sheet.
'
'
'    Dim arrData( ), i, j
'    Dim objExcel, objRS
'    Dim strHeader, strRange
'
'    Const adOpenForwardOnly = 0
'    Const adOpenKeyset      = 1
'    Const adOpenDynamic     = 2
'    Const adOpenStatic      = 3
'
'    ' Define header parameter string for Excel object
'    If blnHeader Then
'        strHeader = "HDR=YES;"
'    Else
'        strHeader = "HDR=NO;"
'    End If
'
'    ' Open the object for the Excel file
'    Set objExcel = CreateObject("ADODB.Connection")
'    ' IMEX=1 includes cell content of any format; 
'    
'    objExcel.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
'                  myXlsFile & ";Extended Properties=""Excel 12.0;IMEX=1;" & _
'                  strHeader & """"
'
'    ' Open a recordset object for the sheet and range
'    Set objRS = CreateObject("ADODB.Recordset")
'    strRange = mySheet & "$" & my1stCell & ":" & myLastCell
'    objRS.Open "Select * from [" & strRange & "]", objExcel, adOpenStatic
'
'    ' Read the data from the Excel sheet
'    i = 0
'    Do Until objRS.EOF
'        ' Stop reading when an empty row is encountered in the Excel sheet
'        If IsNull( objRS.Fields(0).Value ) Or Trim( objRS.Fields(0).Value ) = "" Then Exit Do
'        ' Add a new row to the output array
'        ReDim Preserve arrData(objRS.Fields.Count - 1, i)
'        ' Copy the Excel sheet's row values to the array "row"
'        ' IsNull test credits: Adriaan Westra
'        For j = 0 To objRS.Fields.Count - 1
'       		If j = ReadExcelRowNumber-1 Then     
'	            If IsNull(objRS.Fields(j).Value) Then
'	                arrData(j, i) = ""
'	            Else
'	                arrData(j, i) = Trim(objRS.Fields(j).Value)
'	                Return = Trim(objRS.Fields(ColumnName))
'	                Exit Do
'	            End If
'	    	End If
'        Next
'        ' Move to the next row
'        objRS.MoveNext
'        ' Increment the array "row" number
'        i = i + 1
'    Loop
'
'    ' Close the file and release the objects
'    objRS.Close
'    objExcel.Close
'    Set objRS    = Nothing
'    Set objExcel = Nothing
'
'    ' Return the results
'    ReadExcel_SpecificRowColmnData = Return
'   ' ReadExcel = arrData
'End Function
'***************************************************************************************************************'
''
'Function ReadExcel(ByVal myXlsFile, ByVal mySheet, ByVal my1stCell, ByVal myLastCell, ByVal blnHeader, ByVal ReadExcelRowNumber)
'' Function :  ReadExcel
'' This function reads data from an Excel sheet without using MS-Office
''
'' Arguments:
'' myXlsFile   [string]   The path and file name of the Excel file
'' mySheet     [string]   The name of the worksheet used (e.g. "Sheet1")
'' my1stCell   [string]   The index of the first cell to be read (e.g. "A1")
'' myLastCell  [string]   The index of the last cell to be read (e.g. "D100")
'' blnHeader   [boolean]  True if the first row in the sheet is a header
''
'' Returns:
'' The values read from the Excel sheet are returned in a two-dimensional
'' array; the first dimension holds the columns, the second dimension holds
'' the rows read from the Excel sheet.
'
'
'    Dim arrData( ), i, j
'    Dim objExcel, objRS
'    Dim strHeader, strRange
'
'    Const adOpenForwardOnly = 0
'    Const adOpenKeyset      = 1
'    Const adOpenDynamic     = 2
'    Const adOpenStatic      = 3
'
'    ' Define header parameter string for Excel object
'    If blnHeader Then
'        strHeader = "HDR=YES;"
'    Else
'        strHeader = "HDR=NO;"
'    End If
'
'    ' Open the object for the Excel file
'    Set objExcel = CreateObject("ADODB.Connection")
'    ' IMEX=1 includes cell content of any format; 
'    
'    objExcel.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
'                  myXlsFile & ";Extended Properties=""Excel 12.0;IMEX=1;" & _
'                  strHeader & """"
'
'    ' Open a recordset object for the sheet and range
'    Set objRS = CreateObject("ADODB.Recordset")
'    strRange = mySheet & "$" & my1stCell & ":" & myLastCell
'    objRS.Open "Select * from [" & strRange & "]", objExcel, adOpenStatic
'
'    ' Read the data from the Excel sheet
'    i = 0
'    Do Until objRS.EOF
'        ' Stop reading when an empty row is encountered in the Excel sheet
'        If IsNull( objRS.Fields(0).Value ) Or Trim( objRS.Fields(0).Value ) = "" Then Exit Do
'        ' Add a new row to the output array
'        ReDim Preserve arrData(objRS.Fields.Count - 1, i)
'        ' Copy the Excel sheet's row values to the array "row"
'        ' IsNull test credits: Adriaan Westra
'        For j = 0 To objRS.Fields.Count - 1
'            If IsNull(objRS.Fields(j).Value) Then
'                arrData(j, i) = ""
'            Else
'                arrData(j, i) = Trim(objRS.Fields(j).Value)
'                mailToFieldValue = Trim(objRS.Fields("MailTo"))
'            End If
'        Next
'        ' Move to the next row
'        objRS.MoveNext
'        ' Increment the array "row" number
'        i = i + 1
'    Loop
'
'    ' Close the file and release the objects
'    objRS.Close
'    objExcel.Close
'    Set objRS    = Nothing
'    Set objExcel = Nothing
'
'    ' Return the results
'    ReadExcel = arrData
'End Function
'
'***************************




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