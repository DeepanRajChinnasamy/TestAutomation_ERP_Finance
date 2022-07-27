'********************************************************************************************************************
'Business Component Header 
'******************************************************************************************************************** 
'Objective: The objective of this test automation component is to validate the GED eFashion error message validation
'Functional Area:' 
'UFT Version: ' 
'Created By:       
'Created On:       
'Description:   
' This business component will :
'Change Log:
'Modified On   Modified By     Comments
'Assumptions: 
' 1/. User should have access to Source and Target system
' 2/. SAP - Scripting to be enabled on the Server side for SAP ECC system
'************************************************************************************************************************
'Varaiable Declarations
strExcelPath = "\\AUS-WNASCRMP-03\Share\02. Test Automation\07. Dunning\00.Input\TD_Inputs.xlsx"
strExcelSheetName = "TestData"
strSheetName = "Login"
Set SAPWindowObject = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
Set SAPWindowObject1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
'Create Excel object 
Call CreateExcelObject(strExcelPath,strSheetName,intExcelRowCount,intExcelColumnCount,objExcelSheet,objExcelWorkbook)
'Read all Values from Login Sheet
Call ReadAllValuesFromInputExcel(objExcelSheet)
'Close the current excel object
Call CloseExcelObject(objExcelWorkbook,objExcelSheet)

Call CreateExcelObject(strExcelPath,strExcelSheetName,intExcelRowCount,intExcelColumnCount,objExcelSheet,objExcelWorkbook)
'Read all Values from Login Sheet
Call ReadAllValuesFromInputExcel(objExcelSheet)

'Iterating with respect to customer count
For intCustomerIterator = 1 To  UBOUND(arrCustomerID)
	SAPWindowObject.SAPGuiOKCode("guicomponenttype:=35","name:=okcd").Set "/nF150"
	SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click
'	Call PressEnter()
	SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click
	SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=F150V-LAUFD").Set arrRunOn(intCustomerIterator)
	intValidationIdenty =  RandomNumber(1000,9999)
	SAPWindowObject.SAPGUiEdit("guicomponenttype:=32","name:=F150V-LAUFI").Set  intValidationIdenty
	'Navigate to tab
	SAPWindowObject.SAPGuiTabStrip("name:=DU_TABSTRIP").Select "Parameter"  
	SAPWindowObject.SAPGUiEdit("guicomponenttype:=32","name:=F150V-AUSDT").Set  arrRunOn(intCustomerIterator)
	SAPWindowObject.SAPGUiEdit("guicomponenttype:=32","name:=F150V-GRDAT").Set  arrRunOn(intCustomerIterator)
	SAPWindowObject.SAPGUiEdit("guicomponenttype:=32","name:=RNG_BKRS-LOW").Set arrCompanyCode(intCustomerIterator)
	SAPWindowObject.SAPGUiEdit("guicomponenttype:=32","name:=RNG_SELC-LOW").Set  arrCustomerID(intCustomerIterator)
	SAPWindowObject.SAPGuiTabStrip("name:=DU_TABSTRIP").Select "Additional Log"  
	SAPWindowObject.SAPGUiEdit("guicomponenttype:=32","name:=RNG_LOGC-LOW").Set   arrCustomerID(intCustomerIterator)
	wait 2
	SAPWindowObject.SAPGUiButton("guicomponenttype:=40","enabled:=True","tooltip:=Save.*").Click
	SAPWindowObject.SAPGuiTabStrip("name:=DU_TABSTRIP").Select "Status" 
'	Call ClickButton("Individual dunning notice   (Shift+F1)")
	SAPWindowObject.SAPGUiButton("guicomponenttype:=40","tooltip:=Schedule dunning run   (F7)").Click
	SAPWindowObject1.SAPGUiEdit("guicomponenttype:=32","name:=USR01-SPLD").Set "LOCL"
	SAPWindowObject1.SAPGUiButton("tooltip:=Continue   \(Enter\)","text:=Continue").Click
	SAPWindowObject1.SAPGuiCheckBox("guicomponenttype:=42","name:=F150V-XSTRF").Set "ON"
	SAPWindowObject1.SAPGUiButton("tooltip:=Execute.*").Click
	Call PressEnter()
	strText =  SAPWindowObject.SAPGUiEdit("guicomponenttype:=31","name:=F150V-STEXT","value:=.*generated.*").GetRoProperty("value")
	If InStr(strText,"1") Then
		Call SavePDFToFolder()
	End If
Next
Call CloseExcelObject(objExcelWorkbook,objExcelSheet)
'****************************************************************************************************************
'Name of the Function   :SavePDFToFolder(prm_RunOn,prm_SenderName)
'Author     :DeepanRaj
'Description    :Function to trigger mail batch job in WE15
'Input Parameters    :prm_RunOn,prm_SenderName
'Output Parameters      :NIL
'Creation Date : 26 June 2022
'****************************************************************************************************************
'Function SavePDFToFolder(prm_RunOn,prm_SenderName)
'****************************************************************************************************************
Public Function SavePDFToFolder(prm_RunOn,prm_SenderName)
	Set SAPWindowObject = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
	Set SAPWindowObject1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
	SAPWindowObject.SAPGuiOKCode("guicomponenttype:=35","name:=okcd").Set "/nSOST"
	SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click
	SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAB1").Select "Period"
	SAPWindowObject.SAPGuiEdit("guicomponenttype:=31","name:=G_MAXSEL").Set " "
	SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=REFRICO2").Click
	wait 2
	SAPWindowObject.SAPGuiGrid("guicomponenttype:=201","title:= Send Requests.*").SelectAll()
	SAPwindowObject.SAPGuiToolbar("guicomponenttype:=204","name:=shell").PressButton "&MB_FILTER"
'	SAPWindowPopUpObject.SAPGuiEdit("guicomponenttype:=32","name:=%%DYN006-LOW").Set intSAPDate
	SAPWindowObject1.SAPGuiEdit("guicomponenttype:=32","name:=%%DYN004-LOW").Set prm_SenderName
'	SAPWindowPopUpObject.SAPGuiEdit("guicomponenttype:=32","name:=%%DYN005-LOW").Set StrSystemMailId
	SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","tooltip:=Execute.*").Click
	wait 2
	SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=REFRICO2").Click
	SAPWindowObject.SAPGuiGrid("guicomponenttype:=201","title:= Send Requests.*").SelectAll()
	SAPwindowObject.SAPGuiToolbar("guicomponenttype:=204","name:=shell").PressButton "SHOW"
	SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=SO33_TAB1").Select "Attachments"
	SAPwindowObject.SAPGuiToolbar("guicomponenttype:=204","name:=shell").PressButton "EXPO"
	strPath = "\\AUS-WNASCRMP-03\Share\03. Test Execution\01. R1\DunningPDF\"
	SAPWindowObject1.SAPGuiEdit("guicomponenttype:=32","name:=DY_PATH").Set strPath
	SAPWindowObject1.SAPGuiButton("guicomponenttype:=40","tooltip:=Create New File.*").Click
	
End Function







'****************************************************************************************************************
'Name of the Function   :ReadPDFContentToTextFileSSA(objFSO,objFSOTxt,strPDFFilePath,strTextPath,strTempFileFSOContent) 
'Author     :DeepanRaj
'Description    :Function to ReadPDF to Text file
'Input Parameters    :strPDFFilePath,strTextPath,strMailID
'Output Parameters      :strTempFileFSOContent
'Creation Date :
'****************************************************************************************************************
'Function ReadPDFContentToTextFileSSA(objFSO,objFSOTxt,strPDFFilePath,strTextPath,strTempFileFSOContent) 
'****************************************************************************************************************
Public Function ReadPDFContentToTextFileSSA(objFSO,objFSOTxt,strPDFFilePath,strTextPath,strTempFileFSOContent,intBillingDocumentNumber) 
	SystemUtil.Run strPDFFilePath
	Window("regexpwndtitle:=.*Adobe.*").Activate
	strPDFText =Window("regexpwndtitle:=.*Adobe.*").GetVisibleText
	Set WshShell = CreateObject("WScript.Shell")  
	WshShell.SendKeys "%{F}"        ' Press ALT+F
	WshShell.SendKeys "{h}"              ' Send the corresponding letter which is underlined, for 'save' under 'File' Menu. This opens a 'save a copy' dialog box
	WshShell.SendKeys "{x}" 
	intDunningNumberPDF= intDunningNumber&".*"
	Window("regexpwndtitle:=.*Adobe.*").Dialog("regexpwndtitle:=Save As").WinEdit("regexpwndclass:=Edit","text:="&intDunningNumberPDF).Set strTextPath
	Window("regexpwndtitle:=.*Adobe.*").Dialog("regexpwndtitle:=Save As").WinComboBox("regexpwndclass:=ComboBox","regexpwndtitle:=Text.*","index:=1").Select "All Files (*.*)"
	Window("regexpwndtitle:=.*Adobe.*").Dialog("regexpwndtitle:=Save As").WinComboBox("regexpwndclass:=ComboBox","regexpwndtitle:=All Files.*","index:=1").Select "Text (Accessible) (*.txt)"
	Window("regexpwndtitle:=.*Adobe.*").Dialog("regexpwndtitle:=Save As").WinButton("text:=&Save").Click
	WshShell.SendKeys "{x}"
	Window("regexpwndtitle:=.*Adobe.*").WinObject("text:=AVPageView","index:=0").click
	Set WshShell = CreateObject("WScript.Shell") 
	WshShell.SendKeys "%{F}"        ' Press ALT+F
	WshShell.SendKeys "{x}"              ' Send the corresponding letter which is underlined, for 'save' under 'File' Menu. This opens a 'save a copy' dialog box
	Set WshShell = Nothing
	Set objFSOFile = CreateObject("Scripting.FileSystemObject")
	wait(5)
	Set objFSOFileTxt = objFSOFile.OpenTextFile(strTextPath,1)
	wait(5)
	strTempFileFSOContent = objFSOFileTxt.ReadAll()
	objFSOFileTxt.Close
	Set objFSOFile = Nothing
	Set objFSOFileTxt = Nothing
End Function
'****************************************************************************************************************
'End Function ReadPDFContentToTextFileSSA(objFSO,objFSOTxt,strPDFFilePath,strTextPath,strTempFileFSOContent) 
'****************************************************************************************************************
















'******************************************************************************
'Function Name -LoginApp
'Description -Login to application
'Input - objWd
'Output - null
'Created by - Ankita Desai
'Date -02/23/2022
'******************************************************************************
Public Function LoginApp(strLink,strUserName,strPassword)

	Call CloseBrowser("Firefox")
	'open Application
	'SystemUtil.Run "Firefox.exe",Environment("ENV_URL")
	SystemUtil.Run "chrome.exe",strLink

	Set Obj=Browser("Sign in to your account").Page("Sign in to your account")
	Obj.Sync
	IF Browser("Sign in to wiley").Page("Sign in to wiley").Link("Microsoft Login").Exist(15) then
		Browser("Sign in to wiley").Page("Sign in to wiley").Link("Microsoft Login").Click
	End  IF
	
	If Browser("title:= Sign in to your account.*").Page("title:=Sign in to your account").WebElement("xpath:=//DIV[@id='otherTile']").Exist(5) Then
		Browser("title:= Sign in to your account.*").Page("title:=Sign in to your account").WebElement("xpath:=//DIV[@id='otherTile']").Click
	End If

	If Browser("title:= Sign in to your account.*").Page("title:=Sign in to your account.*").WebEdit("name:=loginfmt.*").Exist(5) Then	
		Browser("title:= Sign in to your account.*").Page("title:=Sign in to your account.*").WebEdit("name:=loginfmt.*").Set strUserName
		Browser("title:= Sign in to your account.*").Page("title:=Sign in to your account.*").WebButton("name:=Next").Click
	End If

	If Browser("title:= Sign in to your account.*").Page("title:=Sign in to your account.*").WebEdit("acc_name:=Enter the password.*").Exist(5) Then
		Browser("title:= Sign in to your account.*").Page("title:=Sign in to your account.*").WebEdit("acc_name:=Enter the password.*").Set  strPassword
		wait 1
		Browser("title:= Sign in to your account.*").Page("title:=Sign in to your account.*").WebButton("name:=Sign in").Click
	End  IF

	If Browser("title:= Sign in to your account.*").Page("title:=Sign in to your account.*").WebButton("name:=Yes").Exist(5) Then
		Browser("title:= Sign in to your account.*").Page("title:=Sign in to your account.*").WebButton("name:=Yes").Click
	End  IF
	wait 1
	Browser("Sign in to wiley").Page("wiley-test - STEP by STIBO").Sync
	
	If Browser("Sign in to wiley").Page("wiley-test - STEP by STIBO").Link("Wiley Test Web UI").Exist(15) Then
		Browser("Sign in to wiley").Page("wiley-test - STEP by STIBO").Link("Wiley Test Web UI").Click
	End If
	
'	IF    Browser("Journal Creation WF V3").Page("Journal Creation WF V3").Exist(10) then 
'		Call WriteDatatToWord (objWd,"STEPBO launchpad screen should be displayed","STEPBO launched Successfully","PASS","")
'		Reporter.ReportEvent micPass,"STEPBO launched Successfully","STEPBO launchpad screen should be displayed"
'	Else
'		 Call WriteDatatToWord (objWd,"STEPBO launchpad screen should be displayed","STEPBO Failed to Launch","FAIL","")
'		  Reporter.ReportEvent micFail,"STEPBO Failed to Launch","Verify Login data"
'		 ExitAction
'	ENd  IF
'	If Environment("CAPTURESCREENSHOT") Then Call CaptureScreenshot(objWd,"Login") End IF
End Function






Public Function ReadAllValuesFromInputExcel(objExcelSheet)
	Set excelUsedRange=objExcelSheet.usedrange
	excelRowCount = excelUsedRange.rows.count
	excelColumnCount=excelUsedRange.Columns.count
	For intColumnLoop = 1 To excelColumnCount
		strFlagForFirstForExit = ""
		If strFlagForFirstForExit = "YES" Then
			Exit For 		
		End If
		For inRowLoop = 2 To excelRowCount - 1
			strFinalRowValue = objExcelSheet.Cells(inRowLoop,1).Value
	    		If strFinalRowValue = "END" Then
	    			strFlagForFirstForExit = "YES"
	    			Exit For
	   	 	End If
	    		strIterationCellValue = objExcelSheet.Cells(inRowLoop,intColumnLoop).Value
	   		strEntireColumnValue = strEntireColumnValue & ";" & strIterationCellValue
		Next
		strExcelColumnName = objExcelSheet.Cells(1,intColumnLoop).Value
		strEntireColumnValue = Replace(strEntireColumnValue,VBCR&VBLF,"")
		strEntireColumnValue = Replace(strEntireColumnValue,VBCR,"")
		strEntireColumnValue = Replace(strEntireColumnValue,VBLF,"")
		strTempExcelColumnName = Replace(strExcelColumnName,"ID_","str")
		Execute strTempExcelColumnName & " =  " & Chr(34) & strEntireColumnValue & Chr(34) 
		Execute "arr" &strExcelColumnName & " = " & "Split(" & strTempExcelColumnName & Chr(44) & Chr (34) & Chr(59) & Chr(34) & ")"
		strEntireColumnValue = ""
	Next
End Function


Public Function CloseExcelObject(objExcelWorkbook,objExcelObject)
	On Error Resume Next
	objExcelWorkbook.Save
	objExcelWorkbook.Close	
	Set objExcelWorkbook = Nothing 
	Set objExcelObject = Nothing
End Function



