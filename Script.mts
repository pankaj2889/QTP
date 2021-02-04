'#################################################################################################
'# Test Case: NV-MT-001-ValidateAllMaintenanceLeftNavigationalLinks
'# Test Case #: 9
'# Steps:
'#  1. Open browser and navigate to web ordering application URL
'#  2. Check all the links on homepage
'#	3. Close the browser
'# Owner: Nimish Srivastava
'# Created on: 12th April 2017
'# Data files: 
'# Libraries - Configuration.qfl, funcGeneral.qfl, funcApplication.qfl
'# Data - MasterTestData.xls, Config.xls and ObjectRepository.tsr
'# Preconditions: 
'# Postconditions: 
'#################################################################################################

Reporter.Filter = dictConfg.Item("intReporterFilter")

''## Data Initialization
'getTestData "SimpleWithHistory","20"

'strURL = dictData.Item("strScanForecastingURL")
On error resume next
Set TSTest = QCUtil.CurrentTestSetTest
strURL = TSTest.Field("TC_USER_01")

If strURL = "" Then
	strURL = dictData.Item("strScanForecastingURL")
End If


strBrowser = dictData.Item("Browser")


'*************************************************************************************************************************************************************************************************************************************
''@BEGIN Step 1 -  Open browser and navigate to scan forecasting application URL
'*************************************************************************************************************************************************************************************************************************************
strDescription = "Open '"& Ucase(strBrowser) & "' browser and navigate to Scan Forecasting application URL."
strExpectedResult = "Scan Forecasting home page should be opened."
Environment.Value("intStepNo") = Environment.Value("intStepNo") + 1
strStepsToReproduce = strStepsToReproduce & vbnewline & Environment.Value("intStepNo") & ") " & strDescription

StepRC = OpenScanForecastingURL(strURL)

If StepRC Then
	StepReporter Environment.Value("intStepNo"),micPass,strDescription,strExpectedResult,additionalInfo
Else
	StepReporter Environment.Value("intStepNo"),micFail, strDescription,strExpectedResult,errDesc
End If

'***********************************************************************************************************************************************************************************************************************************
Dim oDesc
Set oDesc = Description.Create
oDesc("micclass").value = "Link"

'Find all the Links under Maintenance
Set obj = Browser("ScanForecasting").Page("Home Page").WebElement("divReportsLinks").ChildObjects(oDesc)

For i = 0 to obj.Count - 1
	Set obj = Browser("ScanForecasting").Page("Home Page").WebElement("divReportsLinks").ChildObjects(oDesc)
	strLinkText = trim(obj(i).GetROProperty("innerhtml"))
	strDestinationPageTitle = dictLinkPageMapping.Item(strLinkText)
	
	'***********************************************************************************************************************************************************************************************************************************
	''@BEGIN Step 2 -  Click on the link
	'*************************************************************************************************************************************************************************************************************************************
	strDescription = StrFormat("Click on the link '{0}'",array(strLinkText))
	strExpectedResult = StrFormat("'{0}' page should open.",array(strDestinationPageTitle))
	Environment.Value("intStepNo") = Environment.Value("intStepNo") + 1
	strStepsToReproduce = strStepsToReproduce & vbnewline & Environment.Value("intStepNo") & ") " & strDescription
	
	''Click on the link
	obj(i).Highlight
	obj(i).click
	
	StepRC = VerifyDestinationPage(strLinkText,strDestinationPageTitle)
	
	If StepRC Then
		StepReporter Environment.Value("intStepNo"),micPass,strDescription,strExpectedResult,StrFormat("'{0}' page is opened on clicking link '{1}'.",array(strDestinationPageTitle,strLinkText))
	Else
		StepReporter Environment.Value("intStepNo"),micFail, strDescription,strExpectedResult,errDesc
	End If

	'***********************************************************************************************************************************************************************************************************************************
	''@BEGIN Step 3 -  Navigate to home page
	'*************************************************************************************************************************************************************************************************************************************
	strDescription = "Navigate to home page."
	strExpectedResult = "Scan Forecasting homepage should be opened."
	Environment.Value("intStepNo") = Environment.Value("intStepNo") + 1
	strStepsToReproduce = strStepsToReproduce & vbnewline & Environment.Value("intStepNo") & ") " & strDescription
	
	StepRC = NavigateToHomePage()
	
	If StepRC Then
		StepReporter Environment.Value("intStepNo"),micPass,strDescription,strExpectedResult,"Scan Forecasting homepage is opened successfully."
	Else
		StepReporter Environment.Value("intStepNo"),micFail, strDescription,strExpectedResult,errDesc
	End If
Next

'***********************************************************************************************************************************************************************************************************************************

CloseBrowserWindow
