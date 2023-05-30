'With Framework
'########################################################################################################################################################################################################################
'Test Tool Version =UFT14.50
'Mini Project Name = Enquiry Of Building Project
'Author = Amandeep Singh(892428)
'Date Created = 26/02/2021

'########################################################################################################################################################################################################################
'GetFolderPath

	strFolderPath=Environment.Value("TestDir")
	GetFolderPath=strFolderPath
	
'########################################################################################################################################################################################################################

'########################################################################################################################################################################################################################
'LoadAllLibraries

strFunctionallibPath = GetFolderPath & "\FunctionalLibrary"&"\CommonFunctions.vbs"
LoadFunctionLibrary(strFunctionallibPath)
Reporter.ReportEvent micPass," Load Functional Library","Functional Library are added in run time"

'########################################################################################################################################################################################################################

'########################################################################################################################################################################################################################

'LoadRepositories
strObjRepoistoryPath = GetFolderPath & "\ObjectRepository"&"\ObjRepository.tsr"
RepositoriesCollection.Add(strObjRepoistoryPath)
Reporter.ReportEvent micPass,"Load Repositories","Shared Repositories are added in run time"

'########################################################################################################################################################################################################################

'########################################################################################################################################################################################################################
'Business Test Case/Flow 
Call LaunchSite()
Call NavigateToContactPage()
Call EnterContactDetails()
Call ExtractProjectDetailsPrintDetailsInConsole()
Call SubmitForm()
Call GoToHomePage()
Call NavigateToVillasPage()
Call DisplayProjectsWithMoreThanTenUnits()
'Possession Date Should be on or before December Month of current Year(********Feature Not Available On Site********)(Consulted with the BU Internal trainer )
Call CloseBrowser()

'########################################################################################################################################################################################################################
