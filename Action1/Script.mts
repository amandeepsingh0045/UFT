'########################################################################################################################################################################################################################
'Test Tool Version =UFT14.50
'Main Project Name = Hotel Booking
'Authors = Amandeep Singh(892428)

'Date Created = 03/03/2021

'########################################################################################################################################################################################################################
Dim strFolderPath,GetFolderPath,strFunctionallibPath,strObjRepositoryPath,CheckInDate,CheckOutDate

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
strObjRepositoryPath = GetFolderPath & "\ObjectRepository"&"\ObjRepository.tsr"
RepositoriesCollection.Add(strObjRepositoryPath)
Reporter.ReportEvent micPass,"Load Repositories","Shared Repositories are added in run time"
'########################################################################################################################################################################################################################

'########################################################################################################################################################################################################################
'Business Test Case/Flow 
Call Launch_Site()
Call Navigate_To_Hotels()
Call Select_City()
Call Click_On_Calender()


'Logic for Selecting Date 
'########################################################################################################################################################################################################################

	CheckIndate=cdate(Datatable.Value("CheckIn_Date","Global"))
	CheckOutDate=cdate(Datatable.Value("Checkout_Date","Global"))

	Select_Date(CheckIndate)		'Selecting Check In Date
	Select_Date(CheckOutDate)		'Selecting Check Out Date


'########################################################################################################################################################################################################################

Call Navigate_To_Hotel_Results()
Call Validating_City()
Call Validating_Check_In_Date()
Call Validating_Check_Out_Date()
Call Filter_Price() 'THe website is dynamic so maximum price also changes according to date so we tried keeping to close to 10000(Consulted with BU internal trainer)
Call Filter_Rating()
Call Filter_Amenities()
Call Print_Hotel_Names()
Call Capture_Screenshot_Of_Results()
Call Close_Browser()

 @@ hightlight id_;_9897126_;_script infofile_;_ZIP::ssf56.xml_;_
'########################################################################################################################################################################################################################
'End Of Function
'########################################################################################################################################################################################################################
