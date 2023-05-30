'################################################################################################################################################
'Test Description          	 	: Find the Interest Amount for current year
'Test Tool Version               	: UFT 14.5
'Return Type                  		: Null
'Author                                      : Amandeep Singh(892428)
'Date Created                    	: 03/06/2021

'################################################################################################################################################
'Declaring Variables
Option Explicit
Dim strFolderPath,GetFolderPath
Dim strFunctionallibPath
Dim strObjRepoistoryPath
'################################################################################################################################################
'Getting FolderPath
    strFolderPath = Environment.Value("TestDir")
    GetFolderPath = strFolderPath

'################################################################################################################################################
'LoadAllLibraries
    
    strFunctionallibPath = GetFolderPath & "\FunctionalLibrary"&"\CommonFunctions.vbs"
    LoadFunctionLibrary(strFunctionallibPath)
    Reporter.ReportEvent micPass,"Load Functional Library","Functional Libraries are added in Run Time"

'################################################################################################################################################
'LoadRepositories
    
    strObjRepoistoryPath = GetFolderPath & "\ObjectRepository"&"\RepositoryActionmerged_Final.tsr"    
    RepositoriesCollection.Add(strObjRepoistoryPath)
    Reporter.ReportEvent micPass,"Load Repositories","Shared Repositories are added in Run Time"

'################################################################################################################################################
'Test Description : Find the EMI for Car with price of 15 Lac, Interest rate of 9.5% & Tenure 1 year; Display the interest amount & principal amount for one month

Call Navigate_to_Car_Loan_EMI_Calculator()						'Launch Application and Navigate to Car Loan EMI Calculator
Call Enter_Loan_Amount_For_Car_Loan()							'Enter Car Loan Amount
Call Enter_Loan_Tenure_For_Car_Loan()							'Enter Car Loan Tenure
Call Enter_Loan_Interest_For_Car_Loan()							'Enter Car Loan ROI
Call Click_Calculate_Car_Loan(GetFolderPath)						'Calculate Car Loan Emi
Call Printing_Car_Loan_EMI_For_One_Month()					'Printing Car Loan EMI For One Month
Call Printing_First_Month_Principal_and_Interest_For_Car_Loan()	'Printing First Month Principal and Interest For Car Loan
Call Close_Browser()												'Closing the Browser

'====================================================================================================================================================
'Test Description : From Menu, pick Loan Calculator and under EMI calculator, do all UI check for text box & scales;

Call Navigate_to_Car_Loan_EMI_Calculator()	'Launch Application and Navigate to Car Loan EMI Calculator
'Validate Entered Data
Call Enter_Loan_Amount_For_Car_Loan()		'Enter Car Loan Amount
Call Validate_Loan_Amount_Text_Box()		'Validating Car Loan Amount Text Box
Call Enter_Invalid_Loan_Amt()				'Entering Invalid Data For Loan Amount
Call Validate_Loan_Amount_Text_Box()		'Validate Loan Amount Text Box With Invalid Data
'---------------------------------------------------------------------------------------------------------------------------------------
'Validating Car Loan Tenure Text Box
Call Enter_Loan_Tenure_For_Car_Loan()		'Enter Car Loan Tenure
Call Validate_Loan_Tenure_Text_Box()			'Validating Car Loan Tenure Text Box
Call Enter_Invalid_Loan_Tenure()				'Entering Invalid Loan Tenure
Call Validate_Loan_Tenure_Text_Box()			'Validate Loan Tenure Text Box With Invalid Data
'---------------------------------------------------------------------------------------------------------------------------------------
'Validating Car Loan Interest Text Box
Call Enter_Loan_Interest_For_Car_Loan()							'Enter Car Loan ROI
Call Validate_Loan_Interest_Text_Box()							'Validating Car Loan Interest Text Box
Call Validate_Loan_Interest_Text_Box_With_Boundary_Values()		'Validate Loan Interest Test Box Display Error Message With Boundary Values
Call Enter_Invalid_Loan_Interest()									'Enter Invalid Car Loan ROI
Call Validate_Loan_Interest_Text_Box()							'Validate Loan Interest Test Box With InValid Data()
'---------------------------------------------------------------------------------------------------------------------------------------
'Calculate Car Loan Emi
Call Click_Calculate_Car_Loan(GetFolderPath)						'Click Calculate Button
Call Validate_EMI_Per_Month()									'check EMI per month'check EMI per month
Call Print_Page_Provided_Results()								'Printing Page displayed or calculated Values
Call Close_Browser()												'Closing the Browser

'====================================================================================================================================================
'Test Description :  From Menu, pick Home Loan EMI Calculator, fill relevant details & extract all the data from  year on year table & store in excel

Call Launch_Home_Application()								'To launch Application in Browser
Call Navigate_to_EMI_calc_using_Menu()						'Navigate to EMI  calc using Menu
Call Navigate_to_Home_Loan_EMI_Calculator()				'Navigate to Home Loan EMI Calculator
Call Enter_Loan_Amount_For_Home_Loan()					'Enter Home Loan Amount
Call Enter_Loan_Tenure_For_Home_Loan()					'Enter Home Loan Tenure
Call Enter_Loan_Interest_For_Home_Loan_and_Calculate()		'Enter Home Loan ROI
Call Click_Calculate_Home_Loan(GetFolderPath)				'Calculate Home Loan Emi
Call Printing_Home_Loan_EMI_For_One_Month()				'Printing Home Loan EMI For One Month
Call Getting_all_the_data_from_year_on_year_table()			'Getting all the data from  year on year table 
DataTable.Export GetFolderPath & "\My_Data"&"\Data.xlsx"	'Exporting Data to Excel File
Call Close_Browser()											'Closing the Browser
'############################################################################################################################
'############################################################################################################################




