'Get folder path for taking screenshot
Function GetFolderPath()
	GetFolderPath=Environment.Value("TestDir")
End Function

'the screenshot has unique name as it is based on time,used "now" function
Function fnl()
	'Capture Screenshot
	screen=now
	dt1=replace(now,"/","")
	dt2=replace(dt1,":","")
	fnl2=replace(replace(dt2," ",""),"AM","")
	fnl="capture" &fnl2
	
	
End Function
'Launches the site of ishahomes
Sub LaunchSite()
	While Browser("creationtime:=0").Exist(10)
	'Close the browser if already open 
	Browser("creationtime:=0").Close
	Wend
	'take data from Global
	strURL = DataTable.Value("URL","Global")
	
	SystemUtil.Run "chrome.exe",strURL
	'capture screenshot
	FilePath = GetFolderPath & "\ResultScreenshot"&"\"&fnl&".png"
	imgg=Browser("Buy Apartments,Flats&Luxury").Page("Buy Apartments,Flats&Luxury").CaptureBitmap(FilePath,True)
	
	'To get report event in result viewer
	If Browser("Buy Apartments,Flats&Luxury").Page("Buy Apartments,Flats&Luxury").Exist Then
		Reporter.ReportEvent micPass, "Verify Home Page","User Navigated to Home Page"
	else
		Reporter.ReportEvent micFail, "Verify Home Page","User NOT Navigated to Home Page" 
	End If
	
End Sub
'Navigate to the contact us page
Sub NavigateToContactPage()
	Browser("Buy Apartments,Flats&Luxury").Page("Buy Apartments,Flats&Luxury").Sync
	Browser("Buy Apartments,Flats&Luxury").Page("Buy Apartments,Flats&Luxury").Link("Contact us").Click
	
	'To get report event in result viewer
	Reporter.ReportEvent micPass, "Contact Us Page","User is taken to the Contact Us Page"
	
End Sub
'Enter contact details in the from from the global datasheet
Sub EnterContactDetails()
	Name=DataTable.Value("Name","Global")
	Email=DataTable.Value("Email","Global")
	Mobile=DataTable.Value("Phone","Global")
	STCode=DataTable.Value("STCode","Global")
	
	'To get report event in result viewer
	If Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").Exist Then
		Reporter.ReportEvent micPass, "Verify Contact Page","User Navigated to Contact Page"
	else
		Reporter.ReportEvent micFail, "Verify Contact Page","User NOT Navigated to Contact Page" 
	End If
	
	Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").Sync
	
	'Capture Screenshot
	FilePath = GetFolderPath & "\ResultScreenshot"&"\"&fnl&".png"
	imgg1 = Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").CaptureBitmap(FilePath,True)
	
	'Enter contact details
	Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebEdit("sell_do[form][lead][name]").Set Name
	Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebEdit("sell_do[form][lead][email]").Set Email
	'Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebEdit("sell_do[form][lead][phone]").Set "+91"&Mobile
	Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebEdit("sell_do[form][lead][phone]").Set "+"&STCode&Mobile
	Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebList("sell_do[form][lead][project_id").Select "Isha Code Field"
	
	'Capture Screenshot 
	FilePath = GetFolderPath & "\ResultScreenshot"&"\"&fnl&".png"
	imgg2 = Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebButton("Submit").CaptureBitmap(FilePath,True)
	'Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebButton("Submit").Click

End Sub

'Extract project details and display in the console
Sub ExtractProjectDetailsPrintDetailsInConsole()  
	
    no_of_project=Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebList("sell_do[form][lead][project_id").GetROProperty("items count")
    project_sites=Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebList("sell_do[form][lead][project_id").GetROProperty("all items")
    selected_project=Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebList("sell_do[form][lead][project_id").GetROProperty("selection")
    'Displays total number of available projects
	print "no of projects available "&no_of_project
    print "the projects are :"
    temp=Split(project_sites,";")
    For each i in temp
    'Displays all project sites from the webpage
	   print i
    Next
    'Displays the selected project in the console
    print "the selected project is :"&selected_project
End Sub
'Submit the contact details
Sub SubmitForm()

    Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebButton("Submit").Click
End Sub
'Navigate to the home page

Sub GoToHomePage()
	Browser("Buy Apartments,Flats&Luxury").Page("Thank you - IshaHomes").Sync
	
	'To get report event in result viewer
	If Browser("Buy Apartments,Flats&Luxury").Page("Thank you - IshaHomes").Exist Then
		Reporter.ReportEvent micPass, "Verify ThankYou Page","User Navigated to Thankyou Page"
	else
		Reporter.ReportEvent micFail, "Verify ThankYou Page","User NOT Navigated to Thankyou Page" 
	End If
	'Capture Screenshot
	FilePath = GetFolderPath & "\ResultScreenshot"&"\"&fnl&".png"
	imgg3 = Browser("Buy Apartments,Flats&Luxury").Page("Thank you - IshaHomes").CaptureBitmap(FilePath,True)
	
	Browser("Thank you - IshaHomes").Page("Thank you - IshaHomes").Link("back to Homepage").Click
	
End Sub
'Navigate to the Villas page 
Sub NavigateToVillasPage()
	Browser("Buy Apartments,Flats&Luxury").Page("Buy Apartments,Flats&Luxury").Link("Buy Villas").Click
End Sub


Sub DisplayProjectsWithMoreThanTenUnits()
	'To get report event in result viewer
	If Browser("Thank you - IshaHomes").Page("Villas for sale in OMR,").Exist Then
		Reporter.ReportEvent micPass, "Verify Villas for sale Page","User Navigated to  Villas for sale  Page"
	else
		Reporter.ReportEvent micFail, "Verify  Villas for sale  Page","User NOT Navigated to  Villas for sale Page" 
	End If
	'Capture Screenshot
	FilePath = GetFolderPath & "\ResultScreenshot"&"\"&fnl&".png"
	imgg4 = Browser("Thank you - IshaHomes").Page("Villas for sale in OMR,").CaptureBitmap(FilePath,True)
	
	ObjWe=Browser("Thank you - IshaHomes").Page("Villas for sale in OMR,").Sync
	
'ObjWe2=Browser("Thank you - IshaHomes_2").Page("Villas for sale in OMR,").Sync
	'Create object
	Set obj1=Description.Create
	obj1("micclass").Value="WebElement"
	obj1("class").Value="propOffer"
	Set ObjWe=Browser("micclass:=Browser").Page("micclass:=page").ChildObjects(obj1)
	'Store the number of project in an array
	'MsgBox ObjWe.Count
	strunit=""
	For i = 0 To ObjWe.Count-1 Step 1
		str=ObjWe(i).GetROProperty("innertext")
		str1=CInt(Mid(str,15,3))
	
		if(strunit="") then 
			strunit=str1
		else
			strunit= strunit & ";" & str1
		end if
	Next

'Msgbox strunit
	'Create object for extracting address and city
	Set obj2=Description.Create
	obj2("micclass").Value="WebElement"
	obj2("html tag").Value="P"
	obj2("height").Value="25"
	Set ObjWe2=Browser("micclass:=Browser").Page("micclass:=page").ChildObjects(obj2)

'MsgBox ObjWe2.Count
	strArea_City=""
	For i =0 To ObjWe2.Count-1 Step 1
		str2=ObjWe2(i).GetROProperty("innertext")
	
		if(strArea_City="") then 
			strArea_City=str2
		else
			strArea_City= strArea_City & ";" & str2
		end if
	Next
'MsgBox strArea_City

	strunittemp=Split(strunit,";")
	strAreatemp=Split(strArea_City,";")
'Extract the adress and city of projects as per requirement and display on the console
	For i= 0 To Ubound(strunittemp) Step 1
		if strunittemp(i)>=10  Then
			Print "The address and city of projects which has more than 10 units are"&strAreatemp(i)&"the number of available projects :"&strunittemp(i)
		End If
	Next
End Sub

'Close the browser
Sub CloseBrowser() 
	SystemUtil.CloseProcessByName("chrome.exe")
End Sub










'Sub CloseBrowser()
'	SystemUtil.CloseDescendentProcesses
'End Sub
