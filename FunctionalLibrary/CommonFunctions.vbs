

Sub Launch_Site()
	Dim strURL
	
	While Browser("creationtime:=0").Exist(0)
    'Close the browser
    Browser("creationtime:=0").Close
    Wend
	
	strURL = DataTable.Value("URL","Global")
    Systemutil.Run "iexplore.exe",strURL
    wait(10)
    
End Sub

Sub Navigate_To_Hotels()
	
	Browser("Browser").Page("Page").Link("Hotels").Click
	wait(5)

End Sub

Sub Select_City()
	
	Browser("Browser").Page("Page").WebEdit("City").Set DataTable("City", dtGlobalSheet)

End Sub

Sub Click_On_Calender()
	
	Browser("Browser").Page("Page").WebElement("class:=fnt13","index:=0").Click    'Clicking the CheckInDate to display the calender
	
End Sub
	

Sub Select_Date(IDate)
	
	Dim GDate
	Set objdate=Browser("Browser").Page("Page").WebElement("Class:=ui-datepicker-title","index:=0")
	GDate=cDate(objdate.getroproperty("Innertext"))


	While Not ( Month(Idate) = Month(Gdate) And Year(Idate) = Year(GDate) )
	If  ((IDate) < GDate ) Then 
		Browser("Browser").Page("Page").WebElement("Prev").Click
		Set objdate=Browser("Browser").Page("Page").WebElement("Class:=ui-datepicker-title","index:=0")
		GDate=CDate(objdate.getroproperty("Innertext"))
	
	Else
		Browser("Browser").Page("Page").WebElement("Next").Click
		Set objdate=Browser("Browser").Page("Page").WebElement("Class:=ui-datepicker-title","index:=0")
		GDate=CDate(objdate.getroproperty("Innertext"))

	End if

	WEND
	
	Set objDate=Nothing
	Browser("Browser").Page("Page").WebTable("Su").Link("Innertext:="&Day(IDate)).Click

End Sub



Sub Navigate_To_Hotel_Results()
	
	Browser("Browser").Page("Page").WebButton("Search").Click
	wait(10)
	
End Sub

Sub Validating_City()
	
	'Validating City
	
	Dim Dispay_City,Enter_City
	
	Display_City=Browser("Browser").Page("Page").WebEdit("Class:=ui-autocomplete-input").GetROProperty("value")
	Enter_City=Trim(DataTable.Value("City","Global"))
	If Display_City=Enter_City Then
		Reporter.ReportEvent micPass, "City Validation","City Entered and Displayed is matching"
	Else
		Reporter.ReportEvent micFail , "City Validation","City Entered and Displayed is not matching"
	End If

End Sub


Sub Validating_Check_In_Date()
	
	'Validating CheckInDate
	Dim Display_InDate,Enter_InDate
	
	Display_InDate=cDate(Browser("Browser").Page("Page").WebEdit("html id:=txtCheckInDate").GetROProperty("value"))
	Enter_InDate=cDate(DataTable.Value("CheckIn_Date","Global"))
	If Display_City=Enter_City Then
		Reporter.ReportEvent micPass, "Check In Date Validation","Check In Date Entered and Displayed is matching"
	Else
		Reporter.ReportEvent micFail , "Check In Date Validation","Check In Date Entered and Displayed is not matching"
	End If
End Sub

Sub Validating_Check_Out_Date()
	
	'Validating CheckOutDate
	Dim Display_OutDate,Enter_OutDate
	
	Display_OutDate=cDate(Browser("Browser").Page("Page").WebEdit("html id:=txtCheckOutDate").GetROProperty("value"))
	Enter_OutDate=cDate(DataTable.Value("Checkout_Date","Global"))
	If Display_City=Enter_City Then
		Reporter.ReportEvent micPass, "Check Out Date Validation","Check Out Date Entered and Displayed is matching"
	Else
		Reporter.ReportEvent micFail , "Check Out Date Validation","Check Out Date Entered and Displayed is not matching"
	End If

	
End Sub


Sub Filter_Price()
	
	'Filtering Price
	Dim max_value,x_max,y_max,x
	
	max_value=Browser("Browser").Page("Page").WebEdit("Filterup").GetROProperty("VALUE")	'Getting the maximum price in filter option
	x_MAX=Browser("Browser").Page("Page").WebElement("Webelement").GetROProperty("x")		'Getting the x coordinates of the Webelement at max value
	y_max=Browser("Browser").Page("Page").WebElement("Webelement").GetROProperty("y")		'Getting the y coordinates of the Webelement at max value
	
	x_max=x_max+10																			'Adjusting the coordinates
	y_max=y_max+10
	
	Browser("Browser").Page("Page").WebEdit("Filterup").Set 10000							'Setting the value to required 10000 to get the coordinates at 10000
	x=Browser("Browser").Page("Page").WebElement("Webelement").GetROProperty("x")			'Getting the x coordinates of the Webelement at 10000

	x=x+11

	Browser("Browser").Page("Page").WebEdit("Filterup").Set max_value						'Setting the webelement back to max value

	Window("Google Chrome").WinObject("Chrome Legacy Window").Drag x_max,y_max				'Drag from the max value
	Window("Google Chrome").WinObject("Chrome Legacy Window").Drop x,y_max					'Drop at required 10000
	
	wait(5)																					'Waiting for the page to filter


End Sub

Sub Filter_Rating()

	'Filtering 5Star and 4Star Hotels
	Browser("Browser").Page("Page").WebCheckbox("Class Name:=WebCheckBox","name:=selectedRating","value:=5").Click
	Browser("Browser").Page("Page").WebCheckbox("Class Name:=WebCheckBox","name:=selectedRating","value:=4").Click
	wait(2)

End Sub

Sub Filter_Amenities()
	
	'Filtering AC and Parking
	Browser("Browser").Page("Page").WebCheckbox("Class Name:=WebCheckBox","name:=filterAmenity","value:=AC").Click
	Browser("Browser").Page("Page").WebCheckbox("Class Name:=WebCheckBox","name:=filterAmenity","value:=Parking").Click
	wait(10)
	

End Sub

Sub Print_Hotel_Names()
	
	'Printing First Five hotels
	
	Dim Hotel 
	
	Print "The First Five Hotels satisfying the filter criteria are: "
	Set obj=Description.Create
	obj("micclass").value="Webelement"
	obj("Class").value="hoteNme ng-binding"
	
	Set oHotel=Browser("Browser").Page("Page").Childobjects(obj)
	For i = 0 To 4 
		If oHotel(i).Exist Then
			Hotel=oHotel(i).GetRoProperty("innertext")
			print i+1&"." & Hotel
		End If
	Next
	
	Set obj=Nothing
	Set oHotel=Nothing

End Sub

Sub Capture_Screenshot_Of_Results()
	
	
	'capture screenshot
	GetFolderPath=Environment.Value("TestDir")
	screen=now
	dt1=replace(now,"/","")
	dt2=replace(dt1,":","")
	fnl2=replace(replace(dt2," ",""),"AM","")
	fnl="capture" &fnl2
	FilePath = GetFolderPath & "\ResultScreenshot"&"\"&fnl&".png"
	imgg=Browser("Browser").Page("Page").CaptureBitmap(FilePath,True)
	
	
End Sub


Sub Close_Browser() 
	SystemUtil.CloseProcessByName("chrome.exe")
End Sub


