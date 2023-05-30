ExitAction
'Without Framework Raw Project
While Browser("creationtime:=0").Exist(0)
	'Close the browser
	Browser("creationtime:=0").Close
Wend

SystemUtil.Run "Chrome","ishahomes.com"
Browser("Buy Apartments,Flats&Luxury").Page("Buy Apartments,Flats&Luxury").Sync
Browser("Buy Apartments,Flats&Luxury").Page("Buy Apartments,Flats&Luxury").Link("Contact us").Click @@ script infofile_;_ZIP::ssf1.xml_;_
Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").Sync
Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebEdit("sell_do[form][lead][name]").Set "Amandeep Singh" @@ script infofile_;_ZIP::ssf2.xml_;_
Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebEdit("sell_do[form][lead][email]").Set "amandeepsingh0045@gmail.com" @@ script infofile_;_ZIP::ssf3.xml_;_
Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebEdit("sell_do[form][lead][phone]").Set "+918264668498" @@ script infofile_;_ZIP::ssf4.xml_;_

Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebList("sell_do[form][lead][project_id").Select "Isha Code Field"

no_of_project=Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebList("sell_do[form][lead][project_id").GetROProperty("items count")
project_sites=Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebList("sell_do[form][lead][project_id").GetROProperty("all items")
selected_project=Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebList("sell_do[form][lead][project_id").GetROProperty("selection")


print "no of projects available "&no_of_project
print "the projects are :"
temp=Split(project_sites,";")
For each i in temp
	print i
Next
print "the selected project is :"&selected_project

Browser("Buy Apartments,Flats&Luxury").Page("Contact Isha Homes, an").WebButton("Submit").Click

Browser("Buy Apartments,Flats&Luxury").Page("Thank you - IshaHomes").Sync

'Browser("Thank you - IshaHomes").Page("Thank you - IshaHomes").Link("back to Homepage").Click @@ script infofile_;_ZIP::ssf9.xml_;_
'Browser("Thank you - IshaHomes").Page("Villas for sale in OMR,").Link("Buy Villas").Click @@ script infofile_;_ZIP::ssf10.xml_;_
'Browser("Buy Apartments,Flats&Luxury").Page("Are you looking to buy").Sync
'
'Browser("Buy Apartments,Flats&Luxury").Page("Are you looking to buy").WebElement("Total Units : 160 (2 &").Click


'//DIV[1]/DIV[1]

'
Browser("Thank you - IshaHomes").Page("Thank you - IshaHomes").Link("back to Homepage").Click @@ hightlight id_;_Browser("Thank you - IshaHomes 2").Page("Thank you - IshaHomes").Link("back to Homepage")_;_script infofile_;_ZIP::ssf11.xml_;_
'Browser("Thank you - IshaHomes").Page("Villas for sale in OMR,").Link("Buy Villas").Click @@ hightlight id_;_Browser("Thank you - IshaHomes 2").Page("Villas for sale in OMR,").Link("Buy Villas")_;_script infofile_;_ZIP::ssf12.xml_;_
 @@ hightlight id_;_2689828_;_script infofile_;_ZIP::ssf13.xml_;_
'
'
'

Browser("Buy Apartments,Flats&Luxury").Page("Buy Apartments,Flats&Luxury").Link("Buy Villas").Click @@ script infofile_;_ZIP::ssf14.xml_;_
ObjWe=Browser("Thank you - IshaHomes").Page("Villas for sale in OMR,").Sync
'ObjWe2=Browser("Thank you - IshaHomes_2").Page("Villas for sale in OMR,").Sync

Set obj1=Description.Create
obj1("micclass").Value="WebElement"
obj1("class").Value="propOffer"
Set ObjWe=Browser("micclass:=Browser").Page("micclass:=page").ChildObjects(obj1)

MsgBox ObjWe.Count
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

Msgbox strunit

Set obj2=Description.Create
obj2("micclass").Value="WebElement"
obj2("html tag").Value="P"
obj2("height").Value="25"
Set ObjWe2=Browser("micclass:=Browser").Page("micclass:=page").ChildObjects(obj2)

MsgBox ObjWe2.Count
strArea_City=""
For i =0 To ObjWe2.Count-1 Step 1
	str2=ObjWe2(i).GetROProperty("innertext")
	
	if(strArea_City="") then 
		strArea_City=str2
	else
		strArea_City= strArea_City & ";" & str2
	end if
Next
MsgBox strArea_City

strunittemp=Split(strunit,";")
strAreatemp=Split(strArea_City,";")

For i= 0 To Ubound(strunittemp) Step 1
	if strunittemp(i)>=10  Then
		Print strunittemp(i)&strAreatemp(i)
	End If
Next


SystemUtil.CloseProcessByName("chrome.exe") @@ script infofile_;_ZIP::ssf16.xml_;_
