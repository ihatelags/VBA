Dim xl
Dim xlBook
Dim xlBook2
  
with CreateObject("WScript.Shell")
	Set oExec=.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
	sFileSelected = oExec.StdOut.ReadLine
	if sFileSelected = "" then
		msgBox "No file was selected!"
		Wscript.Quit
	end if
end with

set xl = createobject("Excel.Application")

xl.Application.Visible = True
xl.DisplayAlerts = False

Set xlBook2 = xl.Workbooks.Open(sFileSelected, 0, False)

'new user
set new_user_WorkSheet = xlBook2.Worksheets(1)

set new_user_Name = new_user_WorkSheet.Range("C8")
	x = 1
	y = 0
	check = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
	if InStr(check, left(new_user_Name.value,1)) then 
		x = 0
		y = 1
	end if
	new_user_fullname_eng = split(new_user_Name.value, " / ")(x)
	new_user_name_eng = split(new_user_fullname_eng, " ")(1)
	new_user_lastname_eng = split(new_user_fullname_eng, " ")(0)
	new_user_fullname_rus = split(new_user_Name.value, " / ")(y)
	new_user_name_rus = split(new_user_fullname_rus, " ")(1)
	new_user_lastname_rus = split(new_user_fullname_rus, " ")(0)
set new_user_CC = new_user_WorkSheet.Range("C10")
set new_user_Department = new_user_WorkSheet.Range("C11")
set new_user_Title_full = new_user_WorkSheet.Range("C12")
new_user_Title_en = split(new_user_Title_full.value, " / ")(1)
new_user_Title_rus = split(new_user_Title_full.value, " / ")(0)
set new_user_Manager = new_user_WorkSheet.Range("C13")
set new_user_Start_date= new_user_WorkSheet.Range("C14")
set new_user_City = new_user_WorkSheet.Range("C16")

'users
Set xlBook = xl.Workbooks.Open("C:\Users\konkova\BDF Group\IT CIS - General\users.xlsx", 0, False)
'find last row in users
lastrow = xlBook.Worksheets("RU0126").Range("A1").End(-4121).Row
xlBook.Worksheets("RU0126").Range("A1").End(-4121).EntireRow.Insert -4121, 0


set WorkSheet_users = xlBook.Worksheets("RU0126")

set users_Username = WorkSheet_users.Range("A" & lastrow)
set users_Name = WorkSheet_users.Range("B" & lastrow)
set users_lastName = WorkSheet_users.Range("C" & lastrow)
set users_Title = WorkSheet_users.Range("D" & lastrow)
set users_City = WorkSheet_users.Range("E" & lastrow)
set users_Title_rus = WorkSheet_users.Range("F" & lastrow)
set users_Name_rus = WorkSheet_users.Range("G" & lastrow)
set users_LastName_rus = WorkSheet_users.Range("I" & lastrow)
set users_CC = WorkSheet_users.Range("M" & lastrow)
set users_Department = WorkSheet_users.Range("Z" & lastrow)
set users_Manager = WorkSheet_users.Range("N" & lastrow)
set users_Start_date= WorkSheet_users.Range("Q" & lastrow)

users_name.value = new_user_name_eng
users_lastName.value = new_user_lastname_eng
users_username.value = new_user_lastname_eng & left(new_user_name_eng, 1)
users_Title.value = new_user_title
users_Title_rus.value = new_user_title_rus
users_City.value = new_user_City 
users_Name_rus.value = new_user_name_rus 
users_LastName_rus.value = new_user_lastname_rus 
users_CC.value = new_user_CC 
users_Department.value = new_user_Department 
users_Manager.value = new_user_Manager 
users_Start_date.value = new_user_Start_date

'xlBook.save
'xl.ActiveWindow.close True
'xl.Quit

'Set xlBook = Nothing
'Set xl = Nothing

