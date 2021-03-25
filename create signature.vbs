'On Error Resume Next

'Setting up the script to work with the file system.
Set WshShell = WScript.CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

'Connecting to Active Directory to get user’s data.
Set objSysInfo = CreateObject("ADSystemInfo")
Set UserObj = GetObject("LDAP://" & replace(objSysInfo.UserName, "/", "\/"))
strAppData = WshShell.ExpandEnvironmentStrings("%APPDATA%")
SigFolder = StrAppData & "\Microsoft\Signatures\"
SigFile = SigFolder & UserObj.SAMAccountName & " short.htm"
SigFile2 = SigFolder & UserObj.SAMAccountName & " long.htm"


'Setting placeholders for the signature.
strUserName = UserObj.SAMAccountName
strFullName = UserObj.firstname & " " & UserObj.lastname
strTitle = UserObj.title
strDep = UserObj.department & " Department"
strMobile = UserObj.mobile
strEmail = UserObj.mail
strCompany = UserObj.company
strOfficePhone = UserObj.telephoneNumber

'Setting global placeholders for the signature. Those values will be identical for all users - make sure to replace them with the right values!
strCompanyLogo = "https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/simplephoto-with-logo/logo.png"
strCompanyAddress1 = "16 Freedom St, Deer Hill"
strCompanyAddress2 = "58-500 Poland"
strWebsite = "https://www.my-company.com"


'Creating HTM signature file for the user's profile, if the file with such a name is found, it will be overwritten.
Set CreateSigFile = FSO.CreateTextFile (SigFile, True, True)

'Signature’s HTML code.
CreateSigFile.WriteLine "<html><head><TITLE>Email Signature</TITLE>"
CreateSigFile.WriteLine "<META content=" & """text/html; charset=utf-8" & """ http-equiv=" & """Content-Type" & """>"
CreateSigFile.WriteLine "<style>"
CreateSigFile.WriteLine "p.MsoNormal {color:#00136F;margin:0cm; margin-bottom:.0001pt;font-size:11pt;font-family:'Calibri',sans-serif;}"
CreateSigFile.WriteLine "p.MsoNormal2 {color:#00136F;margin:0cm; margin-bottom:.0001pt;font-size:10pt;font-family:'Arial',sans-serif;}"
CreateSigFile.WriteLine "</style>"
CreateSigFile.WriteLine "</head>"
CreateSigFile.WriteLine ""
CreateSigFile.WriteLine "<body lang=RU link=blue vlink='#954F72' style='tab-interval:35.4pt'>"
CreateSigFile.WriteLine "<div>"
CreateSigFile.WriteLine "<p class=MsoNormal2><b><span><o:p>&nbsp;</o:p></span></b></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><b><span>" & strFullName & " | " & strTitle &"</span></b></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><span>" & strDep & "</span></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><span>Beiersdorf Russia | Zemlyanoy Val 9, 105064 Moscow, Russia</span></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><span>Tell.: " & strOfficePhone  & "</span></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><span>Cell phone: " &  strMobile & "</span></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2>&nbsp;</p>"
CreateSigFile.WriteLine "</div></body></html>"

CreateSigFile.Close

'Creating HTM signature file for the user's profile, if the file with such a name is found, it will be overwritten.
files_folder = SigFolder & UserObj.SAMAccountName & " long_files"
If Not fso.FolderExists(files_folder) then FSO.CreateFolder files_folder
FSO.CopyFile "\\mows0001\bdfdata$\ExchangeFiles\!\image002.jpg", files_folder & "\"
FSO.CopyFile "\\mows0001\bdfdata$\ExchangeFiles\!\image004.jpg", files_folder & "\"
Set CreateSigFile = FSO.CreateTextFile (SigFile2, True, True)

'Signature’s HTML code.
CreateSigFile.WriteLine "<html><head><TITLE>Email Signature</TITLE>"
CreateSigFile.WriteLine "<META content=" & """text/html; charset=utf-8" & """ http-equiv=" & """Content-Type" & """>"
CreateSigFile.WriteLine "<style>"
CreateSigFile.WriteLine "p.MsoNormal {color:#00136F;margin:0cm; margin-bottom:.0001pt;font-size:11pt;font-family:'Calibri',sans-serif;}"
CreateSigFile.WriteLine "p.MsoNormal2 {color:#00136F;margin:0cm; margin-bottom:.0001pt;font-size:10pt;font-family:'Arial',sans-serif;}"
CreateSigFile.WriteLine "</style>"
CreateSigFile.WriteLine "</head>"
CreateSigFile.WriteLine ""
CreateSigFile.WriteLine "<body lang=RU link=blue vlink='#954F72' style='tab-interval:35.4pt'>"
CreateSigFile.WriteLine "<div>"
CreateSigFile.WriteLine "<p class=MsoNormal2><b><span><o:p>&nbsp;</o:p></span></b></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><b><span>" & strFullName & "</span></b></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><span>" & strTitle & "</span></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><span><img width=99 height=38 src='" & UserObj.SAMAccountName & " long_files" & "/image002.jpg' alt='BDF'></span>"
CreateSigFile.WriteLine "<p class=MsoNormal2><span>Beiersdorf Russia</span></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><span>" & strDep & "</span></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><span>Zemlyanoy Val 9</span></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><span>105064 Moscow</span></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><span>Russia</span></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><span>Tell.: " & strOfficePhone  & "</span></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><span>Cell phone: " &  strMobile & "</span></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><u><span style='color:blue'>" &  strEmail & "</span></u></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><a href='http://www.'><span>www.nivea.ru</span></a><span style='color:blue'></span></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><a href='http://www.beiersdorf.com'><span>www.beiersdorf.com</span></a><span><br></span></p>"
CreateSigFile.WriteLine "<p class=MsoNormal2><span><img width=177 height=40 src='" & UserObj.SAMAccountName & " long_files" & "/image004.jpg' alt='Logo'></span>"
CreateSigFile.WriteLine "<p class=MsoNormal2>&nbsp;</p>"
CreateSigFile.WriteLine "</div></body></html>"

CreateSigFile.Close

'Applying the signature in Outlook’s settings.
Set objWord = CreateObject("Word.Application")
Set objSignatureObjects = objWord.EmailOptions.EmailSignature

'Setting the signature as default for new messages.
'objSignatureObjects.NewMessageSignature = strUserName & " short"

'Setting the signature as default for replies & forwards.
'objSignatureObjects.ReplyMessageSignature = strUserName & " short"

objWord.Quit