'Modified to update reply section for locations 7/1/16

On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

Set WshShell = CreateObject("WScript.Shell")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strName = objUser.FullName
strTitle = objUser.Description
strOffice2 = objUser.info
StrOffice = objUser.physicalDeliveryOfficeName
StrCompany = objUser.company
strCity = objUser.l
strStreet = objUser.StreetAddress
strLocation = objUser.l
strPostCode = objUser.PostalCode
strPhone = objUser.TelephoneNumber
strMobile = objUser.Mobile
strFax = objUser.FacsimileTelephoneNumber
strEmail = objUser.mail

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSelection.Font.Name = "Arial"
objSelection.Font.Size = 10
objSelection.ParagraphFormat.SpaceAfter = 0
objSelection.TypeText strName
objSelection.TypeText Chr(11)
if (strTitle) Then objSelection.TypeText strTitle & chr(11)
if (strOffice2) Then
  objSelection.TypeText strOffice2
ElseIf (strCompany) = "VTC Corp" Then
  objSelection.TypeText "Vanguard Truck Centers"
ElseIf (strCompany) = "VTC of Atlanta" Then
  objSelection.TypeText "Vanguard Truck Center of Atlanta"
ElseIf (strCompany) = "VTC of Adairsville" Then
  objSelection.TypeText "Vanguard Truck Center of Adairsville"
ElseIf (strCompany) = "VTC of Commerce" Then
  objSelection.TypeText "Vanguard Truck Center of Commerce"
ElseIf (strCompany) = "VTC of Savannah" Then
  objSelection.TypeText "Vanguard Truck Center of Savannah"
ElseIf (strCompany) = "VTC of Hempstead" Then
  objSelection.TypeText "Vanguard Truck Center of Houston East"
ElseIf (strCompany) = "VTC of Houston Bodyshop" Then
  objSelection.TypeText "Vanguard Truck Center of Houston Bodyshop"
ElseIf (strCompany) = "VTC of St. Louis" Then
  objSelection.TypeText "Vanguard Truck Center of St. Louis"
ElseIf (strCompany) = "VTC of Cahokia" Then
  objSelection.TypeText "Vanguard Truck Center of Cahokia"
ElseIf (strCompany) = "VTC of Austin" Then
  objSelection.TypeText "Vanguard Truck Center of Austin"
ElseIf (strCompany) = "VTC of Victoria" Then
  objSelection.TypeText "Vanguard Truck Center of Victoria"
ElseIf (strCompany) = "VTC of El Campo" Then
  objSelection.TypeText "Vanguard Truck Center of Victoria, El Campo"
ElseIf (strCompany) = "VTC of Phoenix" Then
  objSelection.TypeText "Vanguard Truck Center of Phoenix"
ElseIf (strCompany) = "VTC of Flagstaff" Then
  objSelection.TypeText "Vanguard Truck Center of Flagstaff"
ElseIf (strCompany) = "VTC of Tucson" Then
  objSelection.TypeText "Vanguard Truck Center of Tucson"
ElseIf (strCompany) = "VTC of Houston" Then
  objSelection.TypeText "Vanguard Truck Center of Houston"
ElseIf (strCompany) = "VTC of Alvin" Then
  objSelection.TypeText "Vanguard Truck Center of Houston South"
Else
  objSelection.TypeText "Vanguard Truck Centers"
End If
objSelection.TypeText Chr(11)
objSelection.TypeText "Office: " & strPhone
objSelection.TypeText Chr(11)
if (strFax) Then objSelection.TypeText "Fax: " & strFax & Chr(11)
if (strMobile) Then objSelection.TypeText "Cell: " & strMobile & Chr(11)
objSelection.TypeText strEmail & chr(11)
objselection.TypeText "www.vanguardtrucks.com" & chr(11)
objSelection.InlineShapes.AddPicture "C:\users\public\pictures\vanguard\emaillogo.jpg" 

Set objSelection = objDoc.Range()

objSignatureEntries.Add "vtc", objSelection
objSignatureObject.NewMessageSignature = "vtc"

objDoc.Saved = True
objWord.Quit

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSelection.Font.Name = "Arial"
objSelection.Font.Size = 10
objSelection.ParagraphFormat.SpaceAfter = 0
objSelection.TypeText strName
objSelection.TypeText Chr(11)
if (strTitle) Then objSelection.TypeText strTitle & chr(11)
if (strOffice2) Then
  objSelection.TypeText strOffice2
ElseIf (strCompany) = "VTC Corp" Then
  objSelection.TypeText "Vanguard Truck Centers"
ElseIf (strCompany) = "VTC of Atlanta" Then
  objSelection.TypeText "Vanguard Truck Center of Atlanta"
ElseIf (strCompany) = "VTC of Adairsville" Then
  objSelection.TypeText "Vanguard Truck Center of Adairsville"
ElseIf (strCompany) = "VTC of Commerce" Then
  objSelection.TypeText "Vanguard Truck Center of Commerce"
ElseIf (strCompany) = "VTC of Savannah" Then
  objSelection.TypeText "Vanguard Truck Center of Savannah"
ElseIf (strCompany) = "VTC of Hempstead" Then
  objSelection.TypeText "Vanguard Truck Center of Houston East"
ElseIf (strCompany) = "VTC of Houston Bodyshop" Then
  objSelection.TypeText "Vanguard Truck Center of Houston Bodyshop"
ElseIf (strCompany) = "VTC of St. Louis" Then
  objSelection.TypeText "Vanguard Truck Center of St. Louis"
ElseIf (strCompany) = "VTC of Cahokia" Then
  objSelection.TypeText "Vanguard Truck Center of Cahokia"
ElseIf (strCompany) = "VTC of Austin" Then
  objSelection.TypeText "Vanguard Truck Center of Austin"
ElseIf (strCompany) = "VTC of Victoria" Then
  objSelection.TypeText "Vanguard Truck Center of Victoria"
ElseIf (strCompany) = "VTC of El Campo" Then
  objSelection.TypeText "Vanguard Truck Center of Victoria, El Campo"
ElseIf (strCompany) = "VTC of Phoenix" Then
  objSelection.TypeText "Vanguard Truck Center of Phoenix"
ElseIf (strCompany) = "VTC of Flagstaff" Then
  objSelection.TypeText "Vanguard Truck Center of Flagstaff"
ElseIf (strCompany) = "VTC of Tucson" Then
  objSelection.TypeText "Vanguard Truck Center of Tucson"
ElseIf (strCompany) = "VTC of Houston" Then
  objSelection.TypeText "Vanguard Truck Center of Houston"
ElseIf (strCompany) = "VTC of Alvin" Then
  objSelection.TypeText "Vanguard Truck Center of Houston South"
Else
  objSelection.TypeText "Vanguard Truck Centers"
End If
objSelection.TypeText Chr(11)
objSelection.TypeText "Office: " & strPhone
objSelection.TypeText Chr(11)
if (strFax) Then objSelection.TypeText "Fax: " & strFax & Chr(11)
if (strMobile) Then objSelection.TypeText "Cell: " & strMobile & Chr(11)
objSelection.TypeText strEmail & chr(11)
objselection.TypeText "www.vanguardtrucks.com" & chr(11)

Set objSelection = objDoc.Range()

objSignatureEntries.Add "vtc2", objSelection

objSignatureObject.ReplyMessageSignature = "vtc2"

objDoc.Saved = True
objWord.Quit

'Option Explicit
'On Error Resume Next

Const HKEY_CURRENT_USER = &H80000001
Dim strComputer, oReg, strKeyPath, strValueName, strValue

strComputer = "."

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & StrComputer & "\root\default:StdRegProv")

strKeyPath = "Software\Microsoft\Office\14.0\Common\MailSettings"
strValueName = "ReplySignature"
strValue = "vtc2"
oReg.SetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue

strKeyPath = "Software\Microsoft\Office\15.0\Common\MailSettings"
strValueName = "ReplySignature"
strValue = "vtc2"
oReg.SetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue

strKeyPath = "Software\Microsoft\Office\16.0\Common\MailSettings"
strValueName = "ReplySignature"
strValue = "vtc2"
oReg.SetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue

Wscript.quit

