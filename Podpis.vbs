On Error Resume Next
Dim objSysInfo, objUser
Set objSysInfo = CreateObject("ADSystemInfo")
Set objUser = GetObject("LDAP://" & objSysInfo.UserName)

strBlock1 = "Z powa¿aniem / Mit freundlichen Grüßen / Best regards,"
strBlock2 = "Preh Car Connect  Polska Sp. z o.o."
strBlock3 = "ul. Poznañska 4, Siemianice " 
strBlock4 = "55-120 Oborniki Œl¹skie "
strBlock5 = "Polska / Poland"
strBlock6 = "________________________________________________________________"
strBlock12 = "E-Mail: "

strName = objUser.FullName
strTitle = objUser.Title
strPhone = objUser.telephoneNumber
strEmail = objUser.mail

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries


objSelection.Font.Name = "Arial"
objSelection.Font.Size = "10"
objSelection.Font.Color = Black

objSelection.TypeParagraph()
objSelection.TypeParagraph()
objSelection.TypeParagraph()
objSelection.TypeParagraph()
objSelection.TypeText strBlock1
objSelection.TypeParagraph()

objSelection.Font.Bold = 1
objSelection.TypeText strName
objSelection.Font.Bold = 0
objSelection.TypeText CHR(11)
objSelection.TypeText strTitle
objSelection.TypeParagraph()

objSelection.Font.Bold = 1
objSelection.TypeText strBlock2
objSelection.Font.Bold = 0
objSelection.TypeText CHR(11)
objSelection.TypeText strBlock3
objSelection.TypeText CHR(11)
objSelection.TypeText strBlock4
objSelection.TypeText CHR(11)
objSelection.TypeText strBlock5
objSelection.TypeText CHR(11)
objSelection.TypeParagraph()

objSelection.Font.Bold = 1
objSelection.TypeText "Phone: " 
objSelection.Font.Bold = 0
objSelection.TypeText strPhone 
objSelection.TypeText CHR(11)
objSelection.Font.Bold = 1
objSelection.TypeText strBlock12
objSelection.Font.Bold = 0
objselection.font.color = RGB(0, 0, 255)
objSelection.Font.Underline = 1
objSelection.TypeText strEmail
objSelection.Font.Underline = 0
objSelection.TypeParagraph()

objselection.font.color = RGB(200, 200, 200)
objSelection.TypeText strBlock6
objSelection.TypeText CHR(11)


Set objSelection = objDoc.Range()

objSignatureEntries.Add "Podpis", objSelection
objSignatureObject.ReplyMessageSignature = "Podpis"

objDoc.Saved = True
objDoc.Close
objWord.Quit


