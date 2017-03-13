On Error Resume Next
Dim objSysInfo, objUser
Set objSysInfo = CreateObject("ADSystemInfo")
Set objUser = GetObject("LDAP://" & objSysInfo.UserName)


strLogo = "Z:\Dzial IT\Zastrzezony dla dzialu\AVI\New folder\Sc\Logo_CarConnect.jpg"
strBlock1 = "Z powa¿aniem / Mit freundlichen Grüßen / Best regards,"
strBlock2 = "Preh Car Connect  Polska Sp. z o.o."
strBlock3 = "ul. Poznañska 4, Siemianice " 
strBlock4 = "55-120 Oborniki Œl¹skie "
strBlock5 = "www.preh-car-connect.com"
strBlock6 = "________________________________________________________________"
strBlock7 = "Wir bitten um Verständnis dafür, dass die in dieser E-Mail gegebene Information  aus Rechts- und Sicherheitsgründen nicht rechtsverbindlich sein kann. Eine rechtsverbindliche Bestätigung reichen wir Ihnen gerne auf Anforderung in schriftlicher Form nach. Beachten Sie bitte, dass jede Form der unautorisierten Nutzung, Veröffentlichung, Vervielfältigung oder Weitergabe des Inhalts dieser E-Mail nicht gestattet ist. Diese Nachricht ist ausschließlich für den bezeichneten Adressaten oder dessen Vertreter bestimmt. Sollten Sie nicht der vorgesehene Adressat dieser Mail oder dessen Vertreter sein, so bitten wir Sie, sich mit dem Absender der E-Mail in Verbindung zu setzen. For legal and security reasons the information provided in this e-mail is not legally binding. Upon request we would be pleased to provide you with a legally binding confirmation in written form. Any form of unauthorized use, publication, reproduction, copying or disclosure of the content of this e-mail is not permitted. This message is exclusively for the person addressed or their representative. If you are not the intended recipient of this message and its contents, please notify the sender immediately."
strBlock12 = "E-Mail: "

strName = objUser.FullName
strTitle = objUser.Title
strPhone = objUser.telephoneNumber
strEmail = objUser.mail
strMobile = objUser.mobile


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
objSelection.TypeText "Phone: " 
objSelection.Font.Bold = 0
objSelection.TypeText strPhone 
objSelection.TypeText CHR(11)

objSelection.Font.Bold = 1
objSelection.TypeText "Mobile: " 
objSelection.Font.Bold = 0
objSelection.TypeText strMobile
objSelection.TypeText CHR(11)


objSelection.Font.Bold = 1
objSelection.TypeText strBlock12
objSelection.Font.Bold = 0
objselection.font.color = RGB(0, 0, 255)
objSelection.Font.Underline = 1
objSelection.TypeText strEmail
objSelection.Font.Underline = 0
objSelection.TypeText CHR(11)
objSelection.TypeParagraph()

objselection.font.color = RGB(0, 0, 0)
objSelection.TypeText strBlock6


objSelection.TypeParagraph()

objSelection.InlineShapes.AddPicture(strLogo)
objSelection.TypeParagraph()

objSelection.Font.Bold = 1
objSelection.TypeText strBlock2
objSelection.Font.Bold = 0
objSelection.TypeText CHR(11)
objSelection.TypeText strBlock3
objSelection.TypeText CHR(11)
objSelection.TypeText strBlock4
objSelection.TypeText CHR(11)

objselection.font.color = RGB(0, 0, 255)
objSelection.Font.Underline = 1
objSelection.TypeText strBlock5
objSelection.Font.Underline = 0
objselection.font.color = RGB(0, 0, 0)
objSelection.TypeParagraph()

objSelection.Font.Size = "7"
objSelection.TypeText strBlock7
objSelection.TypeText CHR(11)



Set objSelection = objDoc.Range()

UserDataPath = ObjShell.ExpandEnvironmentStrings("%appdata%")
FolderLocation = UserDataPath &"\Microsoft\AD_Sig\"
HTMFileString = FolderLocation & Company & ".htm"
objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Outlook\Options\Mail\EnableLogging" , "0", "REG_DWORD"



objSignatureEntries.Add "Company Signature", objSelection
objSignatureObject.NewMessageSignature = "Company Signature"


objDoc.Saved = True
objDoc.Close
objWord.Quit