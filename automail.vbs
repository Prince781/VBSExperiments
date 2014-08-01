'Automated mailing
Set objOL = CreateObject("Outlook.Application")
Set objMail = objOL.CreateItem(0)
objMail.Subject = "Interesting File of Mine"
'objMail.To = "ambrosiacf@optonline.net"
objMail.Attachments.Add("C:\test.ppt")  'add an attachment
'objMail.Body = "Hi there, Ambrosia, this is a test message automated by vbscript.."
objMail.Display '- shows the email while compilation
'objMail.Send