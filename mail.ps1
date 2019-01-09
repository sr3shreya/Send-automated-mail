$Outlook = New-Object -ComObject Outlook.Application 
#create Outlook MailItem named Mail using CreateItem() method 
$Mail = $Outlook.CreateItem(0) 

#add properties as desired 
$Mail.To = "shreya_s@hcl.com" 
$Mail.Subject = "Production Checkout" 
$Mail.Body = "testing" 
#send message 
$Mail.Send() 
#quit and cleanup 
$Outlook.Quit()