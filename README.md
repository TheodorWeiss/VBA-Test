$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)

$mail.To = "..."
$mail.Subject = "Test"
$mail.Body = "Hallo"

$mail.Send()
