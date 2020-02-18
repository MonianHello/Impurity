Set Wshshell=CreateObject("Wscript.Shell") 
Wscript.sleep 3000
dim s 
do until s=20000
s=s+1 
WshShell.SendKeys "^v^{ENTER}"
loop
