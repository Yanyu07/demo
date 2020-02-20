Set WshShell=WScript.CreateObject("WScript.Shell")
for i =1 to 5
WScript.Sleep 500 
WshShell.SendKeys i
WshShell.SendKeys"woaini"
WshShell.SendKeys"{ }"
WScript.Sleep 100 
WshShell.SendKeys"{ENTER}"
Next