Set objShell = WScript.CreateObject("WScript.Shell")
WScript.Sleep 200
objShell.AppActivate("Scripts")
objShell.SendKeys "{ESC}"
objShell.SendKeys "%F"
objShell.SendKeys "P"
objShell.SendKeys "R"
objShell.SendKeys "~"