Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "cmd"
WScript.Sleep 200
WshShell.SendKeys "color c {enter}"
WScript.Sleep 1200
WshShell.SendKeys "h"
WScript.Sleep 150
WshShell.SendKeys "i"
WScript.Sleep 150
WshShell.SendKeys ","
WScript.Sleep 150
WshShell.SendKeys " "
WScript.Sleep 150
WshShell.SendKeys "h"
WScript.Sleep 150
WshShell.SendKeys "o"
WScript.Sleep 150
WshShell.SendKeys "q"
WScript.Sleep 150
WshShell.SendKeys "{bs}"
WScript.Sleep 150
WshShell.SendKeys "w"
WScript.Sleep 150
WshShell.SendKeys " "
WScript.Sleep 150
WshShell.SendKeys "a"
WScript.Sleep 150
WshShell.SendKeys "r"
WScript.Sleep 150
WshShell.SendKeys "e"
WScript.Sleep 150
WshShell.SendKeys " "
WScript.Sleep 150
WshShell.SendKeys "y"
WScript.Sleep 150
WshShell.SendKeys "o"
WScript.Sleep 150
WshShell.SendKeys "u"
WScript.Sleep 150
WshShell.SendKeys "?"
WScript.Sleep 2500
WshShell.SendKeys "{enter}"
WScript.Sleep 20
WshShell.SendKeys "exit"
WScript.Sleep 20
WshShell.SendKeys "{enter}"
WScript.Sleep 20
