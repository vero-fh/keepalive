Dim objResult

Set objShell = WScript.CreateObject("WScript.Shell")
                         
Do While True
  objResult = objShell.sendkeys("{F15}{F15}")
  Wscript.sleep (10000)
Loop