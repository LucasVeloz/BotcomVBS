Dim startmundo
Set Fso = CreateObject("Scripting.FileSystemObject")
set startmundo = CreateObject("WScript.Shell")
'startmundo.Run("https://www.instagram.com/?hl=pt-br")
Dim x
x=0
Do While (x < 5)
wscript.sleep 500
startmundo.SendKeys""
wscript.sleep 500
startmundo.SendKeys "{ENTER}"
wscript.sleep 500
startmundo.SendKeys "TESTE"
wscript.sleep 500
startmundo.SendKeys "{ENTER}"
x=x+1
Loop    
Wscript.Quit