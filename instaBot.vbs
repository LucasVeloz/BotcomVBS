Dim startmundo
Set Fso = CreateObject("Scripting.FileSystemObject")
set funcao = CreateObject("WScript.Shell")
'funcao.Run("https://www.instagram.com/?hl=pt-br")
Dim x
x=0
Do While (x < 5)
wscript.sleep 500
funcao.SendKeys""
wscript.sleep 500
funcao.SendKeys "{ENTER}"
wscript.sleep 500
funcao.SendKeys "TESTE"
wscript.sleep 500
funcao.SendKeys "{ENTER}"
wscript.sleep 500

x=x+1
Loop    
Wscript.Quit