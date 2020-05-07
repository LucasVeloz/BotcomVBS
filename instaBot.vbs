Dim startmundo
Set Fso = CreateObject("Scripting.FileSystemObject")
set funcao = CreateObject("WScript.Shell")
'funcao.Run("https://www.instagram.com/?hl=pt-br%22)
Dim x
x=0
Do While (x < 420)
wscript.sleep 5000
funcao.SendKeys""
funcao.SendKeys "@kairofilipe"
wscript.sleep 2000
funcao.SendKeys "{ENTER}"
funcao.SendKeys "{ENTER}"
funcao.SendKeys "@nandafefe"
wscript.sleep 2000
funcao.SendKeys "{ENTER}"
funcao.SendKeys "{ENTER}"
funcao.SendKeys "{ENTER}"
funcao.SendKeys "{ENTER}"
funcao.SendKeys "{ENTER}"
wscript.sleep 5000
funcao.SendKeys "{F5}"
wscript.sleep 5000
funcao.SendKeys "{TAB}"
funcao.SendKeys "{TAB}"
wscript.sleep 45000
x=x+1
Loop
Wscript.Quit
