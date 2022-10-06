Set fso=CreateObject("Scripting.FileSystemObject")

filename = "C:\Users\login\Desktop\ffff.txt"
listFile = fso.OpenTextFile(filename).ReadAll
listLines = Split(listFile, vbCrLf)
i = 0
WScript.Sleep 5000

For Each line In listLines
   Set WshShell = WScript.CreateObject("WScript.Shell")
   WScript.Sleep 500
   WshShell.SendKeys line
   WScript.Sleep 100
   WshShell.SendKeys "{ENTER}"

Next
