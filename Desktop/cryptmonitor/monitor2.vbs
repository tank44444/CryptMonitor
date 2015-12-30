WScript.Echo "WARNING! Please copy the file C:\Windows\System32\shutdown.exe to F:\shutdown.exe"
WScript.Echo "WARNING! Please check these file existed C:\1.jpg, C:\1.pdf, C:\1.xlsx, C:\1.pptx, C:\1.docx, C:\1.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("Wscript.Shell")

Dim honeypotLoca(6)
honeypotLoca(0) = "C:\!test\1.jpg"
honeypotLoca(1) = "C:\!test\1.pdf"
honeypotLoca(2) = "C:\!test\1.xlsx"
honeypotLoca(3) = "C:\!test\1.pptx"
honeypotLoca(4) = "C:\!test\1.docx"
honeypotLoca(5) = "C:\!test\1.txt"

'check exist
Dim b(6)

For count = 0 To 5
		If Not fso.FileExists(honeypotLoca(count)) Then
			WScript.Echo "File not exist! " + honeypotLoca(count)
		End If
Next	



Do 
	For count = 0 To 5
		If Not fso.FileExists(honeypotLoca(count)) Then
			WScript.Echo "LLL"
			ws.run "shutdown.exe -s -f -t 0" ,vbhide 
			ws.run "cmd /c F:\shutdown.exe -s -f -t 0" ,vbhide 
		End If		
	Next	
	wscript.sleep 15000
Loop 







