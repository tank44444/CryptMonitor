WScript.Echo "WARNING! Please copy the file C:\Windows\System32\shutdown.exe to F:\shutdown.exe"
WScript.Echo "WARNING! Please check these file existed C:\1.jpg, C:\1.pdf, C:\1.xlsx, C:\1.pptx, C:\1.docx, C:\1.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("Wscript.Shell")

'record the MD5 hash value
Dim md5cmd(6)
md5cmd(0) = "cmd /c C:\!test\fciv.exe -md5 C:\!test\1.jpg"
md5cmd(1) = "cmd /c C:\!test\fciv.exe -md5 C:\!test\1.pdf"
md5cmd(2) = "cmd /c C:\!test\fciv.exe -md5 C:\!test\1.xlsx"
md5cmd(3) = "cmd /c C:\!test\fciv.exe -md5 C:\!test\1.pptx"
md5cmd(4) = "cmd /c C:\!test\fciv.exe -md5 C:\!test\1.docx"
md5cmd(5) = "cmd /c C:\!test\fciv.exe -md5 C:\!test\1.txt"

Dim md5_Temp(6), sRes,sResAr,md5Hex
For count = 0 To 5
	sRes = ws.Exec(md5cmd(count)).Stdout.ReadAll()
	sResAr = split(sRes," ")
	md5_Temp(count) = right(sResAr(6),32)
Next
count = 0

Dim honeypotLoca(6)
honeypotLoca(0) = "C:\!test\1.jpg"
honeypotLoca(1) = "C:\!test\1.pdf"
honeypotLoca(2) = "C:\!test\1.xlsx"
honeypotLoca(3) = "C:\!test\1.pptx"
honeypotLoca(4) = "C:\!test\1.docx"
honeypotLoca(5) = "C:\!test\1.txt"


'check exist
Dim b(6)

Do 
	For count = 0 To 5
		If Not fso.FileExists(honeypotLoca(count)) Then
			ws.run "shutdown.exe -s -f -t 0" ,vbhide 
			ws.run "cmd /c F:\shutdown.exe -s -f -t 0" ,vbhide 
		End If
		
		sRes = ws.Exec(md5cmd(count)).Stdout.ReadAll()
		sResAr = split(sRes," ")
		If Not md5_Temp(count) = right(sResAr(6),32) Then
			ws.run "shutdown.exe -s -f -t 0" ,vbhide 
			ws.run "cmd /c F:\shutdown.exe -s -f -t 0" ,vbhide
		End IF		
	Next	
	wscript.sleep 15000
Loop 







