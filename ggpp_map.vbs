Dim tomorrow, editString, url

tomorrow = DateAdd("d",1,Date)
editString = "18" & Right(String(2, "0") & Month(tomorrow), 2) & Right(String(2, "0") & Day(tomorrow), 2)

url = "https://www.civilprotection.gr/sites/default/gscp_uploads/"& editString &".jpg"
'msgbox(url)
While Not(testUrl(url))
	WScript.Sleep 300000
Wend

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run(chr(34) & "C:\Program Files (x86)\Mozilla Firefox\firefox.exe" & chr(34) & url)

'Set objShell = WScript.CreateObject("Wscript.Shell")
'objShell.SendKeys "{FN}{F6}"

'strCommand = "wmplayer /play /close C:\Windows\media\Ring05.wav"
'objShell.Run strCommand, 0, True

'Wscript.Sleep 1000

function testUrl(url)
    Set o = CreateObject("MSXML2.XMLHTTP")
    on error resume next
    o.open "GET", url, False
    o.send
    if o.Status = 200 then testUrl = True
    on error goto 0 
end function