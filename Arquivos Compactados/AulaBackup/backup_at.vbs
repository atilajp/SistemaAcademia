Dim strlocal
Dim objws

on error resume next

Set objws = CreateObject("wscript.shell")

strlocal = "c:\maestro\backup\backup_at.accdb"

'ch(34) = Aspas 

strlocal = Chr(34) & "MSACCESS.EXE" & Chr(34) & " " & Chr(34) & strlocal & Chr(34)

'0 - oculto
'5 - visivel

ObjWS.Run strlocal, 0,"false"
