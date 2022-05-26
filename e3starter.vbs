Option Explicit

dim oProcessStarter
set oProcessStarter = CreateObject("CT.E3Starter")

Dim oE3 : set oE3 = oProcessStarter.Start( _
"C:\Program Files\Zuken\E3.series_2019\E3.series.exe", _ 
"/multiuser /schema /formboard", -1)

dim E3COMRegistry

on Error Resume Next

If (Err.Numer) then
	Wscript.Echo Err.Description
end If

Err.clear
