	
	Set e3 = CreateObject("CT.Application")
	Set Job = e3.CreateJobObject()
	Dim scriptFile : scriptFile = "C:\Program Files\Zuken\E3.series_2021\scripts\CoverSheet.vbs"
	Dim scriptFile2 : scriptFile2 = "C:\Program Files\Zuken\E3.series_2021\scripts\PartsSheet_complete.vbs"

	Job.PurgeUnused()
	Job.Save

	e3.Run  scriptFile , "" 
	e3.Run  scriptFile2 , "" 

 
	Set e3 = Nothing
	Set Job = Nothing
