	
	Set e3 = CreateObject("CT.Application")
	Set Job = e3.CreateJobObject
	Set Con = Job.CreateConnectionObject
	Set Wire = Job.CreatePinObject
	Set Dev = Job.CreateDeviceObject
	Set Dev2 = Job.CreateDeviceObject
	Set Pin = Job.CreatePinObject
	Set Pin2 = Job.CreatePinObject
	
	dim lista 
	
	Job.GetAllConnectionIds connIds
	
	dev.getviewids vids	
	
	For i = 1 to UBound(connIds)
		Con.SetId connIds(i)	
		Con.GetCoreIds coreIds
		For j = 1 to UBound(coreIds)
			Wire.SetId coreIds(j)
			Pin.SetId Wire.GetEndPinId (1,ret)
			Pin2.SetId Wire.GetEndPinId (2,ret)
			Dev.SetId Pin.GetId
			Dev2.SetId Pin2.GetId			
	lista = lista ...
	& Dev.GetName & ":" & Pin.GetName & "/" & Dev2.GetName & ":" & Pin2.GetName & "-" & Wire.GetName & "-" & wire.GetFitting  &vbCrlf
		
		
		
		Next
	Next
	txtFileName  = Job.GetPath & Job.GetName  & " DE-PARA-FIOS.txt"
	WriteFile lista, txtFileName
	MsgBox lista
	
	
Function WriteFile(list, file)

	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	Set MyFile = ObjFSO.CreateTextFile(file, True)
	MyFile.WriteLine(list)
	MyFile.Close
	
End Function


