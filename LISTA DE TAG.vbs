	
	Set e3 = CreateObject("CT.Application")
	Set Job = e3.CreateJobObject
	Set Dev = Job.CreateDeviceObject
	
	dim lista 
	Set outlist = CreateObject("System.Collections.ArrayList")
	
	Job.GetAllDeviceIds connIds
	For i = 1 to UBound(connIds)
		Dev.SetId connIds(i)
		If InStr(Dev.GetName, "Wires") > 0 then
		Exit for			
		End If
	lista = lista & Dev.GetName & vbCrlf
	
	
	
	outlist.Add Dev.GetName
	
	
	Next
	txtFileName  = Job.GetPath & Job.GetName  & " Tag.txt"
	WriteFile lista, txtFileName
	MsgBox lista
	
	MsgBox outlist.read
	
	
Function WriteFile(list, file)

	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	Set MyFile = ObjFSO.CreateTextFile(file, True)
	MyFile.WriteLine(list)
	MyFile.Close
	
End Function


