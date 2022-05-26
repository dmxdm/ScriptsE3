
	Set e3 = CreateObject("CT.Application")
	Set Job = e3.CreateJobObject
	Set Con = Job.CreateConnectionObject
	Set Wire = Job.CreatePinObject
	Set Dev = Job.CreateDeviceObject
	Set Dev2 = Job.CreateDeviceObject
	Set Pin = Job.CreatePinObject
	Set Pin2 = Job.CreatePinObject
	
	dim lista 
	
	TEXTO = "INFORME A OPCAO DESEJADA" & vbCrlf & "1 - Lista de Tag" & vbCrlf & "2 - Lista De/Para-Fio"
	MODE = InputBox(TEXTO)
	Select Case MODE
		Case "1"
			TagList
		Case "2"
			DeParaList
		Case ""
			MsgBox "Saindo."
		Case Else
			MsgBox "Opcao invalida" & ", saindo."
	End Select
	
	
Function TagList
	
	Job.GetAllDeviceIds connIds
	For i = 1 to UBound(connIds)
		Dev.SetId connIds(i)
		If InStr(Dev.GetName, "Fios") > 0 then
		Exit for			
		End If
	lista = lista & Dev.GetName & vbCrlf
	Next
	txtFileName  = Job.GetPath & Job.GetName  & ".txt"
	WriteFile lista, txtFileName
	
	
End Function
	
	
Function DeParaList
	
	Job.GetAllConnectionIds connIds
	For i = 1 to UBound(connIds)
		Con.SetId connIds(i)	
		Con.GetCoreIds coreIds
		For j = 1 to UBound(coreIds)
			Wire.SetId coreIds(j)
			Pin.SetId Wire.GetEndPinId (1,ret)
			Pin2.SetId Wire.GetEndPinId (2,ret)
			Dev.SetId Pin.GetId
			Dev2.SetId Pin2.GetId			
	lista = lista & Dev.GetName & ":" & Pin.GetName & "/" & Dev2.GetName & ":" & Pin2.GetName & "-" & Wire.GetName & vbCrlf
	Next
	Next
	txtFileName  = Job.GetPath & Job.GetName  & ".txt"
	WriteFile lista, txtFileName
	

	
End Function
	
Function WriteFile(list, file)

	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	Set MyFile = ObjFSO.CreateTextFile(file, True)
	Set WSHShell = WScript.CreateObject("WScript.Shell")
	MyFile.WriteLine(list)
	MyFile.Close
	
	result = MsgBox ("Deseja abrir o arquvio?", vbYesNo + vbQuestion)

	Select Case result
		Case vbYes
			WshShell.Run "notepad.exe " & file
		Case vbNo
			MsgBox "Saindo." 
		End Select
	
End Function




