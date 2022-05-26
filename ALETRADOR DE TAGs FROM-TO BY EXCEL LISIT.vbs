	
	Set e3 = CreateObject("CT.Application")
	Set Job = e3.CreateJobObject
	Set Dev = Job.CreateDeviceObject
	Set Comp = Job.CreateComponentObject

	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open _
	("C:\Users\CIM-TEAM-N09\Desktop\ALETRADOR DE TAGs FROM-TO BY EXCEL LISIT.xlsx")

	Job.GetAllDeviceIds connIds
	
	For i = 1 to UBound(connIds)
		Dev.SetId connIds(i)
		Comp.SetId Dev.GetId
		
		intRow = 2

		Do Until objExcel.Cells(intRow,1).Value = ""
		
		If Comp.GetAttributeValue ("DeviceLetterCode") = objExcel.Cells(intRow, 1).Value then
		numero = Mid(Dev.GetName, Len(Comp.GetAttributeValue ("DeviceLetterCode"))+2)
		Dev.SetName (objExcel.Cells(intRow, 2).Value & numero)	
		end if

		intRow = intRow + 1
		
		Loop	
			
	Next

	objExcel.Quit
