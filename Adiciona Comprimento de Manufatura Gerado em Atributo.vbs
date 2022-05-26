	Set e3 = CreateObject("CT.Application")
	Set Job = e3.CreateJobObject
	Set Con = Job.CreateConnectionObject
	Set Wire = Job.CreatePinObject	
	Const unit = "mm"
	dim lista 
	Job.GetAllConnectionIds connIds
	For i = 1 to UBound(connIds)
		Con.SetId connIds(i)	
		Con.GetCoreIds coreIds
		For j = 1 to UBound(coreIds)
			Wire.SetId coreIds(j)
			lista = round(Wire.GetLength,2)
			Wire.AddAttributeValue "Length", lista & " " & unit			
		Next
	Next
e3.PutInfo 0,"THE LENGTH ATTRIBUTE HAS BEEN APPLIED"