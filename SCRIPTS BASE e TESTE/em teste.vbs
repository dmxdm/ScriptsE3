	
	Set e3 = CreateObject("CT.Application")
	Set Job = e3.CreateJobObject
	Set Sym = Job.CreateSymbolObject
	Set Dev = Job.CreateDeviceObject
	
	
	Job.GetSymbolIds connIds
	'Job.GetSymbolTypeIds connIds
	For i = 1 to UBound(connIds)
		Sym.SetId connIds(i)
		Sym.GetSymbolIds vIds
		For j = 1 to UBound(vIds)
			Sym.SetId vIds(j)
			MsgBox Sym.
			
		Next
	Next	

 

	

