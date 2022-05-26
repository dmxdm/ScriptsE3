	
	Set e3 = CreateObject("CT.Application")
	Set Job = e3.CreateJobObject
	Set Comp = Job.CreateComponentObject
	Set Dev = Job.CreateDeviceObject
	Set Att= Job.CreateAttributeObject

	Job.GetComponentIds compIds

	For i = 1 to UBound(compIds)
		Comp.SetId compIds(i)
		Comp.GetAttributeIds attids
		For j = 1 to UBound(attids)
		Att.SetId attids(j)
		Comp.AddAttributeValue Att.Getname, "" 
		Comp.DeleteAttribute( Att.Getname )
		MsgBox Att.Getname
		Next
	Next
	
	Job.GetDeviceIds devIds

	For k = 1 to UBound(devIds)
		Dev.SetId devIds(k)
		Dev.GetAttributeIds attids2
		For p = 1 to UBound(attids2)
		Att.SetId attids2(p)
		Dev.AddAttributeValue Att.Getname, "" 
		Dev.DeleteAttribute( Att.Getname )
		MsgBox Att.Getname
		Next
	Next
	
	
	Set e3 = nothing
	Set Job = nothing
	Set Comp = nothing
	Set Dev = nothing
	Set Att= nothing