	
	Set e3 = CreateObject("CT.Application")
	Set Job = e3.CreateJobObject
	Set Con = Job.CreateConnectionObject
	Set Signal = Job.CreateSignalObject
	Set Net = Job.CreateNetObject
    Set Ns = Job.CreateNetSegmentObject

	Job.GetConnectionIds connIds

	For i = 1 to UBound(connIds)
		Con.SetId connIds(i)	
		Con.GetNetSegmentIds netIds
		For j = 1 to UBound(netIds)
			Ns.SetId netIds(j)
			Net.SetId netIds(j)
			Signal.SetId Net.GetId
			Signal.SetName Ns.GetAttributeValue("WireNumber")
		Next
	Next

	Set e3 = Nothing
	Set Job = Nothing
	Set Con = Nothing
	Set Signal = Nothing
	Set Net = Nothing
    Set Ns = Nothing
	