    Set e3 = CreateObject("CT.Application")
    Set prj = e3.CreateJobObject
    Set net = prj.CreateNetObject
    Set ns = prj.CreateNetSegmentObject
    Set sht = prj.CreateSheetObject
    Set EmbSht = prj.CreateSheetObject
    e3.ClearOutputWindow	
	shtcnt = prj.GetSheetIds(sIds)
	For k = 1 To shtcnt
	sht.SetId sIds(k)
	If ( sht.GetEmbeddedSheetIds(EShtIds) > 0) Then
		EmbSht.SetId EShtIds(1)		
			netcnt = EmbSht.GetNetIds(netids)
			For i = 1 To netcnt
				net.SetId netids(i)	
				netsegcnt = net.GetNetSegmentIds(netsegids)
				For j = 1 To netsegcnt
					ns.SetId netsegids(j)	
					e3.PutMessage "Name: " & ns.GetId & "Dia: " & ns.GetOuterDiameter, ns.GetId
					width= ns.GetOuterDiameter	
					If EmbSht.isformboard = 1 Then
					ns.SetLineWidth width
					End If
				Next	
			Next
		Else
		e3.PutMessage "Not a Formboard Sheet"
		End If
	Next
    Set e3 = Nothing


