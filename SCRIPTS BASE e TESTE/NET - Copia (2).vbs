    Set e3 = CreateObject("CT.Application")
    Set prj = e3.CreateJobObject
    Set net = prj.CreateNetObject
    Set ns = prj.CreateNetSegmentObject
    Set sht = prj.CreateSheetObject
    Set EmbSht = prj.CreateSheetObject
    Set conn = prj.CreateConnectionObject
	
	e3.ClearOutputWindow	
	
	shtcnt = prj.GetSheetIds(sIds)
	
	For k = 1 To shtcnt
		sht.SetId sIds(k)
			
	If ( sht.GetEmbeddedSheetIds(EShtIds) > 0) Then
		EmbSht.SetId EShtIds(1)
		
		If EmbSht.IsPanel = 1 Then
			conncnt = prj.GetConnectionIds(connIds)
			
			For l = 1 To conncnt
				conn.SetId	connIds(l)
				
				If EmbSht.IsPanel = 1 Then
				msgbox "yes"
				End if
				
			Next	
			'For i = 1 To conncnt
				'msgbox netids(i)
				'net.SetId netids(i)	
				'netsegcnt = net.GetNetSegmentIds(netsegids)
				'For j = 1 To netsegcnt
				'ns.SetId netsegids(j)
				'MsgBox ns.GetLevel 
				'Next
			'Next
		End If		
	End If
	Next
    Set e3 = Nothing


