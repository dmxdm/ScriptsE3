    Set e3 = CreateObject("CT.Application")
    Set prj = e3.CreateJobObject
    Set netseg = prj.CreateNetSegmentObject
    Set sht = prj.CreateSheetObject
	Set wire = prj.CreatePinObject
	
    e3.ClearOutputWindow	
	prj.GetSheetIds sIds
	For k = 1 To UBound(sIds)
	sht.SetId sIds(k)
	sht.GetNetSegmentIds netsegIds
				For j = 1 To UBound(netsegIds)
				netseg.SetId netsegIds(j)	
				netseg.GetCoreIds coreIds
					For i = 1 To UBound(coreIds)
					wire.SetId CoreIds(i)
					MsgBox wire.GetColourDescription & "  "
			
					select case wire.GetColourDescription
					case "Preto"
					netseg.SetLineColour 0
					case "Vermelho"
					netseg.SetLineColour 13
					'case else
					end select

					next			
				Next	
	next			
    Set e3 = Nothing


