	Set e3 = CreateObject("CT.Application")
    Set prj = e3.CreateJobObject
    Set net = prj.CreateNetObject
    Set ns = prj.CreateNetSegmentObject
    Set sht = prj.CreateSheetObject
    Set EmbSht = prj.CreateSheetObject
	set Sym = prj.CreateSymbolObject
	Set wire = prj.CreatePinObject

    e3.ClearOutputWindow	
	
	'PARA FUNCIONAR É NECESSARIO COLOCAR TUDO DE NOME DE COR NA LIGUA DO SOFTWARE!!!!!
	
	shtcnt = prj.GetSheetIds(sIds)
	
	dim lista
	
	For k = 1 To shtcnt
	sht.SetId sIds(k)	
	netcnt = sht.GetNetIds(netids)
		For i = 1 To netcnt
		net.SetId netids(i)	
		netsegcnt = net.GetNetSegmentIds(netsegids)
			For j = 1 To netsegcnt
			ns.SetId netsegids(j)				
			nscnt = ns.GetCoreIds (coreIds)		
				For p = 1 To nscnt
				wire.SetId coreIds(p)
				if nscnt = 1 then
				ns.SetLineStyle 1
					select case wire.GetColourDescription
						case "Black"
						ns.SetLineColour 0
						case "Red"
						ns.SetLineColour 13
						case "Blue"
						ns.SetLineColour 239
						case "Brown"
						ns.SetLineColour 12					
						case "Grey"
						ns.SetLineColour 11
						case "Dark Blue"
						ns.SetLineColour 4
						case "Green-Yellow"
						ns.SetLineStyle 5
						case else
					end select
				else
				'MsgBox wire.GetColourDescription
					if p = 1 then
					ns.SetLineStyle 1
						lista = wire.GetColourDescription
						select case lista
							case "Black"
							ns.SetLineColour 0
							case "Red"
							ns.SetLineColour 13
							case "Blue"
							ns.SetLineColour 239
							case "Brown"
							ns.SetLineColour 12					
							case "Grey"
							ns.SetLineColour 11
							case "Dark Blue"
							ns.SetLineColour 4
							case "Green-Yellow"
							ns.SetLineStyle 5
							case else
						end select
					else
						if lista = wire.GetColourDescription then
						else
						ns.SetLineStyle 4
						ns.SetLineColour 28				
						end if
					end if
				
				
				end if				
				Next	
			Next	
		Next
	Next