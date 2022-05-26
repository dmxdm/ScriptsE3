
	
	Set e3 = CreateObject("CT.Application")
    Set prj = e3.CreateJobObject
    Set net = prj.CreateNetObject
    Set ns = prj.CreateNetSegmentObject
    Set sht = prj.CreateSheetObject
    Set EmbSht = prj.CreateSheetObject
	set Sym = prj.CreateSymbolObject
	set txt = prj.CreateTextObject
	
    e3.ClearOutputWindow	
	

		
	Dim z, cnt, x, y, EShtIds
	Dim LineLength, xat, yat
	Dim WireNumberSymbol
	
	
	Const SYM_VERT		= "SegmentLength_vert"
	Const SYM_HORI		= "SegmentLength_hori"
	
	minDist = 2 * prj.GetAltGridSize

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
					ns.SetManufacturingLength(ns.GetSchemaLength)		
					
					symtextcnt = ns.GetSymbolIds (symtextids)
					For c = 1 To symtextcnt
					Sym.SetId symtextids(c)
					Sym.delete
					
					Next							

				Next	
			Next
		Else
		End If
	Next
	
	
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
					
					cnt = ns.GetLineSegments( EShtIds, x, y )
					
						If CInt(x(1)) = CInt(x(2)) Then				' [16rk]
							LineLength = Abs(y(1) - y(2))
							WireNumberSymbol = SYM_VERT
						ElseIf CInt(y(1)) = CInt(y(2)) Then				' [16rk]
							LineLength = Abs(x(1) - x(2))
							WireNumberSymbol = SYM_HORI

						Else
						
						LineLength = Abs(x(1) - x(2))
						WireNumberSymbol = SYM_HORI
						
						End If
					
					If LineLength >= minDist Then
						xat = (x(2) + ((x(1) - x(2)) / 2))
						yat = (y(2) + ((y(1) - y(2)) / 2))
																	
						Sym.Load WireNumberSymbol, "1"
						Sym.Place EShtIds, xat, yat, 0
						
						textcnt = Sym.GetTextIds(textIds)
						
						For f = 1 To textcnt
						txt.SetId textIds(f)
						txt.SetLocking 0
												
						'MsgBox	"X1: " & x(1) & "    "  & "X2: " & x(2) & vbcrlf & "Y1: " & y(1) & "    "  & "Y2: " & y(2)		
						If CInt(x(1)) = CInt(x(2)) Then				' [16rk]
							txt.SetSchemalocation xat-((ns.GetOuterDiameter*1.1)/2), yat
						ElseIf CInt(y(1)) = CInt(y(2)) Then				' [16rk]
							txt.SetSchemalocation xat, yat+((ns.GetOuterDiameter*1.1)/2)
						ElseIf CInt(y(1)) > CInt(y(2)) and CInt(x(1)) > CInt(x(2)) OR CInt(y(1)) < CInt(y(2)) and CInt(x(1)) < CInt(x(2)) Then				' [16rk]
							txt.SetSchemalocation xat-ns.GetOuterDiameter/1.5, yat+ns.GetOuterDiameter/1.5
						Else				
							txt.SetSchemalocation xat+ns.GetOuterDiameter/1.5, yat+ns.GetOuterDiameter/1.5
						End If
							
						Next

					End If

				Next	
			Next
		Else
		 e3.PutMessage "Sheet " & sht.getname & " is not a Formboard Sheet"
		End If
	Next
	
	
    Set e3 = Nothing
	Set e3 = Nothing
    Set prj = Nothing
    Set net = Nothing
    Set ns = Nothing
    Set sht = Nothing
    Set EmbSht = Nothing
	set Sym = Nothing
	set Sym2 = Nothing



