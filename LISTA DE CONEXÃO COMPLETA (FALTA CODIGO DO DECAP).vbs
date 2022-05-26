
	Set e3 = CreateObject("CT.Application")
	Set Job = e3.CreateJobObject
	Set Con = Job.CreateConnectionObject
	Set Wire = Job.CreatePinObject
	Set Dev = Job.CreateDeviceObject
	Set Dev2 = Job.CreateDeviceObject
	Set Dev3 = Job.CreateDeviceObject
	Set Pin = Job.CreatePinObject
	Set Pin2 = Job.CreatePinObject
	Set Cav= Job.CreateCavitypartObject
	Set Comp= Job.CreateComponentObject
	Set Sig= Job.CreateSignalObject
	Set db = CreateObject( "ADODB.Connection" )

	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False
	Set WorkBook = objExcel.Workbooks.Add()

			objExcel.Cells(1,1).Value = "NOME DO FIO"
			objExcel.Cells(1,2).Value = "COR"
			objExcel.Cells(1,3).Value = "SECAO"
			objExcel.Cells(1,4).Value = "DO CONECTOR"
			objExcel.Cells(1,5).Value = "DO PINO"
			objExcel.Cells(1,6).Value = "DO TERMINAL"
			objExcel.Cells(1,7).Value = "DO SELO"
			objExcel.Cells(1,8).Value = "PARA CONECTOR"
			objExcel.Cells(1,9).Value = "PARA PINO"
			objExcel.Cells(1,10).Value = "PARA TERMINAL"
			objExcel.Cells(1,11).Value = "PARA SELO"
			objExcel.Cells(1,12).Value = "COMPRIMENTO"
			objExcel.Cells(1,13).Value = "CABO"
			objExcel.Cells(1,14).Value = "DECAP A"
			objExcel.Cells(1,15).Value = "DECAP B"
			objExcel.Cells(1,16).Value = "SINAL"
			
	DeParaList
	
Function DeParaList
	
	Job.GetAllConnectionIds connIds
	ReDim lista(UBound(connIds))
	For i = 1 to UBound(connIds)
		Con.SetId connIds(i)	
		Con.GetCoreIds coreIds
		For j = 1 to UBound(coreIds)
			Wire.SetId coreIds(j)
			Pin.SetId Wire.GetEndPinId (1,ret)
			Pin2.SetId Wire.GetEndPinId (2,ret)
			Dev.SetId Pin.GetId
			Dev2.SetId Pin2.GetId			
			Dev3.SetId Wire.GetId
			Sig.SetId Wire.GetId
			Comp.SetId Wire.GetId
			
			objExcel.Cells(i+1,1).Value = Wire.GetName
			objExcel.Cells(i+1,2).Value = Wire.GetColourDescription
			objExcel.Cells(i+1,3).Value = Wire.GetCrossSection
			objExcel.Cells(i+1,4).Value = Dev.GetName
			objExcel.Cells(i+1,5).Value = Pin.GetName
			objExcel.Cells(i+1,6).Value = Pin.GetFitting	
			
			db.Open( e3.GetComponentDatabase )
			Set rs = db.Execute("SELECT AttributeValue FROM ComponentAttribute WHERE AttributeName= 'COMPRIMENTO_DECAPE' and Entry= '"& Pin.GetFitting &"' order by Entry")
			Do Until rs.EOF
			objExcel.Cells(i+1,14).Value = rs(1)
			rs.MoveNext
			Loop
			db.Close	
			
			objExcel.Cells(i+1,7).Value = getcavitypartfromcable( Pin.GetId, Wire.GetId, 2)			
			objExcel.Cells(i+1,8).Value = Dev2.GetName
			objExcel.Cells(i+1,9).Value = Pin2.GetName
			objExcel.Cells(i+1,10).Value = Pin2.GetFitting
			
			db.Open( e3.GetComponentDatabase )
			Set rs = db.Execute("SELECT AttributeValue FROM ComponentAttribute WHERE AttributeName= 'COMPRIMENTO_DECAPE' and Entry= '"& Pin2.GetFitting &"' order by Entry")
			Do Until rs.EOF
			objExcel.Cells(i+1,15).Value = rs(1)
			rs.MoveNext
			Loop
			db.Close
			
			objExcel.Cells(i+1,11).Value = getcavitypartfromcable( Pin2.GetId, Wire.GetId, 2)			
			objExcel.Cells(i+1,12).Value = Wire.GetLength
			objExcel.Cells(i+1,13).Value = Dev3.GetName
			objExcel.Cells(i+1,16).Value = Sig.GetName
			

	
	Next
	Next
	
	objExcel.Columns.Autofit
	objExcel.Columns.HorizontalAlignment = -4108
	objExcel.Visible = True
	objExcel.ActiveWorkbook.SaveAs Job.GetName  & "-LISTA.xlsx"
	
	Set WorkBook = Nothing
	Set objExcel = Nothing

	
End Function
	

Function getcavitypartfromcable( pinid, core, cavtype)						'[20ps]
	
	Dim cavpart, pincavs, pincavcnt, i
	Dim corcavs, corcavcnt, j 
	Dim plugs, plugcnt, k
	
	Pin.SetId pinid 
	pincavcnt =  Pin.getcavitypartids(pincavs, cavtype)				'find all cavities from pin
	
	If cavtype = 1 Then								'if pin terminals are searched
		plugcnt = Pin.getcavitypartids(plugs, 3)				'also the cavityplugs should be listed
		For k = 1 To plugcnt
			ReDim Preserve pincavs(UBound(pincavs) + 1)
			pincavs(UBound(pincavs)) = plugs(k)				'therefore added to array
		next
		pincavcnt = pincavcnt + plugcnt
	End If
	
	Pin.setid core 
	corcavcnt = Pin.GetCavityPartIds(corcavs, cavtype)				'find all cavities from core
	 
	If cavtype = 1 Then								'add plugs if pinterminals are searched
		plugcnt = Pin.getcavitypartids(plugs, 3)
		For k = 1 To plugcnt
			ReDim Preserve corcavs(UBound(corcavs) + 1)
			corcavs(UBound(corcavs)) = plugs(k)
		next
		corcavcnt = corcavcnt + plugcnt
	End If
	
	For i = 1 To pincavcnt								'compare the cavities from pin and core with eachother
		cavpart = pincavs(i)
		For j = 1 To corcavcnt
			If corcavs(j) = cavpart Then  
				Cav.setid cavpart
				getcavitypartfromcable = cav.getvalue			'if equeal it should be listed
			End If
	    Next
	Next

End Function
