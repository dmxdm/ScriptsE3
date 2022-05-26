	
	Set e3 = CreateObject("CT.Application")
	Set Job = e3.CreateJobObject
	Set Dev = Job.CreateDeviceObject
	Set Comp = Job.CreateComponentObject
	Set NetSegment = Job.CreateNetSegmentObject
	Set Sym = Job.CreateSymbolObject
	Set objExcel = CreateObject("Excel.Application")
	Set WorkBook = objExcel.Workbooks.Add()
	objExcel.Visible = False

	Dim symbolIds, sidx
	Dim NetSegment
	
	
			'objExcel.Cells(1,1).Value = "POS."
			objExcel.Cells(1,2).Value = "CANT."
			objExcel.Cells(1,3).Value = "REFERENCIA"
			objExcel.Cells(1,4).Value = "DESCRIPCION"
			objExcel.Cells(1,5).Value = "FABRICANTE"
			objExcel.Cells(1,6).Value = "NOMBRE DISPOSITIVO"
	
	Job.GetAllDeviceIds connIds	
		
	intRow = 1

	
	For i = 1 to UBound(connIds)
		Dev.SetId connIds(i)
		Comp.SetId Dev.GetId
		
			Dev.GetSymbolIds symbolIds,4
			
			If Comp.GetAttributeValue ("DeviceLetterCode") = "FD" or Comp.GetAttributeValue ("DeviceLetterCode") = "CR" then
				For sidx = 1 to Ubound(symbolIds)
				Sym.SetId symbolIds(sidx)
				NetSegment.SetId symbolIds(sidx)
				comp_count = NetSegment.GetLength
				
					If Job.getMeasure()="MM" Then 'mm into M, inch remains inch
					comp_count = comp_count/1000
					End If
				Next
			
			else 
				comp_count = 1
			end if
			
			
			If Dev.GetName <> "Hilos" then	
			
			'objExcel.Cells(intRow+1,1).Value = intRow
			objExcel.Cells(intRow+1,2).Value = comp_count
			objExcel.Cells(intRow+1,3).Value = Comp.GetAttributeValue("ArticleNumber")
			objExcel.Cells(intRow+1,4).Value = Dev.GetComponentName
			objExcel.Cells(intRow+1,5).Value = Comp.GetAttributeValue("Supplier")
			objExcel.Cells(intRow+1,6).Value = Dev.GetName
			
			intRow = intRow + 1
			
			End if
		
			
	Next
	
	
    objExcel.Range("A:F").AutoFilter
	objExcel.Columns.Autofit
	objExcel.Columns.HorizontalAlignment = -4108
	objExcel.Visible = True
	objExcel.ActiveWorkbook.SaveAs Job.GetName  & "-LISTA.xlsx"


	Set e3 = Nothing
	Set Job = Nothing
	Set Dev = Nothing
	Set Comp = Nothing
	Set NetSegment = Nothing
	Set Sym = Nothing
	Set objExcel = Nothing
	Set WorkBook = Nothing