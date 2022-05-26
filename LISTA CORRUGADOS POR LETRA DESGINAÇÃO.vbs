	
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
	
	
			objExcel.Cells(1,1).Value = "POS."
			objExcel.Cells(1,2).Value = "CANT."
			objExcel.Cells(1,3).Value = "REFERENCIA"
			objExcel.Cells(1,4).Value = "DESCRIPCION"
			objExcel.Cells(1,5).Value = "FABRICANTE"
			objExcel.Cells(1,6).Value = "NOMBRE DISPOSITIVO"
	
	Job.GetAllDeviceIds connIds	
	
	For i = 1 to UBound(connIds)
		Dev.SetId connIds(i)
		Comp.SetId Dev.GetId
		
		If Comp.GetAttributeValue ("DeviceLetterCode") = "FD" then
			Dev.GetSymbolIds symbolIds,4
			
			For sidx = 1 to Ubound(symbolIds)
				Sym.SetId symbolIds(sidx)
				NetSegment.SetId symbolIds(sidx)
				comp_count = NetSegment.GetLength
				
					If Job.getMeasure()="MM" Then 'mm into M, inch remains inch
					comp_count = comp_count/1000
					End If
			Next
			
			objExcel.Cells(i+1,1).Value = i
			objExcel.Cells(i+1,2).Value = comp_count
			objExcel.Cells(i+1,3).Value = Comp.GetAttributeValue("ArticleNumber")
			objExcel.Cells(i+1,4).Value = Dev.GetComponentName
			objExcel.Cells(i+1,5).Value = Comp.GetAttributeValue("Supplier")
			objExcel.Cells(i+1,6).Value = Dev.GetName
			
		end if
			
	Next
	
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