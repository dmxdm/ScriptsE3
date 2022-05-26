
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
	Set Att= Job.CreateAttributeObject
	Set Sym= Job.CreateSymbolObject
	Set Txt= Job.CreateTextObject
	Set db = CreateObject( "ADODB.Connection" )
	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False
	Set WorkBook = objExcel.Workbooks.Add()

			objExcel.Cells(1,1).Value = "ETIQUETA"

	Job.GetSymbolIds connIds
	'dim lista
	j=1
	For i = 1 to UBound(connIds)
		Sym.SetId connIds(i)	
		
		if Sym.GetSymbolTypeName = "Etiqueta" then
			Sym.GetTextIds txtids	
			
			For k=1 to UBound(txtids)
			txt.SetId txtids(k)
			objExcel.Cells(j+1,1).Value = Txt.GetText
			Next
			
		j=j+1
		end if
	Next

	objExcel.Columns.Autofit
	objExcel.Columns.HorizontalAlignment = -4108
	objExcel.Visible = True
	objExcel.ActiveWorkbook.SaveAs Job.GetName  & "-LISTA-ETIQUETA.xlsx"
	
	Set WorkBook = Nothing
	Set objExcel = Nothing
	Set e3 = Nothing
	Set Job = Nothing
	Set Con = Nothing
	Set Wire = Nothing
	Set Dev = Nothing
	Set Dev2 = Nothing
	Set Dev3 = Nothing
	Set Pin = Nothing
	Set Pin2 = Nothing
	Set Cav= Nothing
	Set Comp= Nothing
	Set Sig= Nothing
	Set Att= Nothing
	Set Sym= Nothing
	Set Txt= Nothing
	Set db = Nothing