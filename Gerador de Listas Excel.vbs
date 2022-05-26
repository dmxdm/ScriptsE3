
	Set e3 = CreateObject("CT.Application")
	Set Job = e3.CreateJobObject
	Set Con = Job.CreateConnectionObject
	Set Wire = Job.CreatePinObject
	Set Dev = Job.CreateDeviceObject
	Set Dev2 = Job.CreateDeviceObject
	Set Pin = Job.CreatePinObject
	Set Pin2 = Job.CreatePinObject
	Set Cav= Job.CreateCavitypartObject

	dim lista()
	
	TEXTO = "INFORME A OPCAO DESEJADA" & vbCrlf & "1 - Lista de Tag" & vbCrlf & "2 - Lista De/Para-Fio"
	MODE = InputBox(TEXTO)
	Select Case MODE
		Case "1"
			TagList
		Case "2"
			DeParaList
		Case ""
			MsgBox "Saindo."
		Case Else
			MsgBox "Opcao invalida" & ", saindo."
	End Select
	
	
Function TagList
	
	Job.GetAllDeviceIds connIds
	ReDim lista(UBound(connIds))
	
	For i = 1 to UBound(connIds)
		Dev.SetId connIds(i)
		If InStr(Dev.GetName, "Fios") > 0 then
		Exit for			
		End If
	lista(i) = " " & Dev.GetName
	Next
	txtFileName  = Job.GetPath & Job.GetName  & ".xlsx"
	WriteFile lista, txtFileName
	
	
End Function
	
	
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
			lista(i) = " " & Dev.GetName & ":" & Pin.GetName & "/" & Dev2.GetName & ":" & Pin2.GetName & "-" & Wire.GetName			
	
	lista(i) = " " & _
	Wire.GetName _
	& "-" & Wire.GetColourDescription _
	& "-" & Wire.GetCrossSection _
	& "/" & _
	_
	Dev.GetName _
	& ":" & Pin.GetName _
	& "-" & Pin.GetFitting _
	& "-" & getcavitypartfromcable( Pin.GetId, Wire.GetId, 2) _
	& "/" & _
	_
	Dev2.GetName _
	& ":" & Pin2.GetName _ 
	& "-" & Pin2.GetFitting _
	& "-" & getcavitypartfromcable( Pin2.GetId, Wire.GetId, 2) _
	& "/" & _
	_
	Wire.GetLength _	
	& "mm"
	Next
	Next
	txtFileName  = Job.GetPath & Job.GetName  & ".xlsx"
	WriteFile lista, txtFileName
	
End Function
	
Sub WriteFile(List, FileToSave)

	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False
	Set WorkBook = objExcel.Workbooks.Add()
	
	currentLine = 1
	For line = 1 To UBound(List)
		objExcel.Cells(currentLine,1).Value = List(line)
		objExcel.Columns.Autofit
		currentLine = currentLine + 1
	Next
	
	objExcel.Visible = True

	Set WorkBook = Nothing
	Set objExcel = Nothing

End Sub

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
