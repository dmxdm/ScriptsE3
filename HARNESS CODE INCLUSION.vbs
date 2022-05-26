	
	Set e3 = CreateObject("CT.Application")
	Set Job = e3.CreateJobObject
	Set Dev = Job.CreateDeviceObject
	Set Comp = Job.CreateComponentObject
	Set Att= Job.CreateAttributeObject
	
	colUM = 3 ' COLLUMN NUMBER WHERE IT CAN FIND THE DEVICE NAME ''From Device Name''
	DATAcolUM = 1 ' COLLUMN NUMBER WHERE IS THE HARNESS COD FOR ''From Device Name''
	
	colDOIS = 7 ' COLLUMN NUMBER WHERE IT CAN FIND THE DEVICE NAME ''To Device Name''
	DATAcolDOIS = 1 ' COLLUMN NUMBER WHERE IS THE HARNESS COD FOR ''To Device Name''
	
	colTRES = 14 ' COLLUMN NUMBER WHERE IT CAN FIND THE DEVICE NAME ''Cable Name''
	DATAcolTRES = 1 ' COLLUMN NUMBER WHERE IS THE HARNESS COD FOR ''Cable Name''
	
	Row = 2 ' STARTING LINE COUNT
	
	attName = "Function" 'ATTIBUTE NAME
	
	Wh="-" ' CONSIDER THE " - " FROM E3 STANDARD DEVICE NAMING 

Dim strFile
strFile = SelectFile( )

	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open _
	(strFile)

	Job.GetAllDeviceIds connIds
	
	For i = 1 to UBound(connIds)
		Dev.SetId connIds(i)
		Comp.SetId Dev.GetId				
		intRow = Row ' LINE COUNT RESET		
		Do Until objExcel.Cells(intRow,1).Value = ""		
		If Dev.GetName = (Wh & objExcel.Cells(intRow, colUM).Value) then
			Dev.SetAttributeValue attName, objExcel.Cells(intRow, DATAcolUM).Value						
		End if
		If Dev.GetName = (Wh & objExcel.Cells(intRow, colDOIS).Value) then
			Dev.SetAttributeValue attName, objExcel.Cells(intRow, DATAcolDOIS).Value						
		End if
		If Dev.GetName = (Wh & objExcel.Cells(intRow, colTRES).Value) then
			Dev.SetAttributeValue attName, objExcel.Cells(intRow, DATAcolTRES).Value						
		End if
		intRow = intRow + 1		
		Loop				
	Next

	objExcel.Quit
	
	Set e3 = Nothing
	Set Job = Nothing
	Set Dev = Nothing
	Set Comp = Nothing
	Set Att= Nothing


Function SelectFile( )
    ' Cannot define default starting folder.
    '           Uses last folder used with MSHTA.EXE stored in Binary in [HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32].
    '           Dialog title says "Choose file to upload".
    Dim objExec, strMSHTA, wshShell

    SelectFile = ""

    ' For use in HTAs as well as "plain" VBScript:
     strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
             & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
             & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
    ' For use in "plain" VBScript only:
     'strMSHTA = "mshta.exe ""about:<input type=file id=FILE>" _
     '         & "<script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
     '         & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>"""

    Set wshShell = CreateObject( "WScript.Shell" )
    Set objExec = wshShell.Exec( strMSHTA )

    SelectFile = objExec.StdOut.ReadLine( )

    Set objExec = Nothing
    Set wshShell = Nothing
End Function