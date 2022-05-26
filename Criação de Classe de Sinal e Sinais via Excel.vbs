Set e3     = CreateObject("CT.Application")
Set job    = e3.CreateJobObject
Set sig     = job.CreateSignalObject
Set sigcla = job.CreateSignalClassObject

Dim strFile
strFile = SelectFile( )

	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open _
	(strFile)
  
  
intRow = 1 	  
intCol = 1

Do Until objExcel.Cells(1,intCol).Value = ""		
			
		if intRow = 1 then
		sigcla.create(objExcel.Cells(intRow, intCol).Value)
		else
		sig.create(objExcel.Cells(intRow, intCol).Value)
		sigcla.AddSignalId sig.getid 
		End if
		intRow = intRow + 1		
		
		if objExcel.Cells(intRow,1).Value = "" then
		intRow = 1
		intCol = intCol + 1
		end if
		
Loop	  


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
	Set objWorkbook = Nothing
End Function


 
