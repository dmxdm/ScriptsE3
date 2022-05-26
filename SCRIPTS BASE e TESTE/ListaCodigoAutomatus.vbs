' Description: Bill of Material (Excel)
'
' Changes:
'	17.02.2000 CIM-Team	Initial Version
'	10.08.2002 CIM-Team	Multi language
'	10.08.2002 CIM-Team	Add additional parts and fitting parts to the list
'	17.04.2003 CIM-Team	Add PartNumber functionality
'	30.06.2003 CIM-Team	Changes for version 2003, work with template file, new calculation for terminals
'	18.02.2004 CIM-Team	Handling of Articlenumber for Components / Additional parts / Fitting parts
'				Device:			1. Component - ArticleNumber
'								2. if empty: Entry - Name
'								3. if empty: Device - PartNumber
'				AdditionalPart:	1. Component - ArticleNumber
'								2. if empty: Entry - Name
'				FittingPart:	1. Component - ArticleNumber
'								2. if empty: Entry - Name
'	15.04.2004 CIM-Team	Changes for ORACLE
'	18.08.2004 CIM-Team	Added control of internal/external execution
'	24.11.2004 CIM-Team	Added structuring according to assignment and location
'	13.03.2005 CIM-Team	Added additional device attributes for dynamic devices
'	04.04.2005 CIM-Team Added portuguese strings
'	16.07.2015 Gianfranco -> Added Double quotes on COD_AUTOM in query string. This avoid problems when have AddParts.
'	7.08.2015 v1.5 Gianfranco -> Changed position to show Article number on AddParts, now are listed on the Codigo Column
'	14.03.2018 v1.7 Gianfranco 	-> Add List Child Feature.
'															-> Remove Extra Lines and Columns not used
'															-> Fixed Add Part and fitting parts Code_Auttom listings
'	14.03.2018 v1.8 Gianfranco 	-> Fix Add part
'	14.03.2018 v1.9 Gianfranco 	-> Fix Internal Script
'	14.03.2018 v1.9 Gianfranco 	-> Get Att from device too
'	14.03.2018 v1.91 Gianfranco -> Project Assemblies always list children
'	05.04.2018 v1.93 Gianfranco -> Include Add parts for ALL Devics (even not listed ones). Fix Mounting Rail Listing problem.
'	20.04.2018 v1.94 Gianfranco -> Include Devices column
'	- EOH -	
'...

LIST_MOUNT_CHILDS_ATTRIBUTE = "ListarFilhos"


' Connect to application and declare object variables to call methods later
Set App      = ConnectToE3
set Job      = App.CreateJobObject
set Dev      = Job.CreateDeviceObject
set DevPart  = Job.CreateDeviceObject
set Cmp      = Job.CreateComponentObject
set Pin      = Job.CreatePinObject
set Att      = Job.CreateAttributeObject
set File     = CreateObject( "Scripting.FileSystemObject" )
set WshShell = CreateObject("WScript.Shell")

'currentDirectory    = WshShell.CurrentDirectory

'msgbox 	

FileName = App.GetInstallationPath &  "reports\Parts_Automatus.xlt"


Description   = "Description"									' Define component attributes to read
COD_AUTOM     = "COD_AUTOM"
AddPart	      = "AdditionalPart"
PartNumber    = "PartNumber"
ArticleNumber = "ArticleNumber"
Description2  = "Description_Device"									' Define component attributes to read
COD_AUTOM2     = "COD_AUTOM_Device"

message  = """" & App.GetInstallationPath & "scripts\message.vbs" & """"
language = App.GetInstallationLanguage

select case language
	case "01"
		text1 = "Pos."
		text2 = "Número"
		text3 = "No.-Ident."
		text4 = "Descrição"
		text5 = "Código"
		text6 = "Dispositivos"
		text7 = "Lista de Material"
		text8 = "_LdM"
		text9 = "<no higher level assignemnt>"
		text10 = "<no location>"
	case "44"
		text1 = "Pos."
		text2 = "Número"
		text3 = "No.-Ident."
		text4 = "Descrição"
		text5 = "Código"
		text6 = "Dispositivos"
		text7 = "Lista de Material"
		text8 = "_LdM"
		text9 = "<no higher level assignemnt>"
		text10 = "<no location>"
	case "49"
		text1 = "Pos."
		text2 = "Anzahl"
		text3 = "Ident-Nr."
		text4 = "Beschreibung"
		text5 = "Hersteller"
		text6 = "Betriebsmittelkennzeichen"
		text7 = "Stückliste"
		text8 = "_BOM"
		text9 = "<kein Anlagenkennzeichen>"
		text10 = "<kein Ortskennzeichen>"
	case "33"
		text1 = "Pos."
		text2 = "Numéro"
		text3 = "No. d'ident."
		text4 = "Description"
		text5 = "Fournisseur"
		text6 = "Noms d'appareil"
		text7 = "Liste des pièces"
		text8 = "_BOM"
		text9 = "<no higher level assignemnt>"
		text10 = "<no location>"
	case "34"
		text1 = "Pos."
		text2 = "Cant."
		text3 = "Referencia"
		text4 = "Descripción"
		text5 = "Fabricante"
		text6 = "Nombre Dispositivo"
		text7 = "Lista de Materiales"
		text8 = "_BOM"
		text9 = "<no higher level assignemnt>"
		text10 = "<no location>"
	case "39"
		text1 = "Pos."
		text2 = "Quant."
		text3 = "Codice"
		text4 = "Descrizione"
		text5 = "Fornitore"
		text6 = "Sigla dispositivoo"
		text7 = "Lista materiali"
		text8 = "_BOM"
		text9 = "<no higher level assignemnt>"
		text10 = "<no location>"
	Case "55"
		text1 = "Pos."
		text2 = "Número"
		text3 = "No.-Ident."
		text4 = "Descrição"
		text5 = "Código"
		text6 = "Dispositivos"
		text7 = "Lista de Material"
		text8 = "_LdM"
		text9 = "<no higher level assignemnt>"
		text10 = "<no location>"
	case else
		text1 = "POSIÇÃO"
		text2 = "QTD"
		text3 = "REFERÊNCIA"
		text4 = "DESCRIÇÃO"
		text5 = "CÓDIGO"                                	
	                text6 = "DISPOSITIVO"
		text7 = "Lista de Material"
		text8 = ".xls"
		
end select

if Job.GetId = 0 then										' check connection to project
	WshShell.run message & " No_Project"
	set app      = nothing
	set WshShell = nothing
	wscript.quit
end if

JobName = Job.GetName 										' read project name 

nAlls = Job.GetAllDeviceIds (zDevIds)									' get all device ids
ReDim SortFeld1 (nAlls*10+1, 7)
ReDim SortFeld2 (nAlls*10+1, 7)
ReDim DevCount (nAlls*10+1)

n1 = -1
for n = 1 to nAlls										' read information for all

	Dev.SetId zDevIds(n)	
	
	listDevice = false
	
	If Dev.IsAssemblyPart = 1 Then 
		DevPart.SetId Dev.GetRootAssemblyId
		If IsProjectAssembly(DevPart) Or DevPart.GetComponentAttributeValue(LIST_MOUNT_CHILDS_ATTRIBUTE) = "1" Or DevPart.GetAttributeValue(LIST_MOUNT_CHILDS_ATTRIBUTE) = "1" Then
			listDevice = true
		End If
	
	ElseIf Dev.IsAssembly = 1 Then
		If Not IsProjectAssembly(Dev) And Dev.GetComponentAttributeValue(LIST_MOUNT_CHILDS_ATTRIBUTE) <> "1" And Dev.GetAttributeValue(LIST_MOUNT_CHILDS_ATTRIBUTE) <> "1" Then
			listDevice = true
		Else
		End If
	Else 
		listDevice = true
	End If	

	If Dev.GetAttributeValue("EXCLUDE") = "1" Then
		listDevice = false
	End If
			
	If listDevice Then
		n1 = n1 + 1
		SortFeld1(n1,1) = zDevIds(n)
		SortFeld1(n1,2) = Dev.GetName
		SortFeld1(n1,3) = Dev.GetComponentName        
		Cmp.SetId zDevIds(n)
		
		ArticleNum = Cmp.GetAttributeValue (ArticleNumber)
		if ArticleNum <> "" then SortFeld1(n1,3) = ArticleNum
		
		if SortFeld1(n1,3) = "" then
			SortFeld1(n1,3) = Dev.GetAttributeValue (PartNumber)
			if SortFeld1(n1,3) <> "" then SortFeld1(n1,1) = -zDevIds(n)
		end if
		SortFeld1(n1,4) = SortFeld1(n1,2) & SortFeld1(n1,3)
		SortFeld1(n1,5) = Dev.GetLocation
		SortFeld1(n1,6) = Dev.GetAssignment
		SortFeld1(n1,7) = Shift_Text_Left (Dev.GetAssignment, 20) & Dev.GetLocation
		DevCount(n1)    = 1
		Cmp.SetId zDevIds(n)
		
		TotalPins = Dev.GetAllPinIds (PinIdArray)						' check for fitting parts
		for n2 = 1 to TotalPins
			Pin.SetId PinIdArray(n2)
			FittingPart = Pin.GetFitting
			if FittingPart <> "" then
				n1 = n1 + 1
				SortFeld1(n1,1) = -1
				SortFeld1(n1,2) = Dev.GetName
				SortFeld1(n1,3) = FittingPart
				SortFeld1(n1,4) = Shift_Text_Left (SortFeld1(n1,2),20) & SortFeld1(n1,3)
				SortFeld1(n1,5) = Dev.GetLocation
				SortFeld1(n1,6) = Dev.GetAssignment
				SortFeld1(n1,7) = Shift_Text_Left (Dev.GetAssignment, 20) & Dev.GetLocation
				
				DevCount(n1)    = 1
			end if	
		next
	End If
	
	'List Add Parts of ALL Devices
	if Dev.GetAttributeValue (AddPart) <> "" then						' check for additional parts
		CmpAttNum = Dev.GetAttributeIds (AttIds)
		for n2 = 1 to CmpAttNum
			Att.SetId AttIds(n2)
			if Att.GetInternalName = AddPart then
				'msgbox "Add PArt: " & Att.GetValue
				n1 = n1 + 1
				SortFeld1(n1,1) = -1
				SortFeld1(n1,2) = Dev.GetName
				SortFeld1(n1,3) = Att.GetValue       
				SortFeld1(n1,4) = SortFeld1(n1,2) & SortFeld1(n1,3)
				SortFeld1(n1,5) = Dev.GetLocation
				SortFeld1(n1,6) = Dev.GetAssignment
				SortFeld1(n1,7) = Shift_Text_Left (Dev.GetAssignment, 20) & Dev.GetLocation
				DevCount(n1)    = 1
			end if
		next
	end if
next

ret = App.SortArrayByIndex (SortFeld1, n1+1, 8, 8, 5)						' sort by assignment/location and then device designation/component code

for n = 0 to n1-1
	if SortFeld1(n,3) = "" then								' flag devices without componets code
		SortFeld1(n,1) = 0
	elseif SortFeld1(n,4) = SortFeld1(n+1,4) then
		SortFeld1(n,1) = 0								' count componets in terminal blocks
		DevCount(n+1)  = DevCount(n+1) + DevCount(n)
	end if
next

nNew = -1
for n = 0 to n1											' ignore devices without componets code
	if SortFeld1(n,1) <> 0 then
		Dev.SetId SortFeld1(n,1)
		if SortFeld1(n,1) <= -1 or Dev.IsWireGroup = 0 then
			nNew = nNew + 1
			SortFeld2(nNew,1) = SortFeld1(n,1)
			SortFeld2(nNew,2) = SortFeld1(n,2)
			SortFeld2(nNew,3) = SortFeld1(n,3)
			SortFeld2(nNew,4) = DevCount(n)
			SortFeld2(nNew,5) = SortFeld1(n,5)
			SortFeld2(nNew,6) = SortFeld1(n,6)
			SortFeld2(nNew,7) = SortFeld1(n,7)
		end if
	end if
next

if nNew = -1 then
        WshShell.run message & " No_Components"
		set app      = nothing
		set WshShell = nothing
        wscript.quit
end if

ret = App.SortArrayByIndex (SortFeld2, nNew+1, 6, 6, 4)						' Sort by component code

for n = 0 to nNew										' count devices with same
	if SortFeld2(n,4) <> 1 then SortFeld2(n,2) = SortFeld2(n,2) & "(" & SortFeld2(n,4) & ")"
next

for n = 0 to nNew										' count devices with same
	if SortFeld2(n,3) = SortFeld2(n+1,3) and SortFeld2(n,5) = SortFeld2(n+1,5)then
		SortFeld2(n+1,2) = SortFeld2(n,2) & ", " & SortFeld2(n+1,2)
		SortFeld2(n+1,4) = SortFeld2(n,4) + SortFeld2(n+1,4)
		SortFeld2(n,1) = 0
	end if
next

if File.FileExists( FileName ) then								' check for existing file
	set ExcelApp = CreateObject("Excel.Application")
	ExcelApp.Visible = TRUE									' open EXCEL
	set Excel  = ExcelApp.WorkBooks.Open(FileName)
	set Excel  = ExcelApp.ActiveWorkBook.WorkSheets(1)
	Excel.Name = text7
	excelName  = Job.GetPath & Job.GetName  & text8
	if File.FileExists (excelName & ".xls") then File.DeleteFile excelName & ".xls"

	DatabaseDSN = App.GetComponentDatabase
	pos1 = instr(DatabaseDSN, "OraOLEDB")
	if pos1 = 0 then
		Oracle = false
	else
		Oracle = true
	end if
	set db = CreateObject("ADODB.Connection")
	db.Open (DatabaseDSN)

	Excel.Cells(1,1).Value = text5
	Excel.Cells(1,2).Value = text2	 							 ' Write head lines
	Excel.Cells(1,4).Value = text3
	Excel.Cells(1,3).Value = text4
	Excel.Cells(1,5).Value = text6	

	nline = 1
	oldAssignment = "xxxxxxx"
	nComp = 0
	for n = 0 to nNew									' loop for all components
		if Sortfeld2(n,1) <> 0 then
			nComp = nComp + 1
			if SortFeld2(n,6) <> oldAssignment And Trim(SortFeld2(n,6)) <> "" then
				oldAssignment = SortFeld2(n,6)
				nline = nline + 2
				Excel.Range("A"&nline&":E"&nline).Select
				ExcelApp.Selection.Font.Bold           = True
				ExcelApp.Selection.Font.Italic         = True
				ExcelApp.Selection.Interior.ColorIndex = 40
				ExcelApp.Selection.Interior.Pattern    = 1
				text = SortFeld2(n,6)
				if text = "" then text = text9
				Excel.Cells(nline,1).Value = text
				oldLocation   = "xxxxxxx"
			end if 
			if SortFeld2(n,5) <> oldLocation And Trim(SortFeld2(n,5)) <> "" then
				oldLocation = SortFeld2(n,5)
				nline = nline + 1
				Excel.Range("A"&nline&":E"&nline).Select
				ExcelApp.Selection.Font.Bold           = True
				ExcelApp.Selection.Font.Italic         = True
				ExcelApp.Selection.Interior.ColorIndex = 40
				ExcelApp.Selection.Interior.Pattern    = 1
				text = SortFeld2(n,5)
				if text = "" then text = text10
				Excel.Cells(nline,2).Value = text
			end if 
			
			if SortFeld2(n,1) <= -1 then						' Read information for additional parts directly from the database
			
				'MAIN DATA
				if Oracle then
					sql = "SELECT """ & Description & """, """ & COD_AUTOM & """, """ & ArticleNumber & """ FROM ""ComponentData"" WHERE ""ENTRY"" = '" & Replace(SortFeld2(n,3), "'", "''") & "'"
				Else
					'sql = "SELECT " & Description & ", " & ArticleNumber & " FROM ComponentData WHERE ENTRY = '" & Replace(SortFeld2(n,3), "'", "''") & "'"
					sql = "SELECT " & Description & ", '" & COD_AUTOM & "', " & ArticleNumber & " FROM ComponentData WHERE ENTRY = '" & Replace(SortFeld2(n,3), "'", "''") & "'"
				end If
				set rs = db.Execute(sql)					' rs(0) = Description rs(1) = COD_AUTOM
				ValueDescription = ""
				ValueEntry       = ""
				if NOT(rs.EOF) then						' Entry not found	
					if NOT(isNull(rs(0))) then ValueDescription = rs(0)	' Value is empty
					'if NOT(isNull(rs(2))) then ValueCOD_AUTOM    = rs(2)'rs(1)	' Value is empty
					if NOT(isNull(rs(2))) then ValueEntry       = rs(2)	' Value is empty
				end if
				
				'ATTRIBUTE DATA
				
				if Oracle then
					sql = "SELECT """ & "AttributeValue" & """ FROM ""ComponentAttribute"" WHERE ""ENTRY"" = '" & Replace(SortFeld2(n,3), "'", "''") & "'" & " AND ""AttributeName"" = '" & COD_AUTOM & "'"
				Else
					'sql = "SELECT " & Description & ", " & ArticleNumber & " FROM ComponentData WHERE ENTRY = '" & Replace(SortFeld2(n,3), "'", "''") & "'"
					sql = "SELECT AttributeValue FROM ComponentAttribute WHERE ENTRY = '" & Replace(SortFeld2(n,3), "'", "''") & "'" & " AND AttributeName = '" & COD_AUTOM & "'"
				end If
				
				set rs = db.Execute(sql)					' rs(0) = Description rs(1) = COD_AUTOM
				ValueCOD_AUTOM    = ""
				if NOT(rs.EOF) then						' Cod Autom not found	
					if NOT(isNull(rs(0))) then 
						ValueCOD_AUTOM    = rs(0)'rs(1)	' Value is empty
					End If
				end if
				
				if ValueEntry = "" then 
					ValueEntry = SortFeld2(n,3)
					'ValueCOD_AUTOM    = SortFeld2(n,3)
				end if
				if SortFeld2(n,1) < -1 then
					Dev.SetId -SortFeld2(n,1)
					if ValueDescription = "" then ValueDescription = Dev.GetAttributeValue( Description2 )
					if ValueCOD_AUTOM = "" then ValueCOD_AUTOM = Dev.GetAttributeValue( COD_AUTOM2 )
				end if
			else
				Dev.SetId SortFeld2(n,1)
				Cmp.SetId SortFeld2(n,1)
				ValueDescription = Cmp.GetAttributeValue( Description )		' write attribute
				ValueCOD_AUTOM    = Cmp.GetAttributeValue( COD_AUTOM )
				ValueEntry       = SortFeld2(n,3)
			end if
			nline = nline + 1
			'Excel.Cells(nline,1).Value = nComp 					' write position
			Excel.Cells(nline,1).Value = ValueCOD_AUTOM
			Excel.Cells(nline,2).Value = SortFeld2(n,4)
			Excel.Cells(nline,3).Value = ValueDescription				' write attribute
			Excel.Cells(nline,4).Value = ValueEntry				' write component type
			Excel.Cells(nline,5).Value = SortFeld2(n,2)	
			
			
			REM while Len(SortFeld2(n,2)) > 45						' write device names of the component
				REM kommaFound = InStr(38,SortFeld2(n,2),",")
				REM if kommaFound = 0 then
					REM devices        = SortFeld2(n,2)
					REM SortFeld2(n,2) = ""
				REM else
					REM devices        = Left(SortFeld2(n,2),kommaFound)
					REM SortFeld2(n,2) = Mid(SortFeld2(n,2),kommaFound+2)
				REM end if
				REM Excel.Cells(nline,6).Value = "'" & devices
				REM if SortFeld2(n,2) <> "" then nline = nline + 1
			REM wend
			REM if SortFeld2(n,2) <> "" then Excel.Cells(nline,6).Value = "'" & SortFeld2(n,2)
		end if
	next

	ExcelApp.ActiveWorkbook.SaveAs excelName

'	ExcelApp.Quit
	Set Excel    = Nothing
	Set ExcelApp = Nothing

	db.Close 
	set db = nothing
else
	WshShell.run message & " File_Not_Existing", 7, true
	App.PutInfo 0, FileName
end if

set app      = nothing
set WshShell = nothing
wscript.quit

' ----------------------------------------------------------------------------------------------
' check for several E3 processes and if process is running internally or externally

function ConnectToE3
	if InStr(WScript.FullName, "E³") then
		set ConnectToE3 = WScript								' internal
	else
		strComputer = "."
		set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
		set colItems      = objWMIService.ExecQuery("Select * from Win32_Process",,48)
		ProcessCnt = 0
		for each objItem in colItems
			if InStr(objItem.Caption, "E3.series") then ProcessCnt = ProcessCnt + 1
		next
		set objWMIService = Nothing
		set colItems      = Nothing
		if ProcessCnt > 1 then
			MsgBMsgBox  "More than one E3-Application running. Script can't run as external program." & vbCrLf &_
					"Please close the other E3-Applications.", 48
			WScript.Quit
		else
			set ConnectToE3 = CreateObject ("CT.Application")		' external
		end if
	end if
end function

' ------------------------------------------------------------------------------------
function Shift_Text_Left (Text, TotalLength)
	TextLength = len(Text)
	if TextLength < TotalLength then
		Shift_Text_Left = Text & SPACE(TotalLength-TextLength)
	else
		Shift_Text_Left = left(Text,TotalLength)
	end if
end function

' ----------------------------------------------------------------------------------------------

Function IsProjectAssembly(DevAssembly)
	
	If Trim(DevAssembly.GetComponentName) = "" Then
		IsProjectAssembly = True
	Else
		IsProjectAssembly = False
	End If
End Function
