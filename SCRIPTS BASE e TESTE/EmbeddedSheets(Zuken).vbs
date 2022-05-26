	dim e3: Set e3 = ConnectToE3()
	Dim Prj: Set Prj = CreateObject("CT.Job")
	Dim Sht: Set Sht = CreateObject("CT.Sheet")
	Dim EmbSht: Set EmbSht = CreateObject("CT.Sheet")
	
	Dim msg
	msg = Time() & " === Start ..." & Right( WScript.ScriptFullName, 40 )
	msg = msg & " " & String( 72-Len(msg), "=" )
	e3.PutInfo 0, msg
	
	shtCnt = Prj.GetSheetIds(ShtIds)
	For s = 1 To ShtCnt
		Sht.SetId ShtIds(s)
		If ( Sht.GetEmbeddedSheetIds(EShtIds) > 0) Then
			EmbSht.SetId EShtIds(1)
			If(EmbSht.IsFormboard = 1) Then 'Here you can check it with the active embedded sheet !!
				MsgBox "IsFormboard: " & Sht.GetName
			End If
		End If
	Next

	
	msg = Time() & " === End ..." & Right( WScript.ScriptFullName, 40 )
	msg = msg & " " & String( 72-Len(msg), "=" )
	e3.PutInfo 0, msg
	
'==================================================================================================
' Tools...
'--------------------------------------------------------------------------------------------------
Function ConnectToE3

    Dim strComputer, objWMIService, colItems, ProcessCnt, objItem
    Dim disp, viewer, lst, e3Obj

	if InStr(WScript.FullName, "E³") Then
		set ConnectToE3 = WScript										' internal -> connect directly
		
		
		objArg = Wscript.ScriptArguments()		
		cntArg = UBound(objArg) + 1
	else
    	On Error Resume Next											' to skip error, if no dispatcher is installed
		Set disp   = CreateObject("CT.Dispatcher")        				' external
		Set viewer = CreateObject("CT.DispatcherViewer")
		On Error GoTo 0

		Set ConnectToE3 = Nothing

	    If IsObject(disp) Then											' test if E3.Dispatcher is installed
	        ProcessCnt = disp.GetE3Applications(lst)					' read active E3 processes
	        If ProcessCnt > 1 Then										' more than 1 process, ask for the project to connect
	            If viewer.ShowViewer(e3Obj) = True Then												' display dispatcher interface to select process
	                Set ConnectToE3 = e3Obj
				Else
					wscript.quit
				End If
	        Else
	            Set ConnectToE3 = CreateObject("CT.Application")
	        End If
	    Else
			strComputer = "."																	' dispatcher not installed
			set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
			set colItems      = objWMIService.ExecQuery("Select * from Win32_Process",,48)
			ProcessCnt = 0
			for each objItem in colItems
				if InStr(objItem.Caption, "E3.series") then ProcessCnt = ProcessCnt + 1
			next
			set objWMIService = Nothing
			set colItems      = Nothing
			if ProcessCnt > 1 then
				MsgBox  "More than one E3-Application running. Script can't run as external program." & vbCrLf & _
						"Please close the other E3-Applications.", 48
				WScript.Quit
			else
				set ConnectToE3 = CreateObject ("CT.Application")		' external
			end if
        End If
		Set disp   = nothing
		Set viewer = Nothing
		Set objArg = WScript.Arguments
		cntArg = WScript.Arguments.Count
	end If	
end function


'--------------------------------------------------------------------------------------------------
Function DeviceName( id )
	dim dev: set dev = prj.CreateDeviceObject
	dev.SetId id
	DeviceName = dev.GetAssignment & dev.GetLocation & dev.GetName
End Function

'--------------------------------------------------------------------------------------------------
Function PinName( id )
	Dim pin: Set pin = prj.CreatePinObject
	pin.SetId id
	PinName = DeviceName(id) & ":" & pin.GetName
End Function

'--------------------------------------------------------------------------------------------------
Function IsNothing(var)

	IsNothing = true
	If( TypeName(var) = "Nothing" )	Then exit function
	If( TypeName(var) = "Empty" )	Then exit function
	If( TypeName(var) = "Null" )	Then exit function
	If( TypeName(var) = "Unknown" )	Then exit function
	If( TypeName(var) = "Error" )	Then exit function
	
	IsNothing = false
end function

