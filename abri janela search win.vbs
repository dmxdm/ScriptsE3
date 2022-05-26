'Const WINDOW_HANDLE = 0
'Const NO_OPTIONS = 0
'Set objShell = CreateObject("Shell.Application")
'Set objFolder = objShell.BrowseForFolder (WINDOW_HANDLE, "Select a folder:", NO_OPTIONS)     
'Set objFolderItem = objFolder.Self
'strPath = objFolderItem.Path
'objShell.Explore strPath



Option Explicit

Dim strFile

strFile = SelectFile( )

Function SelectFile( )
    ' Cannot define default starting folder.
    '           Uses last folder used with MSHTA.EXE stored in Binary in [HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32].
    '           Dialog title says "Choose file to upload".
    Dim objExec, strMSHTA, wshShell

    SelectFile = ""

    ' For use in HTAs as well as "plain" VBScript:
   ' strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
   '         & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
   '          & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
    ' For use in "plain" VBScript only:
     strMSHTA = "mshta.exe ""about:<input type=file id=FILE>" _
              & "<script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
              & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>"""

    Set wshShell = CreateObject( "WScript.Shell" )
    Set objExec = wshShell.Exec( strMSHTA )

    SelectFile = objExec.StdOut.ReadLine( )

    Set objExec = Nothing
    Set wshShell = Nothing
End Function