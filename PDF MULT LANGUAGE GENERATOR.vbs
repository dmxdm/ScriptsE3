

Set App = CreateObject( "CT.Application" )
Set Prj = App.CreateJobObject
Set Sht = Prj.CreateSheetObject

shtcnt = prj.GetSheetIds( shtids )
Dim NewLanguages(1)

If( shtcnt = 0 ) Then

MsgBox "Não há folha no projeto!"

Else



NewLanguages(1) = "Português"
prj.SetLanguages NewLanguages
ret = prj.ExportPdf(Prj.GetPath & prj.getname & "_POR.pdf", shtids, &h200 +&h7, "xyz" ) 'Version 7 + not modifiable

NewLanguages(1) = "American English"
prj.SetLanguages NewLanguages
ret = prj.ExportPdf(Prj.GetPath & prj.getname & "_ENG.pdf", shtids, &h200 +&h7, "xyz" ) 'Version 7 + not modifiable

NewLanguages(1) = "Russian"
prj.SetLanguages NewLanguages
ret = prj.ExportPdf(Prj.GetPath & prj.getname & "_RUS.pdf", shtids, &h200 +&h7, "xyz" ) 'Version 7 + not modifiable

NewLanguages(1) = "Castellano"
prj.SetLanguages NewLanguages
ret = prj.ExportPdf(Prj.GetPath & prj.getname & "_SPN.pdf", shtids, &h200 +&h7, "xyz" ) 'Version 7 + not modifiable

End If