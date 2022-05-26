Set e3 = CreateObject("CT.Application")
Set Job = e3.CreateJobObject
Set Con = Job.CreateConnectionObject
Set Wire = Job.CreatePinObject
Set Dev = Job.CreateDeviceObject
Set Dev2 = Job.CreateDeviceObject
Set Dev3 = Job.CreateDeviceObject
Set Pin = Job.CreatePinObject
Set Pin2 = Job.CreatePinObject
Set Comp= Job.CreateComponentObject
Set Sig= Job.CreateSignalObject
Set Att= Job.CreateAttributeObject
Set db = CreateObject( "ADODB.Connection" )

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
SetDecapeArrasto

Next
Next


Function SetDecapeArrasto()

db.Open( e3.GetComponentDatabase )

Set decap = db.Execute( "SELECT AttributeValue FROM ComponentAttribute WHERE AttributeName= 'COMPRIMENTO_DECAPE' and Entry= '"& Pin.GetFitting &"' order by Entry" )
Set decap2 = db.Execute( "SELECT AttributeValue FROM ComponentAttribute WHERE AttributeName= 'COMPRIMENTO_DECAPE' and Entry= '"& Pin2.GetFitting &"' order by Entry" )
Set arrasto = db.Execute( "SELECT AttributeValue FROM ComponentAttribute WHERE AttributeName= 'COMPRIMENTO_ARRASTE' and Entry= '"& Pin.GetFitting &"' order by Entry" )
Set arrasto2 = db.Execute( "SELECT AttributeValue FROM ComponentAttribute WHERE AttributeName= 'COMPRIMENTO_ARRASTE' and Entry= '"& Pin2.GetFitting &"' order by Entry" )
Set CodTerm1 = db.Execute( "SELECT AttributeValue FROM ComponentAttribute WHERE AttributeName= 'Codigo_Marcopolo' and Entry= '"& Pin.GetFitting &"' order by Entry" )
Set CodTerm2 = db.Execute( "SELECT AttributeValue FROM ComponentAttribute WHERE AttributeName= 'Codigo_Marcopolo' and Entry= '"& Pin2.GetFitting &"' order by Entry" )
Set Selo1 = db.Execute( "SELECT AttributeValue FROM ComponentAttribute WHERE AttributeName= 'AdditionalPart' and Entry= '"& Pin.GetFitting &"' order by Entry" )
Set Selo2 = db.Execute( "SELECT AttributeValue FROM ComponentAttribute WHERE AttributeName= 'AdditionalPart' and Entry= '"& Pin2.GetFitting &"' order by Entry" )
Set CodSelo1 = db.Execute( "SELECT AttributeValue FROM ComponentAttribute WHERE AttributeName= 'Codigo_Marcopolo' and Entry= '"& Selo1(0) &"' order by Entry" )
Set CodSelo2 = db.Execute( "SELECT AttributeValue FROM ComponentAttribute WHERE AttributeName= 'Codigo_Marcopolo' and Entry= '"& Selo2(0) &"' order by Entry" )


Wire.GetAttributeIds Attids
For k = 1 to UBound(Attids)
Att.Setid Wire.SetAttributeValue("Decape1", decap(0))
Att.Setid Wire.SetAttributeValue("Decape2", decap2(0))
Att.Setid Wire.SetAttributeValue("Arraste1", arrasto(0))
Att.Setid Wire.SetAttributeValue("Arraste2", arrasto2(0))
Att.Setid Wire.SetAttributeValue("CodMarTerm1", CodTerm1(0))
Att.Setid Wire.SetAttributeValue("CodMarTerm2", CodTerm2(0))
Att.Setid Wire.SetAttributeValue("CodMarSelo1", CodSelo1(0))
Att.Setid Wire.SetAttributeValue("CodMarSelo2", CodSelo2(0))
next

db.Close

End Function