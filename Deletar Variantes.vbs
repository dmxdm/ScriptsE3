Set e3Application = CreateObject( "CT.Application" ) 
Set job = e3Application.CreateJobObject()
Set connection = job.CreateConnectionObject()
Set var = job.CreateVariantObject
Set Wire = job.CreatePinObject

connectionCount = job.GetSelectedConnectionIds( connectionIds )        
 
If connectionCount > 0 Then
    For connectionIndex = 1 To connectionCount
        connection.SetId( connectionIds( connectionIndex ) )
        connection.GetCoreIds coreIds
			For j = 1 to UBound(coreIds)
				Wire.SetId coreIds(j)
				Wire.DeleteForced
			Next
		
    Next
End If
 
Set connection = Nothing
Set job = Nothing 
Set e3Application = Nothing
 
