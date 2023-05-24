Sub ETLResearch()
    Dim fieldsSheet As Worksheet
    Dim airlinesSheet As Worksheet
    Dim routesSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim airlinesDict As Object
    Dim reasonDict As Object
    Dim airline As Variant
    Dim reason As String
    Dim numClosedRoutes As Long
    Dim i As Long
    
   
    Set fieldsSheet = ThisWorkbook.Sheets("airports")
    
   
    Set airlinesSheet = ThisWorkbook.Sheets("airlines")
    
   
    Set routesSheet = ThisWorkbook.Sheets("routes")
    
    
    Set targetSheet = ThisWorkbook.Sheets("Results")
    
    
    Set airlinesDict = CreateObject("Scripting.Dictionary")
    Set reasonDict = CreateObject("Scripting.Dictionary")
    
  
    lastRow = airlinesSheet.Cells(airlinesSheet.Rows.Count, 1).End(xlUp).Row
    
   
    For i = 2 To lastRow
        airline = airlinesSheet.Cells(i, 1).Value
        reason = airlinesSheet.Cells(i, 2).Value
        
        airlinesDict(airline) = 0
        reasonDict(airline) = reason
    Next i
    
   
    lastRow = routesSheet.Cells(routesSheet.Rows.Count, 1).End(xlUp).Row
  
    For i = 2 To lastRow
        airline = routesSheet.Cells(i, 3).Value
        
        If Not airlinesDict.Exists(airline) Then
            airlinesDict(airline) = 1
        Else
            airlinesDict(airline) = airlinesDict(airline) + 1
        End If
    Next i
    
    
    targetSheet.Cells.ClearContents
    targetSheet.Cells(1, 1).Value = "Airline Name"
    targetSheet.Cells(1, 2).Value = "Closed Lines"
    targetSheet.Cells(1, 3).Value = "Reason"
    
    i = 2
       For Each airline In airlinesDict
        numClosedRoutes = airlinesDict(airline)
        reason = reasonDict(airline)
        
        targetSheet.Cells(i, 1).Value = airline
        targetSheet.Cells(i, 2).Value = numClosedRoutes
        targetSheet.Cells(i, 3).Value = reason
        
        i = i + 1
    Next airline
    
    
    With targetSheet.Range("A1:C" & targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
    
    
    With targetSheet.Range("A1:C1").Interior
        .Color = RGB(215, 215, 215)
        .Pattern = xlSolid
    End With
    
   
    targetSheet.Columns("A:C").AutoFit
    
  
    targetSheet.Cells(1, 1).Value = "Airline Name"
    targetSheet.Cells(1, 2).Value = "Closed Lines"
    targetSheet.Cells(1, 3).Value = "Reason"
End Sub

