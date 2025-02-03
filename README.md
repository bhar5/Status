# Status


Sub CreateHyperlinks()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim searchRange As Range, foundCell As Range
    Dim sourceColumn As String, targetColumn As String
    
    ' Set the worksheet
    Set ws = ActiveSheet
    
    ' Define source and target columns (adjust as needed)
    sourceColumn = "A" ' Column where you want to create hyperlinks
    targetColumn = "B" ' Column where related values are found

    ' Get last row in the sheet
    lastRow = ws.Cells(ws.Rows.Count, sourceColumn).End(xlUp).Row
    
    ' Loop through each cell in the source column
    For i = 2 To lastRow ' Assuming headers in row 1
        ' Define the search range in the target column
        Set searchRange = ws.Range(targetColumn & "2:" & targetColumn & lastRow)
        
        ' Find matching value in the target column
        Set foundCell = searchRange.Find(ws.Cells(i, sourceColumn).Value, LookAt:=xlWhole)
        
        ' If a match is found, create a hyperlink
        If Not foundCell Is Nothing Then
            ws.Hyperlinks.Add Anchor:=ws.Cells(i, sourceColumn), _
                Address:="", SubAddress:="'" & ws.Name & "'!" & foundCell.Address, _
                TextToDisplay:=ws.Cells(i, sourceColumn).Value
        End If
    Next i
    
    MsgBox "Hyperlinks Created Successfully!", vbInformation
End Sub
