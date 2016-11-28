Sub fliterSheet4()
    Dim startRow As Integer
    Dim endRow As Integer
    Dim startArray() As Integer
    Dim lengthArray() As Integer
    
    startRow = 2
    Length = 1
    iindex = 0
    For Row = 3 To 46
        If (Cells(Row, 1) = Cells(startRow, 1)) Then
            Length = Length + 1
        Else
            ReDim Preserve startArray(iindex + 1)
            ReDim Preserve lengthArray(iindex + 1)
            startArray(iindex) = startRow
            lengthArray(iindex) = Length
            startRow = Row
            Length = 1
            iindex = iindex + 1
        End If
    Next Row

    startArray(iindex) = startRow
    lengthArray(iindex) = Length
    
    Dim path As String
    Dim originSheet As Sheets
    'originSheet = ThisWorkbook.Sheets(4)
    path = ThisWorkbook.path & "\" & "sheet4.xls"
    Workbooks.Add
    

For s = 0 To iindex
    Sheets(s + 1).Name = Cells(startArray(s), 1)
    i = 1
    For r = 0 To lengthArray(s) - 1
        Sheets(s + 1).Cells(i, 1) = Cells(startArray(s) + r, 2)
        i = i + 1
    Next r
    Sheets.Add After:=Sheets(s + 1)
    
Next s
    
    
    'ActiveWorkbook.SaveAs path, True
    
    
    
End Sub

