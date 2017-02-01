Sub DigitalNameAndExpand()
    Dim newRange As Range
    Dim newString As String
    Dim digOpts() As String
    Dim newName As String
    

    Set newRange = Range("C2:C87")
    For Each Cell In newRange
        newString = Cell.Value
        
        digOpts = Split(newString, "|")
        newName = "AT_" & digOpts(0)
        Cell.Offset(0, 1).Value = digOpts(0)
        For i = 1 To UBound(digOpts, 1)
            newName = newName & "+" & digOpts(i)
            If digOpts(i) = "-" Then
                Cell.Offset(0, 1 + i).Value = "'-"
            Else
                Cell.Offset(0, 1 + i).Value = digOpts(i)
            End If
        Next i
        Cell.Offset(0, -1).Value = newName
    Next

End Sub
