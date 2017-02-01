Function ConcatToLeft(ByVal cell_range As Range, _
                    Optional ByVal seperator As String) As String

Dim cell As Range
Dim newString As String

Set cell = cell_range.Cells


Do While cell.Value <> Empty And cell.Column < 200
    If newString <> "" Then
        newString = newString & "|" & cell.Value
    Else
        newString = cell.Value
    End If
    
    Set cell = cell.Offset(0, 1)
    
Loop

ConcatToLeft = newString
