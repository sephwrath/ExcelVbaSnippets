' hello Lindsay - this is a function I wrote for you. Sorry I couldn't write it this morning but I've only just checked my messages.
' hope it's still useful
' use it like this... but only if you don't need the color to update...
'Private Sub Worksheet_Change(ByVal Target As Excel.Range)
'    ReflectCell Target, "A2", "Legend", "F4"
'    ReflectCell Target, "A2", "Cost Work Up", "B30"
'    ReflectCell Target, "A2", "Quote", "A14"
'End Sub
'
' if you do need the color to update use it like this...
'
'Dim oldCell As Range
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'    ReflectCell oldCell, "A2", "Legend", "F4"
'    Set oldCell = Target
'End Sub

Private Sub ReflectCell(Target As Excel.Range, sFrom As String, sWkTo As String, sCellTo As String, Optional col As Boolean = True)
    If Not (Target Is Nothing) Then
        If Not Intersect(Target, Range(sFrom)) Is Nothing Then
            Range(sFrom).Copy (Worksheets(sWkTo).Range(sCellTo))
            
            ' this bit copies the color as well based on what is passed to col - defaults to true so you only need to
            ' set it if you dont want perty colors.
            If col = True Then
                Worksheets(sWkTo).Range(sCellTo).Interior.color = Range(sFrom).Interior.color
            End If
        End If
    End If
End Sub
