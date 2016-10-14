Sub ConvertTextNumberToNumber()
    For Each WS In Sheets
        On Error Resume Next
        For Each r In WS.UsedRange.SpecialCells(xlCellTypeConstants)
            If IsNumeric(r) Then r.Value = Val(r.Value)
        Next
    Next
End Sub
