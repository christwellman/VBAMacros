' src is the range to scan. It must be a single rectangular range (no multiselect).
' dst gives the offset where to paste. Should be a single cell.
' Pasted values will have shape N rows x 1 column, with unknown N.
' src and dst can be in different Worksheets or Workbooks.
Public Sub unique(src As Range, dst As Range)
    Dim cl As Collection
    Dim buf_in() As Variant
    Dim buf_out() As Variant
    Dim val As Variant
    Dim i As Long

    ' It is good practice to catch special cases.
    If src.Cells.Count = 1 Then
        dst.Value = src.Value   ' ...which is not an array for a single cell
        Exit Sub
    End If
    ' read all values at once
    buf_in = src.Value
    Set cl = New Collection
    ' Skip all already-present or invalid values
    On Error Resume Next
    For Each val In buf_in
        cl.Add val, CStr(val)
    Next
    On Error GoTo 0

    ' transfer into output buffer
    ReDim buf_out(1 To cl.Count, 1 To 1)
    For i = 1 To cl.Count
        buf_out(i, 1) = cl(i)
    Next

    ' write all values at once
    dst.Resize(cl.Count, 1).Value = buf_out

End Sub