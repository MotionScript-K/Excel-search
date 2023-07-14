Sub RunScript()
    Dim source As Workbook
    Set source = ActiveWorkbook
    
    Dim lastRow As Long
    lastRow = source.ActiveSheet.Cells(source.ActiveSheet.Rows.Count, "A").End(xlUp).Row
    
    Dim x As Long
    For x = 2712 To lastRow
        Dim targetVal As Variant
        targetVal = source.ActiveSheet.Range("A" & x).Value
        
        On Error Resume Next
        Dim textFinder As Range
        Set textFinder = source.ActiveSheet.Range("D:D").Find(What:=targetVal, LookAt:=xlWhole, MatchCase:=False)
        On Error GoTo 0
        
        If Not textFinder Is Nothing Then
            source.ActiveSheet.Range("B" & x).Value = "YES"
        Else
            source.ActiveSheet.Range("B" & x).Value = "NO"
        End If
    Next x
End Sub
