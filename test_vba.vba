Sub TestMacro()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Sheet1")

    ws.Range("A1").Value = "Hello"
    ws.Range("A2:B5").Sort

    ActiveWorkbook.Save
End Sub
