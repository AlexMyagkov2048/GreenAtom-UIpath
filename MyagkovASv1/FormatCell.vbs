Sub FormatCell()
    
    Range("B:B,E:E,G:G").NumberFormat = "0.00;-0.00;;@"
	Range("A:A,D:D").NumberFormat="dd.MM.yyyy"
	Range("C:C,F:F").NumberFormat = "hh:mm:ss"
    ActiveSheet.Cells.VerticalAlignment = xlCenter
    ActiveSheet.Cells.HorizontalAlignment = xlCenter
    Dim headerRange As Range
    Set headerRange = Range("A1").CurrentRegion.Rows(1)
    headerRange.EntireColumn.AutoFit
	headerRange.NumberFormat="@"

End Sub
