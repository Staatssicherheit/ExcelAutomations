Sub copyToTelephony()
    Dim sourceSheet As Worksheet
    Dim DestinationSheet As Worksheet
    Dim sourceRange As Range
    Dim destinationRange As Range
    Dim sourceRow As Range
    Dim destinationRow As Range
    Dim pcusername As String
    
    Set sourceSheet = ThisWorkbook.Sheets("메인_보고서")
    Set DestinationSheet = ThisWorkbook.Sheets("메인 보고서")
    
    Set sourceRange = sourceSheet.Range("A7:O41")
    
    Set destinationRange = DestinationSheet.Range("A56:O90")
    
    sourceSheet.Columns("B:B").ColumnWidth = 15
    sourceRange.Copy
    destinationRange.PasteSpecial Paste:=xlPasteAll, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    For Each sourceRow In sourceRange.Rows
        Set destinationRow = destinationRange.Rows(sourceRow.Row)
        destinationRow.RowHeight = sourceRow.RowHeight
    Next sourceRow
    
    With DestinationSheet
                    .Columns("B:B").ColumnWidth = 15
                    .Range("I4:J4").UnMerge
                    .Range("I5:J5").UnMerge
                    .Range("I4:J5").Borders.LineStyle = 1
                    .Range("J4") = "컨택센터"
                    .Range("I4:J5").HorizontalAlignment = xlCenter
                    .Range("I4:J5").VerticalAlignment = xlCenter
    End With
    
    pcusername = Environ$("Username")
    
    Select Case pcusername
    Case "DEV", "dev.jiho.kim"
        DestinationSheet.Range("J5") = "사람1"
        sourceSheet.Range("A2") = " 점검자 : 사람1 / 주임"
    Case "SH HONG", "rune.hong"
        DestinationSheet.Range("J5") = "사람2"
        sourceSheet.Range("A2") = " 점검자 : 사람2 / 선임"
    Case "Kim NI"
        DestinationSheet.Range("J5") = "사람3"
        sourceSheet.Range("A2") = " 점검자 : 사람3 / 과장"
    Case "Kwon yonghwon"
        DestinationSheet.Range("J5") = "사람3"
        sourceSheet.Range("A2") = " 점검자 : 사람3 / 과장"
    End Select
            
    
    Application.CutCopyMode = False
    
End Sub
