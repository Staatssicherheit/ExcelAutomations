Sub LoadTelephony()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim lastSheetIndex As Integer
    Dim telephonyFile As String
    Dim filePath As String
    Dim dateFormat As String
    Dim workPath As String
    
    workPath = ThisWorkbook.Path
    
    dateFormat = Format(Date, "yyyymmdd")
        
    telephonyFile = "일일모니터링텔레포니_" & dateFormat & ".xlsx"
    
    filePath = workPath & "\" & telephonyFile
    
    MsgBox filePath
    
    Set sourceWorkbook = Workbooks.Open(filePath)
    
    Set targetWorkbook = ThisWorkbook
    
    For i = 1 To 7
        Set sourceSheet = sourceWorkbook.Sheets(i)
        sourceSheet.Copy Before:=targetWorkbook.Sheets(i)
        Set targetSheet = targetWorkbook.Sheets(i)
        targetSheet.Name = sourceSheet.Name
        targetSheet.Cells.Clear
        sourceSheet.Cells.Copy targetSheet.Cells
    Next i
    
    Worksheets(7).Move After:=Worksheets("OVOC")
    
    sourceWorkbook.Close False
    
End Sub
