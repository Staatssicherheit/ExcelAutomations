Sub loadPreviousCCPulse()
    Dim sourcePath As String
    Dim targetPath As String
    Dim sourceFile As String
    Dim targetFile As String
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    'Dim ThisWorkbook As Workbook
    Dim dateFormat As Date
    Dim currentTime As Date
    Dim rng As Range
    Dim fileName As String

    Dim year As String
    Dim month As String
    Dim day As String
    Dim lastDayofMonth As Date

    Dim targetSheet As Worksheet
    Dim destinationWB As Workbook
    
    
    year = Format(Date, "yyyy")
    month = Format(Date, "mm")
    day = Format(Date, "dd")
    MsgBox year & "년 " & month & "월 " & day & "일"
   
    '전월 마지막 날짜 ex) 2월이면, 24년 01월 31일로 표기됨
    lastDayofMonth = DateSerial(year, month, 0)
    lastDay = Format(DateSerial(year, month, 0), "dd")
    
    currentTime = today

    sourcePath = "K:\Shared files\GCCSHELP 폴더\0.일일점검일지\1.ITC 일일모니터링\"
    
    '일일모니터링_컨택센터_20240521_오후.xlsx
    currentTime = TimeValue("07:00:00")
    If currentTime >= TimeValue("06:30:00") And currentTime <= TimeValue("09:30:00") Then
        If day = 1 Then
            If month = 1 Then
                sourceFile = sourcePath & year - 1 & "년\12월\" & lastDay & "일\" & "일일모니터링_컨택센터_" & year - 1 & "12" & lastDay & "_오후.xlsx"
                MsgBox sourceFile
            ElseIf month >= 2 And month <= 12 Then
                sourceFile = sourcePath & year & "년\" & month - 1 & "월\" & lastDay & "일\" & "일일모니터링_컨택센터_" & year & month - 1 & lastDay & "_오후.xlsx"
                MsgBox sourceFile
            Else
                MsgBox "Invalid Date, Please check month."
            End If
        ElseIf day >= 2 And day <= day Then
            sourceFile = sourcePath & year & "년\" & month & "월\" & day - 1 & "일\" & "일일모니터링_컨택센터_" & year & month & day - 1 & "_오후.xlsx"
            MsgBox sourceFile
        End If
    End If
    
    Set sourceWorkbook = Workbooks.Open(sourceFile)
    Set sourceSheet = sourceWorkbook.Sheets("IVR COUNT")
    sourceSheet.UsedRange.Copy ThisWorkbook.Sheets("IVR COUNT").Range("A1")
   
    '클립보드 초기화
    Application.CutCopyMode = False
    
    sourceWorkbook.Close False

End Sub



