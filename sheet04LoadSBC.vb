Sub sheet04loadingSBC()
    Dim dateFormat As String
    Dim currentTime As Date
    Dim timeSet As String
    
    Dim PreWB As Workbook
    Dim preFilename As String
   
    If Hour(Now) < 12 Then
        dateFormat = Format(Now() - 1, "mm월 dd일 ")
        Sheets("고객사 단말현황").Range("B3") = dateFormat & "07시 내역"
        dateFormat = Format(Now(), "mm월 dd일 ")
        Sheets("고객사 단말현황").Range("B11") = dateFormat & "07시 내역"
    Else
        dateFormat = Format(Now() - 1, "mm월 dd일 ")
        Sheets("고객사 단말현황").Range("B3") = dateFormat & "19시 내역"
        dateFormat = Format(Now(), "mm월 dd일 ")
        Sheets("고객사 단말현황").Range("B11") = dateFormat & "19시 내역"
    End If

    currentTime = Time
    If currentTime >= TimeValue("07:00:00") And currentTime <= TimeValue("16:00:00") Then
        timeSet = "오전"
    ElseIf currentTime >= TimeValue("17:00:00") And currentTime <= TimeValue("23:00:00") Then
        timeSet = "오후"
    End If
    
    preFilename = "K:\Shared files\GCCSHELP 폴더\0.일일점검일지\1.ITC 일일모니터링\" & Format(Date, "yyyy") & "년\" & Format(Date, "mm") & "월\" & Format(Date, "dd") - 1 & "일\" & "일일모니터링_컨택센터_" & Format(Date - 1, "yyyymmdd") & "_" & timeSet & ".xlsx"
    MsgBox preFilename
    
    'Access Workbook
    Set PreWB = Workbooks.Open(preFilename)
    Set PreSheet = PreWB.Sheets("고객사 단말현황")
    PreSheet.Range("E13:X15").Copy ThisWorkbook.Sheets("고객사 단말현황").Range("E5:X7")
    
    Application.CutCopyMode = False
        
    PreWB.Close False
    
End Sub
