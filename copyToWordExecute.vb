Sub copyToWord()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim arrNames() As Variant
    Dim i As Integer
    
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Add
    Set wb = ThisWorkbook
    
    arrNames = Array("메인_보고서", "고객사 단말현황", "고객사점검", "대표번호 녹취내역")

    For i = LBound(arrNames) To UBound(arrNames)
        Set ws = wb.Sheets(arrNames(i))

        Select Case arrNames(i)
            Case "메인보고서"
                Set rng = ws.Range("A2:O41")
            Case "고객사 단말현황"
                Set rng = ws.Range("B3:X24")
            Case "고객사점검"
                Set rng = ws.Range("B2:M49")
            Case Else
                Set rng = ws.UsedRange
        End Select

        rng.Copy
        wdApp.Selection.Paste

        If i = 0 Then
            wdApp.Selection.Font.Bold = True
            wdApp.Selection.TypeText Text:=vbCrLf & "6. 고객사 단말 현황" & vbCrLf
        ElseIf i = 1 Then
            wdApp.Selection.TypeText Text:=vbCrLf & "7. 고객사별 서비스점검 내역" & vbCrLf
        ElseIf i = 2 Then
            wdApp.Selection.TypeText Text:=vbCrLf & "8. 대표번호 녹취내역" & vbCrLf
        End If

        ''If i < UBound(arrNames) Then
        ''    wdApp.Selection.InsertBreak Type:=7
        ''End If
    Next i

    Set rng = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub



