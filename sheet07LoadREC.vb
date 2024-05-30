Sub LoadRecordingMain()
Dim driver As WebDriver

'Dim myID As String
'Dim myPW As String

Dim loginURL As String
Dim searchURL As String
Dim myID As String
Dim myPW As String
Dim excelPath As String

excelPath = ThisWorkbook.Path

''{URL} 변경해서 사용해야함
loginURL = "https://{URL}/index.jsp"
searchURL = "https://{URL}/rec_search/rec_search.jsp"
''https://{URL} 1번서버 Primary
''https://{URL} 2번서버 Secondary

Call checkAccount.Users

Set driver = New WebDriver

driver.AddArgument "--start-maximized"
driver.SetPreference "download.default_directory", excelPath
driver.SetPreference "download.directory_upgrade", True
driver.SetPreference "download.extensions_to_open", ""
driver.SetPreference "download.prompt_for_download", False

driver.Start "Chrome"
driver.Get loginURL

driver.Wait (1500)

driver.FindElementById("login_id").SendKeys checkAccount.recordingUserID
driver.FindElementById("login_pass").SendKeys checkAccount.recordingUserPW
driver.FindElementByCss("#login > div.space03 > button").Click
driver.Get searchURL
driver.Wait (1500)

driver.FindElementByName("user_name").SendKeys "gccshelp"
driver.Wait (1500)
driver.FindElementByName("btn_search").Click
driver.Wait (3000)
driver.FindElementByCss("#grid > div.pq-grid-bottom.ui-widget-header.ui-corner-bottom > div > span:nth-child(9) > select").AsSelect.SelectByIndex 2

driver.Wait (15000)
''wait 대신 쓸 대안법 찾아볼것
driver.FindElementByXPath("/html/body/div[5]/div/div[3]/div[2]/div[2]/div/div[1]/div[2]/button[1]").Click
''driver.FindElementByCss("#grid > div.pq-grid-top.ui-widget-header.ui-corner-top > div.pq-toolbar > button:nth-child(2)").Click
driver.Wait (20000)

Call CopyRecordSheetDataToExisting

End Sub

Sub CopyRecordSheetDataToExisting()
    Dim selectedFile As String
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    
    selectedFile = Application.GetOpenFilename("Excel Files (*.xls;*.xlsx), *.xls;*.xlsx", , "녹취내역 엑셀을 선택해주세요")
    
    If selectedFile <> "False" Then
        Set sourceWorkbook = Workbooks.Open(selectedFile)
        
        On Error Resume Next
        Set targetWorkbook = ThisWorkbook
        On Error GoTo 0
        
        If targetWorkbook Is Nothing Then
            MsgBox "Error : Could not find target Workbook.", vbCritical
            Exit Sub
        End If
        
        Set targetSheet = Nothing
        For Each ws In targetWorkbook.Sheets
            If ws.Name = "대표번호 녹취내역" Then
                Set targetSheet = ws
                Exit For
            End If
        Next ws
        
        If targetSheet Is Nothing Then
            MsgBox "Error: Could not find target sheet.", vbCritical
            sourceWorkbook.Close False
            Exit Sub
        End If
        
        targetSheet.Cells.Clear
        
        
        sourceWorkbook.Sheets(1).UsedRange.Copy targetSheet.Range("A1")
        
        sourceWorkbook.Close False
        
    Else
    
    End If
End Sub
    
