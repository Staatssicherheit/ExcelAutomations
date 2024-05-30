Sub renameDownloadedFileAM()

    Dim filePath As String
    Dim newFilePath As String
    Dim fileName As String
    Dim timestamp As String
    
    filePath = ThisWorkbook.Path & "\"
    fileName = Dir(filePath & "line_pf_daily_*.xls_")
    timestamp = Mid(fileName, 15, 14)
    
    newFilePath = filePath & timestamp & "_오전" & ".xls"
    
    Name filePath & fileName As newFilePath
End Sub

Sub renameDownloadedFilePM()

    Dim filePath As String
    Dim newFilePath As String
    Dim fileName As String
    Dim timestamp As String
    
    filePath = ThisWorkbook.Path & "\"
    fileName = Dir(filePath & "net_term_line_*.xls_")
    timestamp = Mid(fileName, 15, 14)
    
    newFilePath = filePath & timestamp & "_오후" & ".xls"
    
    Name filePath & fileName As newFilePath
End Sub

Sub LineResult()

Dim driver As WebDriver

Dim loginURL As String
Dim lineResultURL As String
Dim lineResultURL2 As String
Dim excelPath As String
Dim eventReport As String
Dim currentTime As Date

currentTime = Time

Call checkAccount.Users

excelPath = ThisWorkbook.Path

''{IPADDRESS} 변경해서 써야함
lineResultURL = "http://{IPADDRESS}/sm_network/net_term_line_list_contract.jsp?menu_cd=M01146?&a_host_name=&a_line_seq=&a_port_type=2&a_view_gu=&m_group_cd=1.1.2&m_host_name=&m_page=0&m_type_gu=2&m_view_all=off&m_item_all=off&m_orderby=3&orderby=3&orderby_old=3&orderby_mode=2&group_cd=1.1.2&chk_item=&contract_yn=N&page=0&host_name=0&line_seq=0&port=0&interval=0&topn=20&avg_yn=y&check_avg_yn=y&max_yn=y&check_max_yn=n&keyword_cond=0&keyword=&item_all=off&current_page=1&max_page=2"
'lineResultURL = "chrome-extension://moffahdcgnjnglbepimcggkjacdmpojc/ieability.html?url=http://{IPADDRESS}/sm_network/net_term_line_list_contract.jsp?menu_cd=M01146?&a_host_name=&a_line_seq=&a_port_type=2&a_view_gu=&m_group_cd=1.1.2&m_host_name=&m_page=0&m_type_gu=2&m_view_all=off&m_item_all=off&m_orderby=3&orderby=3&orderby_old=3&orderby_mode=2&group_cd=1.1.2&chk_item=&contract_yn=N&page=0&host_name=0&line_seq=0&port=0&interval=0&topn=20&avg_yn=y&check_avg_yn=y&max_yn=y&check_max_yn=n&keyword_cond=0&keyword=&item_all=off&current_page=1&max_page=2"
lineResultURL2 = "http://{IPADDRESS}/sm_network/net_term_line_list2_contract.jsp?menu_cd=M01149"
eventReport = "http://{IPADDRESS}/report/fault/report_fault_totlist_in.jsp?menu_cd=M00870"

'lineResultURL 현재회선(야간근무자가 뽑음)
'lineResultURL2는 오전전일

Set driver = New WebDriver

driver.AddArgument "--start-maximized"
driver.AddArgument "--allow-running-insecure-content"
driver.AddArgument "--unsafely-treat-insecure-origin-as-secure=http://{IPADDRESS}"
driver.SetPreference "download.default_directory", excelPath
driver.SetPreference "download.directory_upgrade", True
driver.SetPreference "download.extensions_to_open", ""
driver.SetPreference "download.prompt_for_download", False

driver.Start "Chrome"

If currentTime >= TimeValue("12:00:01") And currentTime <= TimeValue("23:00:00") Then
    driver.Get lineResultURL

    driver.Wait (1500)
    driver.FindElementByName("memberId").SendKeys checkAccount.InfraNixUserID
    driver.FindElementByName("password").SendKeys checkAccount.InfraNixUserPW
    driver.Wait (1500)
    driver.FindElementByXPath("//table/tbody/tr[1]/td[3]/input").Click
    
    driver.Wait (3000)
    driver.FindElementByXPath("//select[@name='search_groupinfo_m_group_cd']").AsSelect.SelectByIndex 2
    driver.Wait (3000)
    driver.FindElementByXPath("//tbody/tr[3]/td/a[3]/img").Click
    driver.Wait (10000)
    Call renameDownloadedFilePM

    Call InsertNetworkData

ElseIf currentTime >= TimeValue("07:00:00") And currentTime <= TimeValue("12:00:00") Then
    driver.Get lineResultURL2

    driver.Wait (1500)
    driver.FindElementByName("memberId").SendKeys checkAccount.InfraNixUserID
    driver.FindElementByName("password").SendKeys checkAccount.InfraNixUserPW
    driver.Wait (1500)
    driver.FindElementByXPath("//table/tbody/tr[1]/td[3]/input").Click

    driver.Wait (3000)
    ''driver.FindElementByXPath("//")
    driver.FindElementByXPath("//select[@name='search_groupinfo_m_group_cd']").AsSelect.SelectByIndex 2
    driver.Wait (3000)
    driver.FindElementByXPath("//select[@name='date_type']").AsSelect.SelectByIndex 2
    driver.FindElementByXPath("//input[@name='check_max_yn']").Click
    driver.FindElementByXPath("//a[@href='javascript:thema_search_search_kuwait();']").Click
    driver.FindElementByXPath("//tbody/tr[3]/td/a[3]/img").Click
    driver.Wait (10000)
    Call renameDownloadedFileAM

    Call InsertNetworkData
End If

End Sub

Sub InsertNetworkData()

    Dim selectedFile As Variant
    Dim wbSource As Workbook
    Dim wsDestination As Worksheet
    Dim currentTime As Date
    Dim dateFormat As String
    Dim dateFormatMinus As String
    
    currentTime = Time
    
    Set wsDestination = ThisWorkbook.Sheets("그룹(네트워크_일간이용량)")
    
    selectedFile = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", Title:="네트워크 사용량 보고서를 선택해주세요")
    
    If selectedFile = False Then
        MsgBox "No file selected. Operation cancelled"
        Exit Sub
    End If
    
    
    
    Set wbSource = Workbooks.Open(selectedFile)
    
    wsDestination.Cells.Clear
    
    wbSource.Sheets(1).UsedRange.Copy wsDestination.Cells(1, 1)
    
        If currentTime >= TimeValue("07:00:00") And currentTime <= TimeValue("12:00:00") Then
    
           dateFormat = Format(Date, "YYYY/MM/DD")
           dateFormatMinus = Format(Date - 1, "YYYY/MM/DD")
           wsDestination.Range("A1").Value = "일별 회선 성능 조회 " & "(" & dateFormatMinus & " ~ " & dateFormat & " )"
        
        ElseIf currentTime >= TimeValue("12:00:01") And currentTime <= TimeValue("23:00:00") Then
            
            dateFormat = Format(Date, "YYYY/MM/DD")
            wsDestination.Range("A1").Value = "일별 회선 성능 조회 " & "(" & dateFormat & " ~ " & dateFormat & " )"
    
        End If
    
wbSource.Close False

End Sub


