Sub otherSendingSMS()

Dim driver As WebDriver

'Dim myID As String
'Dim myPW As String

Dim loginURL As String
Dim bizpurrioURL As String
Dim bizpurrioChkURL As String

Call checkAccount.SMSUser

loginURL = "https://www.bizppurio.com/login"
bizpurrioURL = "https://www.bizppurio.com/message/sms/write"
bizpurrioChkURL = "https://www.bizppurio.com/send/result/list"

Set driver = New WebDriver

driver.Start "Chrome"
driver.Get loginURL

driver.FindElementById("userId").SendKeys checkAccount.SMSUserID
driver.FindElementById("userPwd").SendKeys checkAccount.SMSUserPW
driver.FindElementById("bizwebBtnLogin").Click

driver.Wait (1000)
driver.Get bizpurrioURL

''내문자함 클릭
driver.FindElementByXPath("//*[@id='container']/div[2]/div[1]/fieldset/div[4]/div[1]/div[2]/table/tbody/tr[2]/td/div/div[3]/span[4]").Click
''아랫단 iframe 접근
driver.SwitchToFrame ("bizwebEmoticonIFrame")
''장문 클릭
driver.Wait (1000)
driver.FindElementByCss("div.emoticonTab > ul > li.box_tab.lms_tab").Click
''메시지 클릭
driver.Wait (2000)
driver.FindElementByXPath("//*[@id='cont']/div/div[2]/div[2]/div[2]/ul/li[1]/div").Click
driver.Wait (1000)
''메시지 로드
driver.SwitchToAlert.Accept
''상단부 진입
driver.SwitchToDefaultContent


''동일 문구클릭
driver.FindElementByXPath("//*[@id='container']/div[2]/div[1]/fieldset/div[4]/div[1]/div[2]/table/tbody/tr[1]/td/ul/li[1]").Click
''제목입력
driver.FindElementById("messageContentTitle").SendKeys ("[컨택센터 상황보고]")
''주소록 진입
driver.FindElementById("btnOpenReceiverAddress").Click

driver.Wait (1000)
''팝업 포커스
driver.SwitchToNextWindow
driver.Wait (2000)
''그룹 선택
driver.FindElementByXPath("/html/body/div[1]/div/div/div[2]/div[2]/table/tbody/tr/td[1]/label").Click
''선택완료
driver.FindElementById("btnAddressGroupReceiverSubmit").Click
driver.SwitchToWindowByTitle ("비즈뿌리오-단/장문 발송")
'driver.FindElementByXPath ("//*[@id='container']/div[2]/div[1]/div")
driver.FindElementByXPath("//*[@id='messageSendSubmit']").Click
driver.Wait (4500)
driver.SwitchToAlert.Accept
driver.Wait (5000)

driver.Get bizpurrioChkURL
driver.FindElementByXPath("//*[@id='searchBtn']").Click
driver.Wait (15000)

''driver.SwitchToDefaultContent
''driver.SwitchToAlert.Accept

End Sub


