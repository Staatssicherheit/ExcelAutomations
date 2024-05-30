Sub changingPWonRecord()

Dim driver As WebDriver
Dim firstLineChcker As Selenium.WebElement

Dim loginURL As String
Dim modifyURL As String

''{URL} 변경해서 써야함
loginURL = "{URL}"
modifyURL = "{URL}"
''https://{URL} 1번서버 Primary
''https://{URL} 2번서버 Secondary

Call accountList.Users

Set driver = New WebDriver

driver.AddArgument "--start-maximized"

driver.Start "Chrome"

driver.Get loginURL

''On Error GoTo SecondaryURL

''Wait 시키는법은 여러가지 방법이 있음

driver.Wait (1500)

AppActivate "Recording management system - Chrome"

driver.FindElementById("login_id").SendKeys accountList.adminID
driver.FindElementById("login_pass").SendKeys accountList.adminPW
driver.FindElementByCss("#login > div.space03 > button").Click
driver.Wait (500)
driver.Get modifyURL
driver.Wait (1500)

Dim useridInput As String

AppActivate Application.Caption

useridInput = InputBox("아이디를 입력하세요:", "Input Required")

If useridInput = "" Then
MsgBox "아무것도 입력하지 않았습니다", vbExclamation

End If

AppActivate "녹취관리 시스템 - Chrome"

driver.FindElementByXPath("//*[@id='labelDiv']/input").SendKeys useridInput
driver.FindElementByXPath("//*[@id='outer_search']/div[2]/div[2]/button").Click

Dim imageXPath As String
imageXPath = "//img[contains(@src,'ico_unlock.png')]"

On Error Resume Next
Dim unlockButtonChecker As WebElement
Set unlockButtonChecker = driver.FindElementByXPath(imageXPath)
On Error GoTo 0

Set firstLineChecker = driver.FindElementByXPath("//*[@id='grid']/div[2]/div[1]/div[2]/div/div/table/tbody/tr[2]")

If Not firstLineChecker Is Nothing Then
    Select Case True
        Case unlockButtonChecker Is Nothing
            driver.Get loginURL
            driver.Wait (500)
            driver.FindElementById("login_id").SendKeys useridInput
            driver.FindElementById("login_pass").SendKeys "testtesttest"
            driver.FindElementByCss("#login > div.space03 > button").Click
            driver.Wait (500)
            driver.SwitchToAlert.Accept
            Dim i As Integer
            For i = 1 To 5
                driver.FindElementByCss("#login > div.space03 > button").Click
                driver.SwitchToAlert.Accept
            Next i
            
            driver.FindElementById("login_id").Clear
            driver.FindElementById("login_id").SendKeys accountList.adminID
            driver.FindElementById("login_pass").Clear
            driver.FindElementById("login_pass").SendKeys accountList.adminPW
            driver.FindElementByCss("#login > div.space03 > button").Click
            driver.Wait (500)
            driver.Get modifyURL
            driver.Wait (500)
            driver.FindElementByXPath("//*[@id='labelDiv']/input").SendKeys useridInput
            driver.FindElementByXPath("//*[@id='outer_search']/div[2]/div[2]/button").Click
            driver.Wait (500)
            driver.FindElementByClass("btn_unlock").Click
            driver.Wait (500)
            driver.SwitchToAlert.Accept
            driver.Wait (1500)
            driver.SwitchToAlert.Accept
            driver.FindElementByClass("btn_edit").Click
            driver.Wait (500)
                                                                                                            ''{PASSWORD} 변경해서 써야함
            driver.FindElementByXPath("//*[@id='partRegi']/div[2]/table/tbody/tr[5]/td[2]/input").SendKeys "{PASSWORD}"
            driver.FindElementByXPath("//*[@id='partRegi']/div[3]/button[1]").Click
            driver.Wait (500)
            driver.SwitchToAlert.Accept
            driver.Wait (500)
            
        Case Else
            driver.FindElementByClass("btn_unlock").Click
            driver.Wait (500)
            driver.SwitchToAlert.Accept
            driver.Wait (1500)
            driver.SwitchToAlert.Accept
            driver.Wait (500)
            driver.FindElementByClass("btn_edit").Click
            driver.Wait (500)
                                                                                                            ''{PASSWORD} 변경해서 써야함
            driver.FindElementByXPath("//*[@id='partRegi']/div[2]/table/tbody/tr[5]/td[2]/input").SendKeys "{PASSWORD}"
            driver.FindElementByXPath("//*[@id='partRegi']/div[3]/button[1]").Click
            driver.Wait (500)
            driver.SwitchToAlert.Accept
    End Select
Else
    MsgBox "ID가 존재하지 않습니다"
    
End If

''Case 잠금이 걸려있음
''잠금해제
''비밀번호 수정

''Case 잠금이 걸려있지않음
''해당 계정에 로그인시도를 해서 잠금을 걸어줌
''관리자계정으로 로그인해서 잠금해제
''비밀번호 수정



End Sub
