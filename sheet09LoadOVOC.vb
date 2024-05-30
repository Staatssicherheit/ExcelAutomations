Sub ovocScreenshot()
Dim bot As New WebDriver
Dim wb As Workbook
Dim ws As Worksheet
Dim workPath As String
Dim imagePath As String
Dim statsPage As Object
Dim dateFormat As String
Dim cell As Range
Dim deleteShp As Shape

workPath = ThisWorkbook.Path

Call checkAccount.OVOCUser

Set ws = ThisWorkbook.Sheets("OVOC")

        For Each cell In ws.UsedRange
            cell.Value = ""
        Next cell
        
        For Each deleteShp In ws.Shapes
            If deleteShp.Type = msoPicture Then
                deleteShp.Delete
            End If
        Next deleteShp
            

bot.Start "chrome"
bot.Window.Maximize
''{URL} 변경해서 사용해야함
bot.Get "https://{URL}/web-ui-ovoc/#!/statistics/aggQoeDevices"
bot.Wait 1000
bot.FindElementById("details-button").Click
bot.FindElementById("proceed-link").Click

bot.Wait 1000

bot.FindElementById("login-username").SendKeys checkAccount.ovocUserID
bot.FindElementById("login-password").SendKeys checkAccount.ovocUserPW
bot.FindElementByCss("#login-button").Click

bot.Wait 1000
bot.Get "https://{URL}/web-ui-ovoc/#!/statistics/aggQoeDevices"

bot.Wait 5000

bot.FindElementByCss(".highcharts-series-1:nth-child(1) tspan").Click
bot.Wait 1000
bot.FindElementByCss(".highcharts-series-3 tspan").Click
bot.Wait 1000
bot.FindElementByCss(".highcharts-series-2 tspan").Click
bot.Wait 1000
bot.FindElementByCss(".highcharts-legend-item:nth-child(4) tspan").Click
bot.Wait 1000


bot.FindElementByClass("ac-filter-parent").Click
bot.Wait 500
bot.FindElementByXPath("//span[text()='Last 12 hours']").Click
bot.Wait 500
bot.FindElementById("FilterHeader_Topology").Click
bot.FindElementByXPath("//div/input").SendKeys ("M4K_GRP")
bot.FindElementById("1_anchor").Click
bot.FindElementByXPath("//ac-button/button/div[text()=' Apply']").Click

bot.Wait 10000

imagePath = workPath & "\Screenshot.png"

dateFormat = Format(Date, "yy-mm-dd")
ThisWorkbook.Sheets("OVOC").Range("A1") = dateFormat & " 내역"

Dim img As Image
Set img = bot.FindElementByXPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div[5]/div/div").TakeScreenshot()
img.SaveAs imagePath

    Set ws = ThisWorkbook.Sheets("OVOC")

    Dim addImage As Shape

    Set addImage = ws.Shapes.AddPicture(fileName:=imagePath, _
                                                LinkToFile:=msoFalse, _
                                                SaveWithDocument:=msoTrue, _
                                                Left:=ws.Cells(3, 1).Left, _
                                                Top:=ws.Cells(3, 1).Top, _
                                                Width:=-1, Height:=-1)
    With img
    End With

End Sub
