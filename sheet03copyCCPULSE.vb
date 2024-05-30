Private Declare PtrSafe Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal dwData As LongPtr, ByVal dwExtraInfo As LongPtr)
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare PtrSafe Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal lParam As Long) As Long

'Private Const WM_LBUTTONDOWN = &H201
'Private Const WM_LBUTTONUP = &H202
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4

Sub copyCCPULSE()

    Dim appPath As String
    Dim appHandle As Long
    Dim xlApp As Object
    Dim xlSheet As Object
    Dim screenWidth As Long
    Dim screenHeight As Long
    Dim mouseX As Long
    Dim mouseY As Long
    Dim ccpulseWindow As Long
    Dim buffer As String
    Dim childWindow As Long
    Dim dateFormat As String
    Dim dateFormat2 As String
    'Dim cursorPos As POINTAPi
    
    Dim currentTime As Date

currentTime = Time

screenWidth = 1920
screenHeight = 1080

mouseX = 960
mouseY = 270

    '오후만 뽑음
    If currentTime >= TimeValue("18:01:00") And currentTime <= TimeValue("21:30:00") Then
    
        appPath = "C:\Program Files\GCTI\CCPulse+\CallCenter.exe"
        appHandle = Shell(appPath, vbNormalFocus)
        
        Application.Wait Now + TimeValue("0:00:05")
        ''{PASSWORD} 변경해서 써야함
        SendKeys "{PASSWORD}", True
        SendKeys "{ENTER}", True
        
        Application.Wait Now + TimeValue("0:00:05")
        
        Do
            ccpulseWindow = FindWindow(vbNullString, "CCPulse+ - [GRP_IVR] - TEST - View 1")
            'ccpulseWindow = FindWindow(vbNullString, "CCPulse+ - [GRP_IVR] - [TEST - View 1]")
            DoEvents
            Application.Wait Now + TimeValue("0:00:01")
        Loop Until ccpulseWindow <> 0
        
        AppActivate "CCPulse+ - [GRP_IVR] - TEST - View 1"
        
        SetForegroundWindow appHandle
        
        'AppActivate "CCPulse+ - [GRP_IVR] - [TEST - View 1]"
        
        Application.Wait Now + TimeValue("0:00:05")
        
        'GetCursorPos cursorPos
        
        SetCursorPos mouseX, mouseY
        SetCursorPos mouseX + 1, mouseY + 1
        
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        Application.Wait Now + TimeValue("0:00:01")
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        
        Application.Wait Now + TimeValue("0:00:01")
        SendKeys "^a", True
        Application.Wait Now + TimeValue("0:00:01")
        SendKeys "^c", True
        Application.Wait Now + TimeValue("0:00:01")
        
        'Do
        '    buffer = Space$(255)
        '    GetWindowText ccpulseWindow, buffer, Len(buffer)
        '    If InStr(buffer, "CurrentStateGSGroup") > 0 Then
        '
        '        childWindow = FindWindowEx(ccpulseWindow, 0, vbNullString, "CurrentStateGSGroup")
        '        If childWindow <> 0 Then
        '            mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        '            mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        '
        '            SendKeys "^a", True
        '            Application.Wait Now + TimeValue("0:00:01")
        '            SendKeys "^c", True
        '            Application.Wait Now + TimeValue("0:00:01")
        '
        '            Exit Do
        '        End If
        '    End If
        '    DoEvents
        '    Application.Wait Now + TimeValue("0:00:01")
        'Loop
        
        'SetCursorPos mouseX, mouseY
        
        'mouse_event &H2, mouseX, mouseY, 0, 0
        'mouse_event &H4, mouseX, mouseY, 0, 0
        'SendKeys "^a", True
        'SendKeys "^c", True
        
        Application.Wait Now + TimeValue("0:00:02")
        
        
        Set xlApp = GetObject(, "Excel.Application")
        xlApp.Visible = True
        ''Set xlSheet = xlApp.Workbooks(1).Sheets("IVR COUNT")
        Set xlSheet = ThisWorkbook.Sheets("IVR COUNT")
        xlSheet.Activate
        xlSheet.Cells.Clear
        xlSheet.Range("A3").PasteSpecial
        dateFormat = Format(Date, "yy-mm-dd")
        xlSheet.Range("A1:B2").Merge
        With xlSheet.Range("A1:B2")
                    .Interior.ColorIndex = 6
                    .Borders.LineStyle = 1
                    .Font.Size = 11
                    .Font.Name = "맑은 고딕"
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
        End With
        
        xlSheet.Range("A1") = dateFormat & " 내역"
        
        AppActivate Application.Caption
        Call Shell("taskkill /PID " & appHandle & " /F", vbHide)
        
    End If


End Sub

Sub copyIVRSheetAndSave()
    Dim wbSource As Workbook
    Dim wbNew As Workbook
    Dim ws As Worksheet
    Dim savePath As String
    Dim dateFormat As String

    '오후만 실행 (18시 이후)
    If currentTime >= TimeValue("18:01:00") And currentTime <= TimeValue("21:30:00") Then
        Set wbSource = ThisWorkbook
    
        Set wbNew = Workbooks.Add
        wbSource.Sheets("IVR COUNT").Copy Before:=wbNew.Sheets(1)
        Set ws = wbNew.Sheets(1)
    
        dateFormat1 = Format(Date, "yyyy년")
        
        dateFormat2 = Format(Date, "mm월")
        
        dateFormat3 = Format(Date, "mmdd")
        
        savePath = "K:\Shared files\GCCSHELP 폴더\0.일일점검일지\3.IVR Count\" & dateFormat1 & "\" & dateFormat2 & "\" & dateFormat3 & "_ivr.xlsx"
    
        ws.Parent.SaveAs fileName:=savePath, FileFormat:=51
        wbNew.Close SaveChanges:=False
    End If
End Sub


