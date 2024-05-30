Sub OfficeSideStart()
    '반드시 로컬 컴퓨터에서 작동할 것
    
    Dim currentTime As Date
    
    currentTime = Time
    
    If currentTime >= TimeValue("07:00:00") And currentTime <= TimeValue("16:30:00") Then
        Call LoadTelephony
        Call copyToTelephony
        Call loadPreviousCCPulse
        Call sheet04loadingSBC
    
    ElseIf currentTime >= TimeValue("18:01:00") And currentTime <= TimeValue("21:30:00") Then
        Call LoadTelephony
        Call copyToTelephony
        Call copyIVRSheetAndSave
        Call sheet04loadingSBC
    End If
    
End Sub
