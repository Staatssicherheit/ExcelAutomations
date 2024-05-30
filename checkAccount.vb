Public InfraNixUserID As String
Public InfraNixUserPW As String
Public recordingUserID As String
Public recordingUserPW As String
Public ovocUserID As String
Public ovocUserPW As String
Public SMSUserID As String
Public SMSUserPW As String
Public pcusername As String

Sub Users()
   pcusername = Environ$("Username")
    
    Select Case pcusername
    Case "username0", "vdiname0"
        InfraNixUserID = ""
        InfraNixUserPW = ""
        recordingUserID = ""
        recordingUserPW = ""
        
    Case "username1", "vdiname1"
        InfraNixUserID = ""
        InfraNixUserPW = ""
        recordingUserID = ""
        recordingUserPW = ""
        
    Case "username2", "vdiname2"
        InfraNixUserID = ""
        InfraNixUserPW = ""
        recordingUserID = ""
        recordingUserPW = ""
        
    Case "username3", "vdiname3"
        InfraNixUserID = ""
        InfraNixUserPW = ""
        recordingUserID = ""
        recordingUserPW = ""
            
    End Select
End Sub

Sub SMSUser()
    SMSUserID = ""
    SMSUserPW = ""
End Sub

Sub OVOCUser()
    ovocUserID = ""
    ovocUserPW = ""
End Sub
