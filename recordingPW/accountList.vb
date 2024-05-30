Public adminID As String
Public adminPW As String
Public pcusername As String

Sub Users()

   pcusername = Environ$("Username")
    
Select Case pcusername
    
    ''1st : officepc, 2nd : vdi pc
    Case "", ""
        adminID = ""
        adminPW = ""
        
    Case "", ""
        adminID = ""
        adminPW = ""
        
    Case "", ""
        adminID = ""
        adminPW = ""
        
    Case "", ""
        adminID = ""
        adminPW = ""
        
    Case "", ""
        adminID = ""
        adminPW = ""
        
End Select
    
End Sub
