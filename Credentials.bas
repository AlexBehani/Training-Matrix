Attribute VB_Name = "Credentials"
Option Compare Database
Option Explicit

Public Function CredentialsCheck() As Boolean

If (CUser Is Nothing) Then

    CredentialsCheck = False
Else
    CredentialsCheck = True
End If

End Function
