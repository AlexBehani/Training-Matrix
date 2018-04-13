VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command11_Click()
Call DisableStdOption
End Sub

Private Sub Command12_Click()

    Dim db As Database
    Dim prp As Property

    Set db = CurrentDb

    For Each prp In db.Properties
        Debug.Print prp.Name
    Next prp
    
    Debug.Print "..........................."
    
    MsgBox dmwListAllTables


End Sub



Private Sub Command6_Click()
Dim UserName As String
Dim Password As String
Dim Condition As Integer

UserName = Nz(Me.UserName, "")
Password = Nz(Me.Password, "")
If (UserName = "" Or Password = "") Then
MsgBox "Please enter User name and Password", vbCritical + vbExclamation, "Invalid Input"
Exit Sub
End If

Condition = Login(UserName, Password)
If (Condition = 0) Then

    DoCmd.Close acForm, "Login"
    DoCmd.OpenForm ("MainMenu")
    
ElseIf (Condition = 2) Then

    MsgBox "User Name and Password does not match", vbCritical, "Invalid User name and Password"
    Me.UserName = ""
    Me.Password = ""

ElseIf (Condition = 1) Then
    
    DoCmd.Close acForm, "Login"
    DoCmd.OpenForm ("Update_Password")
End If
End Sub



Function DisableStdOption()
 On Error GoTo Err_DisableStdOption

 ChangeProperty "StartupShowDBWindow", DB_BOOLEAN, False
 ChangeProperty "AllowFullMenus", DB_BOOLEAN, False
 ChangeProperty "AllowBuiltinToolbars", DB_BOOLEAN, False
 ChangeProperty "StartupMenuBar", DB_TEXT, "Empty Menu"

 'turn off the database window in normal use
 DoCmd.SelectObject acTable, , True
 DoCmd.RunCommand acCmdWindowHide

Exit_DisableStdOption:
 Exit Function

Err_DisableStdOption:
 MsgBox "Error #" & Err.Number & ": " & Err.Description
 Resume Exit_DisableStdOption
End Function

Private Sub ExitCommand_Click()
If (MsgBox("Exit?", vbQuestion + vbYesNo, "Exit the Database?") = vbYes) Then DoCmd.Quit
End Sub

Function ChangeProperty(strPropName As String, varPropType As Variant, _
  varPropValue As Variant) As Integer
 On Error GoTo Change_Err
 Dim dbs As Object, prp As Variant
 Const conPropNotFoundError = 3270

 Set dbs = CurrentDb
 dbs.Properties(strPropName) = varPropValue
dbs.Properties.Remove (varPropValue)

 ChangeProperty = True

Change_Bye:
 Exit Function

Change_Err:
 If Err = conPropNotFoundError Then    ' Property not found.
  Set prp = dbs.CreateProperty(strPropName, varPropType, varPropValue)
  dbs.Properties.Append prp
  Resume Next
 Else  ' Unknown error.
  ChangeProperty = False
  Resume Change_Bye
 End If
End Function


