VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_User_Administration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Combo17_AfterUpdate()
Dim strg As String
Dim InPos As Long
'If Nz(Me.Combo17, "") = "" Then
'Exit Sub
'End If
'strg = Me.Combo17
'InPos = InStr(1, strg, " ")
'
'Me.fn_txt = Left(strg, InPos - 1)
'Me.lname_txt = Right(strg, Len(strg) - InPos)
End Sub

Private Sub Command10_Click()
DoCmd.Close acForm, "User_administration"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub Command13_Click()
MsgBox "1 = Administrator -Unlimited Access" & vbCrLf & _
"2 = Managers -Access to Assign Credit functionality" & vbCrLf & _
"3 = Area Owners -Access to Corrective Action functionality" & vbCrLf & "4 = All -Limited Access", vbDefaultButton1, "Note"
End Sub

Private Sub Command20_Click()
MsgBox "Importing Employee's information instead of typing them" _
, vbInformation, "Import First and Last Name"
End Sub

Private Sub Command9_Click()
'On Error GoTo handle_err
Dim Fname As String
Dim Lname As String
'Dim Acss_level As Integer

If Nz(Me.fn_txt, "") = "" Then
    MsgBox "Please type a first name", vbCritical, "ERROR"
    Exit Sub
Else:
    Fname = Me.fn_txt
End If
If Nz(Me.lname_txt, "") = "" Then
    MsgBox "Please type a last name", vbCritical, "ERROR"
    Exit Sub
Else:
    Lname = Me.lname_txt
End If
'If Nz(Me.Access_level, "") = "" Then
'    MsgBox "Please select an Access Level", vbCritical, "ERROR"
'    Exit Sub
'Else:
'    Acss_level = Me.Access_level
'End If
    Call Register_User(Fname, Lname)
    List2.Requery
    Me.fn_txt = ""
    Me.lname_txt = ""
    Me.Access_level = ""
Exit Sub
handle_err:
MsgBox Err.Description
End Sub

Private Sub Form_Load()

'Call Access_Check(FrmAccss, Form.Name)
End Sub

Private Sub List2_DblClick(Cancel As Integer)
On Error GoTo handle_err

Set SUser = New SelectedUser

If Nz(Me.List2, "") = "" Then
    
    Exit Sub
End If

SUser.UserID = Me.List2.Column(2)
SUser.Fname = Me.List2.Column(0)
SUser.Lname = Me.List2.Column(1)


DoCmd.Close acForm, "User_Administration"
DoCmd.OpenForm "User_info"

Exit Sub
handle_err:
MsgBox Err.Description
End Sub
