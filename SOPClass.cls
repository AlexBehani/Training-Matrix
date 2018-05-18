VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SOPClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public SOPName As String
Public CurrentRev As String
Public F_GroupName As String
Public Inactive As Boolean
Public SOPTitle As String
Public DateReleased As Date
Private arr() As Variant
Public RowString As String

Sub LoadForm(SOPN As String, SOPT As String, CRev As String, _
            DReleased As Date, GroupArea As Variant)

Let SOPName = SOPN
Let SOPTitle = SOPT
Let CurrentRev = CRev
Let DateReleased = DReleased
Let arr() = GroupArea
            
End Sub

Sub LoadRevForm(CRev As String, SOPN As String)
Let SOPName = SOPN
Let CurrentRev = CRev
End Sub


Sub SaveSOP()
Dim n As Integer
Dim db As Database
Dim SOP As Recordset

Set db = CurrentDb
Set SOP = db.OpenRecordset("SOP")
For n = LBound(arr) To UBound(arr) - 1
    SOP.AddNew
    With SOP
        !SOPName = SOPName
        !CurrentRev = CurrentRev
        !F_GroupName = arr(n)
        !Inactive = Inactive
        !SOPTitle = SOPTitle
        SOP.Update

    End With
Next n

Set SOP = Nothing
Set db = Nothing


End Sub

Sub GroupRow(SOPN As String)

    Dim Str As String
    Dim db As Database
    Dim SOP As Recordset
    
    Str = ""
    Set db = CurrentDb
    Set SOP = db.OpenRecordset("SELECT F_GroupName FROM SOP WHERE = '" & SOPN & "'")
    SOP.MoveFirst
    Do While Not SOP.EOF
        Str = Str & SOP!F_GroupName & ";"
        SOP.MoveNext
    Loop
    
    Let RowString = Str
    Set db = Nothing
    Set SOP = Nothing
        
    
End Sub

Sub LoadRecord(SOPN As String)

    Dim Str As String
    Dim db As Database
    Dim SOP As Recordset

    
    Str = ""
    Set db = CurrentDb
    Set SOP = db.OpenRecordset("SELECT SOPName, CurrentRev, Inactive, SOPTitle, F_GroupName FROM SOP WHERE SOPName= '" & SOPN & "'")
    SOP.MoveFirst
    
    Let SOPName = SOP!SOPName
    Let CurrentRev = SOP!CurrentRev
    Let Inactive = SOP!Inactive
    Let SOPTitle = SOP!SOPTitle
'    Let DateReleased = SOP!DateReleased

    SOP.MoveFirst
        
    Do While Not SOP.EOF
        Str = Str & SOP!F_GroupName & ";"
        SOP.MoveNext
    Loop
    Str = Left(Str, Len(Str) - 1)
           
    Let RowString = Str
    Set db = Nothing
    Set SOP = Nothing
    
End Sub

Sub DeleteSOP(SOP As String)
Dim db As Database
Dim Str As String

Set db = CurrentDb
Str = "DELETE * FROM SOP WHERE SOPName= '" & SOP & "'"
db.Execute Str

Set db = Nothing

End Sub

Sub UpdateRev(SOPN As String)

Dim db As Database
Dim SOP As Recordset

Set db = CurrentDb
Set SOP = db.OpenRecordset("SELECT CurrentRev FROM SOP WHERE SOPName ='" & SOPN & "'")
If SOP.RecordCount > 0 Then

    SOP.MoveFirst
    Do While Not SOP.EOF
    
        SOP.Edit
        SOP!CurrentRev = CurrentRev
        SOP.Update
        SOP.MoveNext
        
    Loop

End If

Set db = Nothing
Set SOP = Nothing

End Sub

Function UniqueSOP(SOPN As String) As Boolean
Dim db As Database
Dim SOP As Recordset

Set db = CurrentDb
Set SOP = db.OpenRecordset("SELECT id FROM SOP WHERE SOPName ='" & SOPN & "'")
If SOP.RecordCount > 0 Then

    UniqueSOP = False
Else
    UniqueSOP = True

End If

Set db = Nothing
Set SOP = Nothing
End Function

