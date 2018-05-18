VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RevisionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Revision As String
Public SOPName As String
Public DateReleased As Date



Sub LoadRevision(Rev As String, SOPN As String, DReleased As Date)
Let Revision = Rev
Let SOPName = SOPN
Let DateReleased = DReleased
End Sub

Sub SaveRevision()

Dim db As Database
Dim Rev As Recordset

Set db = CurrentDb
Set Rev = db.OpenRecordset("Revision")
    Rev.AddNew
    With Rev
    
    !Revision = Revision
    !SOPName = SOPName
    !DateReleased = DateReleased
    
    End With
    Rev.Update
    

End Sub

Sub LoadRecord(SOP As String)

Dim db As Database
Dim Rev As Recordset

Set db = CurrentDb
Set Rev = db.OpenRecordset("SELECT Revision, SOPName, DateReleased FROM Revision WHERE SOPName = '" & SOP & "'")
Rev.MoveFirst

Let Revision = Rev!Revision
Let DateReleased = Rev!DateReleased

Set db = Nothing
Set Rev = Nothing
End Sub

Sub UpdateRevision()

Dim db As Database
Dim Rev As Recordset

Set db = CurrentDb
Set Rev = db.OpenRecordset("SELECT Revision, SOPName, DateReleased FROM Revision WHERE SOPName ='" & SOPName & "'")
    Rev.MoveFirst
    Rev.Edit
    With Rev
    
    !Revision = Revision
    !SOPName = SOPName
    !DateReleased = DateReleased
    
    End With
    Rev.Update
    

End Sub
