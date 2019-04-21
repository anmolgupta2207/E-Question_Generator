Attribute VB_Name = "Module1"
'Option Explicit
Public ws As Workspace
Public db As Database
Public flag As Integer

Public rsPro As Recordset

Public Sub InitRS()
Set rsPro = db.OpenRecordset("Product", dbOpenDynaset)
rsPro.MoveFirst
End Sub
Public Sub InitProc()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\database")
End Sub

Public Function TextToNull(tmp As String)
If tmp = "" Then
TextToNull Null
Else
TextToNull tmp
End If
End Function


