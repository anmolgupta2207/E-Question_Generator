VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmcustlist 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer List"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   4020
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmcustlist.frx":0000
      Height          =   2535
      Left            =   0
      OleObjectBlob   =   "frmcustlist.frx":0014
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "frmcustlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DBGrid1_DblClick()
    frmpurchorderdtls.txtcno = Data1.Recordset(0)
    frmpurchorderdtls.txtname = Data1.Recordset(1)
    frmpurchorderdtls.txtprodno.SetFocus
    Unload Me
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\database"

End Sub

Private Sub Text1_Change()
Data1.RecordSource = "SELECT custno,custname FROM Customer where custname LIKE '" + Trim(Text1.Text) + "*'"
Data1.Refresh
End Sub

Private Sub Text1_GotFocus()
MDIForm1.StatusBar1.Panels(1).Text = "Type a letter to sort or display items"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DBGrid1.SetFocus
End If
End Sub

Private Sub Text1_LostFocus()
MDIForm1.StatusBar1.Panels(1).Text = ""
End Sub
