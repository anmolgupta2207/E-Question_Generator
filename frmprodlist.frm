VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmprodlist 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Product List"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   5580
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   3660
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmprodlist.frx":0000
      Height          =   2535
      Left            =   120
      OleObjectBlob   =   "frmprodlist.frx":0014
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "frmprodlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Sub DBGrid1_DblClick()
If flag = 1 Then
    frmpurchorderdtls.txtprodno.Text = Data1.Recordset(0)
    frmpurchorderdtls.txtdesc.Text = Data1.Recordset(1)
    frmpurchorderdtls.txtrate.Text = Data1.Recordset(2)
    frmpurchorderdtls.txtqty.SetFocus
    Me.Hide
    'Unload Me
ElseIf flag = 2 Then
    frmBilManual.txtprodno.Text = Data1.Recordset(0)
    frmBilManual.txtdesc.Text = Data1.Recordset(1)
    frmBilManual.txtrate.Text = Data1.Recordset(2)
    frmBilManual.txtqty.SetFocus
    Me.Hide
    'Unload Me
End If
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\database"
End Sub

Private Sub Text1_Change()
Data1.RecordSource = "SELECT prodno,proddesc,prodrate,prodqty,prodreorderlvl FROM Product WHERE proddesc LIKE '" + Trim(Text1.Text) + "*'"
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
