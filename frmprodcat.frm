VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmprodcat 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Product Category Description"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmprodcat.frx":0000
      Height          =   2415
      Left            =   120
      OleObjectBlob   =   "frmprodcat.frx":0014
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "frmprodcat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DBGrid1_DblClick()
    frmproddtls.txtcat = Data1.Recordset(0)
    frmproddtls.cmbcompany.SetFocus
    'Unload Me
    Me.Hide
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\database"
End Sub

Private Sub Text1_Change()
Data1.RecordSource = "SELECT prodcat,cat_desc FROM Product_Cat WHERE prodcat LIKE '" + Trim(Text1.Text) + "*'"
Data1.Refresh
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DBGrid1.SetFocus
End If
End Sub


