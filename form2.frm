VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmsupplist 
   Caption         =   "Supplier List"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   LinkTopic       =   "Form2"
   ScaleHeight     =   4155
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   120
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
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Width           =   3540
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "form2.frx":0000
      Height          =   2535
      Left            =   480
      OleObjectBlob   =   "form2.frx":0014
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
End
Attribute VB_Name = "frmsupplist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DBGrid1_DblClick()
    frmpurchorderdtls.txtsno.Text = Data1.Recordset(0)
    frmpurchorderdtls.txtsname.Text = Data1.Recordset(1)
    frmpurchorderdtls.txtdesc.SetFocus
    Unload Me
End Sub

Private Sub Form_Load()
    Data1.DatabaseName = "c:\cd shop\database"
End Sub

Private Sub Text1_Change()
    Data1.RecordSource = "SELECT suppno,suppname FROM Supplier WHERE suppname LIKE '" + Trim(Text1.Text) + "*'"
    Data1.Refresh
End Sub


