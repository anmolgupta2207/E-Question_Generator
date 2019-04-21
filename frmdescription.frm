VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmdescription 
   Caption         =   "Product List"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Height          =   615
      Left            =   6720
      Picture         =   "frmdescription.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Height          =   615
      Left            =   5400
      Picture         =   "frmdescription.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdLast 
      Height          =   615
      Left            =   4080
      Picture         =   "frmdescription.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Height          =   615
      Left            =   2760
      Picture         =   "frmdescription.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   615
      Left            =   1440
      Picture         =   "frmdescription.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdFirst 
      Height          =   615
      Left            =   120
      Picture         =   "frmdescription.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\CD SHOP\database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Product"
      Top             =   4680
      Width           =   5415
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmdescription.frx":198C
      Height          =   2415
      Left            =   0
      OleObjectBlob   =   "frmdescription.frx":19A0
      TabIndex        =   0
      Top             =   600
      Width           =   11535
   End
End
Attribute VB_Name = "frmdescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim rs1 As Recordset
Dim rs2 As Recordset     'For Querries only

Private Sub cmdfind_click()
Dim x As Long
x = InputBox("Enter Product Code To Be Found:", "OMSAI CD's SHop")
Data1.Recordset.FindFirst "[Prodcode] = " & x
If Data1.Recordset.NoMatch Then
    MsgBox "Product Code Not Found", vbInformation, "OMSAI CD's SHOP"
End If
End Sub
Private Sub cmdOK_Click()
   Dim Reply
   Reply = MsgBox("QUIT(Y/N)?", vbQuestion + vbYesNo, "OMSAI CD's SHOP")
   If Reply = vbYes Then Unload Me
End Sub


Private Sub Form_Load()
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase("c:\cd shop\database", dbOpenDynaset)
  Set rs1 = db.OpenRecordset("Product", dbOpenDynaset)
  Data1.RecordSource = "SELECT * FROM Product;"
  Data1.Refresh
End Sub
Private Sub cmdFirst_Click()
Data1.Recordset.MoveFirst
End Sub

Private Sub cmdlast_Click()
Data1.Recordset.MoveLast
End Sub

Private Sub cmdnext_Click()
On Error GoTo x
Data1.Recordset.MoveNext
Exit Sub
x:
Data1.Recordset.MovePrevious
MsgBox "Last Record Entry", vbInformation, "EOF"
End Sub

Private Sub cmdprev_Click()
On Error GoTo x
Data1.Recordset.MovePrevious
Exit Sub
x:
Data1.Recordset.MoveNext
MsgBox "First Record Entry", vbInformation, "BOF"
End Sub

