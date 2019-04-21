VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmpurchorderlist 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Purchase Orders"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmpurchorderlist.frx":0000
      Height          =   2175
      Left            =   120
      OleObjectBlob   =   "frmpurchorderlist.frx":0014
      TabIndex        =   1
      Top             =   960
      Width           =   3855
   End
End
Attribute VB_Name = "frmpurchorderlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TempRs As Recordset
Private Sub DBGrid1_DblClick()
Set TempRs = db.OpenRecordset("SELECT Purchaseorder.custno,customer.custname,Purchaseorderdtls.Prodno,proddesc,qty,Product.prodrate,purchaseorder.t_amt,purchaseorderdtls.amt FROM Purchaseorder,Purchaseorderdtls,Product,customer WHERE Purchaseorder.purchorderno=Purchaseorderdtls.purchorderno AND Purchaseorderdtls.Prodno=Product.prodno AND purchaseorder.custno = customer.custno AND Purchaseorder.purchorderno='" + Trim(Data1.Recordset(0)) + "'")
frmbill.txtcname.Text = ""
frmbill.txtcno.Text = ""
frmbill.txttotal.Text = ""
frmbill.List1.Clear
frmbill.List2.Clear
frmbill.List3.Clear
frmbill.List4.Clear
frmbill.List5.Clear
TempRs.MoveLast
TempRs.MoveFirst
frmbill.txtcno.Text = TempRs!custno
frmbill.txtcname.Text = TempRs!custname
If TempRs.RecordCount > 0 Then
    Do While Not TempRs.EOF
        frmbill.txtpono.Text = Data1.Recordset(0)
        frmbill.List1.AddItem TempRs!prodno
        frmbill.List2.AddItem TempRs!proddesc
        frmbill.List3.AddItem TempRs!qty
        frmbill.List4.AddItem TempRs!prodrate
        frmbill.List5.AddItem TempRs!amt
        frmbill.txttotal.Text = TempRs!t_amt
        TempRs.MoveNext
    Loop
End If
frmbill.Command1(0).Enabled = True
frmbill.Command1(0).SetFocus
Me.Hide
'Unload Me
End Sub

Private Sub Form_Load()
    Data1.DatabaseName = App.Path & "\database"
    InitProc
End Sub

Private Sub Text1_Change()
    Data1.RecordSource = "SELECT purchorderno,purchorderdate FROM Purchaseorder WHERE purchorderno LIKE '" + Trim(Text1.Text) + "*'"
    Data1.Refresh
End Sub

Private Sub Text1_GotFocus()
MDIForm1.StatusBar1.Panels(1).Text = "Type a letter to sort or display items"
End Sub

Private Sub Text1_LostFocus()
MDIForm1.StatusBar1.Panels(1).Text = ""
End Sub
