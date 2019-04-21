VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Shiv Stationary"
   ClientHeight    =   10230
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12990
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0442
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   9855
      Width           =   12990
      _ExtentX        =   22913
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   10583
            MinWidth        =   10585
            Key             =   "stKey"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "10/24/2010"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "6:28 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12990
      _ExtentX        =   22913
      _ExtentY        =   1111
      ButtonWidth     =   1191
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tasks"
            Key             =   "Task"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AddPro"
                  Text            =   "Add &Product"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AddCust"
                  Text            =   "Add &Customers"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AddPo"
                  Text            =   "Add &PO"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BillPo"
                  Text            =   "Generate Bill by PO"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BillMan"
                  Text            =   "Generate Bill Manually"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View"
            Key             =   "View"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AllCust"
                  Text            =   "All &Customers"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AllPro"
                  Text            =   "All &Products"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            Key             =   "Report"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CustList"
                  Text            =   "&Customer List"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ProList"
                  Text            =   "Product List"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "POToday"
                  Text            =   "Todays POs"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BillToday"
                  Text            =   "Todays Bills"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5520
      Top             =   3360
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuprod 
         Caption         =   "Add &Product "
      End
      Begin VB.Menu mnucust 
         Caption         =   "Add &Customer"
      End
      Begin VB.Menu mnupurchorder 
         Caption         =   "&Purchase Order "
      End
      Begin VB.Menu mnuBill 
         Caption         =   "Add &Bill by PO"
      End
      Begin VB.Menu mnuBillMan 
         Caption         =   "Add Bill &Manually"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnuprodlist 
         Caption         =   "Product List"
      End
      Begin VB.Menu mnucustlist 
         Caption         =   "Customer List"
      End
   End
   Begin VB.Menu mnureports 
      Caption         =   "&Reports"
      Begin VB.Menu mnucustlistr 
         Caption         =   "&Customer List"
      End
      Begin VB.Menu mnuprodlistr 
         Caption         =   "&Product List"
      End
      Begin VB.Menu mnupurchorderr 
         Caption         =   "Todays &POs"
      End
      Begin VB.Menu mnubillre 
         Caption         =   "Todays &Bills"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
MDIForm1.Width = Screen.Width
MDIForm1.Height = Screen.Height
End Sub

Private Sub MDIForm_Terminate()
    'Beep
    MsgBox "THANK YOU FOR USING Shiv Stationary", , "SHIV STATIONARY"
End Sub

Private Sub mnuchallan_Click()
    frmchallan.Show
End Sub

Private Sub mnubill_Click()
   frmbill.Show
End Sub

Private Sub mnuBillMan_Click()
frmBilManual.Show
End Sub

Private Sub mnubillre_Click()
DataReport4.Show
End Sub

Private Sub mnucust_Click()
    frmcustdtls.Show
End Sub

Private Sub mnucustlist_Click()
  frmViewAllCust.Data1.RecordSource = "SELECT * FROM customer"
  frmViewAllCust.Data1.Refresh
  frmViewAllCust.Data1.Caption = "Customer Details"
  frmViewAllCust.Show
End Sub

Private Sub mnucustlistr_Click()
    DataReport1.Show
    
    End Sub

Private Sub mnuexit_Click()
    End
End Sub

Private Sub mnuprod_Click()
    frmproddtls.Show
End Sub

Private Sub mnuProdList_Click()
  frmViewAllPro.Data1.RecordSource = "SELECT * FROM PRODUCT"
  frmViewAllPro.Data1.Refresh
  frmViewAllPro.Data1.Caption = "Product Details"
  frmViewAllPro.Show
End Sub

Private Sub mnuprodlistr_Click()
  DataReport2.Show
End Sub

Private Sub mnupurchorder_Click()
    frmpurchorderdtls.Show
End Sub

Private Sub mnusupplist_Click()
  frmViewAllCust.Data1.RecordSource = "SELECT * FROM Supplier"
  frmViewAllCust.Data1.Refresh
  frmViewAllCust.Data1.Caption = "Supplier Details"
  frmViewAllCust.Show
End Sub

Private Sub mnusupplistr_Click()
   DataReport3.Show
End Sub

Private Sub mnupurchorderr_Click()
DataReport3.Show
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
    Case "AddPro": frmproddtls.Show
    Case "AddCust": frmcustdtls.Show
    Case "AddPo": frmpurchorderdtls.Show
    Case "BillPo": frmbill.Show
    Case "BillMan": frmBilManual.Show
    Case "AllPro": frmViewAllPro.Show
    Case "AllCust": frmViewAllCust.Show
    Case "CustList": DataReport1.Show
    Case "ProdList": DataReport2.Show
    Case "POToday": DataReport3.Show
    Case "BillToday": DataReport4.Show
    Case "Exit": End
End Select
End Sub
