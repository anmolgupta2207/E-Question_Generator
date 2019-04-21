VERSION 5.00
Begin VB.Form frmpurchorderdtls 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order "
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5865
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Purchase Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   5775
      Begin VB.TextBox txtname 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtcno 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   32
         Text            =   " "
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtorderdate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   31
         Text            =   " "
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtorderno 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   30
         Text            =   " "
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Order Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Customer No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Order No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Purchase Order Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   0
      Left            =   0
      TabIndex        =   15
      Top             =   2040
      Width           =   5775
      Begin VB.TextBox txtrate 
         Height          =   285
         Left            =   3480
         TabIndex        =   28
         Text            =   " "
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtqty 
         Height          =   285
         Left            =   2520
         TabIndex        =   2
         Text            =   " "
         Top             =   840
         Width           =   975
      End
      Begin VB.ListBox List4 
         Height          =   840
         Left            =   3480
         TabIndex        =   26
         Top             =   1080
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   840
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
      Begin VB.ListBox List2 
         Height          =   840
         Left            =   1080
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ListBox List3 
         Height          =   840
         Left            =   2520
         TabIndex        =   18
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txttotal 
         Height          =   285
         Left            =   4440
         TabIndex        =   17
         Text            =   " "
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtprodno 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   " "
         Top             =   795
         Width           =   975
      End
      Begin VB.TextBox txtdesc 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Text            =   " "
         Top             =   795
         Width           =   1455
      End
      Begin VB.ListBox List5 
         Height          =   840
         Left            =   4440
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtamt 
         Height          =   285
         Left            =   4440
         TabIndex        =   13
         Text            =   " "
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   11
         Left            =   3480
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prod No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   6
         Left            =   1080
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Qty Ordered"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   12
         Left            =   2520
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3480
         TabIndex        =   22
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   15
         Left            =   4440
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Select Option"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   4440
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   240
         Picture         =   "frmpurchorderdtls.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   4
         Left            =   1440
         Picture         =   "frmpurchorderdtls.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   2040
         Picture         =   "frmpurchorderdtls.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   2640
         Picture         =   "frmpurchorderdtls.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   7
         Left            =   3240
         Picture         =   "frmpurchorderdtls.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   8
         Left            =   4440
         Picture         =   "frmpurchorderdtls.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   9
         Left            =   840
         Picture         =   "frmpurchorderdtls.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   610
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   10
         Left            =   5040
         Picture         =   "frmpurchorderdtls.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   3840
         Picture         =   "frmpurchorderdtls.frx":2210
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label lblFlag 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "frmpurchorderdtls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As Recordset  'for purchase order
Dim Rs1 As Recordset  'for purchase order details
Dim Rs2 As Recordset  'for temp


Private Sub Command1_Click(Index As Integer)
Dim X, Reply
Select Case Index
    Case 0
        Rs.AddNew
        Add1_Rec
        Rs.Update
        Add2_Rec
        Clear_Rec
    Case 3
        X = InputBox("Enter Purchase Order No: ", "SHIV STATIONARY")
        Rs.FindFirst "[purchorderno]='" + Trim(X) + "'"
        If Rs.NoMatch = True Then
            MsgBox "Record Not Found", , "SHIV STATIONARY"
        Else
            Load_Rec
        End If
    Case 4
       On Error GoTo ErrHandler
        Rs.MoveFirst
        Load_Rec
    Case 5
        On Error GoTo ErrHandler
        Rs.MovePrevious
        Load_Rec
    Case 6
        On Error GoTo ErrHandler
        Rs.MoveNext
        Load_Rec
    Case 7
        On Error GoTo ErrHandler
        Rs.MoveLast
        Load_Rec
    Case 8
        Reply = MsgBox("Quit (Y/N)?", vbYesNo + vbQuestion, "SHIV STATIONARY")
        If Reply = vbYes Then
            Unload Me
        End If
    Case 9
        Clear_Rec
    'Case 10
   '     x = InputBox("Enter Purchase Order No: ", "OMSAI CD's SHOP")
    '    DataEnvironment1.PurchOrder_Grouping Trim(x)
    '    DataReport4.Show
End Select
Exit Sub
ErrHandler:
    MsgBox "No More Records(Y/N)?", , "EOF/BOF"
End Sub

Private Sub Form_Load()
InitProc
Set Rs = db.OpenRecordset("Purchaseorder", dbOpenDynaset)
Set Rs1 = db.OpenRecordset("Purchaseorderdtls", dbOpenDynaset)
If Rs.RecordCount > 0 Then
    Rs.MoveLast
End If
txtorderno.Text = Trim("PO") & Format(Trim(Rs.RecordCount + 1), "#0000")
txtorderdate.Text = Date
End Sub



Private Sub txtdesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    flag = 1
    frmprodlist.Show
End If
End Sub

Private Sub txtname_GotFocus()
MDIForm1.StatusBar1.Panels(1).Text = "Press down Arrow key to display Customer Details"
End Sub

Private Sub txtname_LostFocus()
MDIForm1.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
Dim Reply
If KeyAscii = 13 Then
    If Len(Trim(txtqty.Text)) > 0 Then
        txtamt.Text = Val(txtrate.Text) * Val(txtqty.Text)
        Reply = MsgBox("Add row (Y/N)?", vbYesNo + vbQuestion, "Inv Sys")
        If Reply = vbYes Then
            List1.AddItem txtprodno.Text
            List2.AddItem txtdesc.Text
            List3.AddItem txtqty.Text
            List4.AddItem txtrate.Text
            List5.AddItem txtamt.Text
            txttotal.Text = Val(Trim(txtamt.Text)) + Val(Trim(txttotal.Text))
            txtprodno.Text = ""
            txtdesc.Text = ""
            txtrate.Text = ""
            txtqty.Text = ""
            txtamt.Text = ""
            Reply = MsgBox("Do You Want to Continue(Y/N)?", vbYesNo + vbQuestion, "SHIV STATIONARY")
            If Reply = vbYes Then
                txtdesc.SetFocus
            Else
                Command1(0).Enabled = True
                Command1(0).SetFocus
            End If
        End If
    End If
End If
End Sub

Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    frmcustlist.Show
End If
 End Sub

Public Sub Load_Rec()
Dim i
Set Rs2 = db.OpenRecordset("SELECT Purchaseorder.purchorderno,Purchaseorder.purchorderdate,Purchaseorder.custno,Customer.custname,Purchaseorderdtls.prodno,proddesc,prodrate,qty,purchaseorderdtls.amt,purchaseorder.t_amt FROM Purchaseorder, Purchaseorderdtls,Product, Customer WHERE Purchaseorder.purchorderno = Purchaseorderdtls.purchorderno AND Purchaseorderdtls.prodno = Product.Prodno AND Purchaseorder.custno = Customer.custno AND Purchaseorder.purchorderno='" + Rs(0) + "'", dbOpenDynaset)
Rs2.MoveLast
Rs2.MoveFirst
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear

Do Until Rs2.EOF
    txtorderno = Rs2!purchorderno
    txtorderdate = Rs2!purchorderdate
    txtcno = Rs2!custno
    txtname = Rs2!custname
    List1.AddItem Rs2!prodno
    List2.AddItem Rs2!proddesc
    List3.AddItem Rs2!qty
    List4.AddItem Rs2!prodrate
    List5.AddItem Rs2!amt
    txttotal = Rs2!t_amt
    Rs2.MoveNext
Loop
    

End Sub

Public Sub Add1_Rec()
    Rs!purchorderno = txtorderno.Text
    Rs!purchorderdate = txtorderdate.Text
    Rs!t_amt = txttotal.Text
    Rs!custno = txtcno.Text
   End Sub

Public Sub Clear_Rec()
    txtorderno.Text = ""
    txtorderdate.Text = ""
    txtcno.Text = ""
    txtname.Text = ""
    txtprodno.Text = ""
    txtdesc.Text = ""
    txtrate.Text = ""
    txtqty.Text = ""
    txtamt.Text = ""
    txttotal.Text = ""
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    List5.Clear
    txtorderno.Text = Trim("PO") & Format(Trim(Rs.RecordCount + 1), "#0000")
    txtorderdate.Text = Date
    txtname.SetFocus
End Sub


Public Sub Add2_Rec()
Dim i
For i = 0 To List1.ListCount - 1
    Rs1.AddNew
    Rs1!purchorderno = txtorderno.Text
    Rs1!prodno = List1.List(i)
    Rs1!qty = List3.List(i)
    Rs1!amt = List5.List(i)
    Rs1.Update
    Next
End Sub


