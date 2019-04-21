VERSION 5.00
Begin VB.Form frmbill 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Bill"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6090
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
      TabIndex        =   26
      Top             =   4320
      Width           =   5775
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
         Left            =   3720
         Picture         =   "frmchallan.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   615
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
         Left            =   4920
         Picture         =   "frmchallan.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   720
         Picture         =   "frmchallan.frx":076C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   610
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
         Left            =   4320
         Picture         =   "frmchallan.frx":0BAE
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   3120
         Picture         =   "frmchallan.frx":0FF0
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
         Left            =   2520
         Picture         =   "frmchallan.frx":1432
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
         Left            =   1920
         Picture         =   "frmchallan.frx":1874
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   1320
         Picture         =   "frmchallan.frx":1CB6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
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
         Left            =   120
         Picture         =   "frmchallan.frx":20F8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Bill Header"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   5775
      Begin VB.TextBox txtpono 
         Height          =   285
         Left            =   2160
         TabIndex        =   0
         Text            =   " "
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtcname 
         Height          =   285
         Left            =   2160
         TabIndex        =   30
         Text            =   " "
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtcno 
         Height          =   285
         Left            =   2160
         TabIndex        =   27
         Text            =   " "
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtbillno 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbilldate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   22
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Purchaseorder No:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Customer Name:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         TabIndex        =   28
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bill No:"
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
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bill Date:"
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
         TabIndex        =   24
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Bill Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   5775
      Begin VB.ListBox List4 
         Height          =   840
         ItemData        =   "frmchallan.frx":253A
         Left            =   3600
         List            =   "frmchallan.frx":253C
         TabIndex        =   31
         Top             =   840
         Width           =   975
      End
      Begin VB.ListBox List5 
         Height          =   840
         ItemData        =   "frmchallan.frx":253E
         Left            =   4560
         List            =   "frmchallan.frx":2540
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.ListBox List3 
         Height          =   840
         ItemData        =   "frmchallan.frx":2542
         Left            =   2520
         List            =   "frmchallan.frx":2544
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txttotal 
         Height          =   285
         Left            =   4560
         TabIndex        =   13
         Text            =   " "
         Top             =   1800
         Width           =   975
      End
      Begin VB.ListBox List2 
         Height          =   840
         ItemData        =   "frmchallan.frx":2546
         Left            =   1080
         List            =   "frmchallan.frx":2548
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.ListBox List1 
         Height          =   840
         ItemData        =   "frmchallan.frx":254A
         Left            =   120
         List            =   "frmchallan.frx":254C
         TabIndex        =   11
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
         Left            =   3600
         TabIndex        =   32
         Top             =   240
         Width           =   975
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
         Left            =   4560
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   14
         Left            =   3360
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Qty Dispatched"
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
         Index           =   13
         Left            =   2520
         TabIndex        =   18
         Top             =   240
         Width           =   1095
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
         TabIndex        =   17
         Top             =   240
         Width           =   1455
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
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As Recordset
Dim Rs1 As Recordset
Dim Rs2 As Recordset
Private Sub Command1_Click(Index As Integer)
Dim Reply, X
Select Case Index
    Case 0
'        frmprodlist.Data1.Refresh
'        frmprodlist.Data1.Refresh
'        frmprodlist.Data1.Recordset.MoveFirst
        Call InitRS
        Do Until rsPro.EOF
            For i = 0 To List3.ListCount - 1
                If List3.List(i) <= rsPro!prodreorderlvl Then
                    MsgBox "Qty <= Reorder Level", vbExclamation, "Failure..."
                    Exit Sub
                End If
            Next i
            rsPro.MoveNext
        Loop
        Reply = MsgBox("Commit Changes Y/N?", vbYesNo + vbQuestion, "Shiv Stationary")
        If Reply = vbYes Then
            Rs.AddNew
            Rs!billno = txtbillno
            Rs!billdate = txtbilldate
            Rs!purchorderno = txtpono
            Rs!t_amt = txttotal
            Rs.Update
            Add_Rec
        End If
        If Rs.RecordCount > 0 Then
            Command1(4).Enabled = True
            Command1(5).Enabled = True
            Command1(6).Enabled = True
            Command1(7).Enabled = True
        End If
        Command1(0).Enabled = False
        Clear_Rec
    Case 9
        Clear_Rec
    Case 2
    Case 3
        X = InputBox("Enter Bill Order No: ", "SHIV STATIONARY CENTRE")
        Rs.FindFirst "[BillNo]='" + Trim(X) + "'"
        If Rs.NoMatch = True Then
            MsgBox "Bill Not Found", , "SHIV STATIONARY CENTRE"
        Else
            Load_Rec
        End If
    
    Case 4
        Rs.MoveFirst
        On Error GoTo ErrHandler
        Load_Rec
    Case 5
        Rs.MovePrevious
        On Error GoTo ErrHandler
        Load_Rec
    Case 6
        Rs.MoveNext
        On Error GoTo ErrHandler
        Load_Rec
    Case 7
        Rs.MoveLast
        On Error GoTo ErrHandler
        Load_Rec
    Case 8
        Reply = MsgBox("Quit (Y/N)?", vbYesNo + vbQuestion, "SHIV STATIONARY")
        If Reply = vbYes Then
            Unload Me
        End If
   ' Case 10
      '  x = InputBox("Enter Challan No: ", "OMSAI CD's SHOP")
      '  DataEnvironment1.Challan_Grouping Trim(x)
      '  DataReport5.Show
End Select
Exit Sub
ErrHandler:
    MsgBox "No More Bills", , "SHIV STATIONARY"

End Sub

Private Sub Form_Load()
    InitProc
    Set Rs = db.OpenRecordset("Bill", dbOpenDynaset)
    Set Rs1 = db.OpenRecordset("Billdtls", dbOpenDynaset)
    If Rs.RecordCount = 0 Then
        Command1(4).Enabled = False
        Command1(5).Enabled = False
        Command1(6).Enabled = False
        Command1(7).Enabled = False
    End If
    If Rs.RecordCount > 0 Then
        Rs.MoveLast
    End If
    txtbillno.Text = Trim("B") & Format(Trim(Rs.RecordCount + 1), "#0000")
    txtbilldate.Text = Date
End Sub



Private Sub txtpono_GotFocus()
MDIForm1.StatusBar1.Panels(1).Text = "Press down arrow key to display Purchase Orders"
End Sub

Private Sub txtpono_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    frmpurchorderlist.Show
End If
End Sub

Public Sub Clear_Rec()
txtbillno.Text = ""
txtbilldate.Text = ""
txtpono.Text = ""
txtcno.Text = ""
txtcname.Text = ""
txttotal.Text = ""
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
txtbillno.Text = Trim("B") & Format(Trim(Rs.RecordCount + 1), "#0000")
txtbilldate.Text = Date
txtpono.SetFocus
End Sub

Public Sub Add_Rec()
Dim i

For i = 0 To List1.ListCount - 1
    Rs1.AddNew
    Rs1!Bill_no = txtbillno.Text
    Rs1!prodno = List1.List(i)
    Rs1!qty = List3.List(i)
    Rs1!amt = List5.List(i)
    Rs1.Update
    Rs2!prodqty = Rs2!prodqty - List3.List(i)
    Rs2.Update
    'db.Execute "UPDATE Product SET prodqty=prodqty-" + List3.List(i) + " WHERE Prodno='" + Trim(List1.List(i)) + "'"
Next
End Sub

Public Sub Load_Rec()
    Set Rs2 = db.OpenRecordset("select bill.billno,billdate,bill.purchorderno,bill.t_amt,purchaseorder.custno,billdtls.prodno,billdtls.qty,billdtls.amt,customer.custname,product.proddesc,product.prodrate from bill,billdtls,customer,product,purchaseorder where bill.billno=billdtls.bill_no and purchaseorder.custno=customer.custno and billdtls.prodno=product.prodno and bill.purchorderno=purchaseorder.purchorderno and bill.t_amt=purchaseorder.t_amt and billno='" + Trim(Rs(0)) + "'")
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    List5.Clear
    txtbillno.Text = ""
    txtbilldate.Text = ""
    txtcno = ""
    txtcname = ""
    txtpono.Text = ""
    txttotal.Text = ""
    If Rs2.RecordCount > 0 Then
        Rs2.MoveLast
        Rs2.MoveFirst
        txtcno.Text = Rs2!custno
        txtcname.Text = Rs2!custname
        txtbillno.Text = Rs2!billno
        txtbilldate.Text = Rs2!billdate
        txtpono.Text = Rs2!purchorderno
        txttotal.Text = Rs2!t_amt
        Rs2.MoveFirst
        Do Until Rs2.EOF
            List1.AddItem Rs2!prodno
            List2.AddItem Rs2!proddesc
            List3.AddItem Rs2!qty
            List4.AddItem Rs2!prodrate
            List5.AddItem Rs2!amt
            Rs2.MoveNext
        Loop
    End If
End Sub


Private Sub txtpono_LostFocus()
MDIForm1.StatusBar1.Panels(1).Text = ""
End Sub
