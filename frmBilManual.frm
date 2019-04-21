VERSION 5.00
Begin VB.Form frmBilManual 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shiv Stationary"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   5955
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
      Left            =   120
      TabIndex        =   34
      Top             =   3960
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
         Left            =   120
         Picture         =   "frmBilManual.frx":0000
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
         Index           =   3
         Left            =   1320
         Picture         =   "frmBilManual.frx":0442
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
         Index           =   4
         Left            =   1920
         Picture         =   "frmBilManual.frx":0884
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
         Index           =   5
         Left            =   2520
         Picture         =   "frmBilManual.frx":0CC6
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
         Index           =   6
         Left            =   3120
         Picture         =   "frmBilManual.frx":1108
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
         Left            =   4320
         Picture         =   "frmBilManual.frx":154A
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
         Index           =   1
         Left            =   720
         Picture         =   "frmBilManual.frx":198C
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
         Index           =   9
         Left            =   4920
         Picture         =   "frmBilManual.frx":1DCE
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
         Index           =   7
         Left            =   3720
         Picture         =   "frmBilManual.frx":2438
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   615
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
      Height          =   2295
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   5775
      Begin VB.TextBox txtrate 
         Height          =   285
         Left            =   3480
         TabIndex        =   27
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
         TabIndex        =   25
         Top             =   1080
         Width           =   975
      End
      Begin VB.ListBox List2 
         Height          =   840
         Left            =   1080
         TabIndex        =   24
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ListBox List3 
         Height          =   840
         Left            =   2520
         TabIndex        =   23
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txttotal 
         Height          =   285
         Left            =   4440
         TabIndex        =   22
         Text            =   " "
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtprodno 
         Height          =   285
         Left            =   120
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtamt 
         Height          =   285
         Left            =   4440
         TabIndex        =   19
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Left            =   3480
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   240
         Width           =   975
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
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5775
      Begin VB.TextBox txtbilldate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtbillno 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtcname 
         Height          =   285
         Left            =   2160
         TabIndex        =   0
         Text            =   " "
         Top             =   960
         Width           =   1215
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
         TabIndex        =   17
         Top             =   600
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
         TabIndex        =   16
         Top             =   240
         Width           =   1575
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
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmBilManual"
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
        Reply = MsgBox("Commit Changes Y/N?", vbYesNo + vbQuestion, "SHIV STATIONARY")
        If Reply = vbYes Then
            Rs.AddNew
            Rs!billno = txtbillno
            Rs!billdate = txtbilldate
            Rs!t_amt = txttotal
            Rs.Update
            Add_Rec
        End If
        Command1(0).Enabled = False
        Clear_Rec
    Case 1
        Clear_Rec
    Case 3
        Rs.MoveFirst
        Load_Rec
    Case 4
        Rs.MovePrevious
        If Rs.BOF Then Rs.MoveFirst
        Load_Rec
    Case 5
        Rs.MoveNext
        If Rs.EOF Then Rs.MoveLast
        Load_Rec
    Case 6
        Rs.MoveLast
        Load_Rec
    Case 7
        X = InputBox("Enter Bill Order No: ", "SHIV STATIONARY")
        Rs.FindFirst "[BillNo]='" + Trim(X) + "'"
        If Rs.NoMatch = True Then
            MsgBox "Bill Not Found", , "SHIV STATIONARY"
        Else
            Load_Rec
        End If
    Case 8
        Reply = MsgBox("Quit (Y/N)?", vbYesNo + vbQuestion, "SHIV STATIONARY")
        If Reply = vbYes Then
            Unload Me
        End If
    Case 9
        'Show Report
End Select
Exit Sub
End Sub

Private Sub Form_Load()
    InitProc
    Set Rs = db.OpenRecordset("Bill", dbOpenDynaset)
    Set Rs1 = db.OpenRecordset("Billdtls", dbOpenDynaset)
             
    If Rs.RecordCount > 0 Then
        Rs.MoveLast
    End If
    txtbillno.Text = Trim("B") & Format(Trim(Rs.RecordCount + 1), "#0000")
    txtbilldate.Text = Date
End Sub

Public Sub Clear_Rec()
    txtbillno.Text = ""
    txtbilldate.Text = ""
    txtcname.Visible = True
    txtcname.Text = ""
    txtcname.SetFocus
    txttotal.Text = ""
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    List5.Clear
    txtbillno.Text = Trim("B") & Format(Trim(Rs.RecordCount + 1), "#0000")
    txtbilldate.Text = Date
End Sub

Private Sub lblFlag_Click()

End Sub

Private Sub txtdesc_GotFocus()
MDIForm1.StatusBar1.Panels(1).Text = "Press Down Arrow Key to display Products."
End Sub

Private Sub txtdesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
       flag = 2
       frmprodlist.Show
    End If
End Sub

Private Sub txtdesc_LostFocus()
MDIForm1.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub txtqty_GotFocus()
MDIForm1.StatusBar1.Panels(1).Text = "Press Enter key to Add Product"
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
    Dim Reply
    If KeyAscii = 13 Then
        If txtqty <= frmprodlist.Data1.Recordset!prodreorderlvl Then
        MsgBox "Qty less than or equal to Reorderlevel!", vbExclamation, "Failure..."
        txtqty = ""
        txtqty.SetFocus
        Exit Sub
        End If
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
                Reply = MsgBox("Do You Want to Continue(Y/N)?", vbYesNo + vbQuestion, "OMSAI CD's SHOP")
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
Public Sub Add_Rec()
    Dim i
    
    For i = 0 To List1.ListCount - 1
        Rs1.AddNew
        Rs1!Bill_no = txtbillno.Text
        Rs1!prodno = List1.List(i)
        Rs1!qty = List3.List(i)
        Rs1!amt = List5.List(i)
        Rs1.Update
'        Rs2!prodqty = Rs2!prodqty - List3.List(i)
'        Rs2.Update
        db.Execute "UPDATE Product SET prodqty=prodqty-" + List3.List(i) + " WHERE Prodno='" + Trim(List1.List(i)) + "'"
        
    Next
End Sub

Public Sub Load_Rec()
    Set Rs2 = db.OpenRecordset("select bill.billno, bill.billdate, bill.t_amt, billdtls.prodno, billdtls.qty, billdtls.amt, product.prodrate, product.proddesc from bill, billdtls,product where bill.billno = billdtls.bill_no and product.prodno =  billdtls.prodno and bill.billno = '" + Trim(Rs(0)) + "'")
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    List5.Clear
    txtbillno.Text = ""
    txtbilldate.Text = ""
    txtcno = ""
    txtcname = ""
    txttotal.Text = ""
    If Rs2.RecordCount > 0 Then
        Rs2.MoveLast
        Rs2.MoveFirst
        txtcname.Visible = False
        txtbillno.Text = Rs2!billno
        txtbilldate.Text = Rs2!billdate
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


Private Sub txtqty_LostFocus()
MDIForm1.StatusBar1.Panels(1).Text = ""
End Sub
