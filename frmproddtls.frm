VERSION 5.00
Begin VB.Form frmproddtls 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Master"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7185
   Begin VB.Frame Frame2 
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
      Left            =   240
      TabIndex        =   16
      Top             =   3840
      Width           =   6855
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
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
         Picture         =   "frmproddtls.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&View"
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
         Picture         =   "frmproddtls.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   " &Edit"
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
         Index           =   2
         Left            =   1320
         Picture         =   "frmproddtls.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Del"
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
         Left            =   2520
         Picture         =   "frmproddtls.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   3120
         Picture         =   "frmproddtls.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Left            =   3720
         Picture         =   "frmproddtls.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   4320
         Picture         =   "frmproddtls.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Left            =   4920
         Picture         =   "frmproddtls.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   5520
         Picture         =   "frmproddtls.frx":2210
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Left            =   1920
         Picture         =   "frmproddtls.frx":2652
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   6120
         Picture         =   "frmproddtls.frx":2A94
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Product Register"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   6855
      Begin VB.TextBox txtcat 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   " "
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cmbcompany 
         DataField       =   "Prodcompany"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "frmproddtls.frx":2ED6
         Left            =   2520
         List            =   "frmproddtls.frx":2F19
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtpno 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Text            =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtdesc 
         Height          =   285
         Left            =   2520
         TabIndex        =   1
         Text            =   " "
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtqty 
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         Text            =   " "
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtreorderlvl 
         Height          =   285
         Left            =   2520
         TabIndex        =   4
         Text            =   " "
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtrate 
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Text            =   " "
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   5400
         Top             =   240
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Company:"
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
         Left            =   480
         TabIndex        =   28
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Category:"
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
         Left            =   480
         TabIndex        =   27
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblpno 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Product No:"
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
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lbldesc 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Description:"
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
         Left            =   480
         TabIndex        =   12
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Qty On Hand:"
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
         Left            =   480
         TabIndex        =   11
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reorder Level:"
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
         Left            =   480
         TabIndex        =   10
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rate:"
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
         Index           =   5
         Left            =   480
         TabIndex        =   9
         Top             =   2520
         Width           =   1935
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Product's Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   14
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmproddtls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As Recordset     'for product table

Private Sub cmbcompany_Click()
txtdesc.SetFocus
End Sub


Private Sub txtcat_GotFocus()
MDIForm1.StatusBar1.Panels(1).Text = "Press down arrow key to display Product Category"
End Sub

Private Sub txtcat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Trim(txtcat.Text)) = 0 Then
        MsgBox "Blank Field", , "SHIV STATIONARY"
    Else
        cmbcompany.SetFocus
    End If
End If
End Sub
Private Sub txtcat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    frmprodcat.Show
End If
End Sub
Private Sub cmbcompany_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Trim(cmbcompany.Text)) = 0 Then
        MsgBox "Blank Field", , "SHIV STATIONARY"
    Else
        txtdesc.SetFocus
    End If
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim Reply
Select Case Index
    Case 0
        Reply = MsgBox("Commit Changes?(Y/N)", vbQuestion + vbYesNo, "SHIV STATIONARY")
        If Reply = vbYes Then
            Rs.AddNew
            Rs(0) = txtpno.Text
            Add_Rec
            Rs.Update
            Command1(0).Enabled = False
        End If
        Clear_Rec
       
    Case 1
        frmprodlist.Data1.RecordSource = "SELECT * FROM Product"
        frmprodlist.Data1.Refresh
        frmprodlist.Data1.Caption = "Product Details"
        frmprodlist.Show
    Case 2
         Rs.Edit
         Add_Rec
         Rs.Update
         If txtdesc.Text = "" Or txtcat.Text = "" Or cmbcompany.Text = "" Or txtqty.Text = "" Or txtreorderlvl.Text = "" Or txtrate.Text = "" Then
         MsgBox "Please enter record to Edit", vbInformation, "SHIV STATIONARY"
         End If
    Case 3
        Rs.Delete
        Rs.MoveNext
        On Error GoTo ErrHandler
        Rs.MoveLast
        Load_Rec
    Case 4
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
        Rs.MoveLast
        Load_Rec
    Case 8
        Reply = MsgBox("Quit?(Y/N)", vbQuestion + vbYesNo, "SHIV STATIONARY")
        If Reply = vbYes Then
            Unload Me
        End If
    Case 9
        Clear_Rec
    Case 10
        DataReport2.Show
End Select
Exit Sub
ErrorHandler:
    MsgBox "You must enter record", , "Blank record"
ErrHandler:
    MsgBox "End or Beginning of the Database", , "EOF/BOF"
   
End Sub

Private Sub Form_Load()
InitProc
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path + "/database", dbOpenDynaset)
Set Rs = db.OpenRecordset("Product", dbOpenDynaset)
If Rs.RecordCount > 0 Then
    Rs.MoveLast
End If
txtpno.Text = Trim("P") & Format(Trim(Str(Rs.RecordCount + 1)), "#0000")
End Sub

Public Sub Add_Rec()
Rs!prodno = txtpno.Text
Rs!proddesc = txtdesc.Text
Rs!prodcat = txtcat.Text
Rs!prodcompany = cmbcompany.Text
Rs!prodqty = txtqty.Text
Rs!prodreorderlvl = txtreorderlvl.Text
Rs!prodrate = txtrate.Text
End Sub

Public Sub Load_Rec()
txtpno.Text = Rs!prodno
txtdesc.Text = Rs!proddesc
txtcat.Text = Rs!prodcat
cmbcompany.Text = Rs!prodcompany
txtqty.Text = Rs!prodqty
txtreorderlvl.Text = Rs!prodreorderlvl
txtrate.Text = Rs!prodrate
End Sub

Public Sub Clear_Rec()
txtpno.Text = ""
txtdesc.Text = ""
txtcat.Text = ""
cmbcompany.Text = ""
txtqty.Text = ""
txtreorderlvl.Text = ""
txtrate.Text = ""
txtpno.Text = Trim("P") & Format(Trim(Str(Rs.RecordCount + 1)), "#0000")
txtcat.SetFocus
End Sub

Private Sub txtdesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Trim(txtdesc.Text)) = 0 Then
        MsgBox "Description Cannot be left blank", , "SHIV STATIONARY"
    Else
       txtqty.SetFocus
    End If
End If
End Sub



Private Sub txtqty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Trim(txtqty.Text)) = 0 Then
        MsgBox "Blank Field", , "SHIV STATIONARY"
    Else
        txtreorderlvl.SetFocus
    End If
End If
End Sub

Private Sub txtreorderlvl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Trim(txtreorderlvl.Text)) = 0 Then
        MsgBox "Blank Field", , "SHIV STATIONARY"
    Else
        txtrate.SetFocus
    End If
End If
End Sub

Private Sub txtrate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Trim(txtrate.Text)) = 0 Then
        MsgBox "Blank Field", , "SHIV STATIONARY"
    Else
        Command1(0).Enabled = True
        Command1(0).SetFocus
    End If
End If
End Sub
Private Sub Timer1_Timer()
    Label2.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    Label3.Caption = Format(Date$, "dd-mmmm-yyyy")
End Sub
