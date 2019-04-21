VERSION 5.00
Begin VB.Form frmcustdtls 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Details"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6855
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Customer's Register"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   17
      Top             =   600
      Width           =   6855
      Begin VB.TextBox txtcno 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtname 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Text            =   " "
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtaddr 
         Height          =   525
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtphone 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Text            =   " "
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtemail 
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Text            =   " "
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   6120
         Top             =   240
      End
      Begin VB.TextBox txtcity 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Text            =   " "
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Address:"
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
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E-mail"
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
         Left            =   240
         TabIndex        =   22
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Phone:"
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
         Left            =   240
         TabIndex        =   21
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "City:"
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
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   1935
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
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   1935
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
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1935
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
      Left            =   0
      TabIndex        =   7
      Top             =   3600
      Width           =   6855
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
         Left            =   5880
         Picture         =   "frmcustdtls.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Left            =   1680
         Picture         =   "frmcustdtls.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Left            =   5280
         Picture         =   "frmcustdtls.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   4680
         Picture         =   "frmcustdtls.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   4080
         Picture         =   "frmcustdtls.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   3480
         Picture         =   "frmcustdtls.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   2880
         Picture         =   "frmcustdtls.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   2280
         Picture         =   "frmcustdtls.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   1080
         Picture         =   "frmcustdtls.frx":2210
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   480
         Picture         =   "frmcustdtls.frx":2652
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Customer's Details"
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
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   24
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmcustdtls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As Recordset

Private Sub Command1_Click(Index As Integer)
Dim Reply
Select Case Index
    Case 0
        Reply = MsgBox("Commit Changes?(Y/N)", vbQuestion + vbYesNo, "SHIV STATIONARY")
        If Reply = vbYes Then
            Rs.AddNew
            Rs(0) = txtcno.Text
            Add_Rec
            Rs.Update
            Command1(0).Enabled = False
        End If
        Clear_Rec
    Case 1
        frmcustlist.Data1.RecordSource = "SELECT * FROM Customer"
        frmcustlist.Data1.Refresh
        frmcustlist.Data1.Caption = "Customer Details"
        frmcustlist.Show
    Case 2
        If txtname.Text = "" Or txtaddr.Text = "" Or txtcity.Text = "" Or txtphone.Text = "" Or txtemail.Text = "" Then
         MsgBox "Please enter record to Edit", vbInformation, "SHIV STATIONARY"
        End If
        Rs.Edit
        Add_Rec
        Rs.Update
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
        DataReport1.Show
End Select
Exit Sub
ErrHandler:
    MsgBox "End or Beginning of the Database", , "EOF/BOF"
    
End Sub

Private Sub Form_Load()
   InitProc
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path + "/database", dbOpenDynaset)
    Set Rs = db.OpenRecordset("Customer", dbOpenDynaset)
    If Rs.RecordCount > 0 Then
        Rs.MoveLast
    End If
    txtcno.Text = Trim("C") & Format(Trim(Rs.RecordCount + 1), "#0000")
End Sub

Private Sub txtcno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Len(txtcno.Text)) > 0 Then
        txtname.SetFocus
    End If
End If
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Len(txtname.Text)) > 0 Then
        txtaddr.SetFocus
    End If
End If
End Sub

Private Sub txtaddr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Len(txtaddr.Text)) > 0 Then
        txtcity.SetFocus
    End If
End If
End Sub

Private Sub txtcity_GotFocus()
txtaddr.Text = Trim(txtaddr.Text)
End Sub

Private Sub txtcity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Len(txtcity.Text)) > 0 Then
        txtphone.SetFocus
    End If
End If
End Sub

Private Sub txtphone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Len(txtphone.Text)) > 0 Then
        txtemail.SetFocus
    End If
End If
End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Len(txtemail.Text)) > 0 Then
        Command1(0).Enabled = True
        Command1(0).SetFocus
    End If
End If
End Sub

Public Sub Add_Rec()
    Rs(0) = txtcno.Text
    Rs(1) = txtname.Text
    Rs(2) = txtaddr.Text
    Rs(3) = txtcity.Text
    Rs(4) = txtphone.Text
    Rs(5) = txtemail.Text
End Sub

Public Sub Load_Rec()
    txtcno.Text = Rs(0)
    txtname.Text = Rs(1)
    txtaddr.Text = Rs(2)
    txtcity.Text = Rs(3)
    txtphone.Text = Rs(4)
    txtemail.Text = Rs(5)
End Sub

Public Sub Clear_Rec()
    txtcno.Text = ""
    txtname.Text = ""
    txtaddr.Text = ""
    txtcity.Text = ""
    txtphone.Text = ""
    txtemail.Text = ""
    txtcno.Text = Trim("C") & Format(Trim(Rs.RecordCount + 1), "#0000")
    txtname.SetFocus
End Sub

Private Sub Timer1_Timer()
    Label2.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    Label3.Caption = Format(Date$, "dd-mmmm-yyyy")
End Sub
