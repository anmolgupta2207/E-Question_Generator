VERSION 5.00
Begin VB.Form frmsuppdtls 
   Caption         =   "Supplier Details"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   6915
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
      Left            =   0
      TabIndex        =   13
      Top             =   3600
      Width           =   6855
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
         Picture         =   "frmsuppdtls.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Picture         =   "frmsuppdtls.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Picture         =   "frmsuppdtls.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Picture         =   "frmsuppdtls.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Picture         =   "frmsuppdtls.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "frmsuppdtls.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "frmsuppdtls.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "frmsuppdtls.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "frmsuppdtls.frx":2210
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
         Left            =   1920
         Picture         =   "frmsuppdtls.frx":2652
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmsuppdtls.frx":2A94
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Supplier's Register"
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
      TabIndex        =   6
      Top             =   600
      Width           =   6855
      Begin VB.TextBox txtcity 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Text            =   " "
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   6120
         Top             =   240
      End
      Begin VB.TextBox txtpin 
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Text            =   " "
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txtstate 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Text            =   " "
         Top             =   2040
         Width           =   2295
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
      Begin VB.TextBox txtname 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Text            =   " "
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtsno 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Supp No:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "State:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pin Code:"
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
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.Label Label3 
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
      TabIndex        =   26
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Supplier's Details"
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
End
Attribute VB_Name = "frmsuppdtls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As Recordset
Dim Rs1 As Recordset

Private Sub Command1_Click(Index As Integer)
Dim Reply
Select Case Index
    Case 0
        Reply = MsgBox("Commit Changes?(Y/N)", vbQuestion + vbYesNo, "OMSAI CD'S SHOP")
        If Reply = vbYes Then
            Rs.AddNew
            Rs1.AddNew
            Rs1(0) = txtsno.Text
            Add_Rec
            Rs.Update
            Rs1.Update
            Command1(0).Enabled = False
        End If
        Clear_Rec
    Case 1
        Form3.Data1.RecordSource = "SELECT * FROM Supplier"
        Form3.Data1.Refresh
        Form3.Data1.Caption = "Supplier Details"
        Form3.Show
    Case 2
        Rs.Edit
        Add_Rec
        Rs.Update
    Case 3
        Rs.Delete
        Rs.MoveNext
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
        Reply = MsgBox("Quit?(Y/N)", vbQuestion + vbYesNo, "OMSAI CD's SHOP")
        If Reply = vbYes Then
            Unload Me
        End If
    Case 9
        Clear_Rec
    Case 10
        DataReport2.Show
End Select
Exit Sub
ErrHandler:
    MsgBox "End or Beginning of the Database", , "EOF/BOF"
End Sub

Private Sub Form_Load()
   InitProc
   Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase("c:\cd shop\database", dbOpenDynaset)
    Set Rs = db.OpenRecordset("Supplier", dbOpenDynaset)
    Set Rs1 = db.OpenRecordset("Supp_Temp", dbOpenDynaset)
    If Rs1.RecordCount > 0 Then
        Rs1.MoveLast
    End If
    txtsno.Text = Trim("S") & Format(Trim(Rs1.RecordCount + 1), "#0000")
End Sub

Private Sub txtsno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Len(txtsno.Text)) > 0 Then
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
        txtstate.SetFocus
    End If
End If
End Sub

Private Sub txtstate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Len(txtstate.Text)) > 0 Then
        txtpin.SetFocus
    End If
End If
End Sub

Private Sub txtpin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Len(txtpin.Text)) > 0 Then
        Command1(0).Enabled = True
        Command1(0).SetFocus
    End If
End If
End Sub



Public Sub Add_Rec()
    Rs(0) = txtsno.Text
    Rs(1) = txtname.Text
    Rs(2) = txtaddr.Text
    Rs(3) = txtcity.Text
    Rs(4) = txtstate.Text
    Rs(5) = txtpin.Text
End Sub

Public Sub Load_Rec()
    txtsno.Text = Rs(0)
    txtname.Text = Rs(1)
    txtaddr.Text = Rs(2)
    txtcity.Text = Rs(3)
    txtstate.Text = Rs(4)
    txtpin.Text = Rs(5)
End Sub

Public Sub Clear_Rec()
    txtsno.Text = ""
    txtname.Text = ""
    txtaddr.Text = ""
    txtcity.Text = ""
    txtstate.Text = ""
    txtpin.Text = ""
    txtsno.Text = Trim("S") & Format(Trim(Rs1.RecordCount + 1), "#0000")
    txtname.SetFocus
End Sub


