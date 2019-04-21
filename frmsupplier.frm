VERSION 5.00
Begin VB.Form frmsupplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\cd shop\database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Supplier"
      Top             =   6480
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   480
      TabIndex        =   15
      Top             =   240
      Width           =   6615
      Begin VB.Frame Frame3 
         Caption         =   "Supplier Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   6135
         Begin VB.TextBox txtSscode 
            DataField       =   "Suppcode"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1200
            TabIndex        =   0
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtSdate 
            DataField       =   "Suppdate"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   3960
            TabIndex        =   1
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtSname 
            DataField       =   "Suppname"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1200
            TabIndex        =   2
            Top             =   840
            Width           =   4575
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Command1"
            Height          =   495
            Left            =   120
            TabIndex        =   18
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtSaddr 
            DataField       =   "Suppaddr"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1200
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   1320
            Width           =   4575
         End
         Begin VB.TextBox txtSemail 
            DataField       =   "Suppemail"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1200
            TabIndex        =   5
            Top             =   2280
            Width           =   4575
         End
         Begin VB.TextBox txtSphone 
            DataField       =   "Suppphone"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1200
            TabIndex        =   4
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label lblSdate 
            AutoSize        =   -1  'True
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3240
            TabIndex        =   24
            Top             =   360
            Width           =   465
         End
         Begin VB.Label lblSname 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   570
         End
         Begin VB.Label lblScode 
            AutoSize        =   -1  'True
            Caption         =   "Supp code:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   990
         End
         Begin VB.Label lblSaddr 
            AutoSize        =   -1  'True
            Caption         =   "Address:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   780
         End
         Begin VB.Label lblSphone 
            AutoSize        =   -1  'True
            Caption         =   "Phone no:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   20
            Top             =   1800
            Width           =   885
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "E_mail:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   19
            Top             =   2280
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operations:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   360
         TabIndex        =   16
         Top             =   3600
         Width           =   5895
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
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
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "SAVE RECORD"
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdfirst 
            Caption         =   "MoveFirst"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "VIEW PREVIOUS RECORD"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton cmdnext 
            Caption         =   "Move Next"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "VIEW NEXT RECORD"
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdlast 
            Caption         =   "MoveLast"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "VIEW LAST RECORD"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdprev 
            Caption         =   "Move Prev"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "VIEW PREVIOUS RECORD"
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "&Delete"
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
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "DELETE RECORD"
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
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
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "FIND RECORD"
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "C&lear "
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
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "MODIFY RECORD"
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdReport 
            Caption         =   "&Report"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "REPORT"
            Top             =   1680
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmsupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim db As Database
  Dim ws As Workspace
  Dim rs1 As Recordset       'for Supplier Table
  Dim rs3 As Recordset       'for Suppno
  Dim rs2 As Recordset       'for Queries
  Dim I As Integer
  Dim msg As String
  Dim rcnt As Integer
  
Public Sub Clear()
    txtSscode.Text = ""
    txtSdate.Text = ""
    txtSname.Text = ""
    txtSaddr.Text = ""
    txtSphone.Text = ""
    txtSemail.Text = ""
 
  cmdNext.Enabled = False
End Sub


Private Sub cmdclear_Click()
Clear
  cmdPrev.Enabled = True
  
'txtSdate.SetFocus
End Sub

Private Sub CmdDesc_Click()
     frmdescgrid.Show
End Sub

Private Sub cmdfind_click()
Dim x As Long
x = InputBox("Enter Supplier Code To Be Found?", "OMSAI CD's Shop")
Data1.Recordset.FindFirst "[Suppcode] = " & x
If Data1.Recordset.NoMatch Then
    msg = MsgBox("Recod Not Found", vbCritical, "OMSAI CD's Shop")
End If
'Dim X
'X = InputBox("Enter Supplier Code To Be Found?", "OMSAI CD's Shop")
'Set rs2 = db.OpenRecordset("SELECT * FROM Supplier WHERE Suppcode Like '" + CLng(X) + "';")
'    txtSscode.Text = rs2(0)
'    txtSdate.Text = rs2(1)
'    txtSname.Text = rs2(2)
'    txtSaddr.Text = rs2(3)
'    txtSphone.Text = rs2(4)
'    txtSemail.Text = rs2(5)
'msg = MsgBox("Recod Not Found", vbCritical, "OMSAI CD's Shop")
End Sub

Private Sub cmdPcat_Click()
  Frmprodlist.Show
End Sub




Private Sub Form_Load()
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase("c:\cd shop\database", dbOpenDynaset)
  Set rs1 = db.OpenRecordset("Supplier", dbOpenDynaset)
  
  
  'cmdnext.Enabled = False
  
  'Set rs2 = db.OpenRecordset("SELECT * FROM Supplier;")
  'rs2.MoveLast
 
End Sub

Private Sub txtSaddr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If IsNumeric(txtSaddr.Text) = False Then
    txtSphone.SetFocus
  Else
    msg = MsgBox("Invalid entry", vbCritical, "OMSAI CD's Shop")
    txtSaddr.Text = ""
    txtSaddr.SetFocus
  End If
End If
End Sub

Private Sub txtSemail_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If IsNumeric(txtSemail.Text) = False Then
      cmdsave.SetFocus
    Else
      msg = MsgBox("Invalid entry", vbCritical, "OMSAI CD's Shop")
      txtSemail.Text = ""
      txtSemail.SetFocus
    End If
  End If
  End Sub

 Private Sub txtSphone_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If IsNumeric(txtSphone.Text) = True Then
      txtSemail.SetFocus
    Else
      msg = MsgBox("Invalid entry", vbCritical, "Omsai CD's Shop")
      txtSphone.Text = ""
      txtSphone.SetFocus
    End If
  End If
 End Sub
 
Private Sub txtSscode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    If IsNumeric(txtSscode.Text) = True Then
      txtSdate.SetFocus
    Else
      msg = MsgBox("Invalid entry", vbCritical, "Omsai CD's Shop")
      txtSscode.Text = ""
      txtSscode.SetFocus
    End If
  End If
 End Sub

Private Sub cmdsave_Click()
  If txtSdate.Text = " " Or txtSname.Text = " " Or txtSaddr.Text = " " Or txtSphone.Text = " " Or txtSemail.Text = "" Then
    msg = MsgBox("Invalid Record Entry", vbCritical, "Omsai CD's Shop")
  Clear
  Else
    On Error GoTo ErrorHandler
    Data1.Recordset.MoveLast
    Data1.Recordset.AddNew
    msg = MsgBox("Record Added successfully", vbCritical, "Omsai CD's Shop")
    'rs1.AddNew
    'rs1(0) = txtSscode.Text
    'rs1(1) = txtSdate.Text
    'rs1(2) = txtSname.Text
    'rs1(3) = txtSaddr.Text
    'rs1(4) = txtSphone.Text
    'rs1(5) = txtSemail.Text
    'rs1.Update
    
    'rs3.AddNew
    'rs3(0) = rcnt
    'rs3.Update
    Clear
    'msg = MsgBox("Record Added successfully", vbCritical, "Omsai CD's Shop")
    txtSdate.SetFocus
  End If
 Exit Sub
 
ErrorHandler:
  msg = MsgBox("Supplier ID Already Exits ", vbCritical, "Omsai CD's Shop")
  Clear
End Sub

Private Sub cmdfirst_Click()
Data1.Recordset.MoveFirst
End Sub

Private Sub cmdlast_Click()
Data1.Recordset.MoveLast
End Sub

Private Sub cmdnext_Click()
On Error GoTo a
Data1.Recordset.MoveNext
Exit Sub
a:
Data1.Recordset.MovePrevious
MsgBox "This Is The Last Record", vbInformation, "EOF"
End Sub

Private Sub cmdOK_Click()
 Dim Reply
   Reply = MsgBox("QUIT(Y/N)?", vbQuestion + vbYesNo, "OMSAI CD's Shop")
   If Reply = vbYes Then Unload Me
End Sub

Private Sub cmdprev_Click()
On Error GoTo a
Data1.Recordset.MovePrevious
Exit Sub
a:
Data1.Recordset.MoveNext
MsgBox "This Is The First Record", vbInformation, "BOF"
End Sub


