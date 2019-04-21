VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H8000000D&
   Caption         =   "LOGIN"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "New User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   5160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   5160
      TabIndex        =   1
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   12810
      Left            =   120
      Picture         =   "login.frx":0000
      Top             =   0
      Width           =   19200
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Private Sub Command1_Click()
Dim a As Integer
rs.MoveFirst
While Not rs.EOF
If rs.Fields(4) = Text1.Text And rs.Fields(5) = Text2.Text Then
subject.Label2.Caption = rs.Fields(0)
subject.Label4.Caption = rs.Fields(1)
subject.Label6.Caption = rs.Fields(2)
subject.Label9.Caption = rs.Fields(3)
result.Label2.Caption = rs.Fields(0)
result.Label4.Caption = rs.Fields(1)
result.Label6.Caption = rs.Fields(2)
result.Label9.Caption = rs.Fields(3)
a = 1
MsgBox "Login Succesfully...."
subject.Show
Unload Me
End If
rs.MoveNext
Wend
If a <> 1 Then
MsgBox "Invalid UserName Or Password!"
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
student.Show
Unload Me
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("D:\project\EXAM.mdb")
Set rs = db.OpenRecordset("select * from student")
End Sub
