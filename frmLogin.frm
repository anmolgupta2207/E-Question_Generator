VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   9705
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   13935
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0442
   ScaleHeight     =   5734.033
   ScaleMode       =   0  'User
   ScaleWidth      =   13084.21
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   6120
      TabIndex        =   1
      Top             =   4080
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00000000&
      Default         =   -1  'True
      DownPicture     =   "frmLogin.frx":19B0C
      Height          =   750
      Left            =   5280
      Picture         =   "frmLogin.frx":19F4E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Height          =   735
      Left            =   7080
      Picture         =   "frmLogin.frx":1BA92
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   900
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   6120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4560
      Width           =   2325
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "SHIV STATIONARY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3360
      TabIndex        =   6
      Top             =   1680
      Width           =   7095
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00000000&
      Caption         =   "&User Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   4560
      TabIndex        =   0
      Top             =   4080
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00000000&
      Caption         =   "&Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   4560
      TabIndex        =   2
      Top             =   4560
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If txtPassword = "ROSA" Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        Unload Me
        MDIForm1.Show
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

