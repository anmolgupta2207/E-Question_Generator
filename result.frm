VERSION 5.00
Begin VB.Form result 
   Caption         =   "RESULT"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   1200
      ScaleHeight     =   135
      ScaleWidth      =   15
      TabIndex        =   22
      Top             =   600
      Width           =   15
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HOME"
      Height          =   615
      Left            =   8160
      TabIndex        =   21
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "SCORE."
      ForeColor       =   &H8000000D&
      Height          =   3735
      Left            =   840
      TabIndex        =   10
      Top             =   3480
      Width           =   12855
      Begin VB.Label Label19 
         Caption         =   "JAVA"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label18 
         Height          =   495
         Left            =   3240
         TabIndex        =   19
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "C PROGRAMMING"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label16 
         Height          =   495
         Left            =   3240
         TabIndex        =   17
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label15 
         Caption         =   "C# PROGRAMMING"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label14 
         Height          =   495
         Left            =   3240
         TabIndex        =   15
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label13 
         Caption         =   "C++ PROGRAMMING"
         Height          =   495
         Left            =   6840
         TabIndex        =   14
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label12 
         Height          =   495
         Left            =   10080
         TabIndex        =   13
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label11 
         Caption         =   "DATABASE MANAGEMENT SYSTEM"
         Height          =   495
         Left            =   6840
         TabIndex        =   12
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label10 
         Height          =   495
         Left            =   10080
         TabIndex        =   11
         Top             =   1680
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   4560
      TabIndex        =   0
      Top             =   7440
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      Height          =   9855
      Left            =   120
      Picture         =   "result.frx":0000
      ScaleHeight     =   9795
      ScaleWidth      =   15315
      TabIndex        =   23
      Top             =   0
      Width           =   15375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "RESULT"
      Height          =   735
      Left            =   3720
      TabIndex        =   9
      Top             =   480
      Width           =   7815
   End
   Begin VB.Label Label1 
      Caption         =   "STUDENT ID"
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "STUDENT NAME"
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "ADDRESS"
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label6 
      Height          =   495
      Left            =   11160
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "DATE OF BIRTH"
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label9 
      Height          =   495
      Left            =   11160
      TabIndex        =   1
      Top             =   2640
      Width           =   2535
   End
End
Attribute VB_Name = "result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset

Private Sub Command1_Click()
rs.AddNew
rs.Fields(0).Value = Label2.Caption
rs.Fields(1).Value = Label4.Caption
rs.Fields(2).Value = Label6.Caption
rs.Fields(3).Value = Label9.Caption
rs.Fields(4).Value = Label16.Caption
rs.Fields(5).Value = Label12.Caption
rs.Fields(6).Value = Label14.Caption
rs.Fields(7).Value = Label10.Caption
rs.Fields(8).Value = Label18.Caption
rs.Update
PrintForm
Unload Me
End Sub

Private Sub Command2_Click()
start.Show
Unload Me
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("D:\project\EXAM.mdb")
Set rs = db.OpenRecordset("select * from result")
End Sub

