VERSION 5.00
Begin VB.Form student 
   Caption         =   "student"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10680
   LinkTopic       =   "Form2"
   Picture         =   "f.frx":0000
   ScaleHeight     =   7035
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   735
      Left            =   9120
      TabIndex        =   15
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      Height          =   735
      Left            =   6000
      TabIndex        =   13
      Top             =   5760
      Width           =   1935
   End
   Begin VB.PictureBox DTPicker1 
      DataField       =   "dob"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   2115
      TabIndex        =   10
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      DataField       =   "password"
      DataSource      =   "Data1"
      Height          =   525
      Left            =   9480
      TabIndex        =   7
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      DataField       =   "address"
      DataSource      =   "Data1"
      Height          =   525
      Left            =   3720
      TabIndex        =   5
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      DataField       =   "name"
      DataSource      =   "Data1"
      Height          =   525
      Left            =   3720
      TabIndex        =   3
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CLEAR"
      Height          =   735
      Left            =   2760
      TabIndex        =   2
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataField       =   "stud_id"
      DataSource      =   "Data1"
      Height          =   525
      Left            =   3720
      TabIndex        =   1
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "STUDENT DETAILS"
      Height          =   735
      Left            =   2400
      TabIndex        =   14
      Top             =   840
      Width           =   7815
   End
   Begin VB.Label Label7 
      Caption         =   "Password"
      Height          =   255
      Left            =   6840
      TabIndex        =   12
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label6 
      DataField       =   "username"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   9600
      TabIndex        =   11
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Date Of Birth"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Username"
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Student ID"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
End
Attribute VB_Name = "student"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Private Sub Command1_Click()
rs.AddNew
rs.Fields(0).Value = Text1.Text
rs.Fields(1).Value = Text2.Text
rs.Fields(2).Value = Text3.Text
rs.Fields(3).Value = DTPicker1.Value
rs.Fields(4).Value = Label6.Caption
rs.Fields(5).Value = Text4.Text
rs.Update
login.Show
rs.Close
Unload Me
End Sub

Private Sub Command2_Click()
login.Show
Unload Me
End Sub

Private Sub Command3_Click()
Text1.Text = " "
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Label6.Caption = ""

End Sub

Private Sub Form_Load()
Set db = OpenDatabase("D:\project\EXAM.mdb")
Set rs = db.OpenRecordset("select * from student")
End Sub

Private Sub Label6_Click()
Label6.Caption = Replace(Text2.Text, " ", "") + Text1.Text
End Sub
