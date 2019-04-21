VERSION 5.00
Begin VB.Form rdbms 
   Caption         =   "DBMS"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "rdbms.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      Height          =   615
      Left            =   6120
      TabIndex        =   40
      Top             =   8400
      Width           =   1815
   End
   Begin VB.Frame Frame8 
      Caption         =   "Which command is used ti extract data from a database table ?"
      Height          =   975
      Left            =   2160
      TabIndex        =   35
      Top             =   7320
      Width           =   10455
      Begin VB.OptionButton Option32 
         Caption         =   "POP"
         Height          =   255
         Left            =   5880
         TabIndex        =   39
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option31 
         Caption         =   "SELECT"
         Height          =   255
         Left            =   4320
         TabIndex        =   38
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option30 
         Caption         =   "GET"
         Height          =   255
         Left            =   2280
         TabIndex        =   37
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option29 
         Caption         =   "EXTRACT"
         Height          =   195
         Left            =   360
         TabIndex        =   36
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Which command is used to delete a table ?"
      Height          =   1095
      Left            =   2160
      TabIndex        =   30
      Top             =   6120
      Width           =   10455
      Begin VB.OptionButton Option28 
         Caption         =   "DELETE TABLE"
         Height          =   195
         Left            =   6360
         TabIndex        =   34
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option27 
         Caption         =   "REMOVE TABLE"
         Height          =   195
         Left            =   4440
         TabIndex        =   33
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Option26 
         Caption         =   "CLEAR TABLE"
         Height          =   255
         Left            =   2520
         TabIndex        =   32
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option25 
         Caption         =   "DROP TABLE"
         Height          =   255
         Left            =   480
         TabIndex        =   31
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Which command is used delete a record from database table ?"
      Height          =   975
      Left            =   2160
      TabIndex        =   25
      Top             =   5040
      Width           =   10455
      Begin VB.OptionButton Option24 
         Caption         =   "REMOVE"
         Height          =   195
         Left            =   6000
         TabIndex        =   29
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option23 
         Caption         =   "MODIFY"
         Height          =   255
         Left            =   4080
         TabIndex        =   28
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option22 
         Caption         =   "DELETE"
         Height          =   195
         Left            =   2160
         TabIndex        =   27
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option21 
         Caption         =   "DROP"
         Height          =   195
         Left            =   480
         TabIndex        =   26
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Which keyword is used with ALTER command to delete a column ?"
      Height          =   855
      Left            =   2160
      TabIndex        =   20
      Top             =   4200
      Width           =   10455
      Begin VB.OptionButton Option20 
         Caption         =   "CHANGE"
         Height          =   195
         Left            =   6720
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option19 
         Caption         =   "REMOVE"
         Height          =   195
         Left            =   4560
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option18 
         Caption         =   "DELETE"
         Height          =   195
         Left            =   2520
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option17 
         Caption         =   "DROP"
         Height          =   195
         Left            =   480
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Which keyword is used with update command to change the value ?"
      Height          =   975
      Left            =   2160
      TabIndex        =   15
      Top             =   3120
      Width           =   10455
      Begin VB.OptionButton Option16 
         Caption         =   "SET"
         Height          =   255
         Left            =   6480
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option15 
         Caption         =   "ADD"
         Height          =   195
         Left            =   4560
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option14 
         Caption         =   "MODIFY"
         Height          =   195
         Left            =   2640
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option13 
         Caption         =   "CHANGE"
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Which command is used to change data in a table ?"
      Height          =   975
      Left            =   2160
      TabIndex        =   10
      Top             =   2040
      Width           =   10455
      Begin VB.OptionButton Option12 
         Caption         =   "MODIFY"
         Height          =   195
         Left            =   6360
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option11 
         Caption         =   "ALTER"
         Height          =   195
         Left            =   4200
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option10 
         Caption         =   "UPDATE"
         Height          =   195
         Left            =   2520
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option9 
         Caption         =   "CHANGE"
         Height          =   195
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Which command is used to modify column name or table structure ?"
      Height          =   975
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   10455
      Begin VB.OptionButton Option8 
         Caption         =   "ALTER"
         Height          =   195
         Left            =   6720
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option7 
         Caption         =   "CHANGE"
         Height          =   195
         Left            =   4200
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option6 
         Caption         =   "ADD"
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         Caption         =   "MODIFY"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Which command is used to insert a new recordin table"
      Height          =   855
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VB.OptionButton Option4 
         Caption         =   "NEW"
         Height          =   195
         Left            =   5760
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "INSERT"
         Height          =   195
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "INSERT INTO"
         Height          =   195
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ADD"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "rdbms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim marks As Integer
marks = 0
If Option2.Value = True Then marks = marks + 1
If Option8.Value = True Then marks = marks + 1
If Option10.Value = True Then marks = marks + 1
If Option16.Value = True Then marks = marks + 1
If Option17.Value = True Then marks = marks + 1
If Option22.Value = True Then marks = marks + 1
If Option25.Value = True Then marks = marks + 1
If Option31.Value = True Then marks = marks + 1




If marks < 4 Then
MsgBox ("Your are fail ")
Else
MsgBox ("Congratulations ! Your are Pass")
End If
MsgBox ("Marks :  " & marks)
result.Label10.Caption = marks
subject.Option4.Enabled = False
If ((subject.Option1.Enabled = False) And (subject.Option2.Enabled = False) And (subject.Option3.Enabled = False) And (subject.Option4.Enabled = False) And (subject.Option5.Enabled = False)) Then
subject.Command1.Enabled = True
End If
Unload Me
End Sub

