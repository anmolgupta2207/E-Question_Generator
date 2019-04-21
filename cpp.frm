VERSION 5.00
Begin VB.Form cpp 
   Caption         =   "C++"
   ClientHeight    =   9225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11340
   LinkTopic       =   "Form7"
   Picture         =   "cpp.frx":0000
   ScaleHeight     =   9225
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      Height          =   615
      Left            =   5760
      TabIndex        =   8
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Frame Frame8 
      Caption         =   "The string ""HELLO WORD""nedd byte?"
      Height          =   975
      Left            =   1800
      TabIndex        =   7
      Top             =   7800
      Width           =   10455
      Begin VB.OptionButton Option32 
         Caption         =   "none of above"
         Height          =   255
         Left            =   5880
         TabIndex        =   40
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option31 
         Caption         =   "8 byte"
         Height          =   255
         Left            =   4320
         TabIndex        =   39
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option30 
         Caption         =   "12 byte"
         Height          =   255
         Left            =   2280
         TabIndex        =   38
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option29 
         Caption         =   "11 byte"
         Height          =   195
         Left            =   360
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "making class member in accessible to nonmember to they function?"
      Height          =   1095
      Left            =   1800
      TabIndex        =   6
      Top             =   6600
      Width           =   10455
      Begin VB.OptionButton Option28 
         Caption         =   "recursion"
         Height          =   195
         Left            =   6360
         TabIndex        =   36
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option27 
         Caption         =   "redundancy"
         Height          =   195
         Left            =   4200
         TabIndex        =   35
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option26 
         Caption         =   "data hiding"
         Height          =   255
         Left            =   2520
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option25 
         Caption         =   "polymorphism"
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "The argument the determine the state of the object?"
      Height          =   975
      Left            =   1800
      TabIndex        =   5
      Top             =   5520
      Width           =   10455
      Begin VB.OptionButton Option24 
         Caption         =   "stste controll"
         Height          =   195
         Left            =   6000
         TabIndex        =   32
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option23 
         Caption         =   "formate flage"
         Height          =   255
         Left            =   4080
         TabIndex        =   31
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option22 
         Caption         =   "manipulator"
         Height          =   195
         Left            =   2160
         TabIndex        =   30
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option21 
         Caption         =   "class"
         Height          =   195
         Left            =   480
         TabIndex        =   29
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "the end of string is reconsided is?"
      Height          =   855
      Left            =   1800
      TabIndex        =   4
      Top             =   4680
      Width           =   10455
      Begin VB.OptionButton Option20 
         Caption         =   "/sign"
         Height          =   195
         Left            =   6720
         TabIndex        =   28
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option19 
         Caption         =   "$ sign"
         Height          =   195
         Left            =   4560
         TabIndex        =   27
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option18 
         Caption         =   "new line"
         Height          =   195
         Left            =   2520
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option17 
         Caption         =   "null character"
         Height          =   195
         Left            =   480
         TabIndex        =   25
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "most common  operation used construcor is-----------?"
      Height          =   975
      Left            =   1800
      TabIndex        =   3
      Top             =   3600
      Width           =   10455
      Begin VB.OptionButton Option16 
         Caption         =   "none"
         Height          =   255
         Left            =   6480
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option15 
         Caption         =   "assigment"
         Height          =   195
         Left            =   4560
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option14 
         Caption         =   "overloading"
         Height          =   195
         Left            =   2640
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option13 
         Caption         =   "adding"
         Height          =   375
         Left            =   720
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "the continues ststement written by?"
      Height          =   975
      Left            =   1800
      TabIndex        =   2
      Top             =   2520
      Width           =   10455
      Begin VB.OptionButton Option12 
         Caption         =   "anywhere"
         Height          =   195
         Left            =   6360
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option11 
         Caption         =   "outside"
         Height          =   195
         Left            =   4200
         TabIndex        =   19
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option10 
         Caption         =   "nested"
         Height          =   195
         Left            =   2520
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option9 
         Caption         =   "body of loop"
         Height          =   195
         Left            =   600
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "A compound statement does not consista ?"
      Height          =   975
      Left            =   1800
      TabIndex        =   1
      Top             =   1560
      Width           =   10455
      Begin VB.OptionButton Option8 
         Caption         =   "none of above"
         Height          =   195
         Left            =   6720
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option7 
         Caption         =   "expression"
         Height          =   195
         Left            =   4200
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option6 
         Caption         =   "other compound"
         Height          =   315
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         Caption         =   "a sigle"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Redirection redirects"
      Height          =   855
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   10455
      Begin VB.OptionButton Option4 
         Caption         =   "noe of above"
         Height          =   195
         Left            =   5760
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "device to screen"
         Height          =   195
         Left            =   3720
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "file to device"
         Height          =   195
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "file to screen"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      Height          =   30000
      Left            =   0
      Picture         =   "cpp.frx":16E46
      Top             =   0
      Width           =   40005
   End
End
Attribute VB_Name = "cpp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim marks As Integer
marks = 0
If Option1.Value = True Then marks = marks + 1
If Option5.Value = True Then marks = marks + 1
If Option9.Value = True Then marks = marks + 1
If Option15.Value = True Then marks = marks + 1
If Option17.Value = True Then marks = marks + 1
If Option22.Value = True Then marks = marks + 1
If Option28.Value = True Then marks = marks + 1
If Option29.Value = True Then marks = marks + 1


If marks < 4 Then
MsgBox ("Your are fail ")
Else
MsgBox ("Congratulations ! Your are Pass")
End If
MsgBox ("Marks :  " & marks)
subject.Option2.Enabled = False
result.Label12.Caption = marks
If ((subject.Option1.Enabled = False) And (subject.Option2.Enabled = False) And (subject.Option3.Enabled = False) And (subject.Option4.Enabled = False) And (subject.Option5.Enabled = False)) Then
subject.Command1.Enabled = True
End If
Unload Me
End Sub

