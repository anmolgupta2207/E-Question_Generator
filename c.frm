VERSION 5.00
Begin VB.Form c 
   Caption         =   "C"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "c.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "How would you round off a value from 1.66 to 2.0?"
      Height          =   855
      Left            =   2280
      TabIndex        =   36
      Top             =   480
      Width           =   10455
      Begin VB.OptionButton Option1 
         Caption         =   "ceil(1.66)"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "floor(1.66)"
         Height          =   195
         Left            =   2040
         TabIndex        =   39
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "roundup(1.66)"
         Height          =   195
         Left            =   3720
         TabIndex        =   38
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "roundto(1.66)"
         Height          =   195
         Left            =   5760
         TabIndex        =   37
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Which of the following is the correct order of evaluation for the "" z = x + y * z / 4 % 2 - 1  """
      Height          =   975
      Left            =   2280
      TabIndex        =   31
      Top             =   1440
      Width           =   10455
      Begin VB.OptionButton Option5 
         Caption         =   "*/%+-="
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         Caption         =   "=*/%+-"
         Height          =   315
         Left            =   1800
         TabIndex        =   34
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option7 
         Caption         =   "/*%-+="
         Height          =   195
         Left            =   4200
         TabIndex        =   33
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option8 
         Caption         =   "*%/-+="
         Height          =   195
         Left            =   6720
         TabIndex        =   32
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "The keyword used to transfer control from a function back to the calling function is"
      Height          =   975
      Left            =   2280
      TabIndex        =   26
      Top             =   2400
      Width           =   10455
      Begin VB.OptionButton Option9 
         Caption         =   "switch"
         Height          =   195
         Left            =   600
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option10 
         Caption         =   "goto"
         Height          =   195
         Left            =   2520
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option11 
         Caption         =   "goback"
         Height          =   195
         Left            =   4200
         TabIndex        =   28
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option12 
         Caption         =   "return"
         Height          =   195
         Left            =   6360
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "The library function used to find the last occurrence of a character in a string is"
      Height          =   975
      Left            =   2280
      TabIndex        =   21
      Top             =   3480
      Width           =   10455
      Begin VB.OptionButton Option13 
         Caption         =   "strnstr()"
         Height          =   375
         Left            =   720
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option14 
         Caption         =   "laststr()"
         Height          =   195
         Left            =   2640
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option15 
         Caption         =   "strrchr()"
         Height          =   195
         Left            =   4560
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option16 
         Caption         =   "strstr()"
         Height          =   255
         Left            =   6480
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Which header file should you include, if you are going to devlop a function, which can accept variable number of arguments()"
      Height          =   855
      Left            =   2280
      TabIndex        =   16
      Top             =   4560
      Width           =   10455
      Begin VB.OptionButton Option17 
         Caption         =   "varagrg.h"
         Height          =   195
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option18 
         Caption         =   "stdlib.h"
         Height          =   195
         Left            =   2520
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option19 
         Caption         =   "stdio.h"
         Height          =   195
         Left            =   4680
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option20 
         Caption         =   "stdarg.h"
         Height          =   195
         Left            =   6720
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Which of the following function sets n characters of string to a given character?"
      Height          =   975
      Left            =   2280
      TabIndex        =   11
      Top             =   5400
      Width           =   10455
      Begin VB.OptionButton Option21 
         Caption         =   "strinit()"
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option22 
         Caption         =   "strnset()"
         Height          =   195
         Left            =   2160
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option23 
         Caption         =   "strset()"
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option24 
         Caption         =   "strcset()"
         Height          =   195
         Left            =   6000
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Which of the following cannot be checked in a switch-case statement ?"
      Height          =   1095
      Left            =   2280
      TabIndex        =   6
      Top             =   6480
      Width           =   10455
      Begin VB.OptionButton Option25 
         Caption         =   "character"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option26 
         Caption         =   "integer"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option27 
         Caption         =   "float"
         Height          =   195
         Left            =   4200
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option28 
         Caption         =   "enum"
         Height          =   195
         Left            =   6360
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "The library function used to reverse a string is"
      Height          =   975
      Left            =   2280
      TabIndex        =   1
      Top             =   7680
      Width           =   10455
      Begin VB.OptionButton Option29 
         Caption         =   "strstr()"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option30 
         Caption         =   "strrev()"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option31 
         Caption         =   "revstr()"
         Height          =   255
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option32 
         Caption         =   "strreverse()"
         Height          =   255
         Left            =   5880
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      Height          =   615
      Left            =   6240
      TabIndex        =   0
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   12360
      Left            =   0
      Picture         =   "c.frx":C510B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim marks As Integer
marks = 0
If Option3.Value = True Then marks = marks + 1
If Option5.Value = True Then marks = marks + 1
If Option12.Value = True Then marks = marks + 1
If Option15.Value = True Then marks = marks + 1
If Option20.Value = True Then marks = marks + 1
If Option22.Value = True Then marks = marks + 1
If Option27.Value = True Then marks = marks + 1
If Option30.Value = True Then marks = marks + 1




If marks < 4 Then
MsgBox ("Your are fail ")
Else
MsgBox ("Congratulations ! Your are Pass")
End If
MsgBox ("Marks :  " & marks)
subject.Option1.Enabled = False

result.Label16.Caption = marks
If ((subject.Option1.Enabled = False) And (subject.Option2.Enabled = False) And (subject.Option3.Enabled = False) And (subject.Option4.Enabled = False) And (subject.Option5.Enabled = False)) Then
subject.Command1.Enabled = True
End If
Unload Me
End Sub

