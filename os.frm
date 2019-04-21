VERSION 5.00
Begin VB.Form os 
   Caption         =   "C# PROGRAMMING"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "os.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "The two kinds of main memory are:"
      Height          =   855
      Left            =   1920
      TabIndex        =   36
      Top             =   360
      Width           =   10455
      Begin VB.OptionButton Option1 
         Caption         =   "Primary and secondary"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "all of the above"
         Height          =   195
         Left            =   2520
         TabIndex        =   39
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "ROM and RAM"
         Height          =   195
         Left            =   4920
         TabIndex        =   38
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "None of the above"
         Height          =   195
         Left            =   7200
         TabIndex        =   37
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Which of the following assigns the number 5 to the area variable ?"
      Height          =   975
      Left            =   1920
      TabIndex        =   31
      Top             =   1320
      Width           =   10455
      Begin VB.OptionButton Option5 
         Caption         =   "area 1=5"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         Caption         =   "area = 5"
         Height          =   315
         Left            =   1800
         TabIndex        =   34
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option7 
         Caption         =   "area==5"
         Height          =   195
         Left            =   4200
         TabIndex        =   33
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option8 
         Caption         =   "area->5"
         Height          =   195
         Left            =   6720
         TabIndex        =   32
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Access time is the highest in the case of"
      Height          =   975
      Left            =   1920
      TabIndex        =   26
      Top             =   2280
      Width           =   10455
      Begin VB.OptionButton Option9 
         Caption         =   "Floppy disk"
         Height          =   195
         Left            =   600
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option10 
         Caption         =   "cache"
         Height          =   195
         Left            =   2520
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option11 
         Caption         =   "swapping devices"
         Height          =   195
         Left            =   4200
         TabIndex        =   28
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Option12 
         Caption         =   "magnetic disk"
         Height          =   195
         Left            =   6360
         TabIndex        =   27
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "The ___ symbol is used in a flowchart to represent a calculation task."
      Height          =   975
      Left            =   1920
      TabIndex        =   21
      Top             =   3360
      Width           =   10455
      Begin VB.OptionButton Option13 
         Caption         =   "Input"
         Height          =   375
         Left            =   720
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option14 
         Caption         =   "Output"
         Height          =   195
         Left            =   2640
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option15 
         Caption         =   "Process"
         Height          =   195
         Left            =   4560
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option16 
         Caption         =   "Start"
         Height          =   255
         Left            =   6480
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "A sequence of instructions, in a computer language, to get the desired result, is known as"
      Height          =   855
      Left            =   1920
      TabIndex        =   16
      Top             =   4440
      Width           =   10455
      Begin VB.OptionButton Option17 
         Caption         =   "Algorithm"
         Height          =   195
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option18 
         Caption         =   "Decision Table"
         Height          =   195
         Left            =   2520
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option19 
         Caption         =   "Program"
         Height          =   195
         Left            =   4560
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option20 
         Caption         =   "All of the above"
         Height          =   195
         Left            =   6720
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "In an absolute loading scheme , which loader function is accomplished by loader ? "
      Height          =   975
      Left            =   1920
      TabIndex        =   11
      Top             =   5280
      Width           =   10455
      Begin VB.OptionButton Option21 
         Caption         =   "Reallocation"
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option22 
         Caption         =   "Allocation"
         Height          =   195
         Left            =   2160
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option23 
         Caption         =   "Linking"
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option24 
         Caption         =   "Loading"
         Height          =   195
         Left            =   6000
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "A medium for transferring data between two locations is called"
      Height          =   1095
      Left            =   1920
      TabIndex        =   6
      Top             =   6360
      Width           =   10455
      Begin VB.OptionButton Option25 
         Caption         =   "Network"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option26 
         Caption         =   "Communication Channel"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton Option27 
         Caption         =   "Modern"
         Height          =   195
         Left            =   5400
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option28 
         Caption         =   "Bus"
         Height          =   195
         Left            =   7560
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Which of the following CANNOT occur multiple number of times in a program ?"
      Height          =   975
      Left            =   1920
      TabIndex        =   1
      Top             =   7560
      Width           =   10455
      Begin VB.OptionButton Option29 
         Caption         =   "Namespace"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option30 
         Caption         =   "Entrypoint"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option31 
         Caption         =   "Class"
         Height          =   255
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option32 
         Caption         =   "Function"
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
      Left            =   5880
      TabIndex        =   0
      Top             =   8640
      Width           =   1815
   End
End
Attribute VB_Name = "os"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim marks As Integer
marks = 0
If Option3.Value = True Then marks = marks + 1
If Option6.Value = True Then marks = marks + 1
If Option12.Value = True Then marks = marks + 1
If Option15.Value = True Then marks = marks + 1
If Option19.Value = True Then marks = marks + 1
If Option24.Value = True Then marks = marks + 1
If Option26.Value = True Then marks = marks + 1
If Option30.Value = True Then marks = marks + 1




If marks < 4 Then
MsgBox ("Your are fail ")
Else
MsgBox ("Congratulations ! Your are Pass")
End If
MsgBox ("Marks :  " & marks)
result.Label14.Caption = marks
subject.Option3.Enabled = False
If ((subject.Option1.Enabled = False) And (subject.Option2.Enabled = False) And (subject.Option3.Enabled = False) And (subject.Option4.Enabled = False) And (subject.Option5.Enabled = False)) Then
subject.Command1.Enabled = True
End If
Unload Me
End Sub

