VERSION 5.00
Begin VB.Form java 
   Caption         =   "JAVA"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Which method must be defined by a class implementing the java.lang.Runnable interface ?"
      Height          =   855
      Left            =   1560
      TabIndex        =   36
      Top             =   120
      Width           =   10455
      Begin VB.OptionButton Option1 
         Caption         =   "void run()"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Public void run()"
         Height          =   195
         Left            =   2040
         TabIndex        =   39
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Caption         =   "public void start()"
         Height          =   195
         Left            =   4560
         TabIndex        =   38
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "void run (int priority)"
         Height          =   195
         Left            =   6720
         TabIndex        =   37
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Which class does not override the eqals() and hashcode() methods,inheriting them directly from class object ?"
      Height          =   975
      Left            =   1560
      TabIndex        =   31
      Top             =   1080
      Width           =   10455
      Begin VB.OptionButton Option5 
         Caption         =   "java.lang.string"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option6 
         Caption         =   "java.lang.double"
         Height          =   315
         Left            =   2040
         TabIndex        =   34
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option7 
         Caption         =   "java.lang.stringbuffer"
         Height          =   195
         Left            =   4200
         TabIndex        =   33
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option8 
         Caption         =   "java.lang.character"
         Height          =   195
         Left            =   6720
         TabIndex        =   32
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Which one is a valid declaration of a boolean ?"
      Height          =   975
      Left            =   1560
      TabIndex        =   26
      Top             =   2040
      Width           =   10455
      Begin VB.OptionButton Option9 
         Caption         =   "boolean b1=0"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option10 
         Caption         =   "boolean b2='false'"
         Height          =   195
         Left            =   2040
         TabIndex        =   29
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton Option11 
         Caption         =   "boolean b3=false"
         Height          =   195
         Left            =   4320
         TabIndex        =   28
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Option12 
         Caption         =   "boolean b4=no"
         Height          =   195
         Left            =   6840
         TabIndex        =   27
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Which is a reversed word in the java programming language ?"
      Height          =   975
      Left            =   1560
      TabIndex        =   21
      Top             =   3120
      Width           =   10455
      Begin VB.OptionButton Option13 
         Caption         =   "method"
         Height          =   375
         Left            =   720
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option14 
         Caption         =   "native"
         Height          =   195
         Left            =   2640
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option15 
         Caption         =   "subclasses"
         Height          =   195
         Left            =   4560
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option16 
         Caption         =   "reference"
         Height          =   255
         Left            =   6480
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Which one of the following will declare an array and initialize it with five numbers ?"
      Height          =   855
      Left            =   1560
      TabIndex        =   16
      Top             =   4200
      Width           =   10455
      Begin VB.OptionButton Option17 
         Caption         =   "array a = new array(5)"
         Height          =   195
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option18 
         Caption         =   "int[] a={23,22,21,20,19}"
         Height          =   195
         Left            =   2640
         TabIndex        =   19
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton Option19 
         Caption         =   "int a [] = new int [5]"
         Height          =   195
         Left            =   4800
         TabIndex        =   18
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option20 
         Caption         =   "int [5] array"
         Height          =   195
         Left            =   6720
         TabIndex        =   17
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "What is the prototype of the default contructor ?"
      Height          =   975
      Left            =   1560
      TabIndex        =   11
      Top             =   5040
      Width           =   10455
      Begin VB.OptionButton Option21 
         Caption         =   "test()"
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option22 
         Caption         =   "test(void)"
         Height          =   195
         Left            =   2160
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option23 
         Caption         =   "public Test()"
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option24 
         Caption         =   "public Test (void)"
         Height          =   195
         Left            =   6000
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Which is valid declaration of float ?"
      Height          =   1095
      Left            =   1560
      TabIndex        =   6
      Top             =   6120
      Width           =   10455
      Begin VB.OptionButton Option25 
         Caption         =   "float f=1F"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option26 
         Caption         =   "float f = 1.0"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option27 
         Caption         =   "float f = ""1"""
         Height          =   195
         Left            =   4200
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option28 
         Caption         =   "float f = 1.0d"
         Height          =   195
         Left            =   6360
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Which will contain the body of the thread ?"
      Height          =   975
      Left            =   1560
      TabIndex        =   1
      Top             =   7320
      Width           =   10455
      Begin VB.OptionButton Option29 
         Caption         =   "run()"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option30 
         Caption         =   "start()"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option31 
         Caption         =   "stop()"
         Height          =   255
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option32 
         Caption         =   "main()"
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
      Left            =   5520
      TabIndex        =   0
      Top             =   8400
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   13575
      Left            =   0
      Picture         =   "java.frx":0000
      Top             =   0
      Width           =   19200
   End
End
Attribute VB_Name = "java"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim marks As Integer
marks = 0
If Option2.Value = True Then marks = marks + 1
If Option7.Value = True Then marks = marks + 1
If Option11.Value = True Then marks = marks + 1
If Option14.Value = True Then marks = marks + 1
If Option18.Value = True Then marks = marks + 1
If Option23.Value = True Then marks = marks + 1
If Option25.Value = True Then marks = marks + 1
If Option29.Value = True Then marks = marks + 1




If marks < 4 Then
MsgBox ("Your are fail ")
Else
MsgBox ("Congratulations ! Your are Pass")
End If
MsgBox ("Marks :  " & marks)
result.Label18.Caption = marks
subject.Option5.Enabled = False
If ((subject.Option1.Enabled = False) And (subject.Option2.Enabled = False) And (subject.Option3.Enabled = False) And (subject.Option4.Enabled = False) And (subject.Option5.Enabled = False)) Then
subject.Command1.Enabled = True
End If
Unload Me

End Sub

