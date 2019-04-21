VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Height          =   2775
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   6135
      Begin VB.TextBox txtqty 
         Height          =   285
         Left            =   4080
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtdesc 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtrate 
         Height          =   285
         Left            =   2880
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblPUprodcode 
         AutoSize        =   -1  'True
         Caption         =   "Prod Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblPUamt 
         AutoSize        =   -1  'True
         Caption         =   "Amt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   15
         Top             =   360
         Width           =   330
      End
      Begin VB.Label lblPUqty 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4080
         TabIndex        =   14
         Top             =   360
         Width           =   300
      End
      Begin VB.Label lblPUrate 
         AutoSize        =   -1  'True
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   13
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lblPUdesc 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "prod desc"
      Height          =   975
      Left            =   3000
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
      Begin VB.ListBox List1 
         Height          =   450
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "VIEW PREVIOUS RECORD"
      Top             =   4920
      Width           =   1095
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "VIEW LAST RECORD"
      Top             =   4920
      Width           =   1215
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "VIEW NEXT RECORD"
      Top             =   4920
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "VIEW PREVIOUS RECORD"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txttotal 
      Height          =   285
      Left            =   6120
      TabIndex        =   0
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblPUtotal 
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   4320
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
