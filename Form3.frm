VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form3"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14535
   LinkTopic       =   "Form3"
   ScaleHeight     =   6960
   ScaleWidth      =   14535
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   6975
      Left            =   0
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   6915
      ScaleWidth      =   14475
      TabIndex        =   0
      Top             =   0
      Width           =   14535
      Begin VB.CommandButton Command5 
         Caption         =   "New Booking"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         TabIndex        =   9
         Top             =   6000
         Width           =   6615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6360
         Picture         =   "Form3.frx":145B32
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6360
         Picture         =   "Form3.frx":1485A4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6360
         Picture         =   "Form3.frx":14B016
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6360
         Picture         =   "Form3.frx":14DA88
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "JW Marriott  Mumbai"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   6
         Top             =   2640
         Width           =   5415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Hyatt Hotel Mumbai"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   1800
         Width           =   5295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hotel Taj Mumbai"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "List of 5 star hotel in mumbai"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   7455
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()
Form6.Show
Me.Hide
End Sub

Private Sub Command5_Click()
Form8.Show
End Sub
