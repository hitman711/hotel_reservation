VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14580
   LinkTopic       =   "Form6"
   ScaleHeight     =   8160
   ScaleWidth      =   14580
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   8415
      Left            =   0
      Picture         =   "Form6.frx":0000
      ScaleHeight     =   8355
      ScaleWidth      =   15435
      TabIndex        =   0
      Top             =   0
      Width           =   15495
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   720
         Width           =   4455
      End
      Begin VB.PictureBox Picture2 
         Height          =   4815
         Left            =   7680
         Picture         =   "Form6.frx":18454
         ScaleHeight     =   4755
         ScaleWidth      =   5355
         TabIndex        =   6
         Top             =   2880
         Width           =   5415
      End
      Begin VB.Label Label7 
         Caption         =   "Ratan Tata Chairmen of Tata Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   7
         Top             =   6840
         Width           =   5175
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.tajhotels.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   5520
         Width           =   5775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Official website"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   4920
         Width           =   3975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "The Taj Mahal Palace, Mumbai ,Apollo Bunder, Mumbai - 400 001 Maharashtra, India"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1455
         Left            =   9240
         TabIndex        =   3
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   8040
         TabIndex        =   2
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "The Taj Mahal Palace"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   7935
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

End Sub

Private Sub Command1_Click()
'form5 shown
From5.Show
Me.Hide
End Sub
