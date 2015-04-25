VERSION 5.00
Begin VB.Form form4 
   Caption         =   "Form3"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15045
   LinkTopic       =   "Form3"
   ScaleHeight     =   7950
   ScaleWidth      =   15045
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H000000FF&
      Height          =   7935
      Left            =   0
      Picture         =   "loggin_window.frx":0000
      ScaleHeight     =   7875
      ScaleWidth      =   15915
      TabIndex        =   0
      Top             =   120
      Width           =   15975
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C000&
         Caption         =   "Click here to enter "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6600
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You are loggin successfully"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2175
         Left            =   3960
         TabIndex        =   1
         Top             =   600
         Width           =   8535
      End
   End
End
Attribute VB_Name = "form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'form3 shown
Form5.Show
Me.Hide

End Sub
