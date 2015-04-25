VERSION 5.00
Begin VB.Form form1 
   Caption         =   "HOTEL BOOKING SYSTEM"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   14640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton User_login 
      BackColor       =   &H00FFFF80&
      Caption         =   "USER LOG IN"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   3375
   End
   Begin VB.CommandButton New_User 
      BackColor       =   &H00FFFF00&
      Caption         =   "NEW USER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   7455
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   7395
      ScaleWidth      =   14475
      TabIndex        =   1
      Top             =   0
      Width           =   14535
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "HOTEL BOOKING IN MUMBAI"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   975
         Left            =   5520
         TabIndex        =   2
         Top             =   2640
         Width           =   7815
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub New_User_Click()
'register info form shown
form2.Show
Me.Hide
End Sub

Private Sub User_login_Click()
'user_info form shown
form3.Show
Me.Hide
End Sub
