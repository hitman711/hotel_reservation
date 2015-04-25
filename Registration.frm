VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form form2 
   Caption         =   "User Registration Form"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14640
   LinkTopic       =   "Form2"
   ScaleHeight     =   7920
   ScaleWidth      =   14640
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   8
      Top             =   2400
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   8055
      Left            =   0
      Picture         =   "Registration.frx":0000
      ScaleHeight     =   7995
      ScaleWidth      =   15915
      TabIndex        =   0
      Top             =   0
      Width           =   15975
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   10320
         Top             =   6960
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=HRM"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "HRM"
         OtherAttributes =   ""
         UserName        =   "scott"
         Password        =   "tiger"
         RecordSource    =   "HRM"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox Text1 
         Height          =   1935
         Index           =   3
         Left            =   2880
         TabIndex        =   10
         Top             =   3840
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   9
         Top             =   3120
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   7
         Top             =   1440
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000040C0&
         Caption         =   "Registration  complete"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         MaskColor       =   &H000000FF&
         Picture         =   "Registration.frx":1A0E9E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6480
         Width           =   3855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"Registration.frx":1B09E4
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   1095
         Left            =   7080
         TabIndex        =   6
         Top             =   6240
         Width           =   7695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   5
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   3
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   2415
      End
   End
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim gen, flag As String


Private Sub Command1_Click()
cmd.ActiveConnection = cn
cmd.CommandType = adCmdText
cmd.CommandText = " insert into HRM values('" & Text1(0).Text & "','" & Text1(1).Text & "'," & Text1(2).Text & ",'" & Text1(3).Text & "')"
cmd.Execute
rs.Requery
MsgBox ("you have register successfully ")
form3.Show
Me.Hide
End Sub

Private Sub List1_Click()
'makes sure there is a value in all of them
   If (Index) = "" And DontCheck = False Then
      MsgBox "Please enter a value"
      DontCheck = True
      txtInfo(Index).SetFocus
   Else
      DontCheck = False
   End If
   
End Sub



Private Sub Text2_Change()
'makes sure there is a value in all of them
   If Text2(Index) = "" And DontCheck = False Then
      MsgBox "Please enter a value"
      DontCheck = True
      Text2(Index).SetFocus
   Else
      DontCheck = False
   End If
   
End Sub

Private Sub Text3_Change()
'makes sure there is a value in all of them
   If Text3(Index) = "" And DontCheck = False Then
      MsgBox "Please enter a value"
      DontCheck = True
      Text3(Index).SetFocus
   Else
      DontCheck = False
   End If
   
End Sub

Private Sub Text4_Change()
'makes sure there is a value in all of them
   If Text4(Index) = "" And DontCheck = False Then
      MsgBox "Please enter a value"
      DontCheck = True
      Text4(Index).SetFocus
   Else
      DontCheck = False
   End If
   
End Sub


Private Sub Form_Load()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

cn.Open "Provider=OraOLEDB.Oracle.1;user id=scott;password=tiger;data source= "
rs.Open "select * from HRM", cn, adOpenKeyset
Command1.Visible = False
End Sub



Private Sub Text1_Change(Index As Integer)
If (Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(2).Text = "") Then
    Command1.Visible = False
Else
    Command1.Visible = True
End If
End Sub
