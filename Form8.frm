VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form8"
   ScaleHeight     =   7695
   ScaleWidth      =   10575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_delete 
      Caption         =   "Cancle Booking"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5760
      TabIndex        =   28
      Top             =   6960
      Width           =   2895
   End
   Begin VB.CommandButton cmd_nxt 
      Caption         =   "NEXT"
      Height          =   375
      Left            =   9480
      TabIndex        =   27
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmd_prv 
      Caption         =   "PREVIOUS"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox mid 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3840
      TabIndex        =   25
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox child_txt 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7320
      TabIndex        =   9
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox adult_txt 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2880
      TabIndex        =   8
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmd_clear 
      Caption         =   "Clear Form"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9000
      TabIndex        =   23
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox style 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4440
      TabIndex        =   6
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox room 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1440
      TabIndex        =   5
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton cmd_booking 
      Caption         =   "Complete Booking"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   22
      Top             =   6960
      Width           =   3375
   End
   Begin VB.TextBox bed 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7320
      TabIndex        =   7
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox hn 
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   5880
      Width           =   4695
   End
   Begin VB.TextBox co 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5400
      TabIndex        =   4
      Top             =   2280
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2880
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   "DSN=bookin"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "bookin"
      OtherAttributes =   ""
      UserName        =   "scott"
      Password        =   "tiger"
      RecordSource    =   "BOOKIN"
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
   Begin VB.TextBox sur 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7200
      TabIndex        =   2
      Top             =   1020
      Width           =   3015
   End
   Begin VB.TextBox ci 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1200
      TabIndex        =   3
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox name_txt 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   240
      TabIndex        =   1
      Top             =   1020
      Width           =   3015
   End
   Begin VB.Label welcome_lab 
      Alignment       =   1  'Right Justify
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   24
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "Single bed OR double bed"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   21
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Room type"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   20
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label10 
      Caption         =   "Childs"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   19
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Adults"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Hotel Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   17
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Ex. 29_02_2010"
      Height          =   255
      Left            =   8520
      TabIndex        =   16
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   525
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Middle Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   14
      Top             =   525
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Surname"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   13
      Top             =   525
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Check in date "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Check out date"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Rooms"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   3240
      Width           =   2535
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim reg_rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim gen, flag As String

Private Sub cmd_booking_Click()
cmd.ActiveConnection = cn
cmd.CommandType = adCmdText
cmd.CommandText = " insert into bookin values('" _
                    & name_txt.Text & "','" _
                    & mid.Text & "','" _
                    & sur.Text & "','" _
                    & ci.Text & "','" _
                    & co.Text & "'," _
                    & room.Text & ",'" _
                    & style.Text & "','" _
                    & bed.Text & "'," _
                    & adult_txt.Text & "," _
                    & child_txt.Text & ",'" _
                    & hn.Text & "')"
cmd.Execute
reg_rs.Requery
MsgBox ("Booking Complete")
reg_rs.MoveLast
Call ref
End Sub

Private Sub cmd_clear_Click()
name_txt.Text = ""
mid.Text = ""
sur.Text = ""
ci.Text = ""
co.Text = ""
room.Text = ""
style.Text = ""
bed.Text = ""
adult_txt.Text = ""
child_txt.Text = ""
hn.Text = ""
End Sub

Private Sub cmd_delete_Click()
MsgBox ("Are you Sure, You want to delete this Booking?")
cmd.ActiveConnection = cn
cmd.CommandType = adCmdText
cmd.CommandText = "delete from bookin where name = '" & name_txt.Text _
                    & "' AND check_in = '" & ci.Text & "'"
cmd.Execute
reg_rs.Requery
MsgBox ("Booking Canceled")
Call ref
End Sub

Private Sub cmd_nxt_Click()
If Not reg_rs.EOF Then
     reg_rs.MoveNext
   If reg_rs.EOF Then
      reg_rs.MoveLast
   End If
   Call ref
 End If
End Sub

Private Sub cmd_prv_Click()
If Not reg_rs.EOF Then
    reg_rs.MovePrevious
     If reg_rs.BOF Then
        reg_rs.MoveFirst
     End If
     Call ref
  End If
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
Set reg_rs = New ADODB.Recordset
Set cmd = New ADODB.Command

cn.Open "Provider=OraOLEDB.Oracle.1;user id=scott;password=tiger;data source= "
reg_rs.Open "select * from bookin", cn, adOpenKeyset
name_txt.Text = form3.user_rs(0)
welcome_lab.Caption = "Welcome " & form3.user_rs(0) & ","
Call ref
End Sub

Private Sub ref()
If Not reg_rs.EOF Then
name_txt.Text = reg_rs(0)
mid.Text = reg_rs(1)
sur.Text = reg_rs(2)
ci.Text = reg_rs(3)
co.Text = reg_rs(4)
room.Text = reg_rs(5)
style.Text = reg_rs(6)
bed.Text = reg_rs(7)
adult_txt.Text = reg_rs(8)
child_txt.Text = reg_rs(9)
hn.Text = reg_rs(10)
Else
    MsgBox ("No Previous Registation Found")
End If
End Sub


