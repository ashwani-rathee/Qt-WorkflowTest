VERSION 5.00
Begin VB.Form login 
   Caption         =   "Super Login "
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   13140
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4440
      TabIndex        =   4
      Text            =   "Password"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4440
      TabIndex        =   3
      Text            =   "Username"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5880
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5880
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2760
      Width           =   2415
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Conn
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from dbpass", co, adOpenKeyset, adLockOptimistic
rs1.MoveFirst

olduser = Textdcrt(rs1.Fields("uname").Value)
oldpass = Textdcrt(rs1.Fields("pass").Value)

If Trim(Text1.Text) = olduser And Trim(Text2.Text) = oldpass Then
rs1.Close
Unload Me
MDIForm1.cld.Visible = True
MDIForm1.cdp.Visible = True
MDIForm1.quit.Visible = True

Else
MsgBox "Invalid login"
Unload MDIForm1
End If

End Sub

Private Sub Form_Load()
Me.Picture = MDIForm1.Picture
MDIForm1.cld.Visible = False
MDIForm1.cdp.Visible = False
MDIForm1.quit.Visible = False
End Sub
