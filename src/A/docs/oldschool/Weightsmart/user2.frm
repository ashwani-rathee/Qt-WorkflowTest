VERSION 5.00
Begin VB.Form passwd 
   Caption         =   "Password Modification"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12540
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   12540
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   4815
      Left            =   1320
      TabIndex        =   7
      Top             =   1560
      Width           =   9735
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   36
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   34
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   17
         Text            =   "Operator"
         Top             =   4200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   3840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   2480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   2920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check5 
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   3840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   2480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check7 
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   2920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check8 
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Confirm New Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " New Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " User_Id"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   30
         Top             =   3840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " User Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   4200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Old Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "User Creation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Item Master"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2475
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Invoicing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Party Master"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2950
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Dupicate Bill"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Modify Parameter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   2475
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Call List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   2920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Client Call Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5535
      Left            =   11640
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   3255
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4905
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1320
      TabIndex        =   1
      Top             =   6240
      Width           =   9735
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   9735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   5400
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   0
      ToolTipText     =   "Unload Form"
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User  Modify"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   4920
      TabIndex        =   32
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   1320
      TabIndex        =   33
      Top             =   960
      Width           =   9735
   End
End
Attribute VB_Name = "passwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recfound As Boolean
Dim g As Integer
Dim str1 As String
Dim num As Integer

Private Sub dataset()
recfound = False
Set rs = New ADODB.Recordset
rs.Open "Select * from users where username='" & Trim(Text2.Text) & "' and password='" & TextEncr(Trim(Text3.Text)) & "'", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
recfound = True
Command1.Caption = "Modify"
End If
End Sub


Private Sub Command1_Click()
a = MsgBox("Change Password ?", vbOKCancel, "Save Data")
If a = 1 Then
dataset
If recfound = True Then
    'rs.Fields("userid").Value = Trim(Text1.Text)
    rs.Fields("username").Value = Trim(Text2.Text)
    rs.Fields("password").Value = TextEncr(Trim(Text4.Text))
    rs.Fields("utype").Value = Trim(Combo1.Text)
    rs.Fields("privilages").Value = Trim(Check1.Value) + Trim(Check2.Value) + Trim(Check3.Value) + Trim(Check4.Value) + Trim(Check5.Value) + Trim(Check6.Value) + Trim(Check7.Value) + Trim(Check8.Value)
    rs.Update
Else
    MsgBox "Username password mismatch, please try again"
End If
End If

unloctxt
End Sub

Private Sub unloctxt()
Text3.Text = ""
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
Check7.Value = 0
Check8.Value = 0
End Sub



Private Sub Command1_GotFocus()
If Trim(Text2.Text) = "" Then
MsgBox "Please Enter the User name ", vbCritical, "Error"
Text2.SetFocus
Exit Sub
End If



If Trim(Text3.Text) = "" Then
MsgBox "Please Enter the Old Password ", vbCritical, "Error"
Text3.SetFocus
Exit Sub
End If

If Trim(Text4.Text) = "" Then
MsgBox "Please Enter the New Password ", vbCritical, "Error"
Text3.SetFocus
Exit Sub
End If

If Trim(Text5.Text) = "" Then
MsgBox "Please Confirm New Password ", vbCritical, "Error"
Text3.SetFocus
Exit Sub
End If


End Sub


Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = main.Picture

On Error Resume Next
Text2.Text = loginname
End Sub



