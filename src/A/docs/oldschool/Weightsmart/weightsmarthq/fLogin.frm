VERSION 5.00
Begin VB.Form fLogin 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12000
   FillColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame fraLogin 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Caption         =   "Login"
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   5040
      TabIndex        =   4
      Top             =   3360
      Width           =   5055
      Begin VB.ComboBox cmbLoginID 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Close"
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
         Index           =   0
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdLogin 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Login"
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
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtLoginPass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2295
      End
      Begin VB.Image Image2 
         Height          =   1530
         Left            =   3600
         Picture         =   "fLogin.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "DADHWAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   1200
         TabIndex        =   8
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "INSTALLED AND MAINTAINED BY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   825
      End
      Begin VB.Label lblID 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Login ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   0
         Top             =   1800
         Width           =   5175
      End
   End
End
Attribute VB_Name = "fLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dat As Date
Dim runt As Integer


Private Sub cmbConID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtConOldPass.SetFocus
End Sub
Private Sub cmbLoginID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtLoginPass.SetFocus
End Sub

Private Sub cmdClose_Click(Index As Integer)
'    Unload fMinShow
    If cmdClose(Index).Caption = "&Close" Then
        End
    Else
        Unload Me
        fMain.Show
    End If
End Sub

Private Sub cmdLogin_Click()
Dim spass As String
'Unload main
'checkvalid
loginname = ""
logintype = ""
If cmdLogin.Caption = "&Login" Then
    If txtLoginPass = "" Then
       MsgBox "Password cannot be blank", vbInformation, "Incorrect Date Entered"
       txtLoginPass = ""
       txtLoginPass.SetFocus
       Exit Sub
    End If
       Set rs = New ADODB.Recordset
       rs.Open "Select * from users where username ='" & Trim(cmbLoginID) & "'", co1, adOpenKeyset, adLockOptimistic
       If rs.RecordCount > 0 Then
        spass = Textdcrt(rs.Fields("password").Value)
        userpermission = rs.Fields("privilages").Value
        loginname = cmbLoginID
        logintype = rs.Fields("utype").Value
       End If
       If spass <> txtLoginPass Then
       MsgBox "Invalid password.. Please contact Dadhwal Weighing", vbInformation, "Error Password"
       txtLoginPass.SetFocus
       Else

       Unload Me
       
        sumreport.Show

   
       End If
       
End If





End Sub

Private Sub checkvalid()
Set rs2 = New ADODB.Recordset
rs2.Open "Select * from inschk", co1, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
rs2.MoveFirst
dat = CDate(Textdcrt(rs2.Fields("idate").Value))
runt = Textdcrt(rs2.Fields("runtime").Value)
'MsgBox Textdcrt(rs2.Fields("idate").Value)
'MsgBox Textdcrt(rs2.Fields("runtime").Value)
End If

End Sub


Private Sub checkinstall()
Dim s As Integer
s = 0
Set rs2 = New ADODB.Recordset
rs2.Open "Select * from inschk", co1, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
rs2.MoveFirst
s = Textdcrt(rs2.Fields("runtime").Value)
Else
rs2.AddNew
rs2.Fields("idate").Value = TextEncr(CDate(Date))
s = 1
End If

rs2.Fields("runtime").Value = TextEncr(s + 1)
rs2.Update
End Sub



Public Sub dbaspath()
Conn1
Set rs = New ADODB.Recordset
rs.Open "Select * from dbasepath", co1, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
   Else
   MsgBox "You dont have permission for this task", vbInformation, "Contact Dadhwal Weighing"
   End
End If
End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If cmdClose(Index).Caption = "&Close" Then
        End
    Else
        Unload Me
        fMain.Show
    End If
End If
End Sub


Private Sub Form_Load()
Me.Picture = MDIForm1.Picture
Call Conn1
End Sub

Public Function ShowLoginForm(FormName As String)
For I = 0 To 2
    fLogin.cmdClose(I).Caption = "&Cancel"
Next I
Select Case FormName
    Case "NewUser"
        fraLogin.Visible = False
        fraNewUser.Visible = True
        fraChangePass.Visible = False
        fraNewUser.Caption = "New User"
        cmdCreate.Caption = "C&reate"
        txtNewID = ""
        txtNewPass = ""
        txtNewConfPass = ""
        txtNewID.Locked = False
        lblTop.Caption = "New User Creation - Window"
        lblTop1.Caption = "New User Creation - Window"
        fMinShow.lblMess.Caption = "New User Creation Form. For the users who will only have a permission to use the Software with their Login ID and Login Password."
    Case "ChangePass"
        fraLogin.Visible = False
        fraNewUser.Visible = False
        fraChangePass.Visible = True
'        lblTop.Caption = "Changing of Password - Window"
        lblTop1.Caption = "Changing of Password - Window"
        fMinShow.lblMess.Caption = "To change the password of the user, select the Login ID and give the Old Password and the New Password and the confirmation of the New Password."
        
    Case "DeleteUser"
        fraLogin.Visible = True
        fraNewUser.Visible = False
        fraChangePass.Visible = False
        fraNewUser.Caption = "&Delete User"
        cmdLogin.Caption = "&Delete"
        txtLoginPass = ""
        lblTop.Caption = "Deletion of User - Window"
        lblTop1.Caption = "Deletion of User - Window"
        fMinShow.lblMess.Caption = "Delete User Form. This form is to delete the existing user. Select the Login ID and give the valid Password for the selected user and click the delete button to delete the user."
    Case "DiffUser"
        fraLogin.Visible = True
        fraNewUser.Visible = False
        fraChangePass.Visible = False
        fraNewUser.Caption = "&Login"
        cmdLogin.Caption = "&Login"
        txtLoginPass = ""
'        lblTop.Caption = "New User Login - Window"
        lblTop1.Caption = "New User Login - Window"
        fMinShow.lblMess.Caption = "Login as Different User Form. This form is to Log In as a different existing user."
End Select
lblTop1.Width = fLogin.Width
lblTop1.Left = 0

Me.Show
End Function

Private Sub username1()

Set rs = New ADODB.Recordset
rs.Open "Select * from users order by username", co1, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveFirst
End If

While Not rs.EOF

cmbLoginID.AddItem rs.Fields("username").Value

rs.MoveNext
Wend
End Sub


Private Sub Form_Unload(Cancel As Integer)
   ' master.Enabled = True
    'Main.Enabled = True
End Sub

Private Sub txtConNewPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtConConfPass.SetFocus
End Sub

Private Sub txtConOldPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtConNewPass.SetFocus
End Sub

Private Sub txtLoginPass_Change()
    If txtLoginPass.Text <> "" Then
        cmdLogin.Default = True
    Else
        cmdLogin.Default = False
    End If
End Sub

Private Sub txtNewConfPass_Change()
    If txtNewConfPass.Text <> "" Then
        cmdCreate.Default = True
    Else
        cmdCreate.Default = False
    End If
End Sub

Private Sub txtNewConfPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 '   cmdCreate_Click
End If
End Sub

Private Sub txtNewID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNewPass.SetFocus
End Sub

Private Sub txtNewPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNewConfPass.SetFocus
End Sub
