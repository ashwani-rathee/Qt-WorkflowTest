VERSION 5.00
Begin VB.Form fLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WEIGHTSMART"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   12600
   FillColor       =   &H00800000&
   Icon            =   "fLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "fLogin.frx":1A7A
   ScaleHeight     =   8595
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame fraNewUser 
      BackColor       =   &H000000A0&
      Caption         =   "New User"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   4800
      TabIndex        =   20
      Top             =   960
      Width           =   3800
      Begin VB.TextBox txtNewConfPass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   23
         Top             =   1140
         Width           =   1935
      End
      Begin VB.TextBox txtNewPass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   22
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtNewID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   21
         Top             =   300
         Width           =   1935
      End
      Begin VB.CommandButton cmdCreate 
         BackColor       =   &H00C0C0C0&
         Caption         =   "C&reate"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1620
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C0C0C0&
         Cancel          =   -1  'True
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
         Index           =   2
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label6 
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
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   26
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame fraChangePass 
      BackColor       =   &H000000A0&
      Caption         =   "Change Password"
      ForeColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   180
      TabIndex        =   8
      Top             =   3840
      Width           =   3800
      Begin VB.ComboBox cmbConID 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtConConfPass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtConNewPass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1140
         Width           =   1935
      End
      Begin VB.TextBox txtConOldPass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00808080&
         Caption         =   "&Update"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1980
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00808080&
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
         Index           =   1
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1140
         Width           =   1260
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password"
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
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1170
      End
   End
   Begin VB.Frame fraLogin 
      BackColor       =   &H000000A0&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   5040
      TabIndex        =   4
      Top             =   3960
      Width           =   4935
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
         ItemData        =   "fLogin.frx":2F21D
         Left            =   1800
         List            =   "fLogin.frx":2F21F
         TabIndex        =   0
         Top             =   240
         Width           =   1935
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
         Left            =   3000
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
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtLoginPass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
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
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label lblID 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "USERNAME"
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
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   $"fLogin.frx":2F221
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1515
      Left            =   3900
      TabIndex        =   19
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Label lblTop1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   630
      Left            =   1650
      TabIndex        =   7
      Top             =   0
      Width           =   150
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
    Unload fMinShow
    If cmdClose(Index).Caption = "&Close" Then
        End
    Else
        Unload Me
        fMain.Show
    End If
End Sub

Private Sub cmdLogin_Click()
Dim spass As String
Unload main
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
       rs.Open "Select * from users where username ='" & Trim(cmbLoginID) & "'", co, adOpenKeyset, adLockOptimistic
       If rs.RecordCount > 0 Then
      ' Main.Label2.Caption = cmbLoginID.Text
       ' head.Logname = cmbLoginID.Text
        spass = Textdcrt(rs.Fields("password").Value)
        userpermission = rs.Fields("privilages").Value
        loginname = cmbLoginID
      '  username = cmbLoginID
        logintype = rs.Fields("utype").Value
        '''''If Date >= CDate(dat) + 20 Then
                  If demo = True Then
                    'checkinstall
                   If runt > 0 Then
                     
                    If (Date >= CDate(dat) + 25) Or runt > 150 Then
                       'Call Varify
                       ' rs.Fields("Pass").Value = TextEncr(str(Date) + Trim(Mid(cmbLoginID.Text, 1, 5)))
                       ' rs.Update
                    End If
                   
                   End If
                    
                 End If
                 
        '''End If
      
       End If
       If spass <> txtLoginPass Then
       MsgBox "Invalid password.. Please contact Administrator", vbInformation, "Error Password"
       txtLoginPass.SetFocus
       Else

       Unload Me
       
'       checkinstall
'loginname = cmbLoginID.Text
'username = loginname
loadshifts
main.Show

   
       End If
       
End If



End Sub

Sub loadshifts()
Set rs2 = New ADODB.Recordset
rs2.Open "Select * from shift where shiftname='A' or shiftname='a'", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
rs2.MoveFirst
shifta1 = rs2.Fields("shiftstart").Value
shifta2 = rs2.Fields("shiftend").Value
End If
Set rs2 = New ADODB.Recordset
rs2.Open "Select * from shift where shiftname='B' or shiftname='b'", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
rs2.MoveFirst
shiftb1 = rs2.Fields("shiftstart").Value
shiftb2 = rs2.Fields("shiftend").Value
End If
Set rs2 = New ADODB.Recordset
rs2.Open "Select * from shift where shiftname='C' or shiftname='C'", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
rs2.MoveFirst
shiftc1 = rs2.Fields("shiftstart").Value
shiftc2 = rs2.Fields("shiftend").Value
End If
End Sub

Private Sub checkvalid()
Set rs2 = New ADODB.Recordset
rs2.Open "Select * from inschk", co, adOpenKeyset, adLockOptimistic
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
rs2.Open "Select * from inschk", co, adOpenKeyset, adLockOptimistic
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
Conn
Set rs = New ADODB.Recordset
rs.Open "Select * from dbasepath", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
   Else
   MsgBox "You dont have permission for this task", vbInformation, "Contact Administrator"
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




Private Sub Varify()
    Set fsys = CreateObject("Scripting.FileSystemObject")
    If Not fsys.FileExists("C:\Windows\System\SYSDISYS.TXT") Then
        MsgBox "Some errors occured  during the installation" & vbCrLf & "Please contact the vendor Administrator", vbCritical, "Warning"
        End
    End If
End Sub

Private Sub Form_Load()
'Me.Picture = LoadPicture(App.Path + "\home.jpg")
Call Conn
    
    
    fraNewUser.BackColor = RGB(124, 92, 124)
    SetWindowRgn Me.hWnd, lreigon, True
 
    fraLogin.Left = (fLogin.Width * 3 / 4) - (fraLogin.Width / 2)
    
    fraChangePass.Left = (fLogin.Width * 3 / 4) - (fraChangePass.Width / 2)
    fraChangePass.Top = (fLogin.Height / 6) - (fraChangePass.Height / 2)
    fraNewUser.Left = (fLogin.Width * 3 / 4) - (fraNewUser.Width / 2)
    fraNewUser.Top = (fLogin.Height / 6) - (fraNewUser.Height / 2)
    lblTop1.Width = fLogin.Width
    lblTop1.Left = 0  '100
    lblTop1.Top = 0
    
  '  Call username1
   ' cmbLoginID.ListIndex = 0
        fraLogin.Visible = True
        fraNewUser.Visible = False
        fraChangePass.Visible = False

        lblNote.Visible = False
End Sub

Public Function ShowLoginForm(FormName As String)
For i = 0 To 2
    fLogin.cmdClose(i).Caption = "&Cancel"
Next i
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
        'fMinShow.lblMess.Caption = "New User Creation Form. For the users who will only have a permission to use the Software with their Login ID and Login Password."
    Case "ChangePass"
        fraLogin.Visible = False
        fraNewUser.Visible = False
        fraChangePass.Visible = True
'        lblTop.Caption = "Changing of Password - Window"
        lblTop1.Caption = "Changing of Password - Window"
        'fMinShow.lblMess.Caption = "To change the password of the user, select the Login ID and give the Old Password and the New Password and the confirmation of the New Password."
        
    Case "DeleteUser"
        fraLogin.Visible = True
        fraNewUser.Visible = False
        fraChangePass.Visible = False
        fraNewUser.Caption = "&Delete User"
        cmdLogin.Caption = "&Delete"
        txtLoginPass = ""
        lblTop.Caption = "Deletion of User - Window"
        lblTop1.Caption = "Deletion of User - Window"
        'fMinShow.lblMess.Caption = "Delete User Form. This form is to delete the existing user. Select the Login ID and give the valid Password for the selected user and click the delete button to delete the user."
    Case "DiffUser"
        fraLogin.Visible = True
        fraNewUser.Visible = False
        fraChangePass.Visible = False
        fraNewUser.Caption = "&Login"
        cmdLogin.Caption = "&Login"
        txtLoginPass = ""
'        lblTop.Caption = "New User Login - Window"
        lblTop1.Caption = "New User Login - Window"
        'fMinShow.lblMess.Caption = "Login as Different User Form. This form is to Log In as a different existing user."
End Select
lblTop1.Width = fLogin.Width
lblTop1.Left = 0

Me.Show
End Function

Private Sub username1()

Set rs = New ADODB.Recordset
rs.Open "Select * from users order by username", co, adOpenKeyset, adLockOptimistic
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
