VERSION 5.00
Begin VB.Form Users 
   Caption         =   "Users"
   ClientHeight    =   10980
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
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
      Left            =   12480
      TabIndex        =   11
      ToolTipText     =   "Unload Form"
      Top             =   1080
      Width           =   375
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   3240
      TabIndex        =   7
      Top             =   6360
      Width           =   6495
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   5400
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5535
      Left            =   9720
      TabIndex        =   5
      Top             =   1560
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
   Begin VB.Frame Frame2 
      Height          =   4815
      Left            =   3240
      TabIndex        =   0
      Top             =   1560
      Width           =   6495
      Begin VB.CheckBox Check8 
         Height          =   255
         Left            =   4080
         TabIndex        =   32
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check7 
         Height          =   255
         Left            =   4080
         TabIndex        =   30
         Top             =   2920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Left            =   4080
         TabIndex        =   28
         Top             =   2480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check5 
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   2920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   2480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
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
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   840
         Width           =   1455
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
         ItemData        =   "Users.frx":0000
         Left            =   1800
         List            =   "Users.frx":000D
         TabIndex        =   14
         Text            =   "Operator"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
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
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
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
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   2895
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
         TabIndex        =   33
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
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
         TabIndex        =   31
         Top             =   2920
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
         Left            =   2520
         TabIndex        =   29
         Top             =   2475
         Visible         =   0   'False
         Width           =   1575
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
         Left            =   2520
         TabIndex        =   27
         Top             =   2040
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
         TabIndex        =   25
         Top             =   2950
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
         TabIndex        =   23
         Top             =   3360
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
         Left            =   240
         TabIndex        =   21
         Top             =   2480
         Visible         =   0   'False
         Width           =   1335
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
         TabIndex        =   18
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Password"
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
         TabIndex        =   15
         Top             =   840
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
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
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
         Left            =   4800
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
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
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User  Master"
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
      Left            =   6720
      TabIndex        =   17
      Top             =   1080
      Width           =   1755
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   3240
      TabIndex        =   12
      Top             =   960
      Width           =   9735
   End
End
Attribute VB_Name = "Users"
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
rs.Open "Select * from users where userid='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
recfound = True
Command1.Caption = "Modify"
Else
Command1.Caption = "Save"
recfound = False
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.SetFocus
End If

End Sub

Private Sub Command1_Click()
a = MsgBox("Save Records", vbOKCancel, "Save Data")
If a = 1 Then
dataset
If recfound = False Then
rs.AddNew
End If
rs.Fields("userid").Value = Trim(Text1.Text)
rs.Fields("username").Value = Trim(Text2.Text)
rs.Fields("password").Value = TextEncr(Trim(Text3.Text))
rs.Fields("utype").Value = Trim(Combo1.Text)
rs.Fields("privilages").Value = Trim(Check1.Value) + Trim(Check2.Value) + Trim(Check3.Value) + Trim(Check4.Value) + Trim(Check5.Value) + Trim(Check6.Value) + Trim(Check7.Value) + Trim(Check8.Value)
rs.Update
End If

If newcatname = True Then
Unload Me
Exit Sub
End If


autoincr
unloctxt
Text2.SetFocus
Command1.Caption = "Save"
End Sub

Private Sub unloctxt()
Text2.Text = ""
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
MsgBox "Please Enter the User Password ", vbCritical, "Error"
Text3.SetFocus
Exit Sub
End If



If Trim(Combo1.Text) = "" Then
MsgBox "Please Enter the User Type ", vbCritical, "Error"
Combo1.SetFocus
Exit Sub
End If


End Sub

Private Sub Command2_Click()
a = MsgBox("Delete Record", vbOKCancel, "Delete Record ?")
If a = 1 Then

Set rs = New ADODB.Recordset
rs.Open "Select * from sImple where O_NAME='" & Trim(Text2.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
MsgBox "This record can't be deleted because data exists for this operator", vbOKCancel, "Data Found"
Exit Sub
End If


Set rs = New ADODB.Recordset
rs.Open "Select * from sImple where O2_NAME='" & Trim(Text2.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
MsgBox "This record can't be deleted because data exists for this operator", vbOKCancel, "Data Found"
Exit Sub
End If

Set rs = New ADODB.Recordset
rs.Open "Select * from sPECIAL where O_NAME='" & Trim(Text2.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
MsgBox "This record can't be deleted because data exists for this operator", vbOKCancel, "Data Found"
Exit Sub
End If


Set rs = New ADODB.Recordset
rs.Open "Select * from sPECIAL where O2_NAME='" & Trim(Text2.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
MsgBox "This record can't be deleted because data exists for this operator", vbOKCancel, "Data Found"
Exit Sub
End If

Set rs = New ADODB.Recordset
rs.Open "sELECT * FROM USERS WHERE USERID='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveFirst

If rs.Fields("Utype").Value = "Admin" Then
MsgBox "Administrator can not be deleted", vbInformation, "Deletion Error"
Exit Sub
End If

End If




Set rs = New ADODB.Recordset
rs.Open "delete   from USERS  where USERID='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
MsgBox "Records Deleted Successfully", vbInformation, "Record Deleted"
Text2.Text = ""
Text1.Text = ""
Text3.Text = ""
autoincr
Text2.SetFocus
End If

End Sub

Private Sub Command3_Click()
recfound = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Command1.Caption = "Save"
autoincr
Text2.SetFocus
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Picture = main.Picture

'Me.Picture = LoadPicture("logo.bmp")
Conn
autoincr
'Label7.BackColor = fhbc
'Me.BackColor = master.BackColor
unloctxt


If logintype = "Admin" Then
Combo1.Visible = True
Else
Combo1.Visible = False
End If

End Sub

Private Sub autoincr()
Set rs = New ADODB.Recordset
rs.Open "Select * from users order by userid", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveLast
Text1.Text = "U" + IncrBillNo11(rs.Fields("userid").Value)
Else
Text1.Text = "U01"
End If
End Sub

Private Sub username_list()
If List1.ListCount > 0 Then
List1.Clear
List1.Refresh
End If

Set rs = New ADODB.Recordset

If logintype = "Admin" Then

rs.Open "Select * from users order by username", co, adOpenKeyset, adLockOptimistic
Else
rs.Open "Select * from users where utype <> 'Admin' order by username", co, adOpenKeyset, adLockOptimistic

End If


If rs.RecordCount > 0 Then
rs.MoveFirst
End If
g = 0
ReDim tmpcode(rs.RecordCount)
While Not rs.EOF
List1.AddItem rs.Fields("username").Value
tmpcode(g) = rs.Fields("userid").Value
g = g + 1
rs.MoveNext
Wend
End Sub

Private Sub List1_DblClick()
Text2.Text = List1.List(List1.ListIndex)
Dataedit
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text2_Change()
str1 = Text2.Text
Call dd1
For i = 0 To List1.ListCount - 1
      If Trim(Text2.Text) = Left(List1.List(i), Len(Trim(Text2.Text))) Then
                    List1.ListIndex = i
                    Exit Sub
        End If
    Next i

End Sub

Private Sub Text2_GotFocus()
str1 = ""
num = 0
username_list
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
        For i = 0 To List1.ListCount - 1
          If Trim(Text2.Text) = Left(List1.List(i), Len(Trim(Text2.Text))) Then
                    List1.ListIndex = i
                    Exit Sub
          End If
    Next i

End Sub

Public Sub dd1()
If Len(str1) > 1 Then
For X = 1 To Len(str1) + 1
m = Mid(str1, X, X + 1)
If m = " " Then
num = 0
Else
num = num + 1
End If
Next X
End If

If Len(str1) = 0 Then
num = 0
End If
End Sub
Public Sub KeyPress1(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then
If Len(str1) > 1 Then
              If num = 1 Then
                  g = KeyAscii - 32
                  KeyAscii = g
            Else
            g = KeyAscii
            KeyAscii = g
            End If
            
End If


If Len(str1) = 0 Then
  g = KeyAscii - 32
  KeyAscii = g
  End If
End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Text2.Text) <> "" Then
     Dataedit
    ' Command1.SetFocus
    Text3.SetFocus
    End If
  Else
    Call KeyPress1(KeyAscii)
End If

End Sub

Private Sub Dataedit()
Set rs = New ADODB.Recordset
rs.Open "Select * from users where username='" & Trim(Text2.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveFirst
Text1.Text = rs.Fields("userid").Value
Text3.Text = Textdcrt(rs.Fields("password").Value)
Combo1.Text = rs.Fields("utype").Value

 Check1.Value = Mid(rs.Fields("privilages").Value & "", 1, 1)
 Check2.Value = Mid(rs.Fields("privilages").Value & "", 2, 1)
 Check3.Value = Mid(rs.Fields("privilages").Value & "", 3, 1)
 Check4.Value = Mid(rs.Fields("privilages").Value & "", 4, 1)
 Check5.Value = Mid(rs.Fields("privilages").Value & "", 5, 1)
 Check6.Value = Mid(rs.Fields("privilages").Value & "", 6, 1)
 Check7.Value = Mid(rs.Fields("privilages").Value & "", 7, 1)
 Check7.Value = Mid(rs.Fields("privilages").Value & "", 8, 1)
Command1.Caption = "Modify"
Else
Command1.Caption = "Save"
End If
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.SetFocus
End If

End Sub
