VERSION 5.00
Begin VB.Form sfeature 
   Caption         =   "Form1"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   12660
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
      Left            =   9120
      TabIndex        =   16
      ToolTipText     =   "Unload Form"
      Top             =   1680
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Width           =   7935
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "sfeature.frx":0000
         Left            =   2280
         List            =   "sfeature.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "sfeature.frx":0014
         Left            =   2280
         List            =   "sfeature.frx":001E
         TabIndex        =   15
         Text            =   "KG"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   5010
         TabIndex        =   11
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Special"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Simple"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Special"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Simple"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Net Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6090
         TabIndex        =   20
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   "Fst Wt."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label Label5 
         Caption         =   "Delete All Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Measured Unit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Special Min Weights Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Check Empty Platform"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Want Feature"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Software Features"
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
      Left            =   4080
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   7935
   End
End
Attribute VB_Name = "sfeature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recfound As Boolean

Private Sub Check1_Click()
If Check1.Value = 0 Then
Check1.Value = 0
Text1.Text = Val(0)
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 0 Then
Check4.Value = 0
Text2.Text = Val(0)
Else

End If

End Sub

Private Sub Check3_Click()
If Check3.Value = 0 Then
Text1.Text = Val(0)
End If

End Sub

Private Sub Check4_Click()
If Check4.Value = 0 Then
Text2.Text = Val(0)
Else
Check2.Value = 1
End If
End Sub

Private Sub Command1_Click()
a = MsgBox("Want to Save Changes", vbOKCancel, "Save ")
If a = 1 Then
dataset
If recfound = False Then
rs1.AddNew
End If

rs1.Fields("permis").Value = Trim(Check1.Value) + Trim(Check2.Value) + Trim(Check3.Value) + Trim(Check4.Value)
rs1.Fields("smpwtr").Value = Val(Text1.Text)
rs1.Fields("spcwtr").Value = Val(Text2.Text)
If Trim(Combo1.Text) = "" Then
rs1.Fields("munit").Value = "KG"
Else
rs1.Fields("munit").Value = Combo1.Text
End If

rs1.Update
End If


If Trim(Combo2.Text) = "Y" Then
Delete_Records
End If



End Sub

Private Sub Delete_Records()

Set rs = New ADODB.Recordset
rs.Open "delete from cust", co, adOpenKeyset, adLockOptimistic

Set rs = New ADODB.Recordset
rs.Open "delete from cust1", co, adOpenKeyset, adLockOptimistic

Set rs = New ADODB.Recordset
rs.Open "delete from do_master", co, adOpenKeyset, adLockOptimistic

Set rs = New ADODB.Recordset
rs.Open "delete from mater1", co, adOpenKeyset, adLockOptimistic

Set rs = New ADODB.Recordset
rs.Open "delete from simple", co, adOpenKeyset, adLockOptimistic

Set rs = New ADODB.Recordset
rs.Open "delete from special", co, adOpenKeyset, adLockOptimistic

Set rs = New ADODB.Recordset
rs.Open "delete from users where utype<>'Admin'", co, adOpenKeyset, adLockOptimistic
MsgBox "All records are deleted Successfully", vbInformation, "Delete Data Successfully"

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Conn
Me.Picture = main.Picture
datashow
End Sub


Private Sub datashow()
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from featur", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Text1.Text = rs1.Fields("smpwtr").Value
Text2.Text = rs1.Fields("spcwtr").Value
Combo1.Text = rs1.Fields("munit").Value
Check1.Value = Mid(rs1.Fields("permis").Value & "", 1, 1)
Check2.Value = Mid(rs1.Fields("permis").Value & "", 2, 1)
Check3.Value = Mid(rs1.Fields("permis").Value & "", 3, 1)
Check4.Value = Mid(rs1.Fields("permis").Value & "", 4, 1)
End If
End Sub



Private Sub dataset()
recfound = False
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from featur", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
recfound = True
rs1.MoveFirst
End If

End Sub
