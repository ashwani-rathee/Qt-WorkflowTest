VERSION 5.00
Begin VB.Form selmac 
   Caption         =   "Change Machine"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12975
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   12975
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "selmac.frx":0000
      Left            =   7200
      List            =   "selmac.frx":000A
      TabIndex        =   13
      Text            =   "0"
      Top             =   5040
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000A0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3480
      TabIndex        =   0
      Top             =   1320
      Width           =   6495
      Begin VB.CommandButton Command2 
         BackColor       =   &H00404040&
         Cancel          =   -1  'True
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   1
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Machine Selection Form"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   4815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4515
      Left            =   3480
      TabIndex        =   3
      Top             =   2040
      Width           =   6495
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "selmac.frx":0014
         Left            =   3720
         List            =   "selmac.frx":001E
         TabIndex        =   11
         Text            =   "1"
         Top             =   2400
         Width           =   735
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "selmac.frx":0028
         Left            =   3720
         List            =   "selmac.frx":0032
         TabIndex        =   9
         Text            =   "1"
         Top             =   1740
         Width           =   735
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "selmac.frx":003C
         Left            =   3720
         List            =   "selmac.frx":0046
         TabIndex        =   7
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "selmac.frx":0050
         Left            =   3720
         List            =   "selmac.frx":006C
         TabIndex        =   5
         Text            =   "1"
         Top             =   420
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   4
         Top             =   3660
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "GPS Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   720
         TabIndex        =   14
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RFID Tags"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   720
         TabIndex        =   12
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Boom Barrier"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   10
         Top             =   1740
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Use Camera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   8
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Machine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   6
         Top             =   420
         Width           =   2895
      End
   End
End
Attribute VB_Name = "selmac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim machcode As String



Private Sub Command1_Click()
Call Conn
Set rs2 = New ADODB.Recordset
rs2.Open "Select * from oper", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
    rs2.MoveFirst
    rs2.Fields("oname") = Trim(Combo1.Text)
    rs2.Fields("uname") = Trim(Combo2.Text)
    rs2.Fields("pword") = Trim(Combo3.Text) + Trim(Combo4.Text) + Trim(Combo5.Text)
    rs2.Update
Else
    rs2.AddNew
    rs2.Fields("oname") = Trim(Combo1.Text)
    rs2.Fields("uname") = Trim(Combo2.Text)
    rs2.Fields("pword") = Trim(Combo3.Text) + Trim(Combo4.Text) + Trim(Combo5.Text)
    rs2.Update
End If
rs2.Close
MsgBox "Machines updated"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = main.Picture
Call Conn
Set rs2 = New ADODB.Recordset
rs2.Open "Select * from oper", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
rs2.MoveFirst
If IsNumeric(rs2.Fields("oname")) Then
    If Val(rs2.Fields("oname")) > 0 And Val(rs2.Fields("oname")) < 9 Then
        Combo1.Text = rs2.Fields("oname")
    End If
Else
    Combo1.Text = "1"
End If
If IsNumeric(rs2.Fields("uname")) Then
    If Trim(rs2.Fields("uname")) = "0" Or Trim(rs2.Fields("uname")) = "1" Then
        Combo2.Text = rs2.Fields("uname")
    End If
Else
    Combo2.Text = "0"
End If

If IsNumeric(rs2.Fields("pword")) Then
    If Mid(rs2.Fields("pword"), 1, 1) = "0" Or Mid(rs2.Fields("pword"), 1, 1) = "1" Then
        Combo3.Text = Left(rs2.Fields("pword"), 1)
    End If
Else
    Combo3.Text = "1"
End If

If IsNumeric(rs2.Fields("pword")) Then
    If Mid(rs2.Fields("pword"), 2, 1) = "0" Or Mid(rs2.Fields("pword"), 2, 1) = "1" Then
        Combo4.Text = Mid(rs2.Fields("pword"), 2, 1)
    End If
Else
    Combo4.Text = "1"
End If

If IsNumeric(rs2.Fields("pword")) Then
    If Mid(rs2.Fields("pword"), 3, 1) = "0" Or Mid(rs2.Fields("pword"), 3, 1) = "1" Then
        Combo5.Text = Mid(rs2.Fields("pword"), 3, 1)
    End If
Else
    Combo5.Text = "0"
End If


End If
rs2.Close
End Sub

