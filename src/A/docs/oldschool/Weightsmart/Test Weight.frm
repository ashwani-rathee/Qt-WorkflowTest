VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Test Weight"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   13065
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2040
      TabIndex        =   28
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   90439682
      CurrentDate     =   41631
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   90439681
      CurrentDate     =   41631
   End
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      Height          =   255
      Left            =   12120
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3375
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   11415
      Begin VB.TextBox Text19 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Left            =   7680
         TabIndex        =   26
         Text            =   "    WEIGHT"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text18 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Left            =   6240
         TabIndex        =   25
         Text            =   "      TIME"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Left            =   4680
         TabIndex        =   24
         Text            =   "       DATE"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Left            =   2160
         TabIndex        =   23
         Text            =   "    VEHICLE NUMBER"
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Read WT3"
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
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Read WT 2"
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
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Read WT 1"
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
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text12 
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text11 
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text10 
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text9 
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
         Left            =   2160
         TabIndex        =   13
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox Text8 
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text7 
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text6 
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text5 
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
         Left            =   2160
         TabIndex        =   9
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox Text4 
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text3 
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   1455
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   1575
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
         Left            =   2160
         TabIndex        =   5
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Test Report"
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Left            =   600
         TabIndex        =   19
         Text            =   " POSITION 3"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Left            =   600
         TabIndex        =   18
         Text            =   " POSITION 2"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Text            =   " POSITION 1"
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Test Weight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   11415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dosprint()
Set fsys = CreateObject("scripting.filesystemobject")
Set wstream = fsys.OpenTextFile(App.Path & "\" & "r.bat", ForWriting, True)
wstream.WriteLine "type " + App.Path + "\" & "rep.txt  >  prn"
wstream.Close
Shell App.Path + "\" + "r.bat"
End Sub


Private Sub printdata()
Set wstream = fsys.OpenTextFile(App.Path + "\" & "rep.txt", ForWriting, True)
wstream.WriteLine Chr(18) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) + PADC(Trim(main.Label1.Caption), 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & PADC(Trim(main.Label2.Caption), 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & PADC(Trim(main.Label3.Caption), 25) & Chr(27) & "F"
wstream.WriteBlankLines 1
wstream.WriteLine Chr(27) & "E" & Chr(14) & PADC("Test Weight Slip", 28) & Chr(27) & "F"
wstream.WriteLine Chr(18) & String(53, "-")
wstream.WriteLine Chr(27) & "E" & Chr(14) & "Date : " & Format(Date, "dd/mm/yyyy") & Chr(27) & "F"
wstream.WriteLine Chr(27) & "E" & Chr(14) & "Time : " & Format(Time, "hh:mm") & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("Test Weight", 13) & padl(":", 2) & padl(Trim(main.comwt.Caption) & " Kg", 15) & Chr(27) & "F"
wstream.WriteBlankLines 1
wstream.WriteBlankLines 1
wstream.WriteLine Chr(18) & String(53, "-")
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR(COMPAUTH, 20) & Chr(27) & "F" & padl(":", 2) & padl("Weighing Operator", 20)
wstream.WriteBlankLines 2
wstream.Close
Set wstream = Nothing
End Sub


Private Sub Command1_Click()
If Trim(Text1.Text <> "") And Trim(Text2.Text <> "") And Trim(Text3.Text <> "") And Trim(Text4.Text <> "") And Trim(Text5.Text <> "") And Trim(Text6.Text <> "") And Trim(Text7.Text <> "") And Trim(Text8.Text <> "") And Trim(Text9.Text <> "") And Trim(Text10.Text <> "") And Trim(Text11.Text <> "") And Trim(Text12.Text <> "") Then
    'printdata
    'dosprint
    printrep
Else
    MsgBox "Please fill data properly"
End If
End Sub

Sub printrep()
testwt.Sections(1).Controls("Label22").Caption = PADC(Trim(main.Label2.Caption), 27)
testwt.Sections(1).Controls("Label23").Caption = PADC(Trim(main.Label3.Caption), 27)

testwt.Sections(1).Controls("Label31").Caption = Text2.Text
testwt.Sections(1).Controls("Label32").Caption = Text3.Text
testwt.Sections(1).Controls("Label33").Caption = Text1.Text
testwt.Sections(1).Controls("Label34").Caption = Text4.Text + " Kg"

testwt.Sections(1).Controls("Label41").Caption = Text6.Text
testwt.Sections(1).Controls("Label42").Caption = Text7.Text
testwt.Sections(1).Controls("Label43").Caption = Text5.Text
testwt.Sections(1).Controls("Label44").Caption = Text8.Text + " Kg"

testwt.Sections(1).Controls("Label51").Caption = Text10.Text
testwt.Sections(1).Controls("Label52").Caption = Text11.Text
testwt.Sections(1).Controls("Label53").Caption = Text9.Text
testwt.Sections(1).Controls("Label54").Caption = Text12.Text + " Kg"

Set rs1 = New ADODB.Recordset
rs1.Open "SELECT * from party", co, adOpenKeyset, adLockOptimistic
rs1.MoveFirst
Set testwt.DataSource = rs1
While Not rs1.EOF
testwt.Sections(1).Controls("Text6").DataField = "pname"
rs1.MoveNext
Wend
testwt.Show vbModal
End Sub



Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
DTPicker1.Value = Date
DTPicker2.Value = Time

If Trim(Text1.Text) <> "" Then
    Text2.Text = Format(CStr(DTPicker1.Day), "0#") + "/" + Format(CStr(DTPicker1.Month), "0#") + "/" + Format(CStr(DTPicker1.Year), "000#")
    Text3.Text = DTPicker2.Value
    Text4.Text = main.comwt.Caption
    Text1.Enabled = False
    Command3.Enabled = False
Else
    MsgBox "Please fill vehicle number"
    Text1.SetFocus
End If
End Sub

Private Sub Command4_Click()
DTPicker1.Value = Date
DTPicker2.Value = Time
If Trim(Text5.Text) <> "" Then
    Text6.Text = Format(CStr(DTPicker1.Day), "0#") + "/" + Format(CStr(DTPicker1.Month), "0#") + "/" + Format(CStr(DTPicker1.Year), "000#")
    Text7.Text = DTPicker2.Value
    Text8.Text = main.comwt.Caption
    Text5.Enabled = False
    Command4.Enabled = False
Else
    MsgBox "Please fill vehicle number"
    Text5.SetFocus
End If

End Sub

Private Sub Command5_Click()
DTPicker1.Value = Date
DTPicker2.Value = Time
If Trim(Text9.Text) <> "" Then
    Text10.Text = Format(CStr(DTPicker1.Day), "0#") + "/" + Format(CStr(DTPicker1.Month), "0#") + "/" + Format(CStr(DTPicker1.Year), "000#")
    Text11.Text = DTPicker2.Value
    Text12.Text = main.comwt.Caption
    Text9.Enabled = False
    Command5.Enabled = False
Else
    MsgBox "Please fill vehicle number"
    Text9.SetFocus
End If
End Sub

Private Sub Form_Load()
Me.Picture = main.Picture

End Sub

Private Sub Text1_Change()
Text5.Text = Text1.Text
Text9.Text = Text1.Text
End Sub
