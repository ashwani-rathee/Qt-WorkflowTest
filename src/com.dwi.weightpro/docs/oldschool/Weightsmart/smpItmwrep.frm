VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form smpItmwrep 
   Caption         =   "a"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   10785
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
      Left            =   7320
      TabIndex        =   11
      ToolTipText     =   "Unload Form"
      Top             =   1080
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   2400
      TabIndex        =   0
      Top             =   1560
      Width           =   5415
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
         Left            =   2280
         TabIndex        =   9
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Click Me"
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   2520
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   90439681
         CurrentDate     =   39983
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   90439681
         CurrentDate     =   39983
      End
      Begin VB.Label icode 
         Caption         =   "Icode"
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Item"
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
         Left            =   960
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "From Date"
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
         Left            =   960
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Upto Date"
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
         Left            =   960
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Date / Item Wise Report(Simple)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Top             =   1080
      Width           =   3645
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   2400
      TabIndex        =   7
      Top             =   960
      Width           =   5415
   End
End
Attribute VB_Name = "smpItmwrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim g As Integer
Dim str1 As String
Dim totalwt As Double

Private Sub Combo1_Click()
On Error Resume Next
icode.Caption = tmpcode(Combo1.ListIndex)
End Sub

Private Sub Command1_Click()
str1 = ""
totalwt = 0

If Trim(Combo1.Text) <> "" Then
        If Trim(icode.Caption) <> "" Then
        str1 = "Select * from smdtwtrep where SECOND_WT > 0 AND date_out>=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "#  and date_out<=#" & Format(DTPicker2.Value, "mm/dd/yyyy") & "# and tm_code='" & Trim(icode.Caption) & "' order by date_Out,VAL(d_serial)"
        
                spdatrep.Sections(2).Controls("ITEMNAME").Visible = True
                spdatrep.Sections(3).Controls("text7").Visible = False
                spdatrep.Sections(2).Controls("label7").Visible = False
        End If
   Else
        
        spdatrep.Sections(2).Controls("ITEMNAME").Visible = False
        spdatrep.Sections(3).Controls("text7").Visible = True
        spdatrep.Sections(2).Controls("label7").Visible = True
        str1 = "Select * from smdtwtrep where SECOND_WT > 0 AND date_out>=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "#  and date_out<=#" & Format(DTPicker2.Value, "mm/dd/yyyy") & "# order by date_Out,VAL(d_serial)"
End If

Set rs1 = New ADODB.Recordset
rs1.Open str1, co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
End If
spdatrep.Sections(2).Controls("cmpname").Caption = main.Label1.Caption
spdatrep.Sections(2).Controls("compadr1").Caption = main.Label2.Caption
spdatrep.Sections(2).Controls("compadr2").Caption = main.Label3.Caption
spdatrep.Sections(2).Controls("REPNAME").Caption = "From Date " + Format(DTPicker1.Value, "dd/mm/yyyy") & " to " & Format(DTPicker2.Value, "dd/mm/yyyy") + " ( Simple )"

spdatrep.Sections(2).Controls("ITEMNAME").Caption = Combo1.Text

Set spdatrep.DataSource = rs1
While Not rs1.EOF
spdatrep.Sections(3).Controls("text1").DataField = "d_serial"
spdatrep.Sections(3).Controls("text2").DataField = "time_in"
spdatrep.Sections(3).Controls("text3").DataField = "time_out"
spdatrep.Sections(3).Controls("text4").DataField = "date_out"
spdatrep.Sections(3).Controls("text5").DataField = "c_name"
spdatrep.Sections(3).Controls("text6").DataField = "v_no"
spdatrep.Sections(3).Controls("text7").DataField = "m_name"
spdatrep.Sections(3).Controls("text8").DataField = "first_wt"
spdatrep.Sections(3).Controls("text9").DataField = "second_wt"
spdatrep.Sections(3).Controls("text10").DataField = "netwt"
totalwt = totalwt + rs1.Fields("netwt").Value

rs1.MoveNext
Wend
spdatrep.Sections(5).Controls("twt").Caption = totalwt

spdatrep.LeftMargin = 200
spdatrep.RightMargin = 700
spdatrep.Top = 2000
spdatrep.Width = 15360
spdatrep.Height = 9000
spdatrep.Show vbModal
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = main.Picture

DTPicker1.Value = Date
DTPicker2.Value = Date
Conn
itmhlp
If Combo1.ListCount > 0 Then
Combo1.ListIndex = 0
icode.Caption = tmpcode(Combo1.ListIndex)
End If
End Sub

Private Sub itmhlp()
If Combo1.ListCount > 0 Then
    Combo1.Clear
    Combo1.Refresh
End If

Set rs = New ADODB.Recordset
rs.Open "Select * from mater1 order by m_name", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveFirst
End If
g = 0
ReDim tmpcode(rs.RecordCount)
While Not rs.EOF
Combo1.AddItem rs.Fields("m_name").Value
tmpcode(g) = rs.Fields("m_code").Value
g = g + 1
rs.MoveNext
Wend

End Sub

