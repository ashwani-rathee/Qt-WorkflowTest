VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form smpcustomrep 
   Caption         =   "F"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   13875
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14520
      TabIndex        =   11
      ToolTipText     =   "Unload Form"
      Top             =   45
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   15255
      Begin VB.ComboBox Combo2 
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
         Left            =   9840
         TabIndex        =   15
         Text            =   "ALL"
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Export to Excel"
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
         Left            =   10800
         TabIndex        =   14
         Top             =   7200
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Show"
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
         Left            =   12960
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   6375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   11245
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   1
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Click Me"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   7200
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   6600
         TabIndex        =   1
         Text            =   "ALL"
         Top             =   240
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   240
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
         Left            =   3600
         TabIndex        =   4
         Top             =   240
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
         Left            =   12000
         TabIndex        =   17
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Select Item"
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
         Left            =   8640
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Upto Date"
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
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " From Date"
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
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Select Customer"
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
         Left            =   4920
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label acno 
         Caption         =   "acno"
         Height          =   255
         Left            =   8040
         TabIndex        =   5
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Date /Customer Wise Report(Simple)"
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
      Left            =   4410
      TabIndex        =   9
      Top             =   120
      Width           =   4245
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   15375
   End
End
Attribute VB_Name = "smpcustomrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim g As Integer
Dim str1 As String
Dim totalwt As Double



Private Sub Combo1_Click()
On Error Resume Next
acno.Caption = tmpcode(Combo1.ListIndex)
End Sub

Private Sub Combo2_Click()
On Error Resume Next
icode.Caption = tmpcode1(Combo2.ListIndex)

End Sub

Private Sub Command1_Click()
str1 = ""
totalwt = 0

If Trim(Combo1.Text) <> "" Then
        If Trim(acno.Caption) <> "" Then
                str1 = "Select * from smdtwtrep where SECOND_WT > 0 AND date_out>=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "#  and date_out<=#" & Format(DTPicker2.Value, "mm/dd/yyyy") & "# and tc_code='" & Trim(acno.Caption) & "' order by date_Out,VAL(d_serial)"
                spdatrep.Sections(2).Controls("ITEMNAME").Visible = True
                spdatrep.Sections(3).Controls("text5").Visible = False
                spdatrep.Sections(2).Controls("label5").Visible = False
        End If
   Else
        
        spdatrep.Sections(2).Controls("ITEMNAME").Visible = False
        spdatrep.Sections(3).Controls("text5").Visible = True
        spdatrep.Sections(2).Controls("label5").Visible = True
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
spdatrep.Sections(2).Controls("REPNAME").Caption = "From Date " + Format(DTPicker1.Value, "dd/mm/yyyy") & " to " & Format(DTPicker2.Value, "dd/mm/yyyy") + "  " + "(Simple)"

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

Private Sub gridheder()
Grid1.Clear
Grid1.Refresh
Grid1.Rows = 2
Grid1.FormatString = "|^          S.N.         |^    Date In   |^    Date Out   |^     Time In   |^    Time Out   |<   Customer Name |<   Customer Address   |<Destination|< DO No.|^    Date Out|<Vehicle No.|<  RLW   |<M_Code|<   Material Name   |>   Ist Weight    |>    IInd Weight    |>      Net weight  "
End Sub

Private Sub Command2_Click()
SQL = ""
gridheder
SQL = "Select * from SIMPLE where DATE_OUT>=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "#  and DATE_OUT<=#" & Format(DTPicker2.Value, "mm/dd/yyyy") & "#"
If Trim(Combo1.Text) <> "All" Then
    SQL = SQL & " and TC_CODE = '" & CStr(tmpcode(Combo1.ListIndex)) & "'"
End If

If Trim(Combo2.Text) <> "All" Then
    SQL = SQL & " and TM_CODE = '" & CStr(tmpcode1(Combo2.ListIndex)) & "'"
End If

Set rs3 = New ADODB.Recordset
rs3.Open SQL, co, adOpenKeyset, adLockOptimistic
If rs3.RecordCount > 0 Then
    rs3.MoveFirst
End If

While Not rs3.EOF
    Grid1.TextMatrix(Grid1.Row, 1) = rs3.Fields("d_serial").Value
    Grid1.TextMatrix(Grid1.Row, 2) = CDate(rs3.Fields("date_in").Value)
    Grid1.TextMatrix(Grid1.Row, 3) = CDate(rs3.Fields("date_out").Value)
    
    Grid1.TextMatrix(Grid1.Row, 4) = Trim(rs3.Fields("time_in").Value)
    Grid1.TextMatrix(Grid1.Row, 5) = Trim(rs3.Fields("time_out").Value)
    
    Grid1.TextMatrix(Grid1.Row, 6) = Trim(rs3.Fields("c_name").Value)
    Grid1.TextMatrix(Grid1.Row, 7) = Trim(rs3.Fields("c_address").Value)
    
'    Grid1.TextMatrix(Grid1.Row, 8) = Trim(rs3.Fields("Dest.").Value)
'    Grid1.TextMatrix(Grid1.Row, 9) = Trim(rs3.Fields("DO No").Value)
'    Grid1.TextMatrix(Grid1.Row, 10) = CDate(rs3.Fields("DO_date").Value)
    
    Grid1.TextMatrix(Grid1.Row, 11) = rs3.Fields("v_no").Value
'    Grid1.TextMatrix(Grid1.Row, 12) = rs3.Fields("RLW").Value
    
    Grid1.TextMatrix(Grid1.Row, 13) = rs3.Fields("TM_CODE").Value
    Grid1.TextMatrix(Grid1.Row, 14) = rs3.Fields("material").Value
    
    Grid1.TextMatrix(Grid1.Row, 15) = Val(rs3.Fields("first_wt").Value & "")
    Grid1.TextMatrix(Grid1.Row, 16) = Val(rs3.Fields("second_wt").Value & "")
    
    Grid1.TextMatrix(Grid1.Row, 17) = Abs(Val(rs3.Fields("second_wt").Value & "") - Val(rs3.Fields("first_wt").Value & ""))
   
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Row = Grid1.Row + 1
    rs3.MoveNext
    
    
Wend

End Sub

Private Sub Command3_Click()
a = MsgBox("Want to Export", vbOKCancel, "Print")
If a = 1 Then
    Set appxl = CreateObject("Excel.Application")
    Set Book = appxl.workbooks
    Set Wsheet = Book.Add.Worksheets(1)
    appxl.Visible = True
    
    Wsheet.Cells(1, 1) = main.Label1.Caption
    Wsheet.Cells(2, 1) = main.Label2.Caption
    Wsheet.Cells(3, 1) = main.Label3.Caption
    Wsheet.Cells(4, 1) = "From Date " + Format(DTPicker1.Value, "dd/mm/yyyy") & " to " & Format(DTPicker2.Value, "dd/mm/yyyy") + "  " + "(Simple)"
    
    For i = 0 To Grid1.Rows - 1
        For j = 1 To Grid1.Cols - 1
            Grid1.Row = i
            Grid1.Col = j
            Wsheet.Cells(i + 6, j) = Grid1.Text
        Next j
    Next i
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
DTPicker2.Value = Date
Me.Picture = main.Picture

Conn
achlp
If Combo1.ListCount > 0 Then
    Combo1.ListIndex = 0
    acno.Caption = tmpcode(Combo1.ListIndex)
End If

itmhlp
If Combo2.ListCount > 0 Then
    Combo2.ListIndex = 0
    icode.Caption = tmpcode1(Combo2.ListIndex)
End If
End Sub

Private Sub achlp()
If Combo1.ListCount > 0 Then
    Combo1.Clear
    Combo1.Refresh
End If

Set rs2 = New ADODB.Recordset
rs2.Open "Select * from cust1 order by c_name", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
    rs2.MoveFirst
End If
g = 0
ReDim tmpcode(rs2.RecordCount)
Combo1.AddItem "All"
tmpcode(g) = "All"
While Not rs2.EOF
    g = g + 1
    tmpcode(g) = rs2.Fields("c_code").Value
    Combo1.AddItem rs2.Fields("c_name").Value

    rs2.MoveNext
Wend

End Sub

Private Sub itmhlp()
If Combo2.ListCount > 0 Then
    Combo2.Clear
    Combo2.Refresh
End If

Set rs = New ADODB.Recordset
rs.Open "Select * from mater1 order by m_name", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
    rs.MoveFirst
End If
g = 0
ReDim tmpcode1(rs.RecordCount)
Combo2.AddItem "All"
tmpcode1(g) = "All"
While Not rs.EOF
    g = g + 1
    tmpcode1(g) = rs.Fields("m_code").Value
    Combo2.AddItem rs.Fields("m_name").Value
    rs.MoveNext
Wend

End Sub
