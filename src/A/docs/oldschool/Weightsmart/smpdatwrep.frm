VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form smpdatwrep 
   Caption         =   "Form1"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   17400
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
      Left            =   14760
      TabIndex        =   8
      ToolTipText     =   "Unload Form"
      Top             =   240
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   15255
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   6375
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   11245
         _Version        =   393216
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
         Left            =   13320
         TabIndex        =   10
         Top             =   7080
         Width           =   1695
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
         Left            =   5640
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Click Me"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   7200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   90374145
         CurrentDate     =   39983
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   90374145
         CurrentDate     =   39983
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
         Left            =   2640
         TabIndex        =   2
         Top             =   240
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
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Date Wise Report(Simple)"
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
      Left            =   5520
      TabIndex        =   6
      Top             =   240
      Width           =   2985
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   15255
   End
End
Attribute VB_Name = "smpdatwrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str1 As String
Dim totalwt As Double

Private Sub Command1_Click()
str1 = ""
totalwt = 0
str1 = "Select * from smdtwtrep where SECOND_WT > 0 AND date_out>=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "#  and date_out<=#" & Format(DTPicker2.Value, "mm/dd/yyyy") & "# order by date_Out,VAL(d_serial)"
Set rs1 = New ADODB.Recordset
rs1.Open str1, co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
End If
spdatrep.Sections(2).Controls("cmpname").Caption = main.Label1.Caption
spdatrep.Sections(2).Controls("compadr1").Caption = main.Label2.Caption
spdatrep.Sections(2).Controls("compadr2").Caption = main.Label3.Caption
spdatrep.Sections(2).Controls("REPNAME").Caption = "From Date " + Format(DTPicker1.Value, "dd/mm/yyyy") & " to " & Format(DTPicker2.Value, "dd/mm/yyyy") + " ( Simple)"
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
'If Trim(Combo1.Text) <> "All" Then
'    sql = sql & " and TC_CODE = '" & CStr(tmpcode(Combo1.ListIndex)) & "'"
'End If

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
Call Conn
DTPicker1.Value = Date
DTPicker2.Value = Date
Me.Picture = main.Picture

End Sub
