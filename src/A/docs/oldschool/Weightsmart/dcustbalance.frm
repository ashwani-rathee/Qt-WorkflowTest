VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ddobalance 
   Caption         =   "Customer Wise Report"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
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
      Left            =   14640
      TabIndex        =   5
      ToolTipText     =   "Unload Form"
      Top             =   120
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   15255
      Begin VB.CheckBox Check1 
         Caption         =   "Summary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13440
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sorted DO"
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
         Left            =   5280
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Selected DO"
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
         Left            =   6840
         TabIndex        =   11
         Top             =   240
         Width           =   1815
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
         Left            =   10680
         MaxLength       =   11
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text2 
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
         Left            =   8760
         TabIndex        =   9
         Text            =   "Enter DO Number"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   10800
         TabIndex        =   8
         Top             =   6840
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Export To Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   12960
         TabIndex        =   7
         Top             =   6840
         Width           =   2055
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
         Left            =   12240
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   390
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   688
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
         Height          =   390
         Left            =   3600
         TabIndex        =   14
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   688
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
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   6255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   11033
         _Version        =   393216
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DO Balance Report (Special)"
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
      Left            =   5325
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "ddobalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim g As Integer
Dim str1 As String
Dim totalwt As Double
Dim liftedtot As Double


Private Sub gridheder()
Grid1.Clear
Grid1.Refresh
Grid1.Rows = 2
Grid1.FormatString = "|^          S.N.         |^    Date In   |^    Date Out   |^     Time In   |^    Time Out   |<   Customer Name |<   Destination        |<   RLW     |< DO No.|^  DO Date   |<Vehicle No.|<  Material   |<M_Code|<     coll. Code    |>   Ist Weight    |>    IInd Weight    |>      Net weight  |>    Order Qty   |>   Challan_no   "
End Sub

Private Sub Command2_Click()
Dim sql1 As String
SQL = ""
gridheder
SQL = "Select * from spdtwtrep where DATE_OUT>=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "#  and DATE_OUT<=#" & Format(DTPicker2.Value, "mm/dd/yyyy") & "#"

If Trim(Text1.Text) <> "" Then
    SQL = SQL & " and special.do_no = '" & Trim(Text1.Text) & "' order by sl_no desc"
Else
    SQL = SQL & " order by special.do_no, sl_no desc"
End If

Set rs3 = New ADODB.Recordset
rs3.Open SQL, co, adOpenKeyset, adLockOptimistic
If rs3.RecordCount > 0 Then
    rs3.MoveFirst
End If

While Not rs3.EOF
    Grid1.TextMatrix(Grid1.Row, 1) = rs3.Fields("sl_no").Value
    Grid1.TextMatrix(Grid1.Row, 2) = CDate(rs3.Fields("date_in").Value)
    Grid1.TextMatrix(Grid1.Row, 3) = CDate(rs3.Fields("date_out").Value)
    
    Grid1.TextMatrix(Grid1.Row, 4) = Trim(rs3.Fields("time_in").Value)
    Grid1.TextMatrix(Grid1.Row, 5) = Trim(rs3.Fields("time_out").Value)
    
    Grid1.TextMatrix(Grid1.Row, 6) = Trim(rs3.Fields("purchaser").Value)
   
    Grid1.TextMatrix(Grid1.Row, 7) = Trim(rs3.Fields("Dest").Value)
    Grid1.TextMatrix(Grid1.Row, 8) = Trim(rs3.Fields("RLW").Value)
    Grid1.TextMatrix(Grid1.Row, 9) = Trim(rs3.Fields("DO_No").Value)
    Grid1.TextMatrix(Grid1.Row, 10) = Mid(rs3.Fields("do_start_date").Value, 7, 2) + "/" + Mid(rs3.Fields("do_start_date").Value, 5, 2) + "/" + Mid(rs3.Fields("do_start_date").Value, 1, 4)
    
    Grid1.TextMatrix(Grid1.Row, 11) = rs3.Fields("v_no").Value
    Grid1.TextMatrix(Grid1.Row, 12) = "na"
    
    Grid1.TextMatrix(Grid1.Row, 13) = rs3.Fields("TM_CODE").Value
    Grid1.TextMatrix(Grid1.Row, 14) = rs3.Fields("coll_code").Value
    
    Grid1.TextMatrix(Grid1.Row, 15) = Val(rs3.Fields("first_wt").Value & "")
    Grid1.TextMatrix(Grid1.Row, 16) = Val(rs3.Fields("second_wt").Value & "")
    
    Grid1.TextMatrix(Grid1.Row, 17) = Abs(Val(rs3.Fields("second_wt").Value & "") - Val(rs3.Fields("first_wt").Value & ""))
    Grid1.TextMatrix(Grid1.Row, 18) = rs3.Fields("order_qty").Value
    Grid1.TextMatrix(Grid1.Row, 19) = rs3.Fields("challan_no").Value
   
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
    Wsheet.Cells(4, 1) = "From Date " + Format(DTPicker1.Value, "dd/mm/yyyy") & " to " & Format(DTPicker2.Value, "dd/mm/yyyy") + "  " + "(Special)"
    
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

Private Sub dosprint()
Set fsys = CreateObject("scripting.filesystemobject")
Set wstream = fsys.OpenTextFile(App.Path & "\" & "r.bat", ForWriting, True)
wstream.WriteLine "type " + App.Path + "\" & "rep.txt  >  prn"
wstream.Close
Shell App.Path + "\" + "r.bat"
End Sub

Private Sub printdata()
Dim repstr1 As String
Dim repstr2 As String
Dim matercod As String
Dim materbr As Boolean
Dim tdono As String
Dim ftot, stot, ntot, matertot, orderqty As Long

Set wstream = fsys.OpenTextFile(App.Path + "\" & "rep.txt", ForWriting, True)
wstream.WriteLine Chr(18) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) + padl(Trim(main.Label1.Caption), 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & padl(Trim(main.Label2.Caption), 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & padl(Trim(main.Label3.Caption), 25) & Chr(27) & "F"
wstream.WriteBlankLines 1
If Option1.Value = True Then
wstream.WriteLine Chr(27) & "E" & Chr(14) & PADC("DO Wise sorted Report", 28) & Chr(27) & "F"
Else
wstream.WriteLine Chr(27) & "E" & Chr(14) & PADC("Selected DO Report", 28) & Chr(27) & "F"
End If
wstream.WriteBlankLines 1
wstream.WriteLine Chr(27) & "E" & Chr(15) & "From Date: " & CStr(DTPicker1.Value) & "      To Date: " & CStr(DTPicker2.Value) & Chr(27) & "F"
wstream.WriteLine Chr(27) & "E" & Chr(15) & "Customer: " & Text1.Text & "F"
wstream.WriteLine Chr(18) & String(80, "-")

Grid1.Row = 1
Grid1.Col = 9
matercod = Trim(Grid1.Text)
Grid1.Col = 18
orderqty = Grid1.Text

ftot = 0
stot = 0
ntot = 0

repstr2 = ""
repstr2 = repstr2 + PADC("Sl. No.", 12) + " | "
repstr2 = repstr2 + PADC("Dt.In/Out", 10) + " | "
repstr2 = repstr2 + PADC("T.In/Out", 8) + " | "
repstr2 = repstr2 + PADC("Customer", 25) + " | "
repstr2 = repstr2 + PADC("DO No./Dt.", 11) + " | "
repstr2 = repstr2 + PADC("VNo/ RLW", 15) + " | "
repstr2 = repstr2 + PADC("Ch.No./ M.Code", 15) + " | "
repstr2 = repstr2 + PADC("1st/2ndWt", 9) + " | "
repstr2 = repstr2 + PADC("Net Wt.", 9)

If Check1.Value <> 1 Then
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(15) & repstr2 & Chr(27) & "F"
wstream.WriteLine Chr(18) & String(80, "-")
End If

materbr = False
matertot = 0
While Grid1.Row < Grid1.Rows - 1
repstr1 = ""
repstr2 = ""
Grid1.Col = 9
If Grid1.Text <> matercod Then
    wstream.WriteLine Chr(18) & String(80, "-")
    wstream.WriteLine Chr(18) & Chr(15) & Chr(14) & "DO Number: " & matercod
    Set rs3 = New ADODB.Recordset
    rs3.Open "Select sum(Abs(second_wt-first_wt)) from special where do_no='" & Trim(matercod) & "' and second_wt>0", co, adOpenKeyset, adLockOptimistic
    If rs3.RecordCount > 0 Then
        If IsNull(rs3.Fields(0)) Then
            liftedtot = 0
        Else
            liftedtot = CDbl(rs3.Fields(0))
        End If
    Else
        liftedtot = 0
    End If
    rs3.Close
    wstream.WriteLine Chr(18) & Chr(15) & Chr(14) & "Ordered:" & CStr(orderqty) & " Lifted:" & CStr(liftedtot) & " Balance:" & CStr(orderqty - liftedtot)
    wstream.WriteLine Chr(18) & String(80, "-")
    matertot = 0
    materbr = False
End If

Grid1.Col = 1
repstr1 = repstr1 + PADC(Trim(Grid1.Text), 12) + " | "
Grid1.Col = 2
repstr1 = repstr1 + PADC(Trim(Grid1.Text), 10) + " | "
Grid1.Col = 4
repstr1 = repstr1 + PADC(Trim(Grid1.Text), 8) + " | "
Grid1.Col = 6
repstr1 = repstr1 + PADC(Trim(Grid1.Text), 25) + " | "
Grid1.Col = 9
repstr1 = repstr1 + PADC(Trim(Grid1.Text), 11) + " | "
tdono = Trim(Grid1.Text)
Grid1.Col = 11
repstr1 = repstr1 + PADC(Trim(Grid1.Text), 15) + " | "
Grid1.Col = 19
repstr1 = repstr1 + PADC(Trim(Grid1.Text), 15) + " | "
Grid1.Col = 15
repstr1 = repstr1 + PADC(Trim(Grid1.Text), 9) + " | "
If IsNumeric(Grid1.Text) Then
ftot = ftot + Val(Grid1.Text)
End If
Grid1.Col = 17
repstr1 = repstr1 + PADC(Trim(Grid1.Text), 9)
If IsNumeric(Grid1.Text) Then
ntot = ntot + Val(Grid1.Text)
matertot = matertot + Val(Grid1.Text)
End If

repstr2 = repstr2 + PADC("", 12) + " | "
Grid1.Col = 3
repstr2 = repstr2 + PADC(Trim(Grid1.Text), 10) + " | "
Grid1.Col = 5
repstr2 = repstr2 + PADC(Trim(Grid1.Text), 8) + " | "
repstr2 = repstr2 + PADC("", 25) + " | "
Grid1.Col = 10
repstr2 = repstr2 + PADC(Trim(Grid1.Text), 11) + " | "
Grid1.Col = 8
repstr2 = repstr2 + PADC(Trim(Grid1.Text), 15) + " | "
Grid1.Col = 13
repstr2 = repstr2 + PADC(Trim(Grid1.Text), 15) + " | "
Grid1.Col = 9
matercod = Trim(Grid1.Text)
Grid1.Col = 18
orderqty = Grid1.Text

Grid1.Col = 16
repstr2 = repstr2 + PADC(Trim(Grid1.Text), 9) + " | "
If IsNumeric(Grid1.Text) Then
stot = stot + Val(Grid1.Text)
End If

repstr2 = repstr2 + PADC("", 9)
If Check1.Value <> 1 Then
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(15) & repstr1 & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(15) & repstr2 & Chr(27) & "F"
End If
Grid1.Row = Grid1.Row + 1
Wend
If Grid1.Text <> matercod Then
    wstream.WriteLine Chr(18) & String(80, "-")
    wstream.WriteLine Chr(18) & Chr(15) & Chr(14) & "DO Number: " & matercod
        Set rs3 = New ADODB.Recordset
        rs3.Open "Select sum(Abs(second_wt-first_wt)) from special where do_no='" & Trim(matercod) & "' and second_wt>0", co, adOpenKeyset, adLockOptimistic
        If rs3.RecordCount > 0 Then
        If IsNull(rs3.Fields(0)) Then
            liftedtot = 0
        Else
            liftedtot = CDbl(rs3.Fields(0))
        End If
Else
    liftedtot = 0
End If
rs3.Close

    wstream.WriteLine Chr(18) & Chr(15) & Chr(14) & "Ordered:" & CStr(orderqty) & " Lifted:" & CStr(liftedtot) & " Balance:" & CStr(orderqty - liftedtot)
    wstream.WriteLine Chr(18) & String(80, "-")
    matertot = 0
    materbr = False
End If

wstream.WriteBlankLines 2
wstream.Close
Set wstream = Nothing
End Sub

Private Sub Command5_Click()
b = MsgBox("Want to Print Report?", vbOKCancel, "Print ?")
If b = 1 Then
    If camcode = 0 Then
        'printdata
        'dosprint
        printrep
    Else
        printrep
    End If
End If
End Sub


Private Sub printrep()
Dim twt As Double
SQL = ""
spdatrep.Sections(2).Controls("Cmpname").Caption = PADC(Trim(main.Label1.Caption), 27)
spdatrep.Sections(2).Controls("compadr1").Caption = PADC(Trim(main.Label2.Caption), 27)
spdatrep.Sections(2).Controls("compadr2").Caption = PADC(Trim(main.Label3.Caption), 27)
If Option1.Value = True Then
    spdatrep.Sections(2).Controls("repname").Caption = "DO Wise Report (Sorted)"
Else
    spdatrep.Sections(2).Controls("repname").Caption = "DO Wise Report (Selected)"
End If

spdatrep.Sections(2).Controls("repname").Caption = "From Date: " & CStr(DTPicker1.Value) & "      To Date: " & CStr(DTPicker2.Value)
If Option2.Value = True Then
    spdatrep.Sections(2).Controls("repname").Caption = spdatrep.Sections(2).Controls("repname").Caption + "     DO No: " & Text1.Text
End If

SQL = "Select *, second_wt-first_wt as net_wt from spdtwtrep where DATE_OUT>=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "#  and DATE_OUT<=#" & Format(DTPicker2.Value, "mm/dd/yyyy") & "#"

If Trim(Text1.Text) <> "" Then
    SQL = SQL & " and special.do_no = '" & Trim(Text1.Text) & "' order by sl_no desc"
Else
    SQL = SQL & " order by special.do_no, sl_no desc"
End If

Set rs3 = New ADODB.Recordset
rs3.Open SQL, co, adOpenKeyset, adLockOptimistic
If rs3.RecordCount > 0 Then
    rs3.MoveFirst
End If
Set spdatrep.DataSource = rs3

twt = 0
While Not rs3.EOF
 spdatrep.Sections("section1").Controls("text1").DataField = "sl_no"
 spdatrep.Sections("section1").Controls("text2").DataField = "date_in"
 spdatrep.Sections("section1").Controls("text3").DataField = "date_out"
 spdatrep.Sections("section1").Controls("text4").DataField = "time_in"
 spdatrep.Sections("section1").Controls("text5").DataField = "time_out"
 spdatrep.Sections("section1").Controls("text6").DataField = "tc_code"
 spdatrep.Sections("section1").Controls("text16").DataField = "purchaser"
 spdatrep.Sections("section1").Controls("text7").DataField = "do_no"
 spdatrep.Sections("section1").Controls("text8").DataField = "do_date"
 spdatrep.Sections("section1").Controls("text9").DataField = "v_no"
 spdatrep.Sections("section1").Controls("text10").DataField = "rlw"
 spdatrep.Sections("section1").Controls("text11").DataField = "challan_no"
 spdatrep.Sections("section1").Controls("text12").DataField = "grade"
 spdatrep.Sections("section1").Controls("text13").DataField = "first_wt"
 spdatrep.Sections("section1").Controls("text14").DataField = "second_wt"
 spdatrep.Sections("section1").Controls("text15").DataField = "net_wt"
 spdatrep.Sections("section1").Controls("text17").DataField = "coll_code"
 twt = twt + Val(rs3.Fields("net_wt"))
 rs3.MoveNext
Wend
spdatrep.Sections("section5").Controls("twt").Caption = twt
spdatrep.Show vbModal
End Sub


Private Sub Form_Load()
Me.Picture = main.Picture
DTPicker1.Value = Date
DTPicker2.Value = Date


Conn
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    Text1.Text = ""
    Text1.Visible = False
    Text2.Visible = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
    Text1.Visible = True
    Text2.Visible = True
    Text1.Text = ""
    Text1.SetFocus
End If
End Sub
