VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form sumreport 
   Caption         =   "Exeption Reporting"
   ClientHeight    =   9210
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   13380
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   8400
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1560
      Width           =   12375
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
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
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   495
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
      Left            =   6360
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   10320
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   48562177
      CurrentDate     =   42718
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MANUAL REFRESH"
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
      Left            =   9240
      TabIndex        =   0
      Top             =   7560
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   7440
      Width           =   8775
      Begin VB.Label Label2 
         Caption         =   "60 seconds to refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   200
         Width           =   8295
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DAILY SUMMARY REPORT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   7
      Top             =   300
      Width           =   8295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "WB Code"
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
      Left            =   5280
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Area Code"
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
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   9360
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   480
      Top             =   240
      Width           =   12375
   End
End
Attribute VB_Name = "sumreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim enable As Boolean
Dim rctr As Integer

Dim totdis As Long
Dim totveh As Integer

Private Sub backup()
sql = ""

sql = "Select * from special where 1=1"

If Trim(Text2.Text) = "" Then
   sql = "and sl_no like '" + Trim(Text1.Text) + "%'"
Else
   sql = "and sl_no like '" + Trim(Text1.Text) + Trim(Text2.Text) + "%'"
End If

sql = sql + " and date_in>='" & Format(DTPicker1.Value, "yyyy-mm-dd") & "' and date_in<='" & Format(DTPicker2.Value, "yyyy/mm/dd") & "'"
sql = sql + " order by sl_no"

Set fsys = CreateObject("scripting.filesystemobject")
Set wstream = fsys.OpenTextFile(App.Path & "/report.htm", ForWriting, True)

Set rs3 = New ADODB.Recordset
rs3.Open sql, co, adOpenKeyset, adLockOptimistic
If rs3.RecordCount > 0 Then
    rs3.MoveFirst
End If

wstream.WriteLine "<table width=900 cellpadding=0 cellspacing=0>"

While Not rs3.EOF
wstream.WriteLine "<tr style='font-family: arial; font-size: 12pt;'>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid; padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(rs3.Fields("sl_no")) + "<br>" + Left(CStr(rs3.Fields("tag")), 4) + "XXXXXXXX" + Right(CStr(rs3.Fields("tag")), 4)
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(rs3.Fields("date_in")) + "<br>"
        wstream.WriteLine CStr(rs3.Fields("date_out"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(rs3.Fields("time_in")) + "<br>"
        wstream.WriteLine CStr(rs3.Fields("time_out"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(rs3.Fields("tc_code")) + "<br>"
        wstream.WriteLine CStr(rs3.Fields("tm_code"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(rs3.Fields("do_no")) + "<br>"
        wstream.WriteLine CStr(rs3.Fields("challan_no"))
    wstream.WriteLine "</td>"
    
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
    wstream.WriteLine CStr(rs3.Fields("o_name")) + "<br>"
    wstream.WriteLine CStr(rs3.Fields("o2_name"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
    wstream.WriteLine CStr(rs3.Fields("v_no")) + "<br>" + CStr(rs3.Fields("rlw"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
    wstream.WriteLine CStr(rs3.Fields("first_wt")) + "<br>"
    wstream.WriteLine CStr(rs3.Fields("second_wt"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
    wstream.WriteLine CStr(Abs(rs3.Fields("first_wt") - rs3.Fields("second_wt")))
    wstream.WriteLine "</td>"
        
wstream.WriteLine "</tr>"
rs3.MoveNext
Wend
wstream.WriteLine "</table>"
wstream.WriteLine ""
wstream.Close
Shell "Explorer " & App.Path & "\report.htm", vbMaximizedFocus

End Sub


Private Sub Command1_Click()
If co.State = 0 Then
Conn
End If
Command1.Caption = "WAIT"
Command1.Enabled = False

Text1.Text = ""
If Trim(Combo1.Text) = "" Then
    showallarea
ElseIf Trim(Text2.Text) = "" Then
    enable = True
    showarea Trim(Combo1.Text)
Else
    enable = True
    showwb Trim(Combo1.Text), Trim(Text2.Text)
End If
rctr = 0
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
Label9.Caption = "DAILY SUMMARY REPORT: " + Format(DTPicker1.Value, "dd-mm-yyyy")
End Sub

Private Sub Form_Load()
Me.Picture = MDIForm1.Picture
Conn
DTPicker1.Value = Date
Label9.Caption = "DAILY SUMMARY REPORT: " + Format(DTPicker1.Value, "dd-mm-yyyy")
loadarea
rctr = 0
Command1_Click
End Sub

Sub showallarea()
Dim areacode As String
totdis = 0
totveh = 0
enable = False
If Combo1.ListCount = 0 Then
    MsgBox "No areas loaded"
    Exit Sub
End If

Text1.Text = ""
For j = 0 To Combo1.ListCount - 1
    Combo1.ListIndex = j
    areacode = Combo1.Text
    If areacode <> "" Then showarea areacode
Next

Text1.Text = Text1.Text + vbCrLf + "Total vehicles: " + CStr(totveh) + "   Total Dispatch: " + CStr(totdis)

enable = True
    If enable = True Then
        Command1.Caption = "MANUAL REFRESH"
        Command1.Enabled = True
    End If
    Combo1.Text = ""
End Sub

Sub showarea(arco As String)
    On Error GoTo errarea
    Dim nov, disp, disp1 As Long
    nov = 0
    disp = 0
    disp1 = 0
    
    sql = "select count(*) as nov, sum(second_wt) as disp, sum(first_wt) as disp1 from special where second_wt>first_wt and sl_no like '" + arco + "%' and date_out='" + Format(DTPicker1.Value, "yyyy-mm-dd") + "'"
    Set rs3 = New ADODB.Recordset
    rs3.Open sql, co, adOpenKeyset, adLockOptimistic
    If rs3.RecordCount > 0 Then
        nov = rs3.Fields("nov")
        totveh = totveh + nov
        disp1 = rs3.Fields("disp1")
        disp = rs3.Fields("disp")
        totdis = totdis + (disp - disp1)
    End If
    rs3.Close
    
errarea:
    Text1.Text = Text1.Text + "Area: " + arco + "   Vehicles: " + CStr(nov) + "   Dispatch: " + CStr(disp - disp1) + vbCrLf
    If enable = True Then
        Command1.Caption = "MANUAL REFRESH"
        Command1.Enabled = True
    End If
End Sub

Sub showwb(arco As String, wbco As String)
On Error GoTo errwb
    Dim nov, disp, disp1 As Long
    nov = 0
    disp = 0
    disp1 = 0
    
    sql = "select count(*) as nov, sum(second_wt) as disp, sum(first_wt) as disp1 from special where second_wt>first_wt and sl_no like '" + arco + wbco + "%' and date_out='" + Format(DTPicker1.Value, "yyyy-mm-dd") + "'"
    Set rs3 = New ADODB.Recordset
    rs3.Open sql, co, adOpenKeyset, adLockOptimistic
    If rs3.RecordCount > 0 Then
        nov = rs3.Fields("nov")
        disp1 = rs3.Fields("disp1")
        disp = rs3.Fields("disp")
    End If
    rs3.Close
    
errwb:
    Text1.Text = Text1.Text + "Area: " + arco + "   Vehicles: " + CStr(nov) + "   Dispatch: " + CStr(disp - disp1) + vbCrLf
    If enable = True Then
        Command1.Caption = "MANUAL REFRESH"
        Command1.Enabled = True
    End If
End Sub

Sub loadarea()
sql = "select distinct substring(sl_no,1,2) as area from special"
Set rs3 = New ADODB.Recordset
rs3.Open sql, co, adOpenKeyset, adLockOptimistic
If rs3.RecordCount > 0 Then
    rs3.MoveFirst
    Combo1.Clear
    While rs3.EOF = False
        Combo1.AddItem (rs3.Fields("area"))
        rs3.MoveNext
    Wend
End If
rs3.Close
End Sub

Private Sub Timer1_Timer()
rctr = rctr + 1
Label2.Caption = CStr(60 - rctr) + " seconds to refresh"
If rctr = 60 Then
    rctr = 0
    Command1_Click
End If
End Sub
