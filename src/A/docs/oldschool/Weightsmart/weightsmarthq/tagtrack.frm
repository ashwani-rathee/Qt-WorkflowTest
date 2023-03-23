VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form tagtrack 
   Caption         =   "TAG Track"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14895
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   14895
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "DATE WISE REPORT"
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
      TabIndex        =   19
      Top             =   2280
      Width           =   3135
   End
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
      Left            =   13680
      TabIndex        =   14
      ToolTipText     =   "Unload Form"
      Top             =   480
      Width           =   375
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
      Left            =   5640
      TabIndex        =   13
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LAST WEIGHMENT"
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
      Left            =   8400
      TabIndex        =   12
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   720
      TabIndex        =   0
      Top             =   2760
      Width           =   13335
      Begin VB.TextBox Text21 
         Height          =   375
         Left            =   11280
         TabIndex        =   7
         Top             =   180
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text22 
         Height          =   375
         Left            =   11280
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text23 
         Height          =   375
         Left            =   11280
         TabIndex        =   5
         Top             =   1020
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text24 
         Height          =   375
         Left            =   11280
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text25 
         Height          =   375
         Left            =   11280
         TabIndex        =   3
         Top             =   1860
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text26 
         Height          =   375
         Left            =   11280
         TabIndex        =   2
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text27 
         Height          =   375
         Left            =   11280
         TabIndex        =   1
         Top             =   2700
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   12855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   12855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "TAG DETAILS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "LAST WEIGHMENT DETAILS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2460
         Width           =   4335
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4200
      TabIndex        =   20
      Top             =   2280
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
      Format          =   48627713
      CurrentDate     =   42718
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Top             =   2280
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
      Format          =   48627713
      CurrentDate     =   42718
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   23
      Top             =   2355
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   21
      Top             =   2355
      Width           =   735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAG / VEHICLE TRACKER"
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
      Left            =   5640
      TabIndex        =   17
      Top             =   540
      Width           =   3000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Vehicle Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1440
      TabIndex        =   16
      Top             =   1020
      Width           =   11895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4080
      TabIndex        =   15
      Top             =   1740
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   6915
      Left            =   600
      Top             =   960
      Width           =   13575
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   600
      TabIndex        =   18
      Top             =   360
      Width           =   13575
   End
End
Attribute VB_Name = "tagtrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from  tags where v_no = '" + Trim(Text1.Text) + "' order by sno desc", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Text21 = rs1.Fields("tagno")
Text23 = "Tag Cancellation"
Text25 = loginname
Text26 = Left(rs1.Fields("unit"), 2)
Text27 = Right(rs1.Fields("unit"), 2)

Label4.Caption = "Tag No: " + Left(rs1.Fields("tagno"), 4) + "XXXXXXXXXXXXXXXXXX" + Right(rs1.Fields("tagno"), 4) + "  Issue: " + CStr(rs1.Fields("issue")) + "  Expiry: " + CStr(rs1.Fields("expiry")) _
+ vbCrLf + "Customer: " + rs1.Fields("tc_code") + "  Vehicle: " + rs1.Fields("v_no") + "  Material: " + rs1.Fields("tm_code") _
+ vbCrLf + "DO: " + rs1.Fields("do_no") + "  Coll: " + rs1.Fields("coll_code") + "  Valid: " + rs1.Fields("valid") + "  Trips: " + CStr(rs1.Fields("tagtrips")) + "  TripsDone: " + CStr(rs1.Fields("trips_done"))
Else
    MsgBox "No active TAG found for Vehicle...it is not issued or transaction is complete"
    Exit Sub
End If
rs1.Close

Set rs1 = New ADODB.Recordset
rs1.Open "Select * from  special where v_no = '" + Trim(Text1.Text) + "' order by sl_no desc", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Text22 = rs1.Fields("sl_no")
Text23 = "Transaction Cancellation"
Label5.Caption = "Serial No: " + rs1.Fields("sl_no") + "  Date in: " + CStr(rs1.Fields("date_in")) + "  Time In: " + CStr(rs1.Fields("time_in")) _
+ vbCrLf + "Operator: " + rs1.Fields("o_name") + "  First Wt: " + CStr(rs1.Fields("first_wt")) + "  Second Wt: " + CStr(rs1.Fields("second_wt")) + "  Net Wt: " + CStr(Abs(CStr(rs1.Fields("second_wt")) - CStr(rs1.Fields("first_wt")))) _
+ "  Destination: " + rs1.Fields("dest")
End If
rs1.Close

End Sub


Private Sub Command2_Click()
Dim tottri As Integer
Dim totfst As Long
Dim totsec As Long
Dim totnet As Long

On Error Resume Next
sql = ""

If Trim(Text1) = "" Then
   MsgBox "Please enter vehicle number"
   Exit Sub
End If

sql = "Select * from  special where v_no = '" + Trim(Text1.Text) + "'"

Set rs1 = New ADODB.Recordset

sql = sql + " and date_in>='" & Format(DTPicker1.Value, "yyyy-mm-dd") & "' and date_in<='" & Format(DTPicker2.Value, "yyyy/mm/dd") & "'"


Set fsys = CreateObject("scripting.filesystemobject")
Set wstream = fsys.OpenTextFile(App.Path & "/report.htm", ForWriting, True)

'MsgBox sql
Set rs1 = New ADODB.Recordset
rs1.Open sql, co, adOpenKeyset, adLockOptimistic

wstream.WriteLine "<table width=900 cellpadding=0 cellspacing=0>"
wstream.WriteLine "<tr style='font-family: arial; font-size: 12pt;'><td colspan=7 style='text-align: center; font-size: 18pt;'>VEHICLE REPORTS (" + Text1.Text + ")</td></tr>"
wstream.WriteLine "<tr style='font-family: arial; font-size: 12pt;'>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid; padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "sl.no</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "DateIn<br>TimeIn</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "DateOut<br>TimeOut</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "Operator</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "1st-Weight"
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "2nd-Weight"
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
    wstream.WriteLine "Net-Wt</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
    wstream.WriteLine "Destination</td>"
wstream.WriteLine "</tr>"

tottri = 0
totfst = 0
totsec = 0
totnet = 0

While Not rs1.EOF
wstream.WriteLine "<tr style='font-family: arial; font-size: 12pt;'>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid; padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(rs1.Fields("sl_no"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(rs1.Fields("date_in")) + "<br>"
        wstream.WriteLine CStr(rs1.Fields("time_in"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(rs1.Fields("date_out")) + "<br>"
        wstream.WriteLine CStr(rs1.Fields("time_out"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine rs1.Fields("o_name") + "<br>"
        wstream.WriteLine rs1.Fields("o2_name")
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(rs1.Fields("first_wt"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(rs1.Fields("second_wt"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
    If Val(rs1.Fields("second_wt")) > 0 Then
    wstream.WriteLine CStr(Abs(CStr(rs1.Fields("second_wt")) - CStr(rs1.Fields("first_wt"))))
    Else
    wstream.WriteLine "0"
    End If
    wstream.WriteLine "</td>"
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
    wstream.WriteLine rs1.Fields("dest")
    wstream.WriteLine "</td>"
wstream.WriteLine "</tr>"
tottri = tottri + 1
totfst = totfst + rs1.Fields("first_wt")
totsec = totsec + rs1.Fields("second_wt")
totnet = totnet + Abs(rs1.Fields("second_wt") - rs1.Fields("first_wt"))
rs1.MoveNext
Wend

wstream.WriteLine "<tr style='font-family: arial; font-size: 12pt;'>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid; padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='font-weight: bold; border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='font-weight: bold; border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "Trips: " + CStr(tottri) + "</td>"
    wstream.WriteLine "<td style='font-weight: bold; border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(totfst)
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='font-weight: bold; border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(totsec)
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='font-weight: bold; border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
    wstream.WriteLine CStr(totnet) + "</td>"
    wstream.WriteLine "<td style='font-weight: bold; border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
    wstream.WriteLine "</td>"
wstream.WriteLine "</tr>"

wstream.WriteLine "</table>"
wstream.WriteLine ""
wstream.Close
Shell "Explorer " & App.Path & "\report.htm", vbMaximizedFocus
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = MDIForm1.Picture
Conn
DTPicker1.Value = Date
DTPicker2.Value = Date

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 13 Then
Text1.Locked = False
Else
Text1.Locked = True
End If
End Sub



