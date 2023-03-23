VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form exepreport 
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   480
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
      Left            =   8520
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
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
      Left            =   3840
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2160
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
      Format          =   15400961
      CurrentDate     =   42718
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SHOW REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   0
      Top             =   3960
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   2160
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
      Format          =   15400961
      CurrentDate     =   42718
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "EXCEPTION REPORTS"
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
      TabIndex        =   10
      Top             =   600
      Width           =   7815
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
      Left            =   6720
      TabIndex        =   8
      Top             =   1800
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
      Left            =   2040
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date To"
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
      Left            =   6720
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date From"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   4695
      Left            =   1680
      Top             =   480
      Width           =   9975
   End
End
Attribute VB_Name = "exepreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
sql = ""

If Trim(Text1) = "" Then
   MsgBox "Please enter area code"
   Exit Sub
End If

If Trim(Text2.Text) = "" Then
   sql = "Select * from tagexceptions where area like '" + Trim(Text1.Text) + "%'"
Else
   sql = "Select * from tagexceptions where area like '" + Trim(Text1.Text) + "%' and wb like '" + Trim(Text2.Text) + "%'"
End If

Set rs3 = New ADODB.Recordset

sql = sql + " and edate>='" & Format(DTPicker1.Value, "yyyy-mm-dd") & "' and edate<='" & Format(DTPicker2.Value, "yyyy/mm/dd") & "'"


Set fsys = CreateObject("scripting.filesystemobject")
Set wstream = fsys.OpenTextFile(App.Path & "/report.htm", ForWriting, True)

'MsgBox sql
Set rs3 = New ADODB.Recordset
rs3.Open sql, co, adOpenKeyset, adLockOptimistic
If rs3.RecordCount > 0 Then
    rs3.MoveFirst
End If

wstream.WriteLine "<table width=900 cellpadding=0 cellspacing=0>"
wstream.WriteLine "<tr style='font-family: arial; font-size: 12pt;'><td colspan=7 style='text-align: center; font-size: 18pt;'>EXCEPTION REPORTS</td></tr>"
wstream.WriteLine "<tr style='font-family: arial; font-size: 12pt;'>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid; padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "Date/<br>Time</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "Tag No</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "V.No<br>C Code</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "E.Type<br>Sl.No"
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine "Description"
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
    wstream.WriteLine "O Name</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
    wstream.WriteLine "Area/<br>WB"
    wstream.WriteLine "</td>"
wstream.WriteLine "</tr>"



While Not rs3.EOF
wstream.WriteLine "<tr style='font-family: arial; font-size: 12pt;'>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid; padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(rs3.Fields("edate")) + "<br>" + CStr(rs3.Fields("etime"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine Left(CStr(rs3.Fields("tagno")), 4) + "XXXXXXXX" + Right(CStr(rs3.Fields("tagno")), 4)
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(rs3.Fields("v_no")) + "<br>"
        wstream.WriteLine CStr(rs3.Fields("ccode"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(rs3.Fields("type")) + "<br>"
        wstream.WriteLine CStr(rs3.Fields("sl_no"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
        wstream.WriteLine CStr(rs3.Fields("description"))
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
    wstream.WriteLine CStr(rs3.Fields("o_name")) + "<br>"
    wstream.WriteLine "</td>"
    wstream.WriteLine "<td style='border: 1px #dddddd solid;  padding: 10px; font-size: 9pt;'>"
    wstream.WriteLine CStr(rs3.Fields("area")) + "<br>"
    wstream.WriteLine CStr(rs3.Fields("wb"))
    wstream.WriteLine "</td>"
wstream.WriteLine "</tr>"
rs3.MoveNext
Wend
wstream.WriteLine "</table>"
wstream.WriteLine ""
wstream.Close
Shell "Explorer " & App.Path & "\report.htm", vbMaximizedFocus

End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = MDIForm1.Picture
Conn
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub

