VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form caimportt 
   Caption         =   "CA Import"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9360
   LinkTopic       =   "Form2"
   ScaleHeight     =   5625
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   4800
      TabIndex        =   6
      Top             =   1680
      Width           =   4215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   76087297
      CurrentDate     =   40985
   End
   Begin VB.FileListBox File1 
      Height          =   4185
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DOWNLOAD CADATA FROM AREA PC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   4920
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "IMPORT SELECTED CA DATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   4920
      Width           =   4215
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   76087297
      CurrentDate     =   40985
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   76087297
      CurrentDate     =   40985
   End
   Begin VB.Label Label2 
      Caption         =   "Please wait..receiving data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Please wait..importing data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "caimportt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents cFTP As clsFTP
Attribute cFTP.VB_VarHelpID = -1

Function dofound(don As String) As Boolean
dofound = False
Set rs2 = New ADODB.Recordset
rs2.Open "Select * from cadata where do_no='" & Trim(don) & "'", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then dofound = True
rs2.Close
End Function

Private Sub Command1_Click()
Dim MailFileName As String, txt As String, stri As String
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from cadata", co, adOpenKeyset, adLockOptimistic
Label1.Visible = True
MailFileName = "D:\wbdata\cadata\" & Trim(File1.FileName)
Dim i As Integer
i = 1
Open MailFileName For Input As i
   While Not EOF(i)
        Line Input #i, txt
        If dofound(Mid(txt, 61, 11)) = False Then
            rs1.AddNew
            rs1.Fields("unit") = Mid(txt, 1, 4)
            rs1.Fields("purchaser") = Mid(txt, 6, 25)
            rs1.Fields("destination") = Mid(txt, 32, 20)
            rs1.Fields("state_code") = Mid(txt, 53, 2)
            rs1.Fields("grade") = Mid(txt, 56, 4)
            rs1.Fields("do_no") = Mid(txt, 61, 11)
            rs1.Fields("do_date") = Mid(txt, 73, 8)
            rs1.Fields("appl_no") = Mid(txt, 82, 8)
            rs1.Fields("appl_date") = Mid(txt, 91, 8)
            rs1.Fields("do_qty") = Mid(txt, 100, 9)
            rs1.Fields("draft_no1") = Mid(txt, 110, 12)
            rs1.Fields("draft_dt1") = Mid(txt, 123, 8)
            rs1.Fields("draft_amt1") = Mid(txt, 132, 11)
            rs1.Fields("bank1") = Mid(txt, 144, 10)
            rs1.Fields("draft_no2") = Mid(txt, 155, 12)
            rs1.Fields("draft_dt2") = Mid(txt, 168, 8)
            rs1.Fields("draft_amt2") = Mid(txt, 177, 11)
            rs1.Fields("bank2") = Mid(txt, 189, 10)
            rs1.Fields("draft_no3") = Mid(txt, 200, 12)
            rs1.Fields("draft_dt3") = Mid(txt, 213, 8)
            rs1.Fields("draft_amt3") = Mid(txt, 222, 11)
            rs1.Fields("bank3") = Mid(txt, 234, 10)
            rs1.Fields("qtybalance") = Mid(txt, 245, 9)
            rs1.Fields("taxtype") = Mid(txt, 255, 1)
            rs1.Fields("tax_percent") = Mid(txt, 257, 5)
            rs1.Fields("custcd") = Mid(txt, 263, 6)
            rs1.Fields("exc_reg_no") = Mid(txt, 270, 20)
            rs1.Fields("range") = Mid(txt, 291, 15)
            rs1.Fields("division") = Mid(txt, 307, 15)
            rs1.Fields("commissionerate") = Mid(txt, 323, 15)
            rs1.Fields("vat_tin_no") = Mid(txt, 339, 15)
            rs1.Fields("cst_no") = Mid(txt, 355, 15)
            rs1.Fields("basic_rate") = Mid(txt, 371, 8)
            rs1.Fields("royalty") = Mid(txt, 380, 7)
            rs1.Fields("sed") = Mid(txt, 388, 6)
            rs1.Fields("clean_engy_cess") = Mid(txt, 395, 6)
            rs1.Fields("weighment_chg") = Mid(txt, 402, 6)
            rs1.Fields("slc") = Mid(txt, 409, 6)
            rs1.Fields("wrc") = Mid(txt, 416, 6)
            rs1.Fields("bazar_fee") = Mid(txt, 423, 6)
            rs1.Fields("PAN") = Mid(txt, 430, 10)
            rs1.Fields("cent_exc_rate") = Mid(txt, 441, 7)
            rs1.Fields("edu_cess_rate") = Mid(txt, 449, 7)
            rs1.Fields("high_edu_rate") = Mid(txt, 457, 7)
            
            If Mid(Mid(txt, 465, 8), 1, 4) = "2017" Or Mid(Mid(txt, 465, 8), 1, 4) = "2018" Or Mid(Mid(txt, 465, 8), 1, 4) = "2019" Or Mid(Mid(txt, 465, 8), 1, 4) = "2020" Then
                rs1.Fields("do_start_date") = Mid(Mid(txt, 465, 8), 1, 4) + Mid(Mid(txt, 465, 8), 5, 2) + Mid(Mid(txt, 465, 8), 7, 2)
            Else
                rs1.Fields("do_start_date") = Mid(Mid(txt, 465, 8), 5, 4) + Mid(Mid(txt, 465, 8), 3, 2) + Mid(Mid(txt, 465, 8), 1, 2)
            End If
            If Mid(Mid(txt, 474, 8), 1, 4) = "2017" Or Mid(Mid(txt, 474, 8), 1, 4) = "2018" Or Mid(Mid(txt, 474, 8), 1, 4) = "2019" Or Mid(Mid(txt, 474, 8), 1, 4) = "2020" Then
                rs1.Fields("do_end_date") = Mid(Mid(txt, 474, 8), 1, 4) + Mid(Mid(txt, 474, 8), 5, 2) + Mid(Mid(txt, 474, 8), 7, 2)
            Else
                rs1.Fields("do_end_date") = Mid(Mid(txt, 474, 8), 5, 4) + Mid(Mid(txt, 474, 8), 3, 2) + Mid(Mid(txt, 474, 8), 1, 2)
            End If
            rs1.Fields("road_cess") = "0"
            rs1.Fields("ambh_cess") = "0"
            rs1.Fields("other_charges") = "0"
            If Len(txt) > 485 Then
                rs1.Fields("tcs") = Mid(txt, 483, 7)
            Else
                rs1.Fields("tcs") = "0"
            End If
            rs1.Update
        End If
        
        If Label1.ForeColor = &H80& Then
            Label1.ForeColor = &H0&
        Else
            Label1.ForeColor = &H80&
        End If
        DoEvents
   Wend
Close i
Label1.Visible = False
rs1.Close
End Sub

Sub genlist()
DTPicker3.Value = DTPicker1.Value
List1.Clear
While DTPicker3.Value <= DTPicker2.Value
    List1.AddItem ("CA" & arcode & Format(DTPicker3.Day, "0#") & Format(DTPicker3.Month, "0#") & Right(CStr(DTPicker3.Year), 2)) & ".txt"
    DTPicker3.Value = DTPicker3.Value + 1
Wend
If List1.ListCount > 0 Then
List1.ListIndex = 0
End If
End Sub

Private Sub Command2_Click()

Set cFTP = New clsFTP
DoEvents
Label2.Visible = True
If cFTP.OpenConnection(ips(2), unames(2), passws(2)) Then
    cFTP.FTPDirectory = "area" & arcode & "/cadata"
'    bSuccess = cFTP.SimpleFTPGetFile("d:\wbdata\cadata\CA09290212.txt", "CA09290212.txt")
'    If bSuccess Then
''        sError = "Success"
        
        For i = 0 To List1.ListCount - 1
            DoEvents
            Label2.Visible = True
            List1.ListIndex = i
            bSuccess = cFTP.SimpleFTPGetFile("d:\wbdata\cadata\" & Trim(List1.Text), Trim(List1.Text))
        Next
        If bSuccess Then
        MsgBox "CA Data received"
        End If
'    Else
'        MsgBox "Error"
'    End If
    cFTP.CloseConnection
Else
    MsgBox "Headquarter not connected", , CStr(sError)
End If
Label2.Visible = False
File1.Refresh
End Sub

Private Sub DTPicker1_Change()
genlist
End Sub

Private Sub DTPicker2_Change()
genlist
End Sub

Private Sub Form_Load()
File1.Path = "d:\wbdata\cadata"
genlist
End Sub
