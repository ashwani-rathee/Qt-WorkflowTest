VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form doimport 
   BackColor       =   &H00C0E0FF&
   Caption         =   "DO Import"
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
      Format          =   90767361
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
      Caption         =   "DOWNLOAD ALLOTMENT FROM HQ"
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
      Left            =   4680
      TabIndex        =   2
      Top             =   4920
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "IMPORT SELECTED ALLOTMENT"
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
      Format          =   90767361
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
      Format          =   90767361
      CurrentDate     =   40985
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
Attribute VB_Name = "doimport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents cFTP As clsFTP
Attribute cFTP.VB_VarHelpID = -1
Dim dofound As Boolean

'Function dofound(don As String) As Boolean
'dofound = False
'Set rs2 = New ADODB.Recordset
'rs2.Open "Select * from cadata where do_no='" & Trim(don) & "'", co, adOpenKeyset, adLockOptimistic
'If rs2.RecordCount > 0 Then dofound = True
'rs2.Close
'End Function

Private Sub Command1_Click()
Dim MailFileName As String, txt As String, stri As String
Dim arco, dono, dat1, dat2, dat3, dat4, dat5, dat6, dat7, trn1, trn2, trn3, trn4, trn5, trn6, trn7 As String

Set rs1 = New ADODB.Recordset
rs1.Open "Select * from allotment", co, adOpenKeyset, adLockOptimistic
Label1.Visible = True
MailFileName = "D:\wbdata\dodata\" & Trim(File1.FileName)
Dim i As Integer
i = 1
Open MailFileName For Input As i
   While Not EOF(i)
        Line Input #i, txt
        arno = Mid(txt, 1, 4)
        dono = Mid(txt, 6, 11)
        dat1 = Mid(txt, 18, 8)
        dat2 = Mid(txt, 27, 8)
        dat3 = Mid(txt, 36, 8)
        dat4 = Mid(txt, 45, 8)
        dat5 = Mid(txt, 54, 8)
        dat6 = Mid(txt, 63, 8)
        dat7 = Mid(txt, 72, 8)
        trn1 = Mid(txt, 81, 3)
        trn2 = Mid(txt, 85, 3)
        trn3 = Mid(txt, 89, 3)
        trn4 = Mid(txt, 93, 3)
        trn5 = Mid(txt, 97, 3)
        trn6 = Mid(txt, 101, 3)
        trn7 = Mid(txt, 105, 3)
        
        Set rs3 = New ADODB.Recordset
        rs3.Open "Select * from cadata where do_no='" & Trim(dono) & "' and unit='" & Trim(arno) & "'", co, adOpenKeyset, adLockOptimistic
        If rs3.RecordCount > 0 Then
        
        dofound = False
        DTPicker3.Year = Mid(dat1, 5, 4)
        DTPicker3.Month = Mid(dat1, 3, 2)
        DTPicker3.Day = Mid(dat1, 1, 2)
        Set rs2 = New ADODB.Recordset
        rs2.Open "Select * from allotment where do_no='" & Trim(dono) & "' and w_date=#" & Format(DTPicker3.Value, "mm/dd/yyyy") & "#", co, adOpenKeyset, adLockOptimistic
        If rs2.RecordCount > 0 Then
            rs2.Fields("allotment") = Val(trn1)
            rs2.Update
            dofound = True
        End If
        rs2.Close
        If dofound = False Then
            rs1.AddNew
                rs1.Fields("do_no") = Trim(dono)
                rs1.Fields("w_date") = DTPicker3.Value
                rs1.Fields("allotment") = Val(trn1)
            rs1.Update
        End If
        
        dofound = False
        DTPicker3.Year = Mid(dat2, 5, 4)
        DTPicker3.Month = Mid(dat2, 3, 2)
        DTPicker3.Day = Mid(dat2, 1, 2)
        Set rs2 = New ADODB.Recordset
        rs2.Open "Select * from allotment where do_no='" & Trim(dono) & "' and w_date=#" & Format(DTPicker3.Value, "mm/dd/yyyy") & "#", co, adOpenKeyset, adLockOptimistic
        If rs2.RecordCount > 0 Then
            rs2.Fields("allotment") = Val(trn2)
            rs2.Update
            dofound = True
        End If
        rs2.Close
        If dofound = False Then
            rs1.AddNew
                rs1.Fields("do_no") = Trim(dono)
                rs1.Fields("w_date") = DTPicker3.Value
                rs1.Fields("allotment") = Val(trn2)
            rs1.Update
        End If
        
        dofound = False
        DTPicker3.Year = Mid(dat3, 5, 4)
        DTPicker3.Month = Mid(dat3, 3, 2)
        DTPicker3.Day = Mid(dat3, 1, 2)
        Set rs2 = New ADODB.Recordset
        rs2.Open "Select * from allotment where do_no='" & Trim(dono) & "' and w_date=#" & Format(DTPicker3.Value, "mm/dd/yyyy") & "#", co, adOpenKeyset, adLockOptimistic
        If rs2.RecordCount > 0 Then
            rs2.Fields("allotment") = Val(trn3)
            rs2.Update
            dofound = True
        End If
        rs2.Close
        If dofound = False Then
            rs1.AddNew
                rs1.Fields("do_no") = Trim(dono)
                rs1.Fields("w_date") = DTPicker3.Value
                rs1.Fields("allotment") = Val(trn3)
            rs1.Update
        End If
        
        dofound = False
        DTPicker3.Year = Mid(dat4, 5, 4)
        DTPicker3.Month = Mid(dat4, 3, 2)
        DTPicker3.Day = Mid(dat4, 1, 2)
        Set rs2 = New ADODB.Recordset
        rs2.Open "Select * from allotment where do_no='" & Trim(dono) & "' and w_date=#" & Format(DTPicker3.Value, "mm/dd/yyyy") & "#", co, adOpenKeyset, adLockOptimistic
        If rs2.RecordCount > 0 Then
            rs2.Fields("allotment") = Val(trn4)
            rs2.Update
            dofound = True
        End If
        rs2.Close
        If dofound = False Then
            rs1.AddNew
                rs1.Fields("do_no") = Trim(dono)
                rs1.Fields("w_date") = DTPicker3.Value
                rs1.Fields("allotment") = Val(trn4)
            rs1.Update
        End If
        
        dofound = False
        DTPicker3.Year = Mid(dat5, 5, 4)
        DTPicker3.Month = Mid(dat5, 3, 2)
        DTPicker3.Day = Mid(dat5, 1, 2)
        Set rs2 = New ADODB.Recordset
        rs2.Open "Select * from allotment where do_no='" & Trim(dono) & "' and w_date=#" & Format(DTPicker3.Value, "mm/dd/yyyy") & "#", co, adOpenKeyset, adLockOptimistic
        If rs2.RecordCount > 0 Then
            rs2.Fields("allotment") = Val(trn5)
            rs2.Update
            dofound = True
        End If
        rs2.Close
        If dofound = False Then
            rs1.AddNew
                rs1.Fields("do_no") = Trim(dono)
                rs1.Fields("w_date") = DTPicker3.Value
                rs1.Fields("allotment") = Val(trn5)
            rs1.Update
        End If
        
        dofound = False
        DTPicker3.Year = Mid(dat6, 5, 4)
        DTPicker3.Month = Mid(dat6, 3, 2)
        DTPicker3.Day = Mid(dat6, 1, 2)
        Set rs2 = New ADODB.Recordset
        rs2.Open "Select * from allotment where do_no='" & Trim(dono) & "' and w_date=#" & Format(DTPicker3.Value, "mm/dd/yyyy") & "#", co, adOpenKeyset, adLockOptimistic
        If rs2.RecordCount > 0 Then
            rs2.Fields("allotment") = Val(trn6)
            rs2.Update
            dofound = True
        End If
        rs2.Close
        If dofound = False Then
            rs1.AddNew
                rs1.Fields("do_no") = Trim(dono)
                rs1.Fields("w_date") = DTPicker3.Value
                rs1.Fields("allotment") = Val(trn6)
            rs1.Update
        End If
        
        dofound = False
        DTPicker3.Year = Mid(dat7, 5, 4)
        DTPicker3.Month = Mid(dat7, 3, 2)
        DTPicker3.Day = Mid(dat7, 1, 2)
        Set rs2 = New ADODB.Recordset
        rs2.Open "Select * from allotment where do_no='" & Trim(dono) & "' and w_date=#" & Format(DTPicker3.Value, "mm/dd/yyyy") & "#", co, adOpenKeyset, adLockOptimistic
        If rs2.RecordCount > 0 Then
            rs2.Fields("allotment") = Val(trn7)
            rs2.Update
            dofound = True
        End If
        rs2.Close
        If dofound = False Then
            rs1.AddNew
                rs1.Fields("do_no") = Trim(dono)
                rs1.Fields("w_date") = DTPicker3.Value
                rs1.Fields("allotment") = Val(trn7)
            rs1.Update
        End If
        
        End If
        rs3.Close
        
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
    List1.AddItem ("DOA" & arcode & Format(DTPicker3.Day, "0#") & Format(DTPicker3.Month, "0#") & Format(DTPicker3.Year, "000#") & ".txt")
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
If cFTP.OpenConnection(ips(3), unames(3), passws(3)) Then
    cFTP.FTPDirectory = paths(3)
        
        For i = 0 To List1.ListCount - 1
            DoEvents
            Label2.Visible = True
            List1.ListIndex = i
            bSuccess = cFTP.SimpleFTPGetFile("d:\wbdata\dodata\" & Trim(List1.Text), Trim(List1.Text))
        Next
        If bSuccess Then
        MsgBox "DO Allotment Data received"
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
File1.Path = "d:\wbdata\dodata"
DTPicker1.Value = Date
DTPicker2.Value = Date
genlist
End Sub
