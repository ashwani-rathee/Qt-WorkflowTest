VERSION 5.00
Begin VB.Form achead 
   Caption         =   "A/C Head Master"
   ClientHeight    =   10980
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame addfrm 
      Height          =   735
      Left            =   3240
      TabIndex        =   16
      Top             =   7920
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmddel 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   8760
      TabIndex        =   11
      Top             =   1800
      Width           =   3495
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5010
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Selection List"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   3120
      TabIndex        =   8
      Top             =   1800
      Width           =   5655
      Begin VB.TextBox Text2 
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
         Left            =   2160
         TabIndex        =   1
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Text1 
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
         Left            =   2160
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Customer Id"
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
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1980
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Name"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.Frame editfrm 
      Height          =   735
      Left            =   3120
      TabIndex        =   3
      Top             =   6960
      Width           =   5655
      Begin VB.CommandButton cmddel1 
         Caption         =   "&Delete"
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
         Left            =   4680
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
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
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdref 
         Caption         =   "&Refresh"
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
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
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
      Left            =   11760
      TabIndex        =   2
      ToolTipText     =   "Unload Form"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Customer Master"
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
      Left            =   6240
      TabIndex        =   14
      Top             =   1320
      Width           =   2010
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   3120
      TabIndex        =   15
      Top             =   1200
      Width           =   9135
   End
End
Attribute VB_Name = "achead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opedit As Boolean
Dim recfound As Boolean
Dim str1 As String
Dim num As Integer
Dim g As Integer
Dim loaded As Boolean
Dim chkpass As Integer

Private Sub cmdadd_Click()
opedit = False
unlocktxt
addfrm.Visible = False
editfrm.Visible = True
Text2.SetFocus
End Sub

Private Sub cmdcancel_Click()
Text1.Enabled = True
opedit = False
unlocktxt
Text2.SetFocus
End Sub

Private Sub cmddel1_Click()
a = MsgBox("Want to Delete Records", vbOKCancel, "Data Deletion")
If a = 1 Then
Set rs3 = New ADODB.Recordset
rs3.Open "select * from special where tc_code='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs3.RecordCount = 0 Then
Set rs5 = New ADODB.Recordset
rs5.Open "delete  from consignee where c_code='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
MsgBox "Records Deleted Successfully", vbInformation, "Delete Confirmation"
unlocktxt
Text1.Enabled = True
Text2.SetFocus
Else
MsgBox "Record cannot be deleted, record exists in sale file"
End If

End If
End Sub

Private Sub cmdedit_Click()
opedit = True
unlocktxt
addfrm.Visible = False
editfrm.Visible = True
Text2.SetFocus

End Sub

Private Sub cmdref_Click()
unlocktxt
Text1.Enabled = True
If Text1.Locked = False Then
    Text1.SetFocus
Else
    Text2.SetFocus
End If
End Sub

Private Sub cmdsave_Click()
a = MsgBox("Want to Save Data", vbOKCancel, "Save Confirmation")

If Trim(Text1.Text) <> "" And Len(Trim(Text2.Text)) > 10 Then
    If a = 1 Then
        dataset
        If recfound = False Then
            rs1.AddNew
        End If
        rs1.Fields("c_code").Value = Trim(Text1.Text)
        rs1.Fields("c_name").Value = Trim(Text2.Text)

        rs1.Update
        MsgBox "Data Updated Successfully", vbInformation, "Data Updated"
    End If
    unlocktxt
Else
    MsgBox "Please fill all the fields properly", , "Incorrect data"
End If
Text1.Enabled = True
Text2.SetFocus

End Sub

Private Sub dataset()
recfound = False
Set rs1 = New ADODB.Recordset
rs1.Open "select * from consignee where c_code='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
recfound = True
End If

End Sub


Private Sub cmdsave_GotFocus()
If Trim(Text1.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
autoincr
Text2.SetFocus
Exit Sub
End If

If Trim(Text2.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
Text2.SetFocus
Exit Sub
End If

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub EdBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If

End Sub


Private Sub Form_Load()
Me.Picture = main.Picture
Call Conn
unlocktxt
loaded = False
End Sub

Private Sub txtlock()
Text1.Locked = True
Text2.Locked = True
End Sub


Private Sub unlocktxt()
Text1.Text = ""
Text2.Text = ""

Text1.Locked = False
Text2.Locked = False
'autoincr
End Sub


Private Sub autoincr()
'Set rs1 = New ADODB.Recordset
'rs1.Open "Select * from consignee  order by val(c_code)", co, adOpenKeyset, adLockOptimistic
'If rs1.RecordCount > 0 Then
'rs1.MoveLast
'Text1.Text = IncrBillNo1(rs1.Fields("c_code").Value)
'Else
'Text1.Text = "1"
'End If
End Sub


Public Sub dd1()
If Len(str1) > 1 Then
For X = 1 To Len(str1) + 1
m = Mid(str1, X, X + 1)
If m = " " Then
num = 0
Else
num = num + 1
End If
Next X
End If

If Len(str1) = 0 Then
num = 0
End If
End Sub
Public Sub KeyPress1(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then
If Len(str1) > 1 Then
              If num = 1 Then
                  g = KeyAscii - 32
                  KeyAscii = g
            Else
            g = KeyAscii
            KeyAscii = g
            End If
            
End If


If Len(str1) = 0 Then
  g = KeyAscii - 32
  KeyAscii = g
  End If
End If
End Sub

Private Sub List1_DblClick()
'If opedit = True Then
    If chkpass = 1 Then
    Text1.Text = Trim(padl(List1.List(List1.ListIndex), 6))
    datashow
    Text2.SetFocus
    End If
'End If

If chkpass = 2 Then
    Text1.Text = tmpcode1(List1.ListIndex)
    datashow
    Text2.SetFocus
End If
End Sub

Private Sub Text1_Change()
str1 = Text1.Text
Call dd1
For i = 0 To List1.ListCount - 1
      If Trim(Text1.Text) = Left(List1.List(i), Len(Trim(Text1.Text))) Then
                    List1.ListIndex = i
                    Exit Sub
        End If
    Next i
End Sub

Private Sub Text1_GotFocus()
chkpass = 1
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
str1 = ""
num = 0
If loaded = False Then
    Label9.Caption = "Wait, Loading Consignee List..."
    MsgBox "Loading Consignee List can take some time..CLICK OK TO START"
    Id_list
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
Call KeyPress1(KeyAscii)
End If

If KeyAscii = 13 Then
    If Trim(Text1.Text) <> "" Then
         datashow
        If opedit = True Then
        Text1.Text = Trim(padl(List1.List(List1.ListIndex), 5))
        End If
        
        
        If opedit = False And recfound = True Then
        MsgBox "This Id already Exsist ", vbInformation, "Data Duplicated"
        Text1.Enabled = True
        unlocktxt
        Text2.SetFocus
        Exit Sub
        End If
        
        If opedit = True And recfound = False Then
        MsgBox "This Id not Exsist in Master ", vbInformation, "No Data "
        
        Text2.SetFocus
        Exit Sub
        End If
        
    Text2.SetFocus
    End If

End If

End Sub

Private Sub name_list()
If List1.ListCount > 0 Then
List1.Clear
List1.Refresh
End If

Set rs = New ADODB.Recordset
rs.Open "Select * from consignee order by c_name", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveFirst
g = 0

End If
g = 0
ReDim tmpcode1(rs.RecordCount)
While Not rs.EOF
List1.AddItem Trim(rs.Fields("c_name").Value & "")
tmpcode1(g) = rs.Fields("c_code").Value
g = g + 1
rs.MoveNext
Wend
End Sub


Private Sub Id_list()
If List1.ListCount > 0 Then
List1.Clear
List1.Refresh
End If


Set rs = New ADODB.Recordset
rs.Open "Select * from consignee order by C_CODE", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveFirst
End If

g = 0
ReDim tmp(rs.RecordCount)
While Not rs.EOF
List1.AddItem padl(rs.Fields("c_code").Value, 6) & " " & padl(rs.Fields("c_name").Value, 25)

rs.MoveNext
Wend
Label9.Caption = "Selection List"
loaded = True
End Sub

Private Sub datashow()
On Error Resume Next
recfound = False
opedit = False
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from consignee where c_code='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
    Text1.Enabled = False
    rs1.MoveFirst
    opedit = True
    recfound = True
    Text2.Text = rs1.Fields("c_name").Value
End If
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
Call KeyPress1(KeyAscii)
End If

If KeyAscii = 13 Then
cmdsave.SetFocus
End If

End Sub

