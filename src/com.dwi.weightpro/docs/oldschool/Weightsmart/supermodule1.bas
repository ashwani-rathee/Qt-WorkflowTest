Attribute VB_Name = "Module1"
Public co As ADODB.Connection
Public rs As ADODB.Recordset
Public rs1 As ADODB.Recordset
Public rs2 As ADODB.Recordset
Public mdatabase As String


Public userpermission As String
Public loginname As String
Public logintype As String

Public demo As Boolean

Public checkempty1 As Boolean
Public checkempty2  As Boolean
Public chkwt1 As Double
Public chkwt2 As Double
Public chkunit As String
Public compdes As String

Public COMPAUTH As String
Public platformin As Boolean

Option Explicit
Const MAX_PATH& = 260


Public Function Conn()
   Set co = New ADODB.Connection
   If co.State = 1 Then co.Close
   'co.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdatabase & ";Persist Security Info=False"
   
   co.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)}; Dbq=" & App.Path & "\safe.mdb; Uid=Admin; Pwd=rana@safe@123"
   co.Open

   'co.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdatabase & ";Persist Security Info=False;pwd=hsd123;"

'    complete connection string passed to the Open method
'cn = "Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";UserID=" & userID & ";pwd=" & Password
'ADOCon.Open cn
End Function


Public Function TextEncr(Text As String) As String
On Error Resume Next


Dim TextLen As Integer
Dim Char As String
Dim ChartoAsc As Integer
Dim i As Integer

TextLen = Len(Text)
For i = 1 To TextLen
    Char = Mid(Text, i, 1)
    ChartoAsc = Asc(Char)
    ChartoAsc = (((((((ChartoAsc / 2) + 30) / 2) * 3) - 20) * 3) / 2)
    ChartoAsc = ChartoAsc + 57
    Char = Chr(ChartoAsc)
    TextEncr = TextEncr + Char
Next i
End Function

Public Function Textdcrt(Text As String) As String
'On Error Resume Next
Dim TextLen As Integer
Dim Char As String
Dim ChartoAsc As Integer
Dim i As Integer

TextLen = Len(Text)
For i = 1 To TextLen
    Char = VBA.Mid(Text, i, 1)
    ChartoAsc = Asc(Char)
    ChartoAsc = ChartoAsc - 57
    ChartoAsc = (((((((ChartoAsc * 2) / 3) + 20) / 3) * 2) - 30) * 2)
    Char = VBA.Chr(ChartoAsc)
    Textdcrt = Textdcrt + Char
Next i
End Function


