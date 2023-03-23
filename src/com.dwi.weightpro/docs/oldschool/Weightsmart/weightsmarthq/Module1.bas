Attribute VB_Name = "Module1"
Public co As ADODB.Connection
Public co1 As ADODB.Connection
Public rs As ADODB.Recordset
Public rs1 As ADODB.Recordset
Public rs2 As ADODB.Recordset
Public rs3 As ADODB.Recordset
Public rs4 As ADODB.Recordset
Public rs5 As ADODB.Recordset
Public rs6 As ADODB.Recordset


Public Function Conn()
   Set co = New ADODB.Connection
   If co.State = 1 Then co.Close
   'co.ConnectionString = "PROVIDER = MSDASQL;driver={SQL Server};database=weighthq; server=192.168.1.5,1433;uid=bccl;pwd=Edesp@#123;"
   co.ConnectionString = "PROVIDER = MSDASQL;driver={SQL Server};database=weighthq; server=172.20.0.96,1433;uid=bccl;pwd=Edesp@#123;"
   co.Open
End Function


Public Function Conn1()
   Set co1 = New ADODB.Connection
   If co1.State = 1 Then co1.Close
  fhbc = &H404040
  co1.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)}; Dbq=" & App.Path & "\bccl.mdb; Uid=Admin; Pwd=hsd123"
co1.Open
End Function

Public Function TextEncr(Text As String) As String
On Error Resume Next


Dim TextLen As Integer
Dim Char As String
Dim ChartoAsc As Integer
Dim I As Integer

TextLen = Len(Text)
For I = 1 To TextLen
    Char = Mid(Text, I, 1)
    ChartoAsc = Asc(Char)
    ChartoAsc = (((((((ChartoAsc / 2) + 30) / 2) * 3) - 20) * 3) / 2)
    ChartoAsc = ChartoAsc + 57
    Char = Chr(ChartoAsc)
    TextEncr = TextEncr + Char
Next I
End Function

Public Function Textdcrt(Text As String) As String
'On Error Resume Next
Dim TextLen As Integer
Dim Char As String
Dim ChartoAsc As Integer
Dim I As Integer

TextLen = Len(Text)
For I = 1 To TextLen
    Char = VBA.Mid(Text, I, 1)
    ChartoAsc = Asc(Char)
    ChartoAsc = ChartoAsc - 57
    ChartoAsc = (((((((ChartoAsc * 2) / 3) + 20) / 3) * 2) - 30) * 2)
    Char = VBA.Chr(ChartoAsc)
    Textdcrt = Textdcrt + Char
Next I
End Function


Public Function IncrBillNo1(BillNo As String) As String
    Dim intBillNo As Long
    Dim BillLen As Integer
    intBillNo = Val(BillNo)
    intBillNo = intBillNo + 1
    BillLen = Len(CStr(intBillNo))
    IncrBillNo1 = CStr(intBillNo)
End Function

Public Function IncrBillNo4(BillNo As String) As String
    Dim intBillNo As Long
    Dim BillLen As Integer
    intBillNo = Val(Right(BillNo, 4))
    intBillNo = intBillNo + 1
    BillLen = Len(CStr(intBillNo))
    IncrBillNo4 = String(4 - BillLen, "0") & CStr(intBillNo)
End Function

Public Function IncrBillNo(BillNo As String) As String
    Dim intBillNo As Long
    Dim BillLen As Integer
    intBillNo = Val(Right(BillNo, 5))
    intBillNo = intBillNo + 1
    BillLen = Len(CStr(intBillNo))
    IncrBillNo = String(5 - BillLen, "0") & CStr(intBillNo)
End Function



Public Function PADR(Message As String, length As Integer) As String
''''Message = Trim(Message)
''''Message = Space(length - Len(Message)) + Message
   '''PADR = Message
    Dim l As Integer   '****Holds the Length of Message
    Dim Pad As Integer
    Dim str As String
    l = Len(Message)
    If l < length Then
        Pad = length - l
        str = Message & Space(Pad)

    Else
        str = Left(Message, length)
    End If

 
    
    PADR = str
End Function

Public Function PADC(Message As String, length As Integer) As String
    Dim l As Integer   '****Holds the Length of Message
    Dim Pad As Integer
    Dim PL As Integer
    Dim pr As Integer
    Dim str As String
    l = Len(Message)
    If l < length Then
        Pad = length - l
        
        PL = Pad / 2      '**********space towards left of the line
        pr = Pad - PL     '**********space towards right of the line
              
        str = Space(PL) & Message & Space(pr)
    Else
        str = Left(Message, length)
    End If
    PADC = str
End Function

Public Function padl(Nmassage As String, length As Integer)
Dim l As Integer
Dim Pad As String
Dim PL As Integer
Dim pr As Integer
Dim str As String
 l = Len(Nmassage)
If l < length Then
Pad = length - l
str = Nmassage & Space(Pad)
Else
str = Left(Nmassage, length)
End If
padl = str

'''Nmassage = Trim(Nmassage)
'''Nmassage = Nmassage + Space(length - Len(Nmassage))
End Function

Public Function IncrBillNo11(BillNo As String) As String
    Dim intBillNo As Long
    Dim BillLen As Integer
    intBillNo = Val(Right(BillNo, 2))
    intBillNo = intBillNo + 1
    BillLen = Len(CStr(intBillNo))
    IncrBillNo11 = String(2 - BillLen, "0") & CStr(intBillNo)
End Function

