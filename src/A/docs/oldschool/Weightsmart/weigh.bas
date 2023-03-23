Attribute VB_Name = "weigh"
Public Declare Function Net_Connect Lib "uhfreader.dll" (ByVal hostip As String, ByVal hostport As Integer, ByVal readerip As String, ByVal readerport As Integer) As Long
Public Declare Function SetWorkAntenna Lib "uhfreader.dll" (ByVal AntennaID As Integer) As Long
Public Declare Function Inventory Lib "uhfreader.dll" (ByVal Repeat As Byte, ByRef OutData As Byte, ByRef TagNum As Byte) As Integer
Public Declare Function CleanInventory Lib "uhfreader.dll" () As Integer

Public Declare Function rdy_read Lib "rfusbhid.dll" (ByVal MemBank As Byte, ByVal WordAdd As Byte, ByVal WordCnt As Byte, ByVal PassWord As Byte, ByRef TagCount As Byte, ByRef DataLen As Byte, _
        ByRef Data As Byte, ByRef ReadLen As Byte, ByRef AntID As Byte, ByRef ReadCount As Byte) As Integer



Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public co As ADODB.Connection
Public co1 As ADODB.Connection
Public co2 As ADODB.Connection
Public co3 As ADODB.Connection
Public rs As ADODB.Recordset
Public rs1 As ADODB.Recordset
Public rs2 As ADODB.Recordset
Public rs3 As ADODB.Recordset
Public rs4 As ADODB.Recordset
Public rs5 As ADODB.Recordset
Public rs6 As ADODB.Recordset
Public fhbc  As String
Public tmpcode3() As String
Public tmpcode2() As String
Public tmpcode1() As String
Public tmpcode() As String
Public tmp() As String
Public mdatabase As String
Public validtag As Boolean
Public shifta1 As Date
Public shifta2 As Date
Public shiftb1 As Date
Public shiftb2 As Date
Public shiftc1 As Date
Public shiftc2 As Date
Public ips(9) As String
Public paths(9) As String
Public unames(9) As String
Public passws(9) As String
Public wbcode As String
Public tmcode() As String
Public tmcode1() As String
Public arcode As String
Public username As String
Public machcode As String
Public camcode As Integer
Public boomcode As Integer
Public rfidcode As Integer
Public gpscode As Integer

Public datact As Integer

Public fsys  As New FileSystemObject
Public wstream As TextStream
Public COMPSAT As String
Public dbpass As String
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

Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const KEYEVENTF_KEYUP = &H2



Declare Function TerminateProcess _
Lib "kernel32" (ByVal ApphProcess As Long, _
ByVal uExitCode As Long) As Long
Declare Function OpenProcess Lib _
"kernel32" (ByVal dwDesiredAccess As Long, _
ByVal blnheritHandle As Long, _
ByVal dwAppProcessId As Long) As Long
Declare Function ProcessFirst _
Lib "kernel32" Alias "Process32First" _
(ByVal hSnapshot As Long, _
uProcess As PROCESSENTRY32) As Long
Declare Function ProcessNext _
Lib "kernel32" Alias "Process32Next" _
(ByVal hSnapshot As Long, _
uProcess As PROCESSENTRY32) As Long
Declare Function CreateToolhelpSnapshot _
Lib "kernel32" Alias "CreateToolhelp32Snapshot" _
(ByVal lFlags As Long, _
lProcessID As Long) As Long
Declare Function CloseHandle _
Lib "kernel32" (ByVal hObject As Long) As Long

Private Type LUID
lowpart As Long
highpart As Long
End Type

Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
LuidUDT As LUID
Attributes As Long
End Type

Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function GetVersion _
Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess _
Lib "kernel32" () As Long
Private Declare Function OpenProcessToken _
Lib "advapi32" (ByVal ProcessHandle As Long, _
ByVal DesiredAccess As Long, _
TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue _
Lib "advapi32" Alias "LookupPrivilegeValueA" _
(ByVal lpSystemName As String, _
ByVal lpName As String, _
lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges _
Lib "advapi32" (ByVal TokenHandle As Long, _
ByVal DisableAllPrivileges As Long, _
NewState As TOKEN_PRIVILEGES, _
ByVal BufferLength As Long, _
PreviousState As Any, _
ReturnLength As Any) As Long

Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szexeFile As String * MAX_PATH
End Type
'---------------------------------------
Public Function KillApp(myName As String) As Boolean
Const TH32CS_SNAPPROCESS As Long = 2&
Const PROCESS_ALL_ACCESS = 0
Dim uProcess As PROCESSENTRY32
Dim rProcessFound As Long
Dim hSnapshot As Long
Dim szExename As String
Dim exitCode As Long
Dim myProcess As Long
Dim AppKill As Boolean
Dim appCount As Integer
Dim i As Integer
On Local Error GoTo Finish
appCount = 0

uProcess.dwSize = Len(uProcess)
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
rProcessFound = ProcessFirst(hSnapshot, uProcess)
Do While rProcessFound
i = InStr(1, uProcess.szexeFile, Chr(0))
szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
If Right$(szExename, Len(myName)) = LCase$(myName) Then
KillApp = True
appCount = appCount + 1
myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
If KillProcess(uProcess.th32ProcessID, 0) Then
'For debug.... Remove this
End If

End If
rProcessFound = ProcessNext(hSnapshot, uProcess)
Loop
Call CloseHandle(hSnapshot)
Exit Function
Finish:
MsgBox "Error!"
End Function


Public Function SaveFormPic() As Picture
   cctv.SetFocus
   DoEvents
   DoEvents
   keybd_event vbKeyMenu, 0, 0, 0
   keybd_event vbKeySnapshot, 0, 0, 0
   DoEvents
   DoEvents
   keybd_event vbKeySnapshot, 0, KEYEVENTF_KEYUP, 0
   keybd_event vbKeyMenu, 0, KEYEVENTF_KEYUP, 0
   DoEvents
   DoEvents
   Set SaveFormPic = Clipboard.GetData(vbCFBitmap)
   DoEvents
   DoEvents
End Function


'Terminate any application and return an exit code to Windows.
Function KillProcess(ByVal hProcessID As Long, Optional ByVal exitCode As Long) As Boolean
Dim hToken As Long
Dim hProcess As Long
Dim tp As TOKEN_PRIVILEGES


If GetVersion() >= 0 Then

If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) = 0 Then
GoTo CleanUp
End If

If LookupPrivilegeValue("", "SeDebugPrivilege", tp.LuidUDT) = 0 Then
GoTo CleanUp
End If

tp.PrivilegeCount = 1
tp.Attributes = SE_PRIVILEGE_ENABLED

If AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, ByVal 0&) = 0 Then
GoTo CleanUp
End If
End If

hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)
If hProcess Then

KillProcess = (TerminateProcess(hProcess, exitCode) <> 0)
' close the process handle
CloseHandle hProcess
End If

If GetVersion() >= 0 Then
' under NT restore original privileges
tp.Attributes = 0
AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&

CleanUp:
If hToken Then CloseHandle hToken
End If

End Function


Public Function forceno(keya As Integer) As Integer
If keya <> 13 And keya <> 8 Then
If keya < 48 Or keya > 57 Then
forceno = 0
Exit Function
End If
End If
forceno = keya
End Function

Public Function timecheck(cuticode As String) As Boolean
If (Hour(Time) < 6 Or Hour(Time) > 17) And Left(cuticode, 1) <> "0" And Left(cuticode, 1) <> "7" And Left(cuticode, 1) <> "8" _
    And Trim(cuticode) <> "600013" And Trim(cuticode) <> "600037" And Trim(cuticode) <> "600039" And Trim(cuticode) <> "600059" And Trim(cuticode) <> "600183" And Trim(cuticode) <> "600221" And Trim(cuticode) <> "600227" And Trim(cuticode) <> "600248" And Trim(cuticode) <> "600423" And Trim(cuticode) <> "600425" _
    And Trim(cuticode) <> "600462" And Trim(cuticode) <> "600569" And Trim(cuticode) <> "600575" And Trim(cuticode) <> "600608" And Trim(cuticode) <> "600643" And Trim(cuticode) <> "600752" And Trim(cuticode) <> "600813" And Trim(cuticode) <> "601409" And Trim(cuticode) <> "601829" And Trim(cuticode) <> "602276" _
    And Trim(cuticode) <> "602298" And Trim(cuticode) <> "602382" And Trim(cuticode) <> "602774" And Trim(cuticode) <> "602963" And Trim(cuticode) <> "603000" And Trim(cuticode) <> "603094" And Trim(cuticode) <> "603115" And Trim(cuticode) <> "603227" And Trim(cuticode) <> "603228" And Trim(cuticode) <> "603277" _
    And Trim(cuticode) <> "603293" And Trim(cuticode) <> "603294" And Trim(cuticode) <> "603297" And Trim(cuticode) <> "603318" And Trim(cuticode) <> "603334" And Trim(cuticode) <> "603355" And Trim(cuticode) <> "603445" And Trim(cuticode) <> "603468" And Trim(cuticode) <> "603518" And Trim(cuticode) <> "603575" _
    And Trim(cuticode) <> "603588" And Trim(cuticode) <> "603597" And Trim(cuticode) <> "603718" And Trim(cuticode) <> "603745" And Trim(cuticode) <> "603778" And Trim(cuticode) <> "603881" And Trim(cuticode) <> "603885" And Trim(cuticode) <> "603893" And Trim(cuticode) <> "603956" And Trim(cuticode) <> "603979" _
    And Trim(cuticode) <> "603982" And Trim(cuticode) <> "603983" And Trim(cuticode) <> "603996" And Trim(cuticode) <> "603998" And Trim(cuticode) <> "604012" And Trim(cuticode) <> "604032" And Trim(cuticode) <> "604037" And Trim(cuticode) <> "604057" And Trim(cuticode) <> "604068" And Trim(cuticode) <> "604117" _
    And Trim(cuticode) <> "604127" And Trim(cuticode) <> "604134" And Trim(cuticode) <> "604146" And Trim(cuticode) <> "604151" And Trim(cuticode) <> "604152" And Trim(cuticode) <> "604153" And Trim(cuticode) <> "604196" And Trim(cuticode) <> "604390" And Trim(cuticode) <> "604432" And Trim(cuticode) <> "604473" _
    And Trim(cuticode) <> "604484" And Trim(cuticode) <> "604530" And Trim(cuticode) <> "604555" And Trim(cuticode) <> "604558" And Trim(cuticode) <> "604560" And Trim(cuticode) <> "604576" And Trim(cuticode) <> "604585" And Trim(cuticode) <> "604596" And Trim(cuticode) <> "604614" And Trim(cuticode) <> "604622" _
    And Trim(cuticode) <> "604638" And Trim(cuticode) <> "604640" And Trim(cuticode) <> "604641" And Trim(cuticode) <> "604648" And Trim(cuticode) <> "604650" And Trim(cuticode) <> "604651" And Trim(cuticode) <> "604655" And Trim(cuticode) <> "604660" And Trim(cuticode) <> "604148" And Trim(cuticode) <> "604184" _
    And Trim(cuticode) <> "600860" And Trim(cuticode) <> "603243" And Trim(cuticode) <> "603415" And Trim(cuticode) <> "603910" And Trim(cuticode) <> "604539" And Trim(cuticode) <> "604552" And Trim(cuticode) <> "604553" And Trim(cuticode) <> "604557" And Trim(cuticode) <> "604561" And Trim(cuticode) <> "604562" _
    And Trim(cuticode) <> "604187" And Trim(cuticode) <> "600819" And Trim(cuticode) <> "604586" And Trim(cuticode) <> "604090" And Trim(cuticode) <> "603607" And Trim(cuticode) <> "601049" And Trim(cuticode) <> "604296" And Trim(cuticode) <> "604047" And Trim(cuticode) <> "604239" And Trim(cuticode) <> "603854" _
    And Trim(cuticode) <> "603884" And Trim(cuticode) <> "603076" And Trim(cuticode) <> "604095" And Trim(cuticode) <> "603600" And Trim(cuticode) <> "603500" And Trim(cuticode) <> "604085" And Trim(cuticode) <> "603751" And Trim(cuticode) <> "600296" And Trim(cuticode) <> "604625" And Trim(cuticode) <> "603935" _
    And Trim(cuticode) <> "604624" And Trim(cuticode) <> "604549" And Trim(cuticode) <> "604017" And Trim(cuticode) <> "604623" And Trim(cuticode) <> "600013" And Trim(cuticode) <> "600175" And Trim(cuticode) <> "603858" And Trim(cuticode) <> "603432" And Trim(cuticode) <> "604620" And Trim(cuticode) <> "603481" _
    And Trim(cuticode) <> "603766" And Trim(cuticode) <> "604402" And Trim(cuticode) <> "603539" And Trim(cuticode) <> "603549" And Trim(cuticode) <> "604361" And Trim(cuticode) <> "602452" And Trim(cuticode) <> "603773" And Trim(cuticode) <> "604361" And Trim(cuticode) <> "600248" And Trim(cuticode) <> "602452" _
    And Trim(cuticode) <> "603094" And Trim(cuticode) <> "603773" And Trim(cuticode) <> "604361" And Trim(cuticode) <> "600248" And Trim(cuticode) <> "602452" And Trim(cuticode) <> "603094" And Trim(cuticode) <> "603773" And Trim(cuticode) <> "603277" And Trim(cuticode) <> "600569" And Trim(cuticode) <> "600575" _
    And Trim(cuticode) <> "600813" And Trim(cuticode) <> "603227" And Trim(cuticode) <> "604638" And Trim(cuticode) <> "604650" And Trim(cuticode) <> "603893" And Trim(cuticode) <> "300030" And Trim(cuticode) <> "300031" _
Then
    timecheck = False
Else
    timecheck = True
End If
End Function


Public Function Conn()
demo = False

On Error Resume Next
  fhbc = &H404040
  mdatabase = App.Path + "\wajan.mdb"
   Set co = New ADODB.Connection
   If co.State = 1 Then co.Close
   co.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)}; Dbq=" & App.Path & "\safe.mdb; Uid=Admin; Pwd=rana@safe@123"
   co.Open
    Set rs1 = New ADODB.Recordset
    rs1.Open "Select * from dbpass", co, adOpenKeyset, adLockOptimistic
    rs1.MoveFirst
    dbpass = Textdcrt(rs1.Fields("dbpass"))
    rs1.Close
    co.Close
   
   Set co = New ADODB.Connection
   If co.State = 1 Then co.Close
   
   co.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)}; Dbq=" & App.Path & "\wajan.mdb; Uid=Admin; Pwd=" & dbpass
   co.Open
End Function

Public Function Conn1()
   Set co1 = New ADODB.Connection
   If co1.State = 1 Then co1.Close
   co1.ConnectionString = "PROVIDER = MSDASQL;driver={SQL Server};database=weighthq; server=172.20.0.96,1433;uid=bccl;pwd=Edesp@#123;"
   co1.Open
End Function

Public Function Conn3()
   Set co3 = New ADODB.Connection
   If co3.State = 1 Then co3.Close
   co3.ConnectionString = "DSN=BCCLTESTDB;Pwd=test321#;"
   co3.Open
End Function

Public Function Conn2()
   Set co2 = New ADODB.Connection
   If co2.State = 1 Then co2.Close
   co2.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)}; Dbq=" & App.Path & "\safe.mdb; Uid=Admin; Pwd=rana@safe@123"
   co2.Open
End Function

Public Function IncrBillNo11(BillNo As String) As String
    Dim intBillNo As Long
    Dim BillLen As Integer
    intBillNo = Val(Right(BillNo, 2))
    intBillNo = intBillNo + 1
    BillLen = Len(CStr(intBillNo))
    IncrBillNo11 = String(2 - BillLen, "0") & CStr(intBillNo)
End Function

Public Sub Dither(vForm As Form, Border As Boolean)
    Dim intLoop As Integer
    On Error Resume Next
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256

    For intLoop = 0 To 255
       ' vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(10, _
        IIf(240 - intLoop > 100, 240 - intLoop, 40), _
        IIf(245 - intLoop > 100, 245 - intLoop, 80)), B
       
       'Purple Shade
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), _
       RGB(255 - intLoop / 4, 255 - intLoop / 1.5, 255 - intLoop / 8), B

       'Yellow Shade
      '''  vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), _
     '''   RGB(255 - intLoop / 8, 255 - intLoop / 4, 255 - intLoop / 2), B
       
       'Dark Blue Shade
       ' vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), _
       ' RGB(255 - intLoop, 255 - intLoop / 1.3, 255 - intLoop / 16), B
        
       'Green Shade
        '''vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), _
        '''RGB(255 - intLoop / 2, 255 - intLoop / 8, 255 - intLoop / 6), B
    Next intLoop
    
    If Border = True Then vForm.Line (0, 0)-(vForm.ScaleWidth, vForm.ScaleHeight), RGB(0, 0, 0), B
End Sub



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


Public Function IncrBillNo1(BillNo As String) As String
    Dim intBillNo As Double
    Dim BillLen As Integer
    intBillNo = Val(BillNo)
    intBillNo = intBillNo + 1
    BillLen = Len(CStr(intBillNo))
    IncrBillNo1 = CStr(intBillNo)
End Function

Public Function IncrBillNo4(BillNo As String) As String
    Dim intBillNo As Double
    Dim BillLen As Integer
    intBillNo = Val(Right(BillNo, 4))
    intBillNo = intBillNo + 1
    BillLen = Len(CStr(intBillNo))
    IncrBillNo4 = String(4 - BillLen, "0") & CStr(intBillNo)
End Function

Public Function IncrBillNo(BillNo As String) As String
    Dim intBillNo As Double
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





Function word1(a As Integer) As String
Dim b  As Integer
word1 = ""
If a = 1 Then
    word1 = word1 + "One "
    ElseIf a = 2 Then
    word1 = word1 + "Two "
    ElseIf a = 3 Then
    word1 = word1 + "Three "
    ElseIf a = 4 Then
    word1 = word1 + "Four "
    ElseIf a = 5 Then
    word1 = word1 + "Five "
    ElseIf a = 6 Then
    word1 = word1 + "Six "
    ElseIf a = 7 Then
    word1 = word1 + "Seven "
    ElseIf a = 8 Then
    word1 = word1 + "Eight "
    ElseIf a = 9 Then
    word1 = word1 + "Nine "
End If
End Function

Function word2(a As Integer) As String
Dim b As Integer
word2 = ""
If a > 9 Then
If a < 20 Then
    b = a
    If b = 10 Then
        word2 = word2 + "Ten "
    ElseIf b = 11 Then
        word2 = word2 + "Eleven "
    ElseIf b = 12 Then
        word2 = word2 + "Twelve "
    ElseIf b = 13 Then
        word2 = word2 + "Thirteen "
    ElseIf b = 14 Then
        word2 = word2 + "Fourteen "
    ElseIf b = 15 Then
        word2 = word2 + "Fifteen "
    ElseIf b = 16 Then
        word2 = word2 + "Sixteen "
    ElseIf b = 17 Then
        word2 = word2 + "Seventeen "
    ElseIf b = 18 Then
        word2 = word2 + "Eighteen "
    ElseIf b = 19 Then
        word2 = word2 + "Nineteen "
    End If
    a = -1
Else
    b = CInt(Left(CStr(a), 1))
    If b = 2 Then
        word2 = word2 + "Twenty "
    ElseIf b = 3 Then
        word2 = word2 + "Thirty "
    ElseIf b = 4 Then
        word2 = word2 + "Forty "
    ElseIf b = 5 Then
        word2 = word2 + "Fifty "
    ElseIf b = 6 Then
        word2 = word2 + "Sixty "
    ElseIf b = 7 Then
        word2 = word2 + "Seventy "
    ElseIf b = 8 Then
        word2 = word2 + "Eighty "
    ElseIf b = 9 Then
        word2 = word2 + "Ninety "
    End If
    a = CInt(Right(CStr(a), 1))
    word2 = word2 + word1(a)
End If
Else
word2 = word2 + word1(a)
End If
End Function

Function word3(a As Integer) As String
Dim b As Integer
word3 = ""
If a > 99 Then
    b = CInt(Left(CStr(a), 1))
    word3 = word3 + word1(b) + "Hundred "
    a = CInt(Mid(CStr(a), 2, 2))
End If
word3 = word3 + word2(a)
End Function

Function word5(a As Long) As String
word5 = ""
Dim b As Integer
If a > 9999 Then
b = CInt(Left(CStr(a), 2))
word5 = word5 + word2(CInt(b)) + "Thousand "
ElseIf a > 999 Then
b = CInt(Left(CStr(a), 1))
word5 = word5 + word1(CInt(b)) + "Thousand "
End If
a = CInt(Right(CStr(a), 3))
word5 = word5 + word3(CInt(a))
End Function

Function word7(a As Long) As String
word7 = ""
Dim b As Integer
If a > 999999 Then
b = CInt(Left(CStr(a), 2))
word7 = word7 + word2(CInt(b)) + "Lakh "
ElseIf a > 99999 Then
b = CInt(Left(CStr(a), 1))
word7 = word7 + word1(CInt(b)) + "Lakh "
End If
a = CLng(Right(CStr(a), 5))
word7 = word7 + word5(CLng(a))
End Function

Public Function worda(a As Long) As String
worda = ""
Dim b As Long
If a <> 0 Then
If a > 9999999 Then
b = CLng(Left(CStr(a), Len(CStr(a)) - 7))
worda = worda + word7(CInt(b)) + "Crore "
End If
a = CLng(Right(CStr(a), 7))
worda = worda + word7(CLng(a))
Else
worda = "Zero "
End If
End Function


