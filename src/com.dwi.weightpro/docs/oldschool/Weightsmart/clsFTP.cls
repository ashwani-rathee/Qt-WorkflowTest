VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------------------------------------------------------------------------------
' This class is based on the SimpleFTP VB example by Microsoft.
' It was extended by Michael Glaser to be class based and support buffer
' based uploads and downloads with a progress event.
'
' If you found this code useful and would like to support the author, please
' visit the eD.I.Y. Software website at http://www.ediy.co.nz to see if the
' products we have available would be useful to you or your customers.
'
' Please credit me if you use this code in your applications.
'
' If you have any questions or possible improvements to this code, email me: mike@ediy.co.nz
'
' For help on any of the class API functions, an excellent reference is available here:
' http://msdn.microsoft.com/library/default.asp?url=/workshop/networking/wininet/reference/win32_ref_entry.asp
' ------------------------------------------------------------------------------------------------------------------------------
 
 
Private Const MAX_PATH = 260
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const NO_ERROR = 0
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_OFFLINE = &H1000
Private Const INTERNET_FLAG_PASSIVE = &H8000000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
 
Private Const ERROR_NO_MORE_FILES = 18
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Private Const INTERNET_OPTION_CONNECT_TIMEOUT = 5000
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_INVALID_PORT_NUMBER = 0
Private Const INTERNET_SERVICE_FTP = 1
 
Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const FTP_TRANSFER_TYPE_BINARY = &H2
 
Public Enum eFTPTransferTypes
    FTP_ASCII = FTP_TRANSFER_TYPE_ASCII
    FTP_BINARY = FTP_TRANSFER_TYPE_BINARY
End Enum
 
Private Const rDayZeroBias As Double = 109205#   ' Abs(CDbl(#01-01-1601#))
Private Const rMillisecondPerDay As Double = 10000000# * 60# * 60# * 24# / 10000#
 
Private Const BUFFERSIZE = 255
 
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As Currency
    ftLastAccessTime As Currency
    ftLastWriteTime As Currency
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
 
' -- private functions
 
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As Any) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function FTPGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszOldName As String, ByVal lpszNewName As String) As Boolean
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszName As String) As Boolean
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszName As String) As Boolean
 
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal sBuff As String, ByVal Access As Long, ByVal Flags As Long, ByVal Context As Long) As Long
Private Declare Function FTPPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Boolean
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
 
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpLibFileName As String) As Long
 
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Private Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToWrite As Long, dwNumberOfBytesWritten As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToRead As Long, dwNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (ByRef lpdwError As Long, ByVal lpszErrorBuffer As String, ByRef lpdwErrorBufferLength As Long) As Boolean
Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
 
 
 
' -- Private Variables
 
Private hOpen As Long
Private hConnection As Long
Private hFile As Long
Private dwType As Long
Private dwSeman As Long
 
Private szErrorMessage As String, szSimpleErrorMessage As String
 
'Private mDirCol As New cDirList
 
 
Public Event FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)
 
 
'Property Get Directory() As cDirList
'   Set Directory = mDirCol
'End Property
 
Property Get LastErrorMessage() As String
LastErrorMessage = szErrorMessage
End Property
 
Property Get SimpleLastErrorMessage() As String
SimpleLastErrorMessage = szSimpleErrorMessage
End Property
 
Public Function OpenConnection(sServer As String, sUser As String, sPassword As String) As Boolean
If hConnection <> 0 Then
    InternetCloseHandle hConnection
End If
 
If CBool(InternetAutodial(INTERNET_AUTODIAL_FORCE_ONLINE, 0)) Then
    hOpen = InternetOpen("eDIY FTP Client", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hOpen = 0 Then
        ErrorOut Err.LastDllError, "InternetOpen"
    End If
    'InternetSetStatusCallback hOpen, AddressOf FTPCallBack
     
    hConnection = InternetConnect(hOpen, sServer, INTERNET_INVALID_PORT_NUMBER, sUser, sPassword, INTERNET_SERVICE_FTP, dwSeman, 0)
    If hConnection = 0 Then
        ErrorOut Err.LastDllError, "InternetConnect"
        OpenConnection = False
        Exit Function
    Else
        'InternetSetStatusCallback hConnection, AddressOf FTPCallBack
        OpenConnection = True
    End If
Else
    OpenConnection = False
End If
End Function
 
 
Public Sub CloseConnection()
 
If hConnection Then
    InternetCloseHandle hConnection
End If
hConnection = 0
 
If hOpen Then
    InternetCloseHandle hOpen
End If
hOpen = 0
 
End Sub
 
Public Function SimpleFTPPutFile(sLocal As String, sRemote As String) As Boolean
If (FTPPutFile(hConnection, sLocal, sRemote, dwType, 0) = False) Then
    ErrorOut Err.LastDllError, "SimpleFtpPutFile"
    SimpleFTPPutFile = False
    Exit Function
Else
    SimpleFTPPutFile = True
End If
End Function
 
Public Function RenameFTPFile(sExisting As String, sNewName As String) As Boolean
If (FtpRenameFile(hConnection, sExisting, sNewName) = False) Then
    ErrorOut Err.LastDllError, "RenameFTPFile"
    RenameFTPFile = False
    Exit Function
Else
    RenameFTPFile = True
End If
End Function
 
Public Function CreateFTPDirectory(sDirectory As String) As Boolean
If (FtpCreateDirectory(hConnection, sDirectory) = False) Then
    ErrorOut Err.LastDllError, "CreateFTPDirectory"
    CreateFTPDirectory = False
    Exit Function
Else
    CreateFTPDirectory = True
End If
End Function
 
Public Function RemoveFTPDirectory(sDirectory As String) As Boolean
If (FtpRemoveDirectory(hConnection, sDirectory) = False) Then
    ErrorOut Err.LastDllError, "RemoveFTPDirectory"
    RemoveFTPDirectory = False
    Exit Function
Else
    RemoveFTPDirectory = True
End If
End Function
 
Public Function DeleteFTPFile(sRemote As String) As Boolean
If (FtpDeleteFile(hConnection, sRemote) = False) Then
    ErrorOut Err.LastDllError, "DeleteFTPFile"
    DeleteFTPFile = False
    Exit Function
Else
    DeleteFTPFile = True
End If
End Function
 
Public Function FTPUploadFile(sLocal As String, sRemote As String) As Boolean
Dim Data(BUFFERSIZE - 1) As Byte
Dim Written As Long
Dim Size As Long
Dim Sum As Long
Dim lBlock As Long
Dim f As Integer
 
Sum = 0
lBlock = 0
sLocal = Trim$(sLocal)
sRemote = Trim$(sRemote)
 
f = FreeFile()
 
On Error GoTo EH
 
If CBool(LenB(sLocal)) And CBool(LenB(sRemote)) Then
    hFile = FtpOpenFile(hConnection, sRemote, GENERIC_WRITE, dwType, 0)
    If hFile = 0 Then
        ErrorOut Err.LastDllError, "FtpOpenFile:PutFile"
        FTPUploadFile = False
        Exit Function
    End If
     
    Open sLocal For Binary Access Read As #f
    Size = LOF(1)
    For lBlock = 1 To Size \ BUFFERSIZE
        Get #f, , Data
        If (InternetWriteFile(hFile, Data(0), BUFFERSIZE, Written) = 0) Then
            ErrorOut Err.LastDllError, "InternetWriteFile"
            GoTo EH
        End If
        DoEvents
        Sum = Sum + BUFFERSIZE
         
        RaiseEvent FileTransferProgress(Sum, Size)
    Next lBlock
     
    'check for leftovers
    If Size Mod BUFFERSIZE <> 0 Then
        Get #f, , Data
        If (InternetWriteFile(hFile, Data(0), Size Mod BUFFERSIZE, Written) = 0) Then
            ErrorOut Err.LastDllError, "InternetWriteFile2"
            GoTo EH
        End If
    End If
     
    Sum = Size
    Close #f
    RaiseEvent FileTransferProgress(Sum, Size)
    InternetCloseHandle hFile
    FTPUploadFile = True
End If
 
Exit Function
EH:
FTPUploadFile = False
Close #f
End Function
 
Public Function FTPDownloadFile(sLocal As String, sRemote As String) As Boolean
Dim Data(BUFFERSIZE - 1) As Byte ' array of 100 elements 0 to 99
Dim Written As Long
Dim Size As Long
Dim Sum As Long
Dim lBlock As Long
Dim f As Integer
 
FTPDownloadFile = False
 
Sum = 0
lBlock = 0
f = FreeFile()
 
sLocal = Trim$(sLocal)
sRemote = Trim$(sRemote)
 
If CBool(LenB(sLocal)) And CBool(LenB(sRemote)) Then
    Size = GetFTPFileSize(sRemote)
     
    If Size > 0 Then
         
        hFile = FtpOpenFile(hConnection, sRemote, GENERIC_READ, dwType, 0)
         
        If hFile = 0 Then
            ErrorOut Err.LastDllError, "FtpOpenFile:GetFile"
            Exit Function
        End If
         
        Open sLocal For Binary Access Write As #f
        Seek #f, 1
        Sum = 1
         
        For lBlock = 1 To Size \ BUFFERSIZE
            If (InternetReadFile(hFile, Data(0), BUFFERSIZE, Written) = 0) Then
                ErrorOut Err.LastDllError, "InternetReadFile"
                GoTo EH
            End If
            Put #f, , Data
            DoEvents
            Sum = Sum + BUFFERSIZE
            RaiseEvent FileTransferProgress(Sum, Size)
        Next lBlock
         
        'Check for leftovers
        If Size Mod BUFFERSIZE <> 0 Then
            ReDim Data2((Size Mod BUFFERSIZE) - 1) As Byte
            If (InternetReadFile(hFile, Data2(0), Size Mod BUFFERSIZE, Written) = 0) Then
                ErrorOut Err.LastDllError, "InternetReadFile2"
                GoTo EH
            End If
        End If
         
        Put #f, , Data2
        Close #f
         
        Sum = Size
        RaiseEvent FileTransferProgress(Sum, Size)
        InternetCloseHandle hFile
        FTPDownloadFile = True
         
    Else
        ErrorOut -1, "FTPDownloadFile", "FTP File Doesn't Exist"
    End If
     
End If
 
Exit Function
EH:
Close #f
End Function
 
Public Function SimpleFTPGetFile(sLocal As String, sRemote As String) As Boolean
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
If (FTPGetFile(hConnection, sRemote, sLocal, False, FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
    ErrorOut Err.LastDllError, "SimpleFtpGetFile"
    SimpleFTPGetFile = False
    Exit Function
Else
    SimpleFTPGetFile = True
End If
End Function
 
Public Property Get FTPDirectory() As String
Dim szDir As String
 
szDir = String(1024, Chr$(0))
 
If (FtpGetCurrentDirectory(hConnection, szDir, 1024) = False) Then
    ErrorOut Err.LastDllError, "FtpGetCurrentDirectory"
    Exit Property
Else
    FTPDirectory = Left(szDir, InStr(1, szDir, String(1, 0), vbBinaryCompare) - 1)
End If
 
End Property
 
Public Property Let FTPDirectory(sDir As String)
 
If (FtpSetCurrentDirectory(hConnection, sDir) = False) Then
    ErrorOut Err.LastDllError, "FtpSetCurrentDirectory"
End If
 
End Property
 
Public Function GetFTPFileSize(sFile As String) As Long
Dim szDir As String
Dim hFind As Long
Dim nLastError As Long
Dim pData As WIN32_FIND_DATA
 
hFind = FtpFindFirstFile(hConnection, Replace(sFile, " ", "?"), pData, 0, 0)
nLastError = Err.LastDllError
If hFind = 0 Then
    If (nLastError = ERROR_NO_MORE_FILES) Then
        GetFTPFileSize = -1  ' File not found
    Else
        GetFTPFileSize = -2  ' Other error
        ErrorOut Err.LastDllError, "FtpFindFirstFile"
    End If
    Exit Function
End If
 
GetFTPFileSize = pData.nFileSizeLow
InternetCloseHandle hFind
End Function
 
'Public Function GetDirectoryListing(sFilter As String) As cDirList
'    Dim szDir As String
'    Dim hFind As Long
'    Dim nLastError As Long
'    Dim dError As Long
'    Dim ptr As Long
'    Dim pData As WIN32_FIND_DATA
'    Dim sFilename As String
'
'    Set mDirCol = Nothing
'    hFind = FtpFindFirstFile(hConnection, sFilter, pData, 0, 0)
'    nLastError = Err.LastDllError
'    If hFind = 0 Then
'        If (nLastError <> ERROR_NO_MORE_FILES) Then
'            ErrorOut Err.LastDllError, "FtpFindFirstFile"
'        End If
'        Exit Function
'    End If
'
'    dError = NO_ERROR
'    Dim bRet As Boolean
'
'    sFilename = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
'    mDirCol.Add pData.dwFileAttributes, Win32ToVbTime(pData.ftCreationTime), Win32ToVbTime(pData.ftLastAccessTime), Win32ToVbTime(pData.ftLastWriteTime), pData.nFileSizeLow, sFilename
'    Do
'        pData.cFileName = String(MAX_PATH, 0)
'        bRet = InternetFindNextFile(hFind, pData)
'        If Not bRet Then
'            dError = Err.LastDllError
'            If dError = ERROR_NO_MORE_FILES Then
'                Exit Do
'            Else
'                ErrorOut Err.LastDllError, "InternetFindNextFile"
'                InternetCloseHandle (hFind)
'                Exit Function
'            End If
'        Else
'            sFilename = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
'            mDirCol.Add pData.dwFileAttributes, Win32ToVbTime(pData.ftCreationTime), Win32ToVbTime(pData.ftLastAccessTime), Win32ToVbTime(pData.ftLastWriteTime), pData.nFileSizeLow, sFilename
'        End If
'    Loop
'
'    Set GetDirectoryListing = mDirCol
'    InternetCloseHandle (hFind)
'End Function
 
Public Sub SetTransferType(ByVal vType As eFTPTransferTypes)
dwType = vType
End Sub
 
Public Sub SetMode(ByVal bActive As Boolean)
If bActive Then
    dwSeman = 0
Else
    dwSeman = INTERNET_FLAG_PASSIVE
End If
End Sub
 
' -- Private Functions
Private Sub ErrorOut(ByVal dwError As Long, ByVal szFunc As String, _
    Optional ByVal sCustError As String)
 
Dim dwRet As Long
Dim dwTemp As Long
Dim szString As String * 2048
Dim i As Integer
 
If dwError <> -1 Then
    dwRet = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, _
        GetModuleHandle("wininet.dll"), dwError, 0, _
        szString, 256, 0)
Else
    szString = sCustError
End If
 
 
i = InStr(1, szString, vbNullChar)
If i Then
    szString = Left$(szString, i - 1)
End If
i = InStr(1, szString, vbNewLine)
If i Then
    szString = Replace$(szString, vbNewLine, vbNullString)
End If
 
szErrorMessage = szFunc & "() Error - Code: " & dwError & " Message: " & Trim$(szString)
szSimpleErrorMessage = Trim$(szString)
 
 
If dwError = 12003 Then
    'Extended error information was returned
    dwRet = InternetGetLastResponseInfo(dwTemp, szString, 2048)
    szErrorMessage = szString
End If
 
End Sub
 
'Private Function Win32ToVbTime(ft As Currency) As Date
'Dim ftl As Currency
'' Call API to convert from UTC time to local time
'If FileTimeToLocalFileTime(ft, ftl) Then
'    ' Local time is nanoseconds since 01-01-1601
'    ' In Currency that comes out as milliseconds
'    ' Divide by milliseconds per day to get days since 1601
'    ' Subtract days from 1601 to 1899 to get VB Date equivalent
'    Win32ToVbTime = CDate((ftl / rMillisecondPerDay) - rDayZeroBias)
'Else
'    MsgBox Err.LastDllError
'End If
'End Function
 
Private Sub Class_Initialize()
SetTransferType FTP_ASCII 'FTP_BINARY
 
SetMode True
 
hOpen = 0
hConnection = 0
End Sub
 
Private Sub Class_Terminate()
CloseConnection
End Sub
 
 
 
 
' -- If anyone can get the wininet.dll to call the FTPCallback function for
' -- status updates (in a public module), please email mike@ediy.co.nz.
 
'Public Declare Function InternetSetStatusCallback Lib "wininet.dll" (ByVal hInternetSession As Long, ByVal lpfnCallBack As Long) As Long
'
'Public Function FTPCallBack(ByVal hInternet As Long, ByVal dwContext As Long, ByVal dwInternetStatus As Long, ByVal lpvStatusInformation As Long, ByVal dwStatusInformationLength As Long) As Long
'   Debug.Print "Status: " & dwInternetStatus
'End Function
'
'Public Const INTERNET_STATUS_RESOLVING_NAME = 10
'Public Const INTERNET_STATUS_NAME_RESOLVED = 11
'Public Const INTERNET_STATUS_CONNECTING_TO_SERVER = 20
'Public Const INTERNET_STATUS_CONNECTED_TO_SERVER = 21
'Public Const INTERNET_STATUS_SENDING_REQUEST = 30
'Public Const INTERNET_STATUS_REQUEST_SENT = 31
'Public Const INTERNET_STATUS_RECEIVING_RESPONSE = 40
'Public Const INTERNET_STATUS_RESPONSE_RECEIVED = 41
'Public Const INTERNET_STATUS_CTL_RESPONSE_RECEIVED = 42
'Public Const INTERNET_STATUS_PREFETCH = 43
'Public Const INTERNET_STATUS_CLOSING_CONNECTION = 50
'Public Const INTERNET_STATUS_CONNECTION_CLOSED = 51
'Public Const INTERNET_STATUS_HANDLE_CREATED = 60
'Public Const INTERNET_STATUS_HANDLE_CLOSING = 70
'Public Const INTERNET_STATUS_REQUEST_COMPLETE = 100
'Public Const INTERNET_STATUS_REDIRECT = 110
'Public Const INTERNET_STATUS_INTERMEDIATE_RESPONSE = 120
'Public Const INTERNET_STATUS_STATE_CHANGE = 200
 


