VERSION 5.00
Object = "{66FF217E-12A8-45F9-8627-D9289E6943EB}#1.0#0"; "webrec.ocx"
Begin VB.Form cctv 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CCTV Camera"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2860
      Width           =   7335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7335
   End
   Begin DHSURVEILLANCECTRLLib.DHSurveillanceCtrl DHSurveillanceCtrl1 
      Height          =   5745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _Version        =   65536
      _ExtentX        =   12938
      _ExtentY        =   10134
      _StockProps     =   0
      lVideoWindNum   =   4
      SetLanguage     =   0
      SetHostPort     =   8000
      SetLangFromIP   =   "http://192.168.1.240"
      VideoWindBGColor=   ""
      VideoWindBarColor=   ""
      VideoWindTextColor=   ""
      VideoWindBorderColor=   ""
      VideoWindPlayRegionBGColor=   ""
      VideoWindControlRegionBGColor=   ""
      IsShowPreview   =   0
      IsShowWndBtn    =   0
   End
End
Attribute VB_Name = "cctv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Dim ax, bx As Boolean
If Trim(ips(1)) <> "" And IsNumeric(paths(1)) And Trim(unames(1)) <> "" And Trim(passws(1)) <> "" Then
ax = DHSurveillanceCtrl1.LoginDevice(ips(1), paths(1), unames(1), passws(1))
bx = DHSurveillanceCtrl1.ConnectAllChannle
Else
    MsgBox "Please fill camera IP, Path, Username, Password properly"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
DHSurveillanceCtrl1.DisConnectAllChannel
DHSurveillanceCtrl1.LogoutDevice
End Sub
