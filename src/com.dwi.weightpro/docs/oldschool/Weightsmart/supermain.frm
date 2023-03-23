VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "SUPER ADMIN"
   ClientHeight    =   7425
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   11805
   LinkTopic       =   "MDIForm1"
   Picture         =   "supermain.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu cld 
      Caption         =   "        CHANGE LOGIN DETAILS        "
   End
   Begin VB.Menu cdp 
      Caption         =   "        CHANGE DATABASE PASSWORD        "
   End
   Begin VB.Menu quit 
      Caption         =   "        QUIT        "
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cdp_Click()
changedbpass.Show
End Sub

Private Sub cld_Click()
CHANGELOGIN.Show
End Sub

Private Sub quit_Click()
Unload Me
End Sub
