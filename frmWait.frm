VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Begin VB.Form frmWait 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   525
   ClientLeft      =   3150
   ClientTop       =   2985
   ClientWidth     =   3690
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash faWait 
      Height          =   525
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   3675
      _cx             =   6482
      _cy             =   926
      FlashVars       =   ""
      Movie           =   "c:\work\demo\HmXP9000\wait.swf"
      Src             =   "c:\work\demo\HmXP9000\wait.swf"
      WMode           =   "Window"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmWait.Left = (Screen.Width - frmWait.Width) / 2
frmWait.Top = (Screen.Height - frmWait.Height) / 2
faWait.Play
End Sub

