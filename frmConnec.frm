VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmConnec 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3675
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash faWait 
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3675
      _cx             =   6482
      _cy             =   926
      FlashVars       =   ""
      Movie           =   "c:\work\demo\HmXP9000\wait.swf"
      Src             =   "c:\work\demo\HmXP9000\wait.swf"
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
   End
End
Attribute VB_Name = "frmConnec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmConnec.Left = (Screen.Width - frmConnec.Width) / 2
frmConnec.Top = (Screen.Height - frmConnec.Height) / 2

End Sub
