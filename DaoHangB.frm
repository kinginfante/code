VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form DaoHangB 
   Caption         =   "Form2"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   9915
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   7635
      Left            =   -30
      TabIndex        =   0
      Top             =   -60
      Width           =   10275
      _cx             =   18124
      _cy             =   13467
      FlashVars       =   ""
      Movie           =   "c:\work\demo\HMXP9000\flash\daohangb.swf"
      Src             =   "c:\work\demo\HMXP9000\flash\daohangb.swf"
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
      SeamlessTabbing =   -1  'True
   End
End
Attribute VB_Name = "DaoHangB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
DaoHangB.Height = 7830
DaoHangB.Width = 10035
End Sub

Private Sub ShockwaveFlash1_FSCommand(ByVal command As String, ByVal args As String)
Select Case command
Case "Back"
MDI.Visible = False
mod1.Fir = False
Form1.Show
Form1.Fa1.GotoFrame (160)
End Select

End Sub

