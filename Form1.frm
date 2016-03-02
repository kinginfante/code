VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10e.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "您好"
   ClientHeight    =   6240
   ClientLeft      =   7980
   ClientTop       =   4110
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Fa1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8955
      _cx             =   15796
      _cy             =   11033
      FlashVars       =   ""
      Movie           =   " c:\work\demo\hmxp9000\newHm.swf"
      Src             =   " c:\work\demo\hmxp9000\newHm.swf"
      WMode           =   "Window"
      Play            =   "-1"
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, _
ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long

Private hbrush As Long, hdc5 As Long

Private Sub Fa1_FSCommand(ByVal command As String, ByVal args As String)



If command = "mxc" Then
    End
ElseIf command = "wb" Then
   mod1.Lb = command
ElseIf command = "xz" Then
    mod1.Lb = command
ElseIf command = "fw" Then
   mod1.Lb = command
ElseIf command = "gc" Then
   mod1.Lb = command
ElseIf command = "cw" Then
   mod1.Lb = command

ElseIf command = "gl" Then
   mod1.Lb = command
ElseIf command = "login" Then
   frmWait.Show
   frmWait.ZOrder
   frmWait.faWait.Play
   frmLogin.Show
   Form1.Enabled = False
End If
'frmWait.Visible = False
'MsgBox command
End Sub

Private Sub Form_Load()
Dim dx As Long, dy As Long
Dim rx1 As Long, rx2 As Long, ry1 As Long, ry2 As Long
Dim i As Long, j As Long, bcolor As Long
Dim DispCnt As Long
Fa1.Movie = "c:\work\demo\hmxp9000\newHm.swf"

Form1.Left = (Screen.Width - Form1.Width) / 2
Form1.Top = (Screen.Height - Form1.Height) / 2
'If mod1.Fir = True Then
DispCnt = 80 ' 注释：一共Display多少次榘形後才显示Form
hdc5 = GetDC(0)
bcolor = GetBkColor(Me.hdc) '注释：取得form的背景色

'注释：注：之所以不使用me.BackColor的原因是：这个属性不一定使用调色盘，
'注释： 如果使用系统配色，那结果会不对
hbrush = CreateSolidBrush(bcolor) ' 注释：设定笔刷颜色
Call SelectObject(hdc5, hbrush)
dx = Me.Width \ (DispCnt * 3)
dy = Me.Height \ (DispCnt * 4)
j = 1
For i = DispCnt To 1 Step -1
rx1 = (Me.Left + dx * (i - 1)) \ Screen.TwipsPerPixelX
ry1 = (Me.Top + dy * (i - 1)) \ Screen.TwipsPerPixelY
rx2 = rx1 + dx * 2 * j \ Screen.TwipsPerPixelX
ry2 = rx1 + dy * 2 * j \ Screen.TwipsPerPixelY
j = j + 1
Call Rectangle(hdc5, rx1, ry1, rx2, ry2)
Sleep (1)

Next i
Call ReleaseDC(0, hdc5)
Call DeleteObject(hbrush)
Fa1.Play
Form1.ZOrder 0
'End If

'''''''''''Shell ("c:\work\demo\Client.exe"), vbHide


'Form1.Visible = False
'Form1.Visible = True
'Form1.Refresh

End Sub

