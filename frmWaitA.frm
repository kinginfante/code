VERSION 5.00
Begin VB.Form frmWaitA 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   585
   ClientLeft      =   8565
   ClientTop       =   6015
   ClientWidth     =   7470
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   1830
      Top             =   240
   End
   Begin VB.Label lblRun 
      BackColor       =   &H8000000D&
      Height          =   165
      Left            =   60
      TabIndex        =   1
      Top             =   330
      Width           =   195
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "命令已经提交,服务中心正在进行处理,请稍候......"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   7425
   End
End
Attribute VB_Name = "frmWaitA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()

End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Timer2_Timer()
lblRun.Width = lblRun.Width + 265
Me.Cls
If lblRun.Width >= Me.Width Then
    Timer2.Enabled = False
End If
End Sub


