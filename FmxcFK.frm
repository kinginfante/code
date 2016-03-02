VERSION 5.00
Begin VB.Form FmxcFK 
   BackColor       =   &H00C0FFC0&
   Caption         =   "跨区销售"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   5535
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtBL3 
      Height          =   285
      Left            =   3930
      TabIndex        =   14
      Top             =   1710
      Width           =   495
   End
   Begin VB.TextBox txtBL2 
      Height          =   300
      Left            =   3930
      TabIndex        =   13
      Top             =   1147
      Width           =   495
   End
   Begin VB.TextBox txtBL1 
      Height          =   285
      Left            =   3930
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtRen3 
      Height          =   285
      Left            =   2670
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1710
      Width           =   1005
   End
   Begin VB.TextBox txtRen2 
      Height          =   270
      Left            =   2670
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1155
      Width           =   1005
   End
   Begin VB.TextBox txtRen1 
      Height          =   285
      Left            =   2670
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   600
      Width           =   1005
   End
   Begin VB.ComboBox comQy3 
      Height          =   300
      ItemData        =   "FmxcFK.frx":0000
      Left            =   1260
      List            =   "FmxcFK.frx":0019
      TabIndex        =   5
      Top             =   1710
      Width           =   1215
   End
   Begin VB.ComboBox comQy2 
      Height          =   300
      ItemData        =   "FmxcFK.frx":0047
      Left            =   1260
      List            =   "FmxcFK.frx":0060
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   1155
      Width           =   1215
   End
   Begin VB.ComboBox comQy1 
      Height          =   300
      ItemData        =   "FmxcFK.frx":008E
      Left            =   1260
      List            =   "FmxcFK.frx":00A7
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "比例%"
      Height          =   195
      Left            =   3990
      TabIndex        =   11
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "管理者"
      Height          =   255
      Left            =   2730
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "区域"
      Height          =   225
      Left            =   1350
      TabIndex        =   9
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "跨区销售3"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "跨区销售2"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   1230
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "跨区销售1"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   660
      Width           =   945
   End
End
Attribute VB_Name = "FmxcFK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xZ As Integer '选择人员框
Public Sub Qing()
comQy1.Text = FMXC.comQy.Text
comQy2.Text = ""
comQy3.Text = ""
txtRen1.Text = FMXC.txtXYwy.Text: txtRen1.ToolTipText = FMXC.txtXYwy.ToolTipText
txtRen2.Text = "": txtRen2.ToolTipText = ""
txtRen3.Text = "": txtRen3.ToolTipText = ""
txtBL1.Text = ""
txtBL2.Text = ""
txtBL3.Text = ""

End Sub

Private Sub comQy2_Change()
If comQy2.Text = "" Then txtRen2.Text = ""
End Sub

Private Sub comQy3_Change()
If comQy3.Text = "" Then txtRen3.Text = ""
End Sub


Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
Cancel = True
End Sub

Private Sub txtBL2_LostFocus()

txtBL1.Text = 100 - Val(txtBL2.Text) - Val(txtBL3.Text)
End Sub

Private Sub txtBL3_LostFocus()

txtBL1.Text = 100 - Val(txtBL2.Text) - Val(txtBL3.Text)
End Sub

Private Sub txtRen2_DblClick()
Set Ren.XForm = New FmxcFK
Call mod1.RenXz("FmxcFK", Me, 0)
Me.xZ = 2
End Sub


Private Sub txtRen3_DblClick()
Set Ren.XForm = New FmxcFK
Call mod1.RenXz("FmxcFK", Me, 0)
Me.xZ = 3
End Sub


