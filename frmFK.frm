VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFKM 
   Caption         =   "付款方式参考"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6630
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4125
   ScaleWidth      =   6630
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame frmHtrq 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   345
      Left            =   90
      TabIndex        =   16
      Top             =   3330
      Width           =   6255
      Begin MSComCtl2.DTPicker dtgRq 
         Height          =   315
         Left            =   1080
         TabIndex        =   17
         Top             =   0
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   150732801
         CurrentDate     =   39626
      End
      Begin VB.Label Label5 
         Caption         =   "合同日期与付款日期相联系，请正确标明！"
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   2610
         TabIndex        =   19
         Top             =   60
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "合同日期"
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   30
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdGB 
      Caption         =   "关闭"
      Height          =   255
      Left            =   5970
      TabIndex        =   14
      Top             =   3840
      Width           =   585
   End
   Begin VB.TextBox txtYj 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3690
      Width           =   1245
   End
   Begin VB.Frame frmWb 
      Caption         =   "维保"
      Height          =   3255
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3405
      Begin VB.OptionButton optB3 
         Caption         =   "其它"
         Height          =   285
         Left            =   480
         TabIndex        =   11
         Top             =   1890
         Width           =   1335
      End
      Begin VB.OptionButton optB2 
         Caption         =   "分三期。  40%，30%，30%"
         Height          =   315
         Left            =   480
         TabIndex        =   10
         Top             =   1170
         Width           =   2685
      End
      Begin VB.OptionButton optB1 
         Caption         =   "分四期，每期25%"
         Height          =   315
         Left            =   480
         TabIndex        =   9
         Top             =   450
         Width           =   1995
      End
   End
   Begin VB.TextBox txtJJ 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3330
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3720
      Width           =   1395
   End
   Begin VB.Frame frmLp 
      Caption         =   "配件（产品）"
      Height          =   3255
      Left            =   3420
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.OptionButton optA5 
         Caption         =   "货到三个月内     /0.7"
         Height          =   255
         Left            =   330
         TabIndex        =   5
         Tag             =   "0.7"
         Top             =   2280
         Width           =   2235
      End
      Begin VB.OptionButton optA4 
         Caption         =   "货到二个月       /0.8"
         Height          =   255
         Left            =   330
         TabIndex        =   4
         Tag             =   "0.8"
         Top             =   1821
         Width           =   2355
      End
      Begin VB.OptionButton optA3 
         Caption         =   "货到一个月       /0.9"
         Height          =   255
         Left            =   330
         TabIndex        =   3
         Tag             =   "0.9"
         Top             =   1364
         Width           =   2325
      End
      Begin VB.OptionButton optA2 
         Caption         =   "货到一周内付款   /0.95"
         Height          =   255
         Left            =   330
         TabIndex        =   2
         Tag             =   "0.95"
         Top             =   907
         Width           =   2775
      End
      Begin VB.OptionButton optA1 
         Caption         =   "款到发货         /1"
         Height          =   285
         Left            =   330
         TabIndex        =   1
         Tag             =   "1"
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "只适用于纯配件和产品合同"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   420
         TabIndex        =   15
         Top             =   2940
         Width           =   2475
      End
   End
   Begin VB.Label Label2 
      Caption         =   "原价："
      Height          =   285
      Left            =   210
      TabIndex        =   12
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "基准价："
      Height          =   315
      Left            =   2490
      TabIndex        =   6
      Top             =   3720
      Width           =   765
   End
End
Attribute VB_Name = "frmFKM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public YJ As Double '原来的基准价
Dim FK As String
Dim FO As Single

Private Sub cmdGB_Click()
Me.Visible = False

If frmGXBj.Visible = True Then
    frmGXBj.Enabled = True: frmGXBj.ZOrder 0
ElseIf frmWBXNew.Visible = True Then
    frmWBXNew.Enabled = True: frmWBXNew.ZOrder 0
ElseIf FMXC.Visible = True Then
    FMXC.Enabled = True: FMXC.ZOrder 0
    'FMXC.lblFk.Caption = FK
    FMXC.FO = FO
    FMXC.txtHtrq.Text = DateSerial(Year(dtgRq.Value), Month(dtgRq.Value), Day(dtgRq.Value))
    FMXC.cmdSave.Enabled = True
    'Call FMXC.cmdSaveClick
End If
End Sub



Public Sub Jqing()
Me.YJ = 0
txtJJ.Text = "": txtYj.Text = ""
optA1.Value = False
optA2.Value = False
optA3.Value = False
optA4.Value = False
optA5.Value = False
optB1.Value = False
optB2.Value = False
optB3.Value = False
frmLp.Enabled = True
frmWb.Enabled = True
End Sub

Private Sub optA1_Click()
txtJJ.Text = Round(Val(txtYj.Text) / optA1.Tag, 2)
FO = optA1.Tag
FK = optA1.Caption
End Sub

Private Sub optA2_Click()
txtJJ.Text = Round(Val(txtYj.Text) / optA2.Tag, 2)
FO = optA2.Tag
FK = optA2.Caption
End Sub

Private Sub optA3_Click()
txtJJ.Text = Round(Val(txtYj.Text) / optA3.Tag, 2)
FO = optA3.Tag
FK = optA3.Caption
End Sub

Private Sub optA4_Click()
txtJJ.Text = Round(Val(txtYj.Text) / optA4.Tag, 2)
FO = optA4.Tag
FK = optA4.Caption
End Sub

Private Sub optA5_Click()
txtJJ.Text = Round(Val(txtYj.Text) / optA5.Tag, 2)
FO = optA5.Tag
FK = optA5.Caption
End Sub

