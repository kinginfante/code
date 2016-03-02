VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmFP 
   Caption         =   "开票"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdGx 
      Caption         =   "更 新"
      Height          =   345
      Left            =   5250
      TabIndex        =   33
      Top             =   6630
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   1770
      TabIndex        =   32
      Top             =   5250
      Width           =   3435
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   1
      Left            =   1380
      TabIndex        =   28
      Top             =   8160
      Width           =   945
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   0
      Left            =   300
      TabIndex        =   25
      Top             =   8160
      Width           =   945
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删 除"
      Height          =   345
      Left            =   5250
      TabIndex        =   24
      Top             =   6285
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "添 加"
      Height          =   345
      Left            =   5250
      TabIndex        =   23
      Top             =   5940
      Width           =   975
   End
   Begin VB.TextBox txtBz 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   7290
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   3990
      Width           =   7905
   End
   Begin VB.TextBox txtJe 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1800
      TabIndex        =   20
      Top             =   4050
      Width           =   3435
   End
   Begin VB.ComboBox comLx 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   19
      Top             =   4680
      Width           =   3465
   End
   Begin VB.ComboBox txtKhmc 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1500
      TabIndex        =   6
      ToolTipText     =   "请在列表中选择客户"
      Top             =   30
      Width           =   3345
   End
   Begin VB.TextBox txtXMMC 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6750
      TabIndex        =   5
      Top             =   0
      Width           =   3465
   End
   Begin VB.TextBox txtHtze 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6750
      TabIndex        =   4
      Top             =   450
      Width           =   3465
   End
   Begin VB.TextBox txtHtbh 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11670
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   3315
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      Height          =   585
      Left            =   13140
      Picture         =   "frmFP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8430
      Width           =   675
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "提交"
      Height          =   585
      Left            =   13830
      Picture         =   "frmFP.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8430
      Width           =   705
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14550
      Picture         =   "frmFP.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8430
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgFk 
      Height          =   1875
      Left            =   270
      TabIndex        =   13
      Top             =   1560
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   3307
      _Version        =   393216
      FillStyle       =   1
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgYk 
      Height          =   1875
      Left            =   7290
      TabIndex        =   16
      Top             =   1560
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   3307
      _Version        =   393216
      FillStyle       =   1
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label9 
      Caption         =   "备注:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   900
      TabIndex        =   31
      Top             =   5310
      Width           =   735
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   1380
      TabIndex        =   30
      Top             =   8580
      Width           =   945
   End
   Begin VB.Label lblQM 
      Caption         =   "开票人"
      Height          =   225
      Index           =   1
      Left            =   1470
      TabIndex        =   29
      Top             =   7890
      Width           =   915
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   300
      TabIndex        =   27
      Top             =   8580
      Width           =   945
   End
   Begin VB.Label lblQM 
      Caption         =   "开票申请"
      Height          =   225
      Index           =   0
      Left            =   390
      TabIndex        =   26
      Top             =   7890
      Width           =   915
   End
   Begin VB.Label Label8 
      Caption         =   "历史备注:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6000
      TabIndex        =   21
      Top             =   4050
      Width           =   1155
   End
   Begin VB.Label Label6 
      Caption         =   "开票类型:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   420
      TabIndex        =   18
      Top             =   4740
      Width           =   1185
   End
   Begin VB.Label Label5 
      Caption         =   "开票金额:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   420
      TabIndex        =   17
      Top             =   4110
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C000C0&
      BorderWidth     =   3
      X1              =   30
      X2              =   15240
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label3 
      Caption         =   "已开票据明细"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7290
      TabIndex        =   15
      Top             =   1140
      Width           =   1635
   End
   Begin VB.Label Label2 
      Caption         =   "预计收款明细"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   300
      TabIndex        =   14
      Top             =   1170
      Width           =   1995
   End
   Begin VB.Label Label7 
      Caption         =   "项目名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5310
      TabIndex        =   12
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "合同总金额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   510
      Width           =   1395
   End
   Begin VB.Label Label4 
      Caption         =   "合同性质"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   300
      TabIndex        =   10
      Top             =   525
      Width           =   1125
   End
   Begin VB.Label Label25 
      Caption         =   "合同编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10470
      TabIndex        =   9
      Top             =   60
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "客户名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   300
      TabIndex        =   8
      Top             =   90
      Width           =   1065
   End
   Begin VB.Label lblHtxz 
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   7
      Top             =   510
      Width           =   3315
   End
End
Attribute VB_Name = "frmFP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub KPQing()
txtXMMC.Text = ""
txtKhmc.Text = ""
txtHtbh.Text = ""
txtHtze.Text = ""
lblHtxz.Caption = ""
txtJe.Text = ""
comLx.Text = ""
txtBz.Text = ""

cmdQm(0).Caption = ""
lblTm(0).Caption = ""
cmdQm(1).Caption = ""
lblTm(1).Caption = ""

Set dtgFk.DataSource = Nothing
Set dtgYk.DataSource = Nothing
End Sub

Private Sub cmdBack_Click()
frmFP.Visible = False
If frmWbNew.Visible = True Then
    frmWbNew.Enabled = True
    frmWbNew.ZOrder 0
End If
End Sub

Private Sub cmdQm_Click(Index As Integer)
Dim oo As Integer
Dim tt As String
Dim Zid As Long
On Error Resume Next
If cmdQm(Index).Caption <> "" Or lblLcRen.Caption = "" Then
    Exit Sub
End If
If Not (lblLcRen.Caption = mod1.DName And lblLcUid.Caption = mod1.DHid) And lblLc.Caption <> (Index + 1) Then
    Exit Sub
End If

If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If


If lblLcUid.Caption <> mod1.DHid Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If

Dim Zi As Integer
Zi = MsgBox("是否确认签字?", vbYesNo)
If Zi = vbNo Then Exit Sub

    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "xtzxAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@yid").Value = 74
    mod1.cmd.Parameters("@lc").Value = lblLc.Caption
    mod1.cmd.Parameters("@bh").Value = lblGid.Caption
    mod1.cmd.Parameters("@ywy").Value = mod1.DName
    mod1.cmd.Parameters("@uid").Value = mod1.DHid
    mod1.cmd.Parameters("@bz").Value = ""
    mod1.cmd.Execute
    Zid = mod1.cmd.Parameters("@Zid").Value
    Set cmd = Nothing

cmdQm(Index).Caption = mod1.DName
lblTm(Index).Caption = mod1.DQda
lblLc.Caption = lblLc.Caption + 1
lblLcRen.Caption = ""
lblLcUid.Caption = ""
If Dialog.Visible = True Then
    Call mod1.refEnvent

End If
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
End Sub
