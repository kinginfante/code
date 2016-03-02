VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGyDetail 
   BackColor       =   &H00FFFFC0&
   Caption         =   "供应商详情"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.ComboBox comPJ 
      Height          =   300
      ItemData        =   "frmGyDetail.frx":0000
      Left            =   6270
      List            =   "frmGyDetail.frx":000D
      TabIndex        =   61
      Top             =   3900
      Width           =   2025
   End
   Begin VB.ComboBox comJS 
      Height          =   300
      ItemData        =   "frmGyDetail.frx":001A
      Left            =   6270
      List            =   "frmGyDetail.frx":002A
      TabIndex        =   59
      Top             =   4350
      Width           =   2055
   End
   Begin VB.CheckBox chkF 
      BackColor       =   &H00FFFFC0&
      Caption         =   "是否一般纳税人"
      Height          =   195
      Left            =   5640
      TabIndex        =   57
      Top             =   3420
      Value           =   2  'Grayed
      Width           =   1695
   End
   Begin VB.Frame frmKP 
      BackColor       =   &H00FFFFC0&
      Caption         =   "开票资料"
      Height          =   2175
      Left            =   60
      TabIndex        =   46
      Top             =   3030
      Width           =   5325
      Begin VB.TextBox txtKPZH 
         Height          =   300
         Left            =   1230
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox txtKPKH 
         Height          =   300
         Left            =   1230
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   1410
         Width           =   3975
      End
      Begin VB.TextBox txtKPDD 
         Height          =   300
         Left            =   1230
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   1020
         Width           =   3975
      End
      Begin VB.TextBox txtKPNS 
         Height          =   315
         Left            =   1230
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtKPMc 
         Height          =   315
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   210
         Width           =   3975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "帐号"
         Height          =   195
         Index           =   2
         Left            =   570
         TabIndex        =   54
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "开户行"
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   53
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "地址电话"
         Height          =   195
         Left            =   210
         TabIndex        =   51
         Top             =   1050
         Width           =   795
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "纳税人识别号"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   690
         Width           =   1125
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "公司名称"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   48
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "审核"
      Height          =   645
      Left            =   8490
      Picture         =   "frmGyDetail.frx":0052
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   8460
      Width           =   675
   End
   Begin VB.Frame frmRen 
      BackColor       =   &H00FFFFC0&
      Caption         =   "联系人编辑"
      Height          =   825
      Left            =   5520
      TabIndex        =   39
      Top             =   2490
      Width           =   2475
      Begin VB.CommandButton cmdRdel 
         Caption         =   "删除"
         Height          =   405
         Left            =   1770
         TabIndex        =   42
         Top             =   270
         Width           =   645
      End
      Begin VB.CommandButton cmdRgx 
         Caption         =   "更新"
         Height          =   405
         Left            =   930
         TabIndex        =   41
         Top             =   270
         Width           =   645
      End
      Begin VB.CommandButton cmdRadd 
         Caption         =   "添加"
         Height          =   405
         Left            =   90
         TabIndex        =   40
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.TextBox txtBz 
      Height          =   1245
      Left            =   6270
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   38
      Text            =   "frmGyDetail.frx":0494
      Top             =   4770
      Width           =   2055
   End
   Begin VB.TextBox txtGdw 
      Height          =   270
      Left            =   3780
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   5730
      Width           =   1455
   End
   Begin VB.TextBox txtZw 
      Height          =   270
      Left            =   3780
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   30
      TabIndex        =   27
      Top             =   7290
      Visible         =   0   'False
      Width           =   6315
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1365
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   300
         Width           =   4965
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   5220
         TabIndex        =   30
         Top             =   480
         Width           =   705
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   5220
         TabIndex        =   29
         Top             =   870
         Width           =   675
      End
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1320
         Width           =   735
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   2235
      Left            =   8370
      TabIndex        =   26
      Top             =   1500
      Visible         =   0   'False
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   3942
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtDj 
      Height          =   300
      Left            =   1320
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   1260
      Width           =   3975
   End
   Begin VB.TextBox txtYZ 
      Height          =   315
      Left            =   1320
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   810
      Width           =   3975
   End
   Begin VB.CommandButton cmdXZ 
      BackColor       =   &H0080FF80&
      Caption         =   "关联"
      Height          =   645
      Left            =   12240
      Picture         =   "frmGyDetail.frx":049A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "选择人员"
      Top             =   8340
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   810
      Top             =   30
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   30
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H00FFFFC0&
      Caption         =   "修改"
      Height          =   645
      Left            =   9870
      Picture         =   "frmGyDetail.frx":059C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "修改"
      Top             =   8460
      Width           =   645
   End
   Begin VB.CommandButton cmdCreate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "新建"
      Height          =   645
      Left            =   9210
      Picture         =   "frmGyDetail.frx":08A6
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8460
      Width           =   645
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "保存"
      Height          =   645
      Left            =   10530
      Picture         =   "frmGyDetail.frx":0CE8
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8460
      Width           =   645
   End
   Begin VB.ComboBox txtJFw 
      Height          =   300
      Left            =   1320
      TabIndex        =   14
      Top             =   2130
      Width           =   3975
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFC0&
      Caption         =   "返回"
      Height          =   645
      Left            =   11190
      Picture         =   "frmGyDetail.frx":1352
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8460
      Width           =   645
   End
   Begin VB.TextBox txtADR 
      Height          =   315
      Left            =   1320
      TabIndex        =   12
      Top             =   2565
      Width           =   3975
   End
   Begin VB.TextBox txtFax 
      Height          =   315
      Left            =   3750
      TabIndex        =   11
      Top             =   1680
      Width           =   1545
   End
   Begin VB.TextBox txtMdw 
      Height          =   270
      Left            =   1320
      TabIndex        =   10
      Top             =   5715
      Width           =   1725
   End
   Begin VB.TextBox txtDw 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtLXR 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   5310
      Width           =   1725
   End
   Begin VB.TextBox txtMC 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   2955
      Left            =   5490
      TabIndex        =   25
      Top             =   330
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   5212
      _Version        =   393216
      BackColor       =   16777152
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   3015
      Left            =   0
      TabIndex        =   32
      Top             =   6090
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   5318
      _Version        =   393216
      BackColor       =   15728356
      ForeColor       =   8404992
      Rows            =   15
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   16777152
      ForeColorFixed  =   0
      BackColorBkg    =   15728356
      GridColorFixed  =   8404992
      GridColorUnpopulated=   8404992
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgCJ 
      Height          =   7515
      Left            =   8460
      TabIndex        =   43
      Top             =   330
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   13256
      _Version        =   393216
      BackColor       =   16777152
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "评级"
      Height          =   255
      Left            =   5610
      TabIndex        =   60
      Top             =   3960
      Width           =   555
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "结算方式"
      Height          =   195
      Left            =   5490
      TabIndex        =   58
      Top             =   4410
      Width           =   765
   End
   Begin VB.Label lblTx 
      BackStyle       =   0  'Transparent
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   8550
      TabIndex        =   45
      Top             =   7980
      Width           =   4275
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "备注"
      Height          =   255
      Left            =   5490
      TabIndex        =   37
      Top             =   4860
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "电话"
      Height          =   255
      Left            =   3330
      TabIndex        =   35
      Top             =   5760
      Width           =   525
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "职务"
      Height          =   255
      Left            =   3300
      TabIndex        =   33
      Top             =   5340
      Width           =   405
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "税务登记"
      Height          =   285
      Left            =   180
      TabIndex        =   24
      Top             =   1350
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "营业执照"
      Height          =   315
      Left            =   180
      TabIndex        =   21
      Top             =   855
      Width           =   1005
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   315
      Left            =   540
      TabIndex        =   19
      Top             =   90
      Width           =   315
   End
   Begin VB.Label lblGid 
      BackStyle       =   0  'Transparent
      Caption         =   "lblGid"
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   90
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGyDetail.frx":1454
      Height          =   285
      Left            =   180
      TabIndex        =   8
      Top             =   2610
      Width           =   1245
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGyDetail.frx":1464
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGyDetail.frx":1472
      Height          =   225
      Left            =   3270
      TabIndex        =   6
      Top             =   1770
      Width           =   405
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGyDetail.frx":147C
      Height          =   255
      Left            =   210
      TabIndex        =   5
      Top             =   5790
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGyDetail.frx":1486
      Height          =   255
      Left            =   150
      TabIndex        =   4
      Top             =   1755
      Width           =   765
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGyDetail.frx":1494
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   5340
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "公司名称"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   405
      Width           =   1275
   End
End
Attribute VB_Name = "frmGyDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim timZm As Integer '(1保存)
Dim liD As Long
Dim LL As String '录入者
Dim LLUid As String
Dim LCRen As String
Dim LCUid As String
Dim Lc As Integer
Dim Fwid As Long
Public Sub dtgPFF()
Dim oo As Integer
For oo = 1 To dtgP.Rows - 1
    dtgP.RowHeight(oo) = dtgP.RowHeight(0) * 2
Next
dtgP.Clear
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "日期": dtgP.Col = 1: dtgP.Text = "姓名": dtgP.Col = 2: dtgP.Text = "职能": dtgP.Col = 3: dtgP.Text = "评审建议": dtgP.Col = 4: dtgP.Text = "审核":
dtgP.ColWidth(0) = 1665
dtgP.ColWidth(1) = 1005
dtgP.ColWidth(2) = 0
 dtgP.ColWidth(3) = 4290: dtgP.ColWidth(4) = 1035
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub
Public Sub Qing()

Me.txtMc.Text = ""

Me.txtDw.Text = ""
Me.chkF.Value = 2
Me.txtFax.Text = ""
Me.txtJFw.Text = ""
Me.txtADR.Text = ""
Me.txtYZ.Text = ""
Me.txtDj.Text = ""
Me.txtKPMc.Text = ""
Me.txtKPNS.Text = ""
Me.txtKPDD.Text = ""
Me.txtKPKH.Text = ""
Me.txtKPZH.Text = ""
Me.lblGid.Caption = ""
Me.txtBz.Text = ""
Call Me.dtgbrFF
Call Me.dtgPFF
frmRen.Visible = False
Call RenQing
LL = ""
LLUid = ""
LCRen = ""
LCUid = ""
Lc = 1
Fwid = 0
comJS.Text = ""
Me.comPJ.Text = ""
End Sub
Public Sub Bound(Gid As Long)
Dim tt As String
Dim Ra
Dim Rb
Dim RC
Dim RD
Dim Rz
Dim Lz As Integer
On Error Resume Next
tt = "select mc,dw,fax,fw,adr,yz,dj,kpns,bz,lcren,lcuid,lc,fwid,ll,lluid,kpdd,kpkh,kpzh,kzf,js,pj from gymxc where gid=" & Gid & ";" & _
    "select name,zw,lid from gyRen where gid=" & Gid & " order by Lid;" & _
    "select top 1 name,zw,mdw,gdw,lid from gyRen where gid=" & Gid & ";" & _
    "select trq,ywy,zn,bz,tf from pizu where bh='" & Gid & "' and yid=92 order by pid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rz = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
On Error Resume Next

Me.txtMc.Text = Ra(0, 0)
Me.txtDw.Text = Ra(1, 0)
Me.txtFax.Text = Ra(2, 0)
Me.txtJFw.Text = Ra(3, 0)
Me.txtADR.Text = Ra(4, 0)
Me.txtYZ.Text = Ra(5, 0)
Me.txtDj.Text = Ra(6, 0)
Me.txtKPMc.Text = Ra(0, 0)
Me.txtBz.Text = Ra(8, 0)
LCRen = Ra(9, 0)
LCUid = Ra(10, 0)
Lc = Ra(11, 0)
        lblTX.Caption = "流程至：" & LCRen
        If Lc = 100 Then lblTX.Caption = "审核完毕!"
Fwid = Ra(12, 0)
LL = Ra(13, 0)
LLUid = Ra(14, 0)
Me.lblGid.Caption = Gid
Me.txtKPNS.Text = Ra(7, 0)
Me.txtKPDD.Text = Ra(15, 0)
Me.txtKPKH.Text = Ra(16, 0)
Me.txtKPZH.Text = Ra(17, 0)
Me.chkF.Value = 2
Me.chkF.Value = Ra(18, 0)
Me.comJS.Text = Ra(19, 0)
Me.comPJ.Text = Ra(20, 0)
Call Me.RenBound(Rb)
Call Me.RDBound(RC)
Lz = UBound(Rz, 2) + 1
Call Me.QMBound(Rz, Lz)
Exit Sub
frmGyER1:
MsgBox "出错!"
End Sub

Public Sub QMBound(Rz, Lz As Integer)
Dim ii As Integer: Dim oo As Integer
On Error Resume Next
Call dtgPFF
dtgP.Rows = Lz + 20

For oo = 1 To Lz + 1
    dtgP.Row = oo
    For ii = 0 To 5
        dtgP.Col = ii
        dtgP.Text = Rz(ii, oo - 1)
        If ii = 3 Then
            If Len(Rz(ii, oo - 1)) > 16 Then
                dtgP.RowHeight(oo) = UpInt(Len(Rz(ii, oo - 1)) / 16) * dtgP.RowHeight(oo)
            End If
        End If
        If ii = 4 Then
            If dtgP.Text = "True" Then
                dtgP.Text = "同意"
            ElseIf dtgP.Text = "False" Then
                dtgP.Text = "驳回"
            End If

        End If
    Next
Next
For oo = 1 To Lz + 1
    dtgP.Row = oo
    dtgP.Col = 4
            If dtgP.Text = "驳回" Then
                For ii = 0 To 5
                    dtgP.Col = ii
                    dtgP.CellForeColor = &HFF&
                Next
            End If
Next
End Sub
Private Sub cmdBack_Click()
Me.Visible = False
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
ElseIf FmxcXJ.Visible = True Then
    FmxcXJ.Enabled = True
    FmxcXJ.ZOrder 0
End If
End Sub

Private Sub cmdCreate_Click()
If mod1.DName = "马晓聪" Or mod1.DName = "" Or mod1.DName = "" Or mod1.DName = "朱婷婷" Then
    Call Me.Qing
    cmdSave.Enabled = True
End If
End Sub

Private Sub cmdDing_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
'''''If Lc = 1 And OptT1.Value = True Then
'''''    If txtMC.Text = "" Then
'''''        MsgBox "没有输入公司名称！"
'''''        txtMC.SetFocus
'''''        Exit Sub
'''''    End If
'''''    If txtDw.Text = "" Then
'''''        MsgBox "没有输入单位名称！"
'''''        txtDw.SetFocus
'''''        Exit Sub
'''''    End If
'''''    If chkF.Value = 2 Then
'''''        MsgBox "没有确认是一般纳税人！"
'''''        txtMC.SetFocus
'''''        Exit Sub
'''''    End If
'''''    If txtFax.Text = "" Then
'''''        MsgBox "没有输入传真！"
'''''        txtFax.SetFocus
'''''        Exit Sub
'''''    End If
'''''    If txtJFw.Text = "" Then
'''''        MsgBox "没有输入经营范围！"
'''''        txtJFw.SetFocus
'''''        Exit Sub
'''''    End If
'''''    If txtADR.Text = "" Then
'''''        MsgBox "没有输入供货方地址！"
'''''        txtADR.SetFocus
'''''        Exit Sub
'''''    End If
'''''    If txtYZ.Text = "" Then
'''''        MsgBox "没有输入营业执照号！"
'''''        txtYZ.SetFocus
'''''        Exit Sub
'''''    End If
'''''    If txtDj.Text = "" Then
'''''        MsgBox "没有输入税务登记！"
'''''        txtDj.SetFocus
'''''        Exit Sub
'''''    End If
'''''    If txtKPMc.Text = "" Or txtKPNS.Text = "" Or txtKPDD.Text = "" Or txtKPKH.Text = "" Or txtKPZH.Text = "" Then
'''''            MsgBox "没有输入开票资料！"
'''''            Exit Sub
'''''    End If
'''''End If


If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If

timZm = 6 '签字
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "供应商资料"
    mod1.cmd.Parameters("@NBLX") = "签字"
    mod1.cmd.Parameters("@bh") = Val(lblGid.Caption)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = LL
    mod1.cmd.Parameters("@mt2") = LLUid
    mod1.cmd.Parameters("@mt3") = txtMc.Text
    
    mod1.cmd.Parameters("@mlt1") = txtQM.Text '评审建议
    mod1.cmd.Parameters("@mm1") = Lc
    mod1.cmd.Parameters("@mm2") = Fwid
    If OptT1.Value = True Then
        mod1.cmd.Parameters("@mb1") = 1 '同意
    Else
        mod1.cmd.Parameters("@mb1") = 0 '拒绝
    End If
    mod1.cmd.Parameters("@md1") = Null
     Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        cmdDing.Enabled = False
    
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If

    
Set mod1.cmd = Nothing
frmQm.Visible = False
End Sub

Private Sub cmdMod_Click()
If mod1.DName = "" Or mod1.DName = "马晓聪" Or mod1.DName = "" Or mod1.DName = "吴金荣" Then
    cmdSave.Enabled = True
    If Val(lblGid.Caption) > 0 Then
        frmRen.Visible = True
    End If
End If
End Sub

Private Sub cmdNQ_Click()
Dim ii As Integer
Dim tt As String
Dim Ra
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号

Dim oo As Integer
On Error Resume Next


If lblTX.Caption = "审核完毕!" Then Exit Sub
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If



If LCUid <> mod1.DHid Then
        MsgBox "此处应由" & LCRen & "签字! 请您不要再点"
        Exit Sub
End If

frmQm.Visible = True
If Lc = 1 Then
    optT2.Enabled = False
    OptT1.Value = True
    
Else
    OptT1.Enabled = True
    optT2.Enabled = True
    OptT1.Value = False
    optT2.Value = False
End If

End Sub

Private Sub cmdRadd_Click()
Call RenQing
End Sub

Private Sub cmdRgx_Click()
Dim tt As String
If txtLXR.Text = "" Or txtZw.Text = "" Or txtMdw.Text = "" Or txtGdw.Text = "" Then
    MsgBox "供应商资料不全!"
    Exit Sub
End If

On Error Resume Next

timZm = 2 '保存
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "供应商资料"
    mod1.cmd.Parameters("@NBLX") = "人员保存"
    mod1.cmd.Parameters("@bh") = lblGid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtLXR.Text
    mod1.cmd.Parameters("@mt2") = txtZw.Text
    mod1.cmd.Parameters("@mt3") = txtMdw.Text
    mod1.cmd.Parameters("@mt4") = txtGdw.Text
    
    mod1.cmd.Parameters("@mm1") = liD
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        cmdSave.Enabled = False
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False
        
        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
        
    End If

    
Set mod1.cmd = Nothing
End Sub


Private Sub cmdSave_Click()
Dim tt As String
If txtMc.Text = "" Or txtJFw.Text = "" Then
    MsgBox "供应商资料不全!"
    Exit Sub
End If
'''''    If chkF.Value = 2 Then
'''''        MsgBox "没有确认是否一般！"
'''''        Exit Sub
'''''    End If
On Error Resume Next

timZm = 1 '保存
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "供应商资料"
    mod1.cmd.Parameters("@NBLX") = "保存"
    mod1.cmd.Parameters("@bh") = lblGid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtMc.Text
    mod1.cmd.Parameters("@mt2") = txtDw.Text
    mod1.cmd.Parameters("@mt3") = txtFax.Text
    mod1.cmd.Parameters("@mt4") = txtJFw.Text
    mod1.cmd.Parameters("@mt5") = txtADR.Text
    mod1.cmd.Parameters("@mt6") = txtYZ.Text
    mod1.cmd.Parameters("@mt7") = txtDj.Text
    mod1.cmd.Parameters("@mt8") = txtKPNS.Text
    mod1.cmd.Parameters("@mt9") = txtKPDD.Text
    mod1.cmd.Parameters("@mt10") = txtKPKH.Text
    mod1.cmd.Parameters("@mt11") = txtKPZH.Text
    mod1.cmd.Parameters("@mt12") = comJS.Text
    mod1.cmd.Parameters("@mt13") = comPJ.Text '评级
    mod1.cmd.Parameters("@mlt1") = txtBz.Text
    mod1.cmd.Parameters("@mb1") = chkF.Value
    mod1.cmd.Parameters("@md1") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        cmdSave.Enabled = False
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False
        
        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
        
    End If

    
Set mod1.cmd = Nothing
End Sub

Private Sub cmdXZ_Click()
If frmHPZL.Visible = True Then
Select Case frmHPZL.GyId
Case 1
    frmHPZL.txtGy1.Text = txtMc.Text
    frmHPZL.txtGy1.ToolTipText = lblGid.Caption
Case 2
    frmHPZL.txtGy2.Text = txtMc.Text
    frmHPZL.txtGy2.ToolTipText = lblGid.Caption
Case 3
    frmHPZL.txtGY3.Text = txtMc.Text
    frmHPZL.txtGY3.ToolTipText = lblGid.Caption
End Select
End If
frmGY.Visible = False
frmGyDetail.Visible = False
End Sub

Private Sub dtgBr_Click()
Dim RC
dtgN.Row = dtgBr.Row
dtgN.Col = 2
liD = Val(dtgN.Text)
If liD = 0 Then Exit Sub
tt = "select name,zw,mdw,gdw,lid from gyRen where lid=" & liD
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
RC = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Call Me.RDBound(RC)
End Sub

Private Sub Form_Load()
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
Cancel = True
End Sub

Private Sub timQuit_Timer()
Dim Rz
Dim Lz As Integer
Dim oo As Integer
Dim ii As Integer
Dim tt As String
Dim Rb
On Error Resume Next
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0

If timZm = 1 Then '保存
ElseIf timZm = 2 Then '人员保存
    tt = "select name,zw,lid from gyRen where gid=" & Val(lblGid.Caption) & " order by Lid"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rb = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Call Me.RenBound(Rb)
ElseIf timZm = 6 Then '签字
    tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & lblGid.Caption & "' and yid=92 order by pid desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rz = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Lz = UBound(Rz, 2) + 1
    Call QMBound(Rz, Lz)
End If
timQuit.Enabled = False
Me.WindowState = 0
Me.ZOrder 0
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2,mt3 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdSave.Enabled = False
        lblGid.Caption = mod1.WP.Fields("mm1").Value
    ElseIf timZm = 2 Then
        liD = mod1.WP.Fields("mm1").Value
    ElseIf timZm = 6 Then
                Lc = mod1.WP.Fields("mm1").Value
                Fwid = mod1.WP.Fields("mm2").Value
                LCRen = mod1.WP.Fields("mt1").Value
                LCUid = mod1.WP.Fields("mt2").Value
                lblTX.Caption = "下一流程,将跳至" & mod1.WP.Fields("mt3").Value & ": " & LCRen
                If Lc = 100 Then lblTX.Caption = "审核完毕!"
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    timWait.Enabled = False
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
'''''    If timZm = 2 Then
'''''        cmdSave.Enabled = False
'''''    ElseIf timZm = 11 Then
'''''        txtHtbh.Text = ""
'''''        lblHtxz.Caption = ""
'''''    End If
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0

    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub


Private Sub txtDj_Change()
txtKPNS.Text = txtDj.Text
End Sub

Private Sub txtJFw_Change()
If Len(txtJFw.Text) > 200 Then
    MsgBox "超出字数!"
End If
End Sub

Public Sub dtgbrFF()
dtgBr.Clear: dtgN.Clear
dtgBr.Rows = 20
dtgBr.Cols = 3
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "联系人": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "职务": dtgBr.CellFontBold = True

dtgN.Rows = 20
dtgN.Cols = 3
dtgBr.ColWidth(1) = 2000
dtgBr.ColWidth(2) = 0
End Sub

Public Sub RenQing()
txtLXR.Text = ""
txtMdw.Text = ""
txtZw.Text = ""
txtGdw.Text = ""
liD = 0
End Sub

Public Sub RenBound(Rb)
On Error Resume Next
Dim Lb As Integer
Dim oo As Integer
Lb = UBound(Rb, 2) + 1
Call Me.dtgbrFF
For oo = 1 To Lb
    dtgBr.Row = oo: dtgN.Row = oo
    dtgBr.Col = 0: dtgBr.Text = Rb(0, oo - 1): dtgN.Col = 0: dtgN.Text = Rb(0, oo - 1)
    dtgBr.Col = 1: dtgBr.Text = Rb(1, oo - 1): dtgN.Col = 1: dtgN.Text = Rb(1, oo - 1)
    dtgBr.Col = 2: dtgBr.Text = Rb(2, oo - 1): dtgN.Col = 2: dtgN.Text = Rb(2, oo - 1)
    dtgBr.Col = 3: dtgBr.Text = Rb(3, oo - 1): dtgN.Col = 3: dtgN.Text = Rb(3, oo - 1)
Next
End Sub

Public Sub RDBound(RC)
    On Error Resume Next
    txtLXR.Text = RC(0, 0)
    txtZw.Text = RC(1, 0)
    txtMdw.Text = RC(2, 0)
    txtGdw.Text = RC(3, 0)
    liD = RC(4, 0)
End Sub

Public Sub dtgCJFF()
dtgCJ.Cols = 3
dtgCJ.Rows = 50
dtgCJ.Clear

End Sub
