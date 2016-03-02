VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FmxcNewZX 
   BackColor       =   &H00C0FFC0&
   Caption         =   "合同执行状况"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15060
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   15060
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "返回"
      Height          =   765
      Left            =   14280
      Picture         =   "FmxcNewZX.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8040
      Width           =   585
   End
   Begin TabDlg.SSTab TabMxc 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12648384
      TabCaption(0)   =   "销售类"
      TabPicture(0)   =   "FmxcNewZX.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "销售订单类"
      TabPicture(1)   =   "FmxcNewZX.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   8655
         Left            =   -74980
         TabIndex        =   2
         Top             =   320
         Width           =   15135
         Begin VB.Frame Frame7 
            BackColor       =   &H00C0FFC0&
            Caption         =   "采购付款"
            Height          =   2775
            Left            =   6840
            TabIndex        =   50
            Top             =   5880
            Width           =   3375
            Begin VB.CommandButton Command6 
               Caption         =   "删除"
               Height          =   255
               Left            =   360
               TabIndex        =   55
               Top             =   1890
               Width           =   825
            End
            Begin VB.CommandButton Command5 
               Caption         =   "添加"
               Height          =   255
               Left            =   360
               TabIndex        =   54
               Top             =   1560
               Width           =   825
            End
            Begin VB.CommandButton Command4 
               Caption         =   "更新"
               Height          =   255
               Left            =   360
               TabIndex        =   53
               Top             =   2220
               Width           =   825
            End
            Begin VB.TextBox Text4 
               Height          =   270
               Left            =   1440
               TabIndex        =   52
               Text            =   "Text2"
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox Text3 
               Height          =   270
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   51
               Text            =   "Text1"
               Top             =   480
               Width           =   1815
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   255
               Left            =   1440
               TabIndex        =   56
               Top             =   480
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Format          =   102957057
               CurrentDate     =   42128
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "采购付款日期"
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "采购付款金额 "
               Height          =   255
               Left            =   120
               TabIndex        =   57
               Top             =   1080
               Width           =   1215
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0FFC0&
            Caption         =   "采购订单"
            Height          =   2775
            Left            =   3360
            TabIndex        =   41
            Top             =   5880
            Width           =   3495
            Begin VB.CommandButton Command3 
               Caption         =   "删除"
               Height          =   255
               Left            =   1440
               TabIndex        =   61
               Top             =   2460
               Width           =   825
            End
            Begin VB.CommandButton Command2 
               Caption         =   "添加"
               Height          =   255
               Left            =   240
               TabIndex        =   60
               Top             =   2460
               Width           =   825
            End
            Begin VB.CommandButton Command1 
               Caption         =   "更新"
               Height          =   255
               Left            =   2520
               TabIndex        =   59
               Top             =   2460
               Width           =   825
            End
            Begin VB.TextBox txtCJE 
               Height          =   270
               Left            =   1440
               TabIndex        =   49
               Text            =   "Text2"
               Top             =   2040
               Width           =   1815
            End
            Begin VB.TextBox txtCSL 
               Height          =   270
               Left            =   1440
               TabIndex        =   47
               Text            =   "Text1"
               Top             =   1520
               Width           =   1815
            End
            Begin VB.TextBox txtCBH 
               Height          =   270
               Left            =   1440
               TabIndex        =   43
               Text            =   "Text2"
               Top             =   1000
               Width           =   1815
            End
            Begin VB.TextBox txtGymc 
               Height          =   270
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   42
               Text            =   "Text1"
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "采购订单金额 "
               Height          =   255
               Left            =   240
               TabIndex        =   48
               Top             =   2040
               Width           =   1095
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "采购数量"
               Height          =   375
               Left            =   480
               TabIndex        =   46
               Top             =   1500
               Width           =   735
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "采购供应商"
               Height          =   255
               Left            =   360
               TabIndex        =   45
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "采购订单号"
               Height          =   255
               Left            =   360
               TabIndex        =   44
               Top             =   1080
               Width           =   1095
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "要货"
            Height          =   2775
            Left            =   0
            TabIndex        =   32
            Top             =   5880
            Width           =   3375
            Begin VB.TextBox txtYRQ 
               Height          =   270
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   37
               Text            =   "Text1"
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtYSL 
               Height          =   270
               Left            =   1200
               TabIndex        =   36
               Text            =   "Text2"
               Top             =   1080
               Width           =   1815
            End
            Begin VB.CommandButton cmdYGx 
               Caption         =   "更新"
               Height          =   255
               Left            =   360
               TabIndex        =   35
               Top             =   2220
               Width           =   825
            End
            Begin VB.CommandButton cmdYadd 
               Caption         =   "添加"
               Height          =   255
               Left            =   360
               TabIndex        =   34
               Top             =   1560
               Width           =   825
            End
            Begin VB.CommandButton cmdYDel 
               Caption         =   "删除"
               Height          =   255
               Left            =   360
               TabIndex        =   33
               Top             =   1890
               Width           =   825
            End
            Begin MSComCtl2.DTPicker dtpYRQ 
               Height          =   255
               Left            =   1200
               TabIndex        =   38
               Top             =   480
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Format          =   102957057
               CurrentDate     =   42128
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "要货数量"
               Height          =   255
               Left            =   360
               TabIndex        =   40
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "要货日期"
               Height          =   255
               Left            =   360
               TabIndex        =   39
               Top             =   480
               Width           =   1095
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN2 
            Height          =   1335
            Left            =   11520
            TabIndex        =   31
            Top             =   6240
            Visible         =   0   'False
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   2355
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgDC 
            Height          =   5655
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   15165
            _ExtentX        =   26749
            _ExtentY        =   9975
            _Version        =   393216
            BackColor       =   16777152
            FixedCols       =   0
            BackColorFixed  =   15728356
            BackColorBkg    =   16777152
            WordWrap        =   -1  'True
            SelectionMode   =   1
            AllowUserResizing=   3
            PictureType     =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "合同评审单"
         Height          =   8655
         Left            =   20
         TabIndex        =   1
         Top             =   320
         Width           =   15135
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "开票"
            Height          =   2775
            Left            =   3840
            TabIndex        =   15
            Top             =   5760
            Width           =   3855
            Begin VB.CommandButton cmdA 
               Caption         =   "添加"
               Height          =   255
               Left            =   240
               TabIndex        =   29
               Top             =   1560
               Width           =   825
            End
            Begin VB.CommandButton cmdD 
               Caption         =   "删除"
               Height          =   255
               Left            =   240
               TabIndex        =   28
               Top             =   1890
               Width           =   825
            End
            Begin VB.TextBox txtWJSJE 
               Enabled         =   0   'False
               Height          =   270
               Left            =   7440
               TabIndex        =   27
               Text            =   "Text2"
               Top             =   1680
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.TextBox txtJSJE 
               Height          =   270
               Left            =   7440
               TabIndex        =   25
               Text            =   "Text2"
               Top             =   1080
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CommandButton cmdG 
               Caption         =   "更新"
               Height          =   255
               Left            =   240
               TabIndex        =   23
               Top             =   2220
               Width           =   825
            End
            Begin VB.TextBox txtWKPJE 
               Enabled         =   0   'False
               Height          =   270
               Left            =   1440
               TabIndex        =   22
               Text            =   "Text2"
               Top             =   1680
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.TextBox txtKPRQ 
               Height          =   270
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   19
               Text            =   "Text1"
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtKPJE 
               Height          =   270
               Left            =   1440
               TabIndex        =   17
               Text            =   "Text2"
               Top             =   1080
               Width           =   1815
            End
            Begin MSComCtl2.DTPicker dtpKPRQ 
               Height          =   255
               Left            =   1440
               TabIndex        =   20
               Top             =   480
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Format          =   102957057
               CurrentDate     =   42128
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "未结算金额"
               Height          =   255
               Left            =   6360
               TabIndex        =   26
               Top             =   1680
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "结算金额"
               Height          =   255
               Left            =   6360
               TabIndex        =   24
               Top             =   1080
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "未开票金额"
               Height          =   255
               Left            =   1800
               TabIndex        =   21
               Top             =   1440
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "开票日期"
               Height          =   255
               Left            =   360
               TabIndex        =   18
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "开票金额"
               Height          =   255
               Left            =   360
               TabIndex        =   16
               Top             =   1080
               Width           =   1095
            End
         End
         Begin VB.Timer timWait 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   9480
            Top             =   5880
         End
         Begin VB.Timer timQuit 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   9480
            Top             =   6360
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN1 
            Height          =   735
            Left            =   10920
            TabIndex        =   14
            Top             =   5880
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   1296
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0FFC0&
            Caption         =   "开单"
            Height          =   2775
            Left            =   0
            TabIndex        =   5
            Top             =   5760
            Width           =   3735
            Begin VB.CommandButton cmdDe 
               Caption         =   "删除"
               Height          =   255
               Left            =   360
               TabIndex        =   13
               Top             =   1890
               Width           =   825
            End
            Begin VB.CommandButton cmdAd 
               Caption         =   "添加"
               Height          =   255
               Left            =   360
               TabIndex        =   12
               Top             =   1560
               Width           =   825
            End
            Begin VB.CommandButton cmdGx 
               Caption         =   "更新"
               Height          =   255
               Left            =   360
               TabIndex        =   11
               Top             =   2220
               Width           =   825
            End
            Begin VB.TextBox txtKDJE 
               Height          =   270
               Left            =   1440
               TabIndex        =   9
               Text            =   "Text2"
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox txtKDRQ 
               Height          =   270
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   7
               Text            =   "Text1"
               Top             =   480
               Width           =   1815
            End
            Begin MSComCtl2.DTPicker dtpKDRQ 
               Height          =   255
               Left            =   1440
               TabIndex        =   8
               Top             =   480
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Format          =   102957057
               CurrentDate     =   42128
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "开单日期"
               Height          =   255
               Left            =   360
               TabIndex        =   10
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "开单金额"
               Height          =   255
               Left            =   360
               TabIndex        =   6
               Top             =   1080
               Width           =   1095
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
            Height          =   5655
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   15165
            _ExtentX        =   26749
            _ExtentY        =   9975
            _Version        =   393216
            BackColor       =   16777152
            FixedCols       =   0
            BackColorFixed  =   15728356
            BackColorBkg    =   16777152
            WordWrap        =   -1  'True
            SelectionMode   =   1
            AllowUserResizing=   3
            PictureType     =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
   End
End
Attribute VB_Name = "FmxcNewZX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Fid As Long
Dim Kid As Long
Dim KPid As Long
Dim Jid As Long '到帐收款明细
Dim Aid As Long '财务到帐单号
Dim timZm As Integer

Dim ACid As Long '执行单id
Dim YHid As Long '要货id


Public Hid As Long


Private Sub cmdA_Click()
On Error Resume Next

If txtKPRQ.Text = "" Or Val(txtKPJE.Text) = 0 Then
    Exit Sub
End If

     timZm = 2
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "执行报表"
    mod1.cmd.Parameters("@NBLX") = "开票编辑"
    mod1.cmd.Parameters("@bh") = Trim(Str(Hid))
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "添加"
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtKPJE.Text)
    mod1.cmd.Parameters("@mm2") = KPid
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = txtKPRQ.Text
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        cmdDel.Enabled = False
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
End Sub

Private Sub cmdAd_Click()
On Error Resume Next

If txtKDRQ.Text = "" Or Val(txtKDJE.Text) = 0 Then
    Exit Sub
End If

     timZm = 1
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "执行报表"
    mod1.cmd.Parameters("@NBLX") = "开单编辑"
    mod1.cmd.Parameters("@bh") = Trim(Str(Hid))
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "添加"
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtKDJE.Text)
    mod1.cmd.Parameters("@mm2") = Kid
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = txtKDRQ.Text
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        cmdDel.Enabled = False
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
End Sub

Private Sub cmdD_Click()
On Error Resume Next

Dim ii As Integer
If KPid = 0 Then
    Exit Sub
End If
ii = MsgBox("是否删除此项开单记录？", vbYesNo + vbQuestion, "请确认")
If ii = vbNo Then Exit Sub

     timZm = 2
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "执行报表"
    mod1.cmd.Parameters("@NBLX") = "开票编辑"
    mod1.cmd.Parameters("@bh") = Trim(Str(Hid))
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "删除"
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtKPJE.Text)
    mod1.cmd.Parameters("@mm2") = KPid
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = txtKPRQ.Text
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        cmdDel.Enabled = False
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
End Sub

Private Sub cmdDe_Click()
On Error Resume Next
Dim ii As Integer
If Kid = 0 Then
    Exit Sub
End If
ii = MsgBox("是否删除此项开单记录？", vbYesNo + vbQuestion, "请确认")
If ii = vbNo Then Exit Sub

     timZm = 1
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "执行报表"
    mod1.cmd.Parameters("@NBLX") = "开单编辑"
    mod1.cmd.Parameters("@bh") = Trim(Str(Hid))
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "删除"
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtKDJE.Text)
    mod1.cmd.Parameters("@mm2") = Kid
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = txtKDRQ.Text
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        cmdDel.Enabled = False
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
End Sub

Private Sub cmdG_Click()
On Error Resume Next

Dim ii As Integer
If KPid = 0 Then
    Exit Sub
End If

     timZm = 2
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "执行报表"
    mod1.cmd.Parameters("@NBLX") = "开票编辑"
    mod1.cmd.Parameters("@bh") = Trim(Str(Hid))
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "更新"
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtKPJE.Text)
    mod1.cmd.Parameters("@mm2") = KPid
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = txtKPRQ.Text
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        cmdDel.Enabled = False
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
End Sub

Private Sub cmdGx_Click()
On Error Resume Next
Dim ii As Integer
If Kid = 0 Then
    Exit Sub
End If

     timZm = 1
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "执行报表"
    mod1.cmd.Parameters("@NBLX") = "开单编辑"
    mod1.cmd.Parameters("@bh") = Trim(Str(Hid))
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "更新"
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtKDJE.Text)
    mod1.cmd.Parameters("@mm2") = Kid
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = txtKDRQ.Text
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        cmdDel.Enabled = False
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
End Sub


Private Sub cmdGx1_Click()
On Error Resume Next
Dim ii As Integer
If Fid = 0 Then
    Exit Sub
End If

     timZm = 2
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "执行报表"
    mod1.cmd.Parameters("@NBLX") = "票收编辑"
    mod1.cmd.Parameters("@bh") = Fid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "更新"
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtKPJE.Text)
    mod1.cmd.Parameters("@mm2") = Val(txtJSJE.Text)
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = txtKPRQ.Text
    If Val(txtKPJE.Text) = 0 Then
        mod1.cmd.Parameters("@md1") = "1777-7-7"
    End If
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        cmdDel.Enabled = False
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
End Sub


Private Sub cmdYadd_Click()
Dim tt As String
Dim Ra
Dim TSL As Single
On Error Resume Next
If ACid = 0 Then Exit Sub

If txtYRQ.Text = "" Or Val(txtYSL.Text) = 0 Then
    Exit Sub
End If

'判断是否数量超出
tt = "Select sum(ysl) from hta150510 where acid=" & ACid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
If IsNull(Ra(0, 0)) = True Then
    TSL = 0
Else
    TSL = Val(Ra(0, 0))
End If
If TSL + Val(txtYSL.Text) > Val(txtYSL.ToolTipText) Then
    MsgBox "超出数额！"
    Exit Sub
End If

     timZm = 3
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "执行报表"
    mod1.cmd.Parameters("@NBLX") = "要货编辑"
    mod1.cmd.Parameters("@bh") = Trim(Str(ACid))
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "添加"
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtYSL.Text)
    mod1.cmd.Parameters("@mm2") = YHid
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = txtYRQ.Text
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        cmdDel.Enabled = False
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
End Sub

Private Sub cmdYdel_Click()
On Error Resume Next
Dim ii As Integer
If YHid = 0 Then Exit Sub
ii = MsgBox("是否删除此条记录? ", vbYesNo + vbQuestion, "请注意")
If ii = vbNo Then Exit Sub

If txtYRQ.Text = "" Or Val(txtYSL.Text) = 0 Then
    Exit Sub
End If

     timZm = 3
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "执行报表"
    mod1.cmd.Parameters("@NBLX") = "要货编辑"
    mod1.cmd.Parameters("@bh") = Trim(Str(ACid))
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "删除"
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtYSL.Text)
    mod1.cmd.Parameters("@mm2") = YHid
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = txtYRQ.Text
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        cmdDel.Enabled = False
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
End Sub


Private Sub cmdYGx_Click()
Dim tt As String
Dim Ra
Dim TSL As Single
On Error Resume Next
Dim ii As Integer
If YHid = 0 Then Exit Sub

If txtYRQ.Text = "" Or Val(txtYSL.Text) = 0 Then
    Exit Sub
End If

'判断是否数量超出
tt = "Select sum(ysl) from hta150510 where acid=" & ACid & " and yhid<>" & YHid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.EOF = False Then
    Ra = mod1.HTP.GetRows
Else
    Ra(0, 0) = 0
End If
mod1.HTP.Close
Set mod1.HTP = Nothing
If IsNull(Ra(0, 0)) = True Then
    TSL = 0
Else
    TSL = Val(Ra(0, 0))
End If
If TSL + Val(txtYSL.Text) > Val(txtYSL.ToolTipText) Then
    MsgBox "超出数额！"
    Exit Sub
End If

     timZm = 3
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "执行报表"
    mod1.cmd.Parameters("@NBLX") = "要货编辑"
    mod1.cmd.Parameters("@bh") = Trim(Str(ACid))
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "更新"
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtYSL.Text)
    mod1.cmd.Parameters("@mm2") = YHid
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = txtYRQ.Text
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        cmdDel.Enabled = False
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
End Sub


Private Sub dtgBr_Click()
dtgN1.Row = dtgBr.Row
dtgN1.Col = 11: Fid = Val(dtgN1.Text)
dtgN1.Col = 12: Kid = Val(dtgN1.Text)
dtgN1.Col = 13: KPid = Val(dtgN1.Text)
dtgN1.Col = 4: txtKDRQ.Text = dtgN1.Text
dtgN1.Col = 5: txtKDJE.Text = dtgN1.Text
dtgN1.Col = 6: txtKPRQ.Text = dtgN1.Text
dtgN1.Col = 7: txtKPJE.Text = dtgN1.Text
'''''dtgN1.Col = 8: txtWKPJE.Text = dtgN1.Text
'''''dtgN1.Col = 9: txtJSJE.Text = dtgN1.Text
'''''dtgN1.Col = 10: txtWJSJE.Text = dtgN1.Text
End Sub

Private Sub dtgDC_Click()
dtgN2.Row = dtgDC.Row
dtgN2.Col = 17: ACid = Val(dtgN2.Text)
dtgN2.Col = 18: YHid = Val(dtgN2.Text)
dtgN2.Col = 19: Cgid = Val(dtgN2.Text)
dtgN2.Col = 20: Gid = Val(dtgN2.Text)
dtgN2.Col = 21: Hid = Val(dtgN2.Text)

dtgN2.Col = 3: txtYRQ.Text = dtgN2.Text
dtgN2.Col = 4: txtYSL.Text = dtgN2.Text
dtgN2.Col = 2: txtYSL.ToolTipText = dtgN2.Text

End Sub

Private Sub dtpkdRq_CloseUp()
dtpKDRQ.Visible = False
txtKDRQ.Text = DateSerial(Year(dtpKDRQ.Value), Month(dtpKDRQ.Value), Day(dtpKDRQ.Value))
txtKDRQ.Visible = True
End Sub

Private Sub dtpKPRQ_CloseUp()
dtpKPRQ.Visible = False
txtKPRQ.Visible = True
txtKPRQ.Text = DateSerial(Year(dtpKPRQ.Value), Month(dtpKPRQ.Value), Day(dtpKPRQ.Value))

End Sub


Private Sub dtpYRQ_CloseUp()
txtYRQ.Text = DateSerial(Year(dtpYRQ.Value), Month(dtpYRQ.Value), Day(dtpYRQ.Value))
txtYRQ.Visible = True
dtpYRQ.Visible = False
End Sub

Private Sub Form_Load()
Me.Left = 0: Me.Top = 0
Me.Width = mod1.FWidth + 500
Me.Height = mod1.FHeight
TabMxc.Width = Me.Width
TabMxc.Height = Me.Height
End Sub

Public Sub dtgbrFF()
dtgBr.Clear
dtgBr.Cols = 16
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "项目名称": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "合同金额": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "应收日期": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "应收金额": dtgBr.CellFontBold = True
dtgBr.Col = 4: dtgBr.Text = "开单日期": dtgBr.CellFontBold = True
dtgBr.Col = 5: dtgBr.Text = "开单金额": dtgBr.CellFontBold = True
dtgBr.Col = 6: dtgBr.Text = "开票日期": dtgBr.CellFontBold = True
dtgBr.Col = 7: dtgBr.Text = "开票金额": dtgBr.CellFontBold = True
dtgBr.Col = 8: dtgBr.Text = "未开票金额": dtgBr.CellFontBold = True
dtgBr.Col = 9: dtgBr.Text = "已结算金额": dtgBr.CellFontBold = True
dtgBr.Col = 10: dtgBr.Text = "未结算金额": dtgBr.CellFontBold = True
dtgBr.Col = 11: dtgBr.Text = Fid
dtgBr.Col = 12: dtgBr.Text = Kid
dtgBr.Col = 13: dtgBr.Text = KPid
dtgBr.Col = 14: dtgBr.Text = Jid
dtgBr.Col = 15: dtgBr.Text = Aid
dtgBr.Col = 1: dtgBr.Row = 1: dtgBr.Text = "合计：": dtgBr.CellFontBold = True
dtgBr.ColWidth(0) = 2550
dtgBr.ColWidth(2) = 1100: dtgBr.ColWidth(4) = 1100: dtgBr.ColWidth(6) = 1100: dtgBr.ColWidth(8) = 1100: dtgBr.ColWidth(9) = 1100: dtgBr.ColWidth(10) = 1100
dtgBr.ColWidth(11) = 0
dtgBr.ColWidth(12) = 0
dtgBr.ColWidth(13) = 0
dtgBr.ColWidth(14) = 0
dtgBr.ColWidth(15) = 0
dtgN1.Clear
dtgN1.Cols = 16
End Sub

Public Sub Bound1(Hid As Long)
Dim tt As String
Dim Ra, Rb, RC, RD, RE
Dim La As Integer: Dim Lb As Integer: Dim Ld As Integer: Dim Le As Integer
Dim oo As Integer
Dim hg As Single
Dim HG1 As Single '开单合计
Dim HG2 As Single '开票合计
Dim HG3 As Single '未开票合计
Dim HG4 As Single '结算金额
Dim HG5 As Single '未结算金额
Dim HTZE As Single
Call Me.dtgbrFF
On Error Resume Next
tt = "select rq,yingfJe,fid from htping1 where htbh='" & Trim(Str(Hid)) & "' order by fid;" & _
    "SELECT KDRQ,KDJE,Kid FROM HTKD WHERE HID=" & Hid & " order by kid;" & _
    "select xmmc,htqy,htqy1,htze from htping where hid=" & Hid & ";" & _
    "select kprq,kpje,kpid from htkp where hid=" & Hid & ";" & _
    "select rq,je,jid,Aid from htADetail where delf=1 and hid=" & Hid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows '商务条款
Set mod1.HTP = mod1.HTP.NextRecordset
RD = mod1.HTP.GetRows '开票明细
Set mod1.HTP = mod1.HTP.NextRecordset
RE = mod1.HTP.GetRows '收款明细
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1: Lb = UBound(Rb, 2) + 1: Ld = UBound(RD, 2) + 1: Le = UBound(RE, 2) + 1
dtgBr.Rows = La + 20: dtgN1.Rows = La + 20
hg = 0: HG1 = 0: HG2 = 0: HG3 = 0: HG4 = 0: HG5 = 0
For oo = 2 To La + 2
    dtgBr.Row = oo
    dtgBr.Col = 2: dtgBr.Text = Ra(0, oo - 2)
    dtgBr.Col = 3: dtgBr.Text = Ra(1, oo - 2)
'''''    dtgBr.Col = 6: dtgBr.Text = Ra(2, oo - 2)
'''''    If Ra(2, oo - 2) = "1777-7-7" Then
'''''        dtgBr.Text = ""
'''''    End If
'''''    dtgBr.Col = 7: dtgBr.Text = Ra(3, oo - 2)
'''''    dtgBr.Col = 8: dtgBr.Text = Val(Ra(1, oo - 2)) - IIf(Ra(3, oo - 2), Ra(3, oo - 2), 0)
'''''    HG3 = HG3 + Val(dtgBr.Text)
'''''    dtgBr.Col = 9: dtgBr.Text = Ra(4, oo - 2)
'''''    dtgBr.Col = 10: dtgBr.Text = Val(Ra(3, oo - 2)) - Val(Ra(4, oo - 2))
'''''    HG5 = HG5 + Val(dtgBr.Text)
    dtgBr.Col = 11: dtgBr.Text = Ra(2, oo - 2) 'fid
    
    hg = hg + Val(Ra(1, oo - 2)): HG2 = HG2 + Val(Ra(3, oo - 2)): HG4 = HG4 + Val(Ra(4, oo - 2))
    
    dtgN1.Row = oo
    dtgN1.Col = 2: dtgN1.Text = Ra(0, oo - 2)
    dtgN1.Col = 3: dtgN1.Text = Ra(1, oo - 2)
'''''    dtgN1.Col = 6: dtgN1.Text = Ra(2, oo - 2)
'''''    If Ra(2, oo - 2) = "1777-7-7" Then
'''''        dtgN1.Text = ""
'''''    End If
'''''    dtgN1.Col = 7: dtgN1.Text = Ra(3, oo - 2)
'''''    dtgN1.Col = 8: dtgN1.Text = Val(Ra(1, oo - 2)) - IIf(Ra(3, oo - 2), Ra(3, oo - 2), 0)
'''''    dtgN1.Col = 9: dtgN1.Text = Ra(4, oo - 2)
'''''    dtgN1.Col = 10: dtgN1.Text = IIf(Ra(3, oo - 2), Ra(3, oo - 2), 0) - IIf(Ra(4, oo - 2), Ra(4, oo - 2), 0)
    dtgN1.Col = 11: dtgN1.Text = Ra(2, oo - 2)

Next

'合计
dtgBr.Row = 1: dtgBr.Col = 3: dtgBr.Text = hg: dtgBr.CellFontBold = True
'''''''dtgBr.Col = 7: dtgBr.Text = HG2: dtgBr.CellFontBold = True '开票金额合计
'''''''dtgBr.Col = 8: dtgBr.Text = HG3: dtgBr.CellFontBold = True
'''''''dtgBr.Col = 9: dtgBr.Text = HG4: dtgBr.CellFontBold = True '结算金额合计
'''''''dtgBr.Col = 10: dtgBr.Text = HG5: dtgBr.CellFontBold = True


'开单绑定
For oo = 2 To Lb + 2
    dtgBr.Row = oo
    dtgBr.Col = 4: dtgBr.Text = Rb(0, oo - 2)
    dtgBr.Col = 5: dtgBr.Text = Rb(1, oo - 2)
    HG1 = HG1 + Val(Rb(1, oo - 2))
    dtgBr.Col = 12: dtgBr.Text = Rb(2, oo - 2)
    
    dtgN1.Row = oo
    dtgN1.Col = 4: dtgN1.Text = Rb(0, oo - 2)
    dtgN1.Col = 5: dtgN1.Text = Rb(1, oo - 2)
    dtgN1.Col = 12: dtgN1.Text = Rb(2, oo - 2)
Next
'合计
dtgBr.Row = 1: dtgBr.Col = 5: dtgBr.Text = HG1: dtgBr.CellFontBold = True

'项目名称与维保年限
dtgBr.Row = 2: dtgBr.Col = 0: dtgBr.Text = RC(0, 0)
dtgBr.Row = 2: dtgBr.Col = 1: dtgBr.Text = RC(3, 0): HTZE = RC(3, 0)
dtgBr.Row = 3: dtgBr.Text = RC(1, 0) & RC(2, 0)

'开票绑定
For oo = 2 To Ld - 1 + 2
    dtgBr.Row = oo
    dtgBr.Col = 6: dtgBr.Text = RD(0, oo - 2)
    dtgBr.Col = 7: dtgBr.Text = RD(1, oo - 2)
    HG2 = HG2 + Val(RD(1, oo - 2))
    dtgBr.Col = 8: dtgBr.Text = HTZE - HG2
    
    dtgBr.Col = 13: dtgBr.Text = RD(2, oo - 2)
    
    dtgN1.Row = oo
    dtgN1.Col = 6: dtgN1.Text = RD(0, oo - 2)
    dtgN1.Col = 7: dtgN1.Text = RD(1, oo - 2)
    dtgN1.Col = 8: dtgBr.Text = HTZE - HG2
    dtgN1.Col = 13: dtgN1.Text = RD(2, oo - 2)
Next
'合计
dtgBr.Row = 1: dtgBr.Col = 7: dtgBr.Text = HG2: dtgBr.CellFontBold = True

'收款绑定
For oo = 2 To Le - 1 + 2
    dtgBr.Row = oo
    dtgBr.Col = 9: dtgBr.Text = RE(1, oo - 2)
    HG4 = HG4 + Val(RE(1, oo - 2))
    dtgBr.Col = 10: dtgBr.Text = HG2 - HG4
    
    dtgBr.Col = 14: dtgBr.Text = RE(2, oo - 2) 'jid
    dtgBr.Col = 15: dtgBr.Text = RE(3, oo - 2) 'Aid
    dtgN1.Row = oo
    dtgN1.Col = 9: dtgN1.Text = RE(1, oo - 2)
    dtgN1.Col = 10: dtgN1.Text = HG2 - HG4
    dtgN1.Col = 14: dtgN1.Text = RD(2, oo - 2)
    dtgN1.Col = 15: dtgN1.Text = RD(3, oo - 2)
Next
'合计
dtgBr.Row = 1: dtgBr.Col = 9: dtgBr.Text = HG4: dtgBr.CellFontBold = True
End Sub

Public Sub edQing()
txtKDRQ.Text = ""
txtKDJE.Text = ""
dtpKDRQ.Value = mod1.DQda
dtpKDRQ.Visible = False
txtKDRQ.Visible = True

txtKPRQ.Text = ""
txtKPJE.Text = ""
dtpKPRQ.Value = mod1.DQda
txtJSJE.Text = ""
txtWJSJE.Text = ""

txtYRQ.Text = ""
txtYSL.Text = ""
txtYSL.ToolTipText = ""

txtGymc.Text = ""
txtGymc.ToolTipText = ""
txtCBH.Text = ""
txtCSL.Text = ""
txtCJE.Text = ""
End Sub

Private Sub Form_Resize()
TabMxc.Width = Me.Width
dtgDC.Width = Me.Width
dtgBr.Width = Me.Width
Frame1.Width = Me.Width
Frame2.Width = Me.Width
cmdBack.Left = Me.Width - cmdBack.Width - 500
End Sub

Private Sub Frame1_Click()
dtpKPRQ.Visible = False
txtKPRQ.Visible = True
dtpKDRQ.Visible = False
txtKDRQ.Visible = True
End Sub

Private Sub Frame4_Click()
dtpKPRQ.Visible = False
txtKPRQ.Visible = True
dtpKDRQ.Visible = False
txtKDRQ.Visible = True
End Sub

Private Sub Frame5_Click()
dtpKPRQ.Visible = False
txtKPRQ.Visible = True
dtpKDRQ.Visible = False
txtKDRQ.Visible = True
End Sub

Private Sub timQuit_Timer()
Dim Rz
Dim Lz As Integer
Dim Rb
Dim Lb As Integer
Dim RD
Dim Ld As Integer
On Error Resume Next
Dim ii As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '如果为开单编辑
    Call Me.Bound1(Hid)
ElseIf timZm = 2 Then '开票编辑
    Call Me.Bound1(Hid)
ElseIf timZm = 3 Then '要货编辑
    Call Me.Bound2(Hid)
End If
timQuit.Enabled = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim Bid As Long
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mm2,mt2,mt1,mt3,mt4 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    mod1.Ti = 5
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    timWait.Enabled = False
    If timZm = 1 Then '开单编辑
        tt = tt
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    timWait.Enabled = False
    Unload frmWaitA
    Me.Enabled = True
    Exit Sub
'''''    If timZm = 1 Then
'''''        NiceButton1.Enabled = False
'''''    End If
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
'''''    If timZm = 1 Then
'''''        NiceButton1.Enabled = False
'''''    End If
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub

Private Sub txtJSJE_Change()
txtWJSJE.Text = Val(txtKPJE.Text) - Val(txtJSJE.Text)
End Sub

Private Sub txtKDRQ_Click()
txtKDRQ.Visible = False
dtpKDRQ.Visible = True
End Sub


Private Sub txtKPJE_Change()
dtgN1.Col = 3
txtWKPJE.Text = Val(dtgN1.Text) - Val(txtKPJE.Text)
txtWJSJE.Text = Val(txtKPJE.Text) - Val(txtJSJE.Text)
End Sub

Private Sub txtKPRQ_Click()
dtpKPRQ.Visible = True
txtKPRQ.Visible = False
End Sub



Public Sub dtgDCFF()
dtgDC.Clear: dtgN2.Clear
dtgDC.Cols = 22
dtgDC.Row = 0
dtgDC.Col = 0: dtgDC.Text = "货品编码": dtgDC.CellFontBold = True
dtgDC.Col = 1: dtgDC.Text = "货品名称": dtgDC.CellFontBold = True
dtgDC.Col = 2: dtgDC.Text = "合同数量": dtgDC.CellFontBold = True
dtgDC.Col = 3: dtgDC.Text = "要货日期": dtgDC.CellFontBold = True
dtgDC.Col = 4: dtgDC.Text = "要货数量": dtgDC.CellFontBold = True
dtgDC.Col = 5: dtgDC.Text = "采购供应商": dtgDC.CellFontBold = True
dtgDC.Col = 6: dtgDC.Text = "采购订单号": dtgDC.CellFontBold = True
dtgDC.Col = 7: dtgDC.Text = "采购数量": dtgDC.CellFontBold = True
dtgDC.Col = 8: dtgDC.Text = "采购订单金额": dtgDC.CellFontBold = True
dtgDC.Col = 9: dtgDC.Text = "采购付款日期": dtgDC.CellFontBold = True
dtgDC.Col = 10: dtgDC.Text = "采购付款金额": dtgDC.CellFontBold = True
dtgDC.Col = 11: dtgDC.Text = "采购未付款金额": dtgDC.CellFontBold = True
dtgDC.Col = 12: dtgDC.Text = "采购收货日期": dtgDC.CellFontBold = True
dtgDC.Col = 13: dtgDC.Text = "采购收货数量": dtgDC.CellFontBold = True
dtgDC.Col = 14: dtgDC.Text = "采购收货金额": dtgDC.CellFontBold = True
dtgDC.Col = 15: dtgDC.Text = "发货日期": dtgDC.CellFontBold = True
dtgDC.Col = 16: dtgDC.Text = "发货数量": dtgDC.CellFontBold = True
dtgDC.Col = 17: dtgDC.Text = ACid
dtgDC.Col = 18: dtgDC.Text = YHid
dtgDC.Col = 19: dtgDC.Text = Cgid
dtgDC.Col = 20: dtgDC.Text = Gid
dtgDC.Col = 21: dtgDC.Text = Hid

dtgDC.Row = 1: dtgDC.Col = 2: dtgDC.Text = "合计：": dtgDC.CellFontBold = True
dtgDC.ColWidth(1) = 4950
dtgDC.ColWidth(5) = 3540
dtgDC.ColWidth(6) = 1300
dtgDC.ColWidth(7) = 1300
dtgDC.ColWidth(8) = 1300
dtgDC.ColWidth(9) = 1500
dtgDC.ColWidth(10) = 1500
dtgDC.ColWidth(11) = 1500
dtgDC.ColWidth(12) = 1500
dtgDC.ColWidth(13) = 1500
dtgDC.ColWidth(17) = 0
dtgDC.ColWidth(18) = 0
dtgDC.ColWidth(19) = 0
dtgDC.ColWidth(20) = 0
dtgDC.ColWidth(21) = 0

dtgN2.Cols = 22
dtgDC.FixedCols = 3
End Sub

Public Sub Bound2(Hid As Long)
Dim oo As Long
Dim ii As Integer
Dim tt As String
Dim Ra, Rb
Dim La As Long
Dim OBh As String
Dim CF As Boolean
Call Me.dtgDCFF
tt = "select bh,mc,sl,yrq,ysl,gymc,cbh,csl,cje,'',0,0,'',0,0,'',0,acid,yhid,cgid,gid,hid from hta150510 where hid=" & Hid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows '采购单
'''Set mod1.HTP = mod1.HTP.NextRecordset
'''Rb = mod1.HTP.GetRows
'''Set mod1.HTP = mod1.HTP.NextRecordset
'''RC = mod1.HTP.GetRows '商务条款
'''Set mod1.HTP = mod1.HTP.NextRecordset
'''RD = mod1.HTP.GetRows '开票明细
'''Set mod1.HTP = mod1.HTP.NextRecordset
'''RE = mod1.HTP.GetRows '收款明细
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgDC.Rows = La + 20: dtgN2.Rows = La + 20
On Error Resume Next
For oo = 2 To La + 2 - 1
    dtgDC.Row = oo: dtgN2.Row = oo: CF = False
    For ii = 0 To 21
    
        dtgDC.Col = ii: dtgDC.Text = Ra(ii, oo - 2)
        dtgN2.Col = ii: dtgN2.Text = Ra(ii, oo - 2)
        If ii >= 0 And ii <= 2 Then
            If ii = 0 Then
                dtgN2.Col = ii
                If OBh = dtgN2.Text Then
                    CF = True
                Else
                    OBh = dtgN2.Text
                End If
            End If
            If CF = True Then
                dtgDC.Text = ""
            End If
        End If
    Next
Next
End Sub

Private Sub txtYRQ_Click()
dtpYRQ.Visible = True
txtYRQ.Visible = False
dtpYRQ.Value = mod1.DQda
If txtYRQ.Text <> "" Then
    dtpYRQ.Value = txtYRQ.Text
End If
End Sub


