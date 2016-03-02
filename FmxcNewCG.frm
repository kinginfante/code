VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FmxcNewCG 
   Caption         =   "销售报表"
   ClientHeight    =   9090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   ScaleHeight     =   9090
   ScaleWidth      =   15210
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1950
      Top             =   0
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   420
   End
   Begin VB.Frame frmMod 
      Caption         =   "编辑"
      Height          =   5445
      Left            =   9420
      TabIndex        =   14
      Top             =   90
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton cmdZD 
         Caption         =   "整单完成"
         Height          =   615
         Left            =   4140
         TabIndex        =   28
         Top             =   4680
         Width           =   675
      End
      Begin VB.TextBox txtHtbh 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   4230
         Width           =   2745
      End
      Begin VB.TextBox txtBz 
         Height          =   525
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Text            =   "FmxcNewCG.frx":0000
         Top             =   2280
         Width           =   2775
      End
      Begin VB.ComboBox comW 
         Height          =   300
         ItemData        =   "FmxcNewCG.frx":0006
         Left            =   2040
         List            =   "FmxcNewCG.frx":0010
         TabIndex        =   23
         Top             =   2910
         Width           =   2775
      End
      Begin VB.TextBox txtddGG 
         Height          =   300
         Left            =   2040
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtGZHZ 
         Height          =   300
         Left            =   2040
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtDHQK 
         Height          =   300
         Left            =   2040
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1830
         Width           =   2775
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "提交"
         Height          =   585
         Left            =   4110
         Picture         =   "FmxcNewCG.frx":0022
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3390
         Width           =   675
      End
      Begin VB.Label lblHtbh 
         Caption         =   "整单完成"
         Height          =   285
         Left            =   840
         TabIndex        =   26
         ToolTipText     =   "双击选择"
         Top             =   4290
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "备注"
         Height          =   255
         Left            =   180
         TabIndex        =   24
         Top             =   2340
         Width           =   1395
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "状态"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   2955
         Width           =   1395
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "订单给供应商"
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   883
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "盖章回传情况"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   1376
         Width           =   1395
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "到货情况"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   1869
         Width           =   1395
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H008080FF&
         Height          =   465
         Left            =   120
         Top             =   330
         Visible         =   0   'False
         Width           =   4965
      End
   End
   Begin VB.ComboBox comWCF 
      Height          =   300
      ItemData        =   "FmxcNewCG.frx":068C
      Left            =   10290
      List            =   "FmxcNewCG.frx":0696
      TabIndex        =   13
      Text            =   "未完成"
      Top             =   8820
      Width           =   1005
   End
   Begin VB.ComboBox comZT 
      Height          =   300
      ItemData        =   "FmxcNewCG.frx":06A8
      Left            =   3360
      List            =   "FmxcNewCG.frx":06B2
      TabIndex        =   11
      Text            =   "未完成"
      Top             =   8790
      Width           =   1065
   End
   Begin VB.ComboBox comLX 
      Height          =   300
      ItemData        =   "FmxcNewCG.frx":06C6
      Left            =   5670
      List            =   "FmxcNewCG.frx":06D6
      TabIndex        =   7
      Text            =   "年度"
      Top             =   8790
      Width           =   1365
   End
   Begin VB.TextBox txtZ 
      Height          =   270
      Left            =   7200
      TabIndex        =   6
      Top             =   8790
      Width           =   2085
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "查询"
      Height          =   315
      Left            =   11610
      TabIndex        =   5
      Top             =   8790
      Width           =   975
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制"
      Height          =   255
      Left            =   12630
      TabIndex        =   4
      Top             =   8820
      Width           =   1035
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      Height          =   285
      Left            =   13860
      TabIndex        =   3
      Top             =   8820
      Width           =   945
   End
   Begin VB.ComboBox comG 
      Height          =   300
      ItemData        =   "FmxcNewCG.frx":06FC
      Left            =   1080
      List            =   "FmxcNewCG.frx":0706
      TabIndex        =   2
      Text            =   "合同评审单"
      Top             =   8790
      Width           =   1605
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   1635
      Left            =   6600
      TabIndex        =   1
      Top             =   5580
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   2884
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgC 
      Height          =   8685
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15319
      _Version        =   393216
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      Caption         =   "状态"
      Height          =   285
      Left            =   9510
      TabIndex        =   12
      Top             =   8850
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "状态"
      Height          =   285
      Left            =   2880
      TabIndex        =   10
      Top             =   8820
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "查询方式"
      Height          =   315
      Left            =   4650
      TabIndex        =   9
      Top             =   8820
      Width           =   885
   End
   Begin VB.Label lblddd 
      Caption         =   "类型"
      Height          =   285
      Left            =   390
      TabIndex        =   8
      Top             =   8820
      Width           =   465
   End
End
Attribute VB_Name = "FmxcNewCG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim companyId As Integer
Dim Hid As Long
Dim Did As Long
Dim Bid As Long
Dim ETT As String
Dim Fl As String '所选择的类型,(合同还是追加单)
Dim timZm As Integer
Public Sub Bound(Ra, La As Long)
Dim oo As Long
Dim ii As Long
Dim ZBQ As String
Dim THTBH As String
On Error Resume Next
dtgC.Visible = False
Call Me.dtgCFF
Me.dtgC.Rows = La + 100
Me.dtgN.Rows = La + 100
For oo = 1 To La
    dtgC.Row = oo
    dtgN.Row = oo
    For ii = 0 To 14
        dtgC.Col = ii
        dtgN.Col = ii
        dtgC.Text = Replace(Ra(ii, oo - 1), Chr(13) + Chr(10), "")
        dtgN.Text = Ra(ii, oo - 1)
        If ii = 2 Then
            dtgC.Text = DateSerial(Year(Ra(ii, oo - 1)), Month(Ra(ii, oo - 1)), Day(Ra(ii, oo - 1)))
        End If
        If ii = 14 Then '判断是否显示人工
            ZBQ = Ra(ii, oo - 1)
            If IsNull(Ra(3, oo - 1)) = True Then
                dtgC.Col = 3: dtgN.Col = 3
                dtgC.Text = ZBQ: dtgN.Text = ZBQ
            End If
        End If

    Next
        If oo = 1 Then THTBH = Ra(0, 0)
        If oo > 1 Then
            If Ra(0, oo - 1) = THTBH Then
                dtgC.Col = 0: dtgC.Text = ""
'                dtgN.Col = 0: dtgN.Text = ""
                dtgC.Col = 1: dtgC.Text = ""
'                dtgN.Col = 1: dtgN.Text = ""
                dtgC.Col = 2: dtgC.Text = ""
'                dtgN.Col = 2: dtgN.Text = ""
            Else
                THTBH = Ra(0, oo - 1)
            End If
        End If
        dtgC.RowHeight(oo) = dtgC.RowHeight(0)
        If dtgC.Col = 3 Then

            frmZu.lblDtg.Width = dtgC.ColWidth(3)
            dtgN.Col = 3
            frmZu.lblDtg.Caption = dtgN.Text

                dtgC.RowHeight(oo) = frmZu.lblDtg.Height * 2

        End If
Next
dtgC.Visible = True
End Sub
Private Sub cmdC_Click()
Dim tt As String
Dim WCF As Integer
Dim Ra
Dim La As Long
If comWCF.Text = "未完成" Then
    WCF = 0
Else
    WCF = 1
End If
If comG.Text = "合同评审单" Then
    Select Case comLX.Text
    Case "年度"
        If WCF = 0 Then
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid,zbq" & _
            " from fmxcCGBx where year(ddrq)=" & txtZ.Text & " and wcf=" & WCF & " order by ddrq desc,htbh"
        Else
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid,zbq" & _
            " from fmxcCGB where year(ddrq)=" & txtZ.Text & " and wcf=" & WCF & " order by ddrq desc,htbh"
        End If
                tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid,zbq" & _
            " from fmxcCGB where year(ddrq)=" & txtZ.Text & " and wcf=" & WCF & " order by ddrq desc,htbh"
    Case "合同编号"
        If WCF = 0 Then
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid,zbq" & _
            " from fmxcCGBx where htbh like '%" & txtZ.Text & "%' and wcf=" & WCF & " order by ddrq desc,htbh"
        Else
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid,zbq" & _
            " from fmxcCGB where htbh like '%" & txtZ.Text & "%' and wcf=" & WCF & " order by ddrq desc,htbh"
        End If
                tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid,zbq" & _
            " from fmxcCGB where htbh like '%" & txtZ.Text & "%' and wcf=" & WCF & " order by ddrq desc,htbh"
    Case "业务员"
        If WCF = 0 Then
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid,zbq" & _
            " from fmxcCGBx where ywy='" & txtZ.Text & "' and wcf=" & WCF & " order by ddrq desc,htbh"
        Else
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid,zbq" & _
            " from fmxcCGB where ywy='" & txtZ.Text & "' and wcf=" & WCF & " order by ddrq desc,htbh"
        End If
                tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid,zbq" & _
            " from fmxcCGB where ywy='" & txtZ.Text & "' and wcf=" & WCF & " order by ddrq desc,htbh"
    Case "货品编码"
        If WCF = 0 Then
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid,zbq" & _
            " from fmxcCGBx where ljbh='" & txtZ.Text & "' and wcf=" & WCF & " order by ddrq desc,htbh"
        Else
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid,zbq" & _
            " from fmxcCGB where ljbh='" & txtZ.Text & "' and wcf=" & WCF & " order by ddrq desc,htbh"
        End If
                tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid,zbq" & _
            " from fmxcCGB where ljbh='" & txtZ.Text & "' and wcf=" & WCF & " order by ddrq desc,htbh"
    End Select
ElseIf comG.Text = "成本追加单" Then
    Select Case comLX.Text
    Case "年度"
        If WCF = 0 Then
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid" & _
            " from fmxcCGBZuix where year(ddrq)=" & txtZ.Text & " and wcf=" & WCF & " order by ddrq desc,htbh"
        Else
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid" & _
            " from fmxcCGBZui where year(ddrq)=" & txtZ.Text & " and wcf=" & WCF & " order by ddrq desc,htbh"
        End If
                tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid" & _
            " from fmxcCGBZui where year(ddrq)=" & txtZ.Text & " and wcf=" & WCF & " order by ddrq desc,htbh"
    Case "合同编号"
        If WCF = 0 Then
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid" & _
            " from fmxcCGBZuix where htbh like '%" & txtZ.Text & "%' and wcf=" & WCF & " order by ddrq desc,htbh"
        Else
         tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid" & _
            " from fmxcCGBZui where htbh like '%" & txtZ.Text & "%' and wcf=" & WCF & " order by ddrq desc,htbh"
        End If
                 tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid" & _
            " from fmxcCGBZui where htbh like '%" & txtZ.Text & "%' and wcf=" & WCF & " order by ddrq desc,htbh"
    Case "业务员"
        If WCF = 0 Then
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid" & _
            " from fmxcCGBZuix where ywy='" & txtZ.Text & "' and wcf=" & WCF & " order by ddrq desc,htbh"
        Else
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid" & _
            " from fmxcCGBZui where ywy='" & txtZ.Text & "' and wcf=" & WCF & " order by ddrq desc,htbh"
        End If
                tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid" & _
            " from fmxcCGBZui where ywy='" & txtZ.Text & "' and wcf=" & WCF & " order by ddrq desc,htbh"
    Case "货品编码"
        If WCF = 0 Then
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid" & _
            " from fmxcCGBZuix where ljbh='" & txtZ.Text & "' and wcf=" & WCF & " order by ddrq desc,htbh"
        Else
        tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid" & _
            " from fmxcCGBZui where ljbh='" & txtZ.Text & "' and wcf=" & WCF & " order by ddrq desc,htbh"
        End If
                tt = "select htbh,xmmc,ddrq,ljmc+'('+ljbh+')',sl,mc,ddgg,gzhz,dhqk,ywy,bz,lid,wcf,bid" & _
            " from fmxcCGBZui where ljbh='" & txtZ.Text & "' and wcf=" & WCF & " order by ddrq desc,htbh"
    End Select

End If
ETT = tt
On Error Resume Next
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
Call Me.Bound(Ra, La)
End Sub

Private Sub cmdCopy_Click()
dtgC.FixedRows = 0
dtgC.FixedCols = 0

dtgC.Col = 0
dtgC.Row = 0
dtgC.ColSel = 11
dtgC.RowSel = dtgC.Rows - 50
Clipboard.Clear
Clipboard.SetText dtgC.Clip
dtgC.FixedRows = 1
dtgC.FixedCols = 1
End Sub

Private Sub cmdMod_Click()
If mod1.DName <> "乔继敏" And mod1.DName <> "顾珣" And mod1.DName <> "陈文超" Then Exit Sub
If Hid = 0 And comG.Text = "合同评审单" Then Exit Sub
If frmMod.Visible = False Then
    frmMod.Visible = True
    cmdSave.Enabled = True
Else
    frmMod.Visible = False
End If
End Sub

Private Sub cmdSave_Click()
If Did = 0 And Hid = 0 Then Exit Sub

'If Hid = 0 Then Exit Sub
timZm = 1 '保存
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "采购报表"
    mod1.cmd.Parameters("@NBLX") = "保存"
    mod1.cmd.Parameters("@bh") = Did
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtddGG.Text
    mod1.cmd.Parameters("@mt2") = txtGZHZ.Text
    mod1.cmd.Parameters("@mt3") = txtDHQK.Text
    mod1.cmd.Parameters("@mt5") = Fl
    
    mod1.cmd.Parameters("@mlt1") = txtBz.Text
    mod1.cmd.Parameters("@mm1") = Bid
    mod1.cmd.Parameters("@mt11") = Hid
    If comW.Text = "未完成" Then
        mod1.cmd.Parameters("@mb1") = 0
    Else
        mod1.cmd.Parameters("@mb1") = 1
    End If
    If txtHtbh.Text = "" Then
        mod1.cmd.Parameters("@mb2") = 0
    Else
        mod1.cmd.Parameters("@mb2") = 1
    End If
        mod1.cmd.Parameters("@md1") = Null
'Exit Sub
   
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 1 Then '保存
            cmdSave.Enabled = False
        End If
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


cmdSave.Enabled = False
End Sub

Private Sub cmdZD_Click()
    If Fl = "合同评审单" Then
        Me.dtgN.Col = 0
        txtHtbh.Text = Me.dtgN.Text
        Me.dtgN.Col = 13
        txtHtbh.ToolTipText = Val(dtgN.Text)
        
        txtHtbh.Text = txtHtbh.Text & "(" & txtHtbh.ToolTipText & ")"
    
    Else
        Me.dtgN.Col = 0
        txtHtbh.Text = Me.dtgN.Text
        Me.dtgN.Col = 13
        txtHtbh.ToolTipText = Val(dtgN.Text)
        
        txtHtbh.Text = txtHtbh.Text & "(" & txtHtbh.ToolTipText & ")"
    End If
Bid = Val(txtHtbh.ToolTipText)

If Bid = 0 Then Exit Sub

timZm = 2 '整单完成"
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "采购报表"
    mod1.cmd.Parameters("@NBLX") = "整单完成"
    mod1.cmd.Parameters("@bh") = Did
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt5") = Fl
    
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Bid
    If comW.Text = "未完成" Then
        mod1.cmd.Parameters("@mb1") = 0
    Else
        mod1.cmd.Parameters("@mb1") = 1
    End If
        mod1.cmd.Parameters("@md1") = Null
'Exit Sub
   
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 1 Then '保存
            cmdSave.Enabled = False
        End If
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


cmdSave.Enabled = False
End Sub

Private Sub dtgC_Click()
dtgN.Row = dtgC.Row
dtgN.Col = 0
Hid = Val(Right(dtgN.Text, 5))
    txtHtbh.Text = ""
    dtgN.Col = 6: txtddGG.Text = dtgN.Text
    dtgN.Col = 7: txtGZHZ = dtgN.Text
    dtgN.Col = 8: txtDHQK.Text = dtgN.Text
    dtgN.Col = 10: txtBz.Text = dtgN.Text
    dtgN.Col = 12:
    If dtgN.Text = "True" Then
        comW.Text = "完成"
    Else
        comW.Text = "未完成"
    End If
    dtgN.Col = 11: Did = Val(dtgN.Text)
    dtgN.Col = 13: Bid = Val(dtgN.Text)
Fl = comG.Text

End Sub

Private Sub dtgC_DblClick()
Dim tt As String
Dim xZ As String


Dim Bid As Long
Dim ZL As String
Dim Ra
'Dim Lid As String
On Error Resume Next



frmWait.Visible = True
frmWait.ZOrder 0
frmWait.Refresh

Exit Sub

'''''If dtgC.Col > 0 Then
'''''
'''''    Exit Sub
'''''End If
If NewF = 0 Then
    If xZ = "C. 维保合同" Or xZ = "D. 维修合同" Then
    'mod1.comJZ = False
    wbHTP.Visible = False
    Call modHt.wbQing
    
    
    tt = "Select * from htping where hid=" & Hid
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Call modHt.wbBound
    
    
    '打开材料表
    tt = "Select * from htSale where htbh='" & wbHTP.txtHtbh.Text & "'"
    wbMx.adoRGF.Recordset.Close
    wbMx.adoRGF.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set wbMx.dtgSale.DataSource = wbMx.adoRGF
    wbMx.lblChg.Caption = wbHTP.txtClcb1.Text
    
    '打开应收款表
    tt = "Select * from htping1 where htBh='" & wbHTP.lblHid.Caption & "' order by rq"
    frmFuK.adoHpt.Recordset.Close
    frmFuK.adoHpt.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set wbMx.dtgFk.DataSource = frmFuK.adoHpt
    
    '打开佣金表
    tt = "Select * from Yongjin where htBh='" & wbHTP.txtHtbh.Text & "' order by yId"
    frmYj.adoYj.Recordset.Close
    frmYj.adoYj.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set frmYj.dtgYJ.DataSource = frmYj.adoYj
    
    ''打开出工信息表(如果为评审阶段则不显示）
    'If wbHTP.optZ.Value = True Or wbHTP.optW.Value = True Then
    '    tt = "Select max(gzb.rq),max(gzb.wxWorker),sum(workXX.wTime),max(bhid)" & _
    '    "max(htbh) from gzb cross join workXX where gzb.bhid=workXX.bhid and gzb.htBh='" & _
    '    wbHTP.txtHtbh.Text & "' group by gzb.bhid"
    '    form2Htp.adoGzb.Recordset.Close
    '    form2Htp.adoGzb.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    '    Set wbMx.dtgGzb.DataSource = form2Htp.adoGzb
    'End If
    wbHTP.txtYj1.Visible = False
    wbHTP.txtYj2.Visible = False
    wbHTP.txtLr1.Visible = False
    wbHTP.txtLr2.Visible = False
    wbHTP.lblTcBe.Visible = False
    wbHTP.txtTcBe.Visible = False
    wbHTP.UpDa.Visible = False
    wbHTP.lblYj.Visible = False
    wbHTP.lblLr.Visible = False
    wbHTP.lblTC.Visible = False
    wbHTP.Visible = True
    Exit Sub
    End If
    
    
    
    
    
    
    
    
    
    
    '购销合同
    
    form2Htp.Visible = True
    mod1.workTt = ""
    mod1.workTt = "Select * from htPing where hid=" & Hid
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open mod1.workTt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    form2Htp.lblHtxz.Caption = ""
    
    Call modHt.htQing
    Call modHt.htBound '绑定合同评审单字段
    

    
    
    '打开收款表
    
    
    tt = "Select * from htPing1 where htBh='" & form2Htp.lblHid.Caption & "' order by rq"
    frmFuK.adoHpt.Recordset.Close
    frmFuK.adoHpt.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    
    
    Set frmFuK.dtgFk.DataSource = frmFuK.adoHpt
    
    'ft = "Select * from yiFk Where htBh='" & frmFuK.adoHpt.Recordset.Fields("htBh").Value & _
    '"' and yingRQ='" & frmFuK.adoHpt.Recordset.Fields("rq").Value & "' order by yiRq"
    'frmFuK.adoYf.Recordset.Close
    'frmFuK.adoYf.Recordset.Open ft, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'Set frmFuK.dtgYf.DataSource = frmFuK.adoYf
    
    '打开产品表
    tt = ""
    tt = "Select * from htSale Where htBh='" & form2Htp.txtHtbh.Text & "'"
    form2Htp.adoSale.Recordset.Close
    form2Htp.adoSale.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set form2Htp.dtgSale.DataSource = form2Htp.adoSale
    Set form2Htp.dtgYJ.DataSource = form2Htp.adoSale
    Set form2Htp.dtgZj.DataSource = form2Htp.adoSale
    
    ''打开“取自库存表”
    'tt = "Select * from kcJa where htBh='" & form2Htp.txtHtbh.Text & "'"
    'form2Htp.adoKu.Recordset.Close
    'form2Htp.adoKu.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'Set form2Htp.dtgKu.DataSource = form2Htp.adoKu
    
    ''打开采购表
    'ft = "Select * from CG Where htbh='" & form2Htp.txtHtbh.Text & "' and khmc<>'库存'"
    'frmAdo.adoTmp.Recordset.Close
    'frmAdo.adoTmp.Recordset.Open ft, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'Set form2Htp.dtgCG.DataSource = frmAdo.adoTmp
    
    '打开佣金表
    tt = "Select * from Yongjin where htBh='" & form2Htp.txtHtbh.Text & "' order by yId"
    frmYj.adoYj.Recordset.Close
    frmYj.adoYj.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set frmYj.dtgYJ.DataSource = frmYj.adoYj
    
    
    
    
    form2Htp.tabHt.TabEnabled(1) = True
    form2Htp.tabHt.TabEnabled(2) = True
    'End If
    
    
    
    
    
    
    
    form2Htp.tabHt.Tab = 0
    htBrowG.MousePointer = 0
    
    
        '佣金、利润2、提成不显示
        form2Htp.txtYj1.Visible = False
        form2Htp.txtYj2.Visible = False
        form2Htp.txtLr1.Visible = False
        form2Htp.txtLr2.Visible = False
        'form2Htp.txtTc1.Visible = False
        'form2Htp.txtTc2.Visible = False
        form2Htp.lblYj.Visible = False
        form2Htp.lblLr2.Visible = False
        'form2Htp.lblTc.Visible = False
ElseIf NewF = 1 Then
        Call modHt.NewQing
        Call modHt.NewLocked
        Call modHt.NewBound(Hid)
'            '设置流程按钮
'        If (frmWbNew.lblHtxz = "维保" And frmWbNew.txtHtze > 50000) Or Val(frmWbNew.txtHtze.Text) > 10000 Then
'            Call modHt.HtLcBut(63)
'        Else
'            Call modHt.HtLcBut(62)
'        End If
        frmWbNew.Visible = True
ElseIf NewF = 2 Then
        Call modNewHT.NewMQing
        Call modNewHT.NewLocked
        Call modNewHT.NewMBound(Hid)
        FMXC.lblMQM(0).Visible = True
        FMXC.lblMTm(0).Visible = True
        FMXC.cmdMQm(0).Visible = True
        
ElseIf NewF = 3 Or NewF = 5 Or NewF = 7 Then
        Call modNewHT.NewMQing
        Call modNewHT.NewLocked
        Call modNewHT.NewB(Hid)
        FMXC.lblMQM(0).Visible = True
        FMXC.lblMTm(0).Visible = True
        FMXC.cmdMQm(0).Visible = True
ElseIf NewF = 6 Or NewF = 8 Then
    Call FmxcNew.Bound(Hid)
    FmxcNew.Show
    FmxcNew.ZOrder 0
End If
End Sub

Private Sub Form_Click()
frmMod.Visible = False
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth + 500
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0


End Sub

Public Sub dtgCFF()
dtgC.Clear
dtgC.Cols = 15
dtgC.Row = 0
dtgC.Col = 0: dtgC.Text = "合同编号": dtgC.CellFontBold = True
dtgC.Col = 1: dtgC.Text = "项目名称": dtgC.CellFontBold = True
dtgC.Col = 2: dtgC.Text = "日期": dtgC.CellFontBold = True
dtgC.Col = 3: dtgC.Text = "内容": dtgC.CellFontBold = True
dtgC.Col = 4: dtgC.Text = "数量": dtgC.CellFontBold = True
dtgC.Col = 5: dtgC.Text = "供应商": dtgC.CellFontBold = True
dtgC.Col = 6: dtgC.Text = "订单给供应商": dtgC.CellFontBold = True
dtgC.Col = 7: dtgC.Text = "盖章回传情况": dtgC.CellFontBold = True
dtgC.Col = 8: dtgC.Text = "到货情况": dtgC.CellFontBold = True
dtgC.Col = 9: dtgC.Text = "采购": dtgC.CellFontBold = True
dtgC.Col = 10: dtgC.Text = "备注": dtgC.CellFontBold = True
dtgC.Col = 11: dtgC.Text = "lid": dtgC.CellFontBold = True
dtgC.Col = 12: dtgC.Text = "wcf": dtgC.CellFontBold = True
dtgC.Col = 13: dtgC.Text = "bid": dtgC.CellFontBold = True
'

dtgN.Clear
dtgN.Cols = 15
dtgN.Row = 0
dtgN.Col = 0: dtgN.Text = "合同编号": dtgN.CellFontBold = True
dtgN.Col = 1: dtgN.Text = "项目名称": dtgN.CellFontBold = True
dtgN.Col = 2: dtgN.Text = "日期": dtgN.CellFontBold = True
dtgN.Col = 3: dtgN.Text = "内容": dtgN.CellFontBold = True
dtgN.Col = 4: dtgN.Text = "数量": dtgN.CellFontBold = True
dtgN.Col = 5: dtgN.Text = "供应商": dtgN.CellFontBold = True
dtgN.Col = 6: dtgN.Text = "订单给供应商": dtgN.CellFontBold = True
dtgN.Col = 7: dtgN.Text = "盖章回传情况": dtgN.CellFontBold = True
dtgN.Col = 8: dtgN.Text = "到货情况": dtgN.CellFontBold = True
dtgN.Col = 9: dtgN.Text = "采购": dtgN.CellFontBold = True
dtgN.Col = 10: dtgN.Text = "备注": dtgN.CellFontBold = True

dtgC.ColWidth(0) = 2400
dtgC.ColWidth(1) = 2955
dtgC.ColWidth(2) = 1500
dtgC.ColWidth(3) = 2520
dtgC.ColWidth(5) = 3060
dtgC.ColWidth(6) = 1305
dtgC.ColWidth(7) = 1305
dtgC.ColWidth(11) = 0
dtgC.ColWidth(12) = 0
dtgC.ColWidth(13) = 0
dtgC.ColWidth(14) = 0
End Sub

Private Sub Form_Resize()
   dtgC.Width = Me.Width - 200
End Sub

Private Sub lblHtbh_DblClick()
If txtHtbh.Text = "" Then
    If Fl = "合同评审单" Then
        Me.dtgN.Col = 0
        txtHtbh.Text = Me.dtgN.Text
        Me.dtgN.Col = 13
        txtHtbh.ToolTipText = Val(dtgN.Text)
        
        txtHtbh.Text = txtHtbh.Text & "(" & txtHtbh.ToolTipText & ")"
    
    Else
    
    End If
Else
    txtHtbh.Text = ""
End If
End Sub

Private Sub timQuit_Timer()
Dim oo As Integer
Dim tt As String
Dim ii As Integer
Dim Ra
Dim La As Long
Dim RC
Dim Lc As Integer
On Error Resume Next
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0

If timZm = 1 Or timZm = 2 Then '保存,整单完成
    cmdSave.Enabled = True

    frmMod.Visible = False
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open ETT, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    On Error Resume Next
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2) + 1
    Call Me.Bound(Ra, La)
    Me.ZOrder 0

End If
timQuit.Enabled = False
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


    ElseIf timZm = 5 Then

    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    timWait.Enabled = False
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdSave.Enabled = False
    End If
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdSave.Enabled = False
    End If
    Exit Sub

End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub


