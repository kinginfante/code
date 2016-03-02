VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Ren 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "人员选择"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   4515
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   945
      Left            =   3240
      TabIndex        =   12
      Top             =   4050
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1667
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame frmBj 
      Caption         =   "上级部门"
      Height          =   525
      Left            =   30
      TabIndex        =   9
      Top             =   5040
      Width           =   2595
      Begin VB.CommandButton cmdBack 
         Caption         =   "返回"
         Height          =   285
         Left            =   1890
         TabIndex        =   11
         Top             =   180
         Width           =   675
      End
      Begin VB.Label lblBJ 
         Caption         =   "Label2"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   870
         TabIndex        =   10
         Top             =   210
         Width           =   975
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgRen 
      Height          =   5505
      Left            =   30
      TabIndex        =   8
      Top             =   30
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   9710
      _Version        =   393216
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      SelectionMode   =   1
      BorderStyle     =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame frmQy 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1185
      Left            =   3090
      TabIndex        =   5
      Top             =   420
      Visible         =   0   'False
      Width           =   1365
      Begin VB.ComboBox comQy 
         Height          =   300
         ItemData        =   "Ren.frx":0000
         Left            =   30
         List            =   "Ren.frx":0002
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "区域选择:"
         Height          =   285
         Left            =   60
         TabIndex        =   7
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdXz 
      Caption         =   "选  择"
      Height          =   405
      Left            =   3180
      TabIndex        =   1
      Top             =   5100
      Width           =   1065
   End
   Begin MSComctlLib.TreeView txt1 
      Height          =   5475
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   9657
      _Version        =   393217
      Style           =   7
      HotTracking     =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblBm 
      Caption         =   "Label1"
      Height          =   405
      Left            =   3090
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblUid 
      Caption         =   "lblUid"
      Height          =   375
      Left            =   3090
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblRen 
      Caption         =   "lblRen"
      Height          =   345
      Left            =   3120
      TabIndex        =   2
      Top             =   1830
      Visible         =   0   'False
      Width           =   1155
   End
End
Attribute VB_Name = "Ren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OpenForm As String
Public XForm As Form
Dim OBm As String

Private Sub cmdBack_Click()
Dim tt As String
Dim GGlId As Integer
On Error Resume Next
Dim Bm As String
Dim Bmid As Integer
Dim ii As Integer: Dim oo As Integer
Dim Ra: Dim ua: Dim Rb: Dim ub: Dim RC
Dim ji As Integer
Bm = lblBj.Caption
GGlId = Val(frmBJ.ToolTipText)

lblBj.Tag = lblBj.Tag - 1
If lblBj.Tag = 0 Then
    lblBj.Tag = 1
    Exit Sub
End If
    
lblBj.Caption = ""
lblBj.ToolTipText = ""
frmBJ.ToolTipText = ""
If lblBj.Tag = 1 Then
    tt = "select bm,bmid,ji from bm where ji=" & lblBj.Tag & " and zzf=1;"
    If frmKhbrG.Visible = True Then
        tt = "select bm,bmid,ji from bm where ji=" & lblBj.Tag & " and left(bm,2)='业务' and zzf=1;"
    End If
Else
    tt = "declare @gglid int,@bm nvarchar(20);" & _
        "select bm,bmid,ji,gglid from bm where  gglid=" & GGlId & " and zzf=1;" & _
        "select @gglid=gglid from bm where gglid=" & GGlId & ";" & _
        "select @bm=bm from bm where bmid=@gglid;" & _
        "SELECT dbo.worker.UserName, dbo.worker.UserId,'" & Bm & "' FROM dbo.worker INNER JOIN dbo.RLA ON dbo.worker.UserId = dbo.RLA.Auid WHERE" & _
        " ((dbo.RLA.bm1 = @bm) OR (dbo.worker.BM = @bm) or (dbo.rla.bm2=@bm) or (dbo.rla.bm3=@bm)) and worker.zzf=1 order by worker.bmjl desc;" & _
        "select bm,gglid from bm where bmid=@gglid"
        
End If
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    Set mod1.HTP = mod1.HTP.NextRecordset
    Rb = mod1.HTP.GetRows
    Set mod1.HTP = mod1.HTP.NextRecordset
    RC = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    ua = UBound(Ra, 2)
    ub = UBound(Rb, 2)
    dtgRen.Clear
    If ua >= 0 Then
        For oo = 0 To ua
            dtgRen.Row = oo
            For ii = 1 To 4
                dtgRen.Col = ii
                dtgRen.Text = Ra(ii - 1, oo)
                dtgRen.CellForeColor = &H8000000D
            Next
        Next
    End If
    If ub >= 0 Then
        For oo = ua + 1 To ub + ua + 1
            dtgRen.Row = oo
            For ii = 1 To 4
                dtgRen.Col = ii
                dtgRen.Text = Rb(ii - 1, oo - ua - 1)
            Next
        Next
    End If
    Call NB
lblBj.Caption = RC(0, 0)
frmBJ.ToolTipText = RC(1, 0)
If lblBj.Tag = 1 Then
    frmBJ.Visible = False
End If

End Sub

Private Sub cmdXZ_Click()
On Error Resume Next
Dim tt As String
Dim ii As Integer: Dim oo As Integer
Dim Ra: Dim ua
Dim Rb
Dim adoZ As Object
If lblRen.Caption <> "" And lblUid.Caption <> "" Or lblBM.Caption <> "" Then
    Select Case OpenForm
    Case "frmPeiView"
        tt = "select name,cq,cq1,zt,zfy,userZw,uid from peiView3 where bm='" & lblBM.Caption & "' and nd=" & Val(frmPeiView.lblYear.Caption) & _
            "select username from worker where zzf=1 and bm='" & lblBM.Caption & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        Set mod1.HTP = mod1.HTP.NextRecordset
        Rb = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        Call frmPeiView.BRBMBound(Ra, Rb)
        frmPeiView.Enabled = True
        frmPeiView.ZOrder 0
        frmPeiView.Frame1.Visible = True
    Case "Dialog"
        Dialog.lblZZ.Caption = lblRen.Caption
        Dialog.lblZZ.ToolTipText = lblUid.Caption
        Dialog.opt1.Value = True
        Call mod1.refEnvent(1)
        Dialog.cmdBJ.Enabled = True

    Case "frmGzbN"
        frmGzbN.lblRen.Caption = lblRen.Caption
        frmGzbN.lblRen.ToolTipText = lblUid.Caption
        Call frmGzbN.WeekDate(mod1.DQda, lblUid.Caption)
    Case "frmBxBrow"
        tt = "Select convert (nvarchar(25),frq,1) as 日期范围 ,convert (nvarchar(25),lrq,1) as 日期范围 ,HG as 金额,bxid as 报销单编号,convert (nvarchar(25),qrq,1) as 签收日期 from FyDBrow " & _
       "where Uid='" & lblUid.Caption & "' and not (hg is null)"
        frmBxBrow.AdoBxBro.Close
        frmBxBrow.AdoBxBro.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'        tt = "Select convert (nvarchar(25),frq,1) as 日期范围 ,convert (nvarchar(25),lrq,1) as 日期范围 ,HG as 金额,bxid as 报销单编号,convert (nvarchar(25),qrq,1) as 签收日期 from FyDBrow " & _
'       "where Uid='" & mod1.DHid & "' and ywy='" & mod1.DName & "' and not (hg is null)"
'        frmBxBrow.AdoBxBro.Close
'        frmBxBrow.AdoBxBro.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmBxBrow.mga.DataSource = frmBxBrow.AdoBxBro
        'PK = "<起 始 期  |<截 至 期  |>  金 额 |^ 报 销 单 编 号|> 签收日期 "
        'PK = "^  日期范围|^  日期范围|>  金 额 |^ 报 销 单 编 号|> 签收日期 "
         frmBxBrow.mga.ColWidth(0) = 500
        'frmBxBrow.mga.FormatString = PK
        frmBxBrow.mga.MergeRow(0) = True
        frmBxBrow.mga.MergeCells = flexMergeRestrictAll
        frmBxBrow.lblFw.Caption = lblRen.Caption
        frmBxBrow.lblFw.ToolTipText = lblUid.Caption
    Case "frmZu"
        Call frmOL.Tbound(lblUid.Caption)
        MDI.timFl.Enabled = False
        frmOL.Show
        frmOL.Left = frmZu.Left
        frmOL.Top = 0
        frmOL.ZOrder 0
        frmOL.Caption = "您正在和" & lblRen.Caption & "交谈"
        Set mod1.HTP = CreateObject("adodb.recordset")
        tt = "select imgid from worker where userid='" & lblUid.Caption & "'"
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        frmOL.img1.Picture = frmZu.ImageList2.ListImages(Ra(0, 0)).Picture
        frmOL.lbl1.Caption = lblRen.Caption
        frmOL.lbl1.ToolTipText = lblUid.Caption
        'frmOL.img2.Picture = frmZu.ImageList2.ListImages(frmZu.tb1.Buttons(frmZu.meIndex).Image).Picture
            frmOL.img2.Picture = frmZu.NR(frmZu.meIndex).PictureNormal
        frmOL.lbl2.Caption = mod1.DName
        frmOL.txt1.SelStart = Len(frmOL.txt1.Text)
        frmOL.txt1.SelLength = 0
        frmOL.txt2.Text = ""
        frmOL.txt2.SetFocus
'''''    Case "bView"
'''''        If lblBm.Caption <> "" Then
'''''            bView.lblFw.Caption = lblBm.Caption
'''''            bView.lblFw.ToolTipText = ""
'''''
'''''        Else
'''''
'''''            bView.lblFw.Caption = lblRen.Caption
'''''            bView.lblFw.ToolTipText = lblUid.Caption
'''''
'''''        End If
'''''        fyBB.ZOrder 0
'''''    Case "b1"
'''''        Call b1.KPIQing
'''''        Call b1.KPIBound(lblRen.Caption, lblUid.Caption, b1.txtM.Value)
    Case "frmRen"
        frmRen.lblGGL.Caption = lblRen.Caption
        frmRen.lblGGL.ToolTipText = lblUid.Caption
    Case "FmxcFK"
        If FmxcFK.xZ = 2 Then
            FmxcFK.txtRen2.Text = lblRen.Caption
            FmxcFK.txtRen2.ToolTipText = lblUid.Caption
        Else
            FmxcFK.txtRen3.Text = lblRen.Caption
            FmxcFK.txtRen3.ToolTipText = lblUid.Caption
        End If
    Case "frmFYBX"
        If frmFYBX.dtgNx.Visible = False Then

    '        frmFYBX.adoF2.Recordset.MoveFirst
    '        Do While Not frmFYBX.adoF2.Recordset.EOF
            tt = "RenOpenA('" & lblRen.Caption & "','" & lblUid.Caption & "')"
            Set mod1.HTP = CreateObject("adodb.recordset")
            mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
                frmFYBX.adoF2.Recordset.Update "ywy", lblRen.Caption
                frmFYBX.adoF2.Recordset.Update "YwYUid", lblUid.Caption
                frmFYBX.adoF2.Recordset.Update "qy", mod1.HTP.Fields("qy").Value
                frmFYBX.adoF2.Recordset.Update "bm", mod1.HTP.Fields("bm").Value
                frmFYBX.adoF2.Recordset.Update "dep", mod1.HTP.Fields("bmid").Value
                
        Else
            frmFYBX.lblGui.Caption = lblRen.Caption
            frmFYBX.lblGuid.Caption = lblUid.Caption
            frmFYBX.lblBM.Caption = lblBM.Caption
        End If

'            frmFYBX.adoF2.Recordset.MoveNext
'        Loop
'        If frmFYBX.lblGui.Caption <> "" Then
'            frmFYBX.cmdGui.Visible = False
'        End If
    Case "frmGzBG"
       
        tt = "Select atime,khqc,xmfy,NewF,gid from xmgz where ywy='" & lblRen.Caption & "' and aTime>='" & modXmGz.FR & _
        "' and aTime <='" & modXmGz.LR & "' and lb=1 order by aTime"
        frmGzBG.adoXm.Close
        frmGzBG.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        
        Set frmGzBG.dtgXmgz.DataSource = frmGzBG.adoXm
        
        tt = "Select atime,khqc,newF,gid from xmgz where ywy='" & lblRen.Caption & "' and aTime>='" & modXmGz.FR & _
        "' and aTime <='" & modXmGz.LR & "' and lb=0 order by aTime"
        frmGzBG.adoJi.Close
        frmGzBG.adoJi.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmGzBG.dtgJi.DataSource = frmGzBG.adoJi
    
'         tt = "select xmmc as 项目名称,khjb as 项目平台,xid as 编号,xmfy as 费用,xid from xmzl where ywy='" & lblRen.Caption & "' order by 项目名称"
'        Set frmGzBG.AdoKh = CreateObject("adodb.recordset")
'        frmGzBG.AdoKh.Close
'        frmGzBG.AdoKh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'        Set frmGzBG.dtgKH.DataSource = frmGzBG.AdoKh
        frmGzBG.lblYwy.Caption = lblRen.Caption
        frmGzBG.lblFw.Caption = lblRen.Caption
    Case "frmKhbrG"
        If frmKhbrG.XuanRen = 1 Then

                If lblRen.Caption = "" Then
                    tt = "Select * from XmView where 部门='" & Trim(lblBM.Caption) & "' order by 业务员"
                    frmKhbrG.lblFw.Caption = Trim(lblBM.Caption)
                Else
                    tt = "Select * from XmView where 业务员='" & Trim(lblRen.Caption) & "' and uid='" & Trim(lblUid.Caption) & "'"
                    frmKhbrG.lblFw.Caption = Trim(lblRen.Caption)
                    
                End If
            If mod1.Qy = "北京" Then
                    tt = "Select * from XmView where 业务员='" & Trim(lblRen.Caption) & "'"
                    frmKhbrG.lblFw.Caption = Trim(lblRen.Caption)
            End If
            Set frmKhbrG.adoKhBr = CreateObject("adodb.recordset")
            frmKhbrG.adoKhBr.Close
            frmKhbrG.adoKhBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
            Set frmKhbrG.dtgKh.DataSource = Nothing
            Set frmKhbrG.dtgKh.DataSource = frmKhbrG.adoKhBr
            If frmKhbrG.adoKhBr.RecordCount > 0 Then
                frmKhbrG.dtgKh.FixedRows = 0
                frmKhbrG.dtgKh.MergeCol(4) = True
                frmKhbrG.dtgKh.MergeCol(12) = True
                frmKhbrG.dtgKh.MergeCol(14) = True
                frmKhbrG.dtgKh.MergeCells = 3
                frmKhbrG.dtgKh.FixedRows = 1
            End If
        ElseIf frmKhbrG.XuanRen = 2 Then
            If lblRen.Caption = "" Then
                Exit Sub
            End If
            frmKhbrG.lblYwy.Caption = Trim(lblRen.Caption)
            frmKhbrG.lblYwy.ToolTipText = Trim(lblUid.Caption)
        End If
    Case "htBrowG"
        If lblBM.Caption <> "" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF  from htView1 where 部门='" & _
                Trim(lblBM.Caption) & "' and 合同编号<>'HMNEW' order by htrq desc"
            htBrowG.lblFw.Caption = Trim(lblBM.Caption)
        Else
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF  from htView1 where 项目归属人='" & _
                Trim(lblRen.Caption) & "' and xuid='" & Trim(lblUid.Caption) & "' and 合同编号<>'HMNEW'  order by htrq desc"
            htBrowG.lblFw.Caption = Trim(lblRen.Caption)
        End If
        If mod1.ZT = "HBData" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF  from htView1 where 项目归属人='" & _
                Trim(lblBM.Caption) & "' and 合同编号<>'HMNEW'  order by htrq desc"
            htBrowG.lblFw.Caption = Trim(lblBM.Caption)
        End If
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        ua = UBound(Ra, 2)
        htBrowG.dtgBr.Visible = False
        htBrowG.dtgBr.Clear
        htBrowG.dtgBr.Row = 0: htBrowG.dtgBr.Col = 1: htBrowG.dtgBr.Text = "项目归属人"
        htBrowG.dtgBr.Col = 2: htBrowG.dtgBr.Text = "项目名称": htBrowG.dtgBr.Col = 3: htBrowG.dtgBr.Text = "合同日期": htBrowG.dtgBr.Col = 4: htBrowG.dtgBr.Text = "合同性质"
        htBrowG.dtgBr.Col = 5: htBrowG.dtgBr.Text = "合同金额": htBrowG.dtgBr.Col = 6: htBrowG.dtgBr.Text = "合同编号": htBrowG.dtgBr.Col = 7: htBrowG.dtgBr.Text = "状态"
        For oo = 1 To ua + 1
            htBrowG.dtgBr.Row = oo
            For ii = 1 To 11
                htBrowG.dtgBr.Col = ii
                htBrowG.dtgBr.Text = Trim(Ra(ii - 1, oo - 1))
            Next
        Next
        htBrowG.dtgBr.Visible = True
    Case "frmBxV"
        If comQy.Text = "全公司" Then
            If lblBM.Caption <> "" Then
                tt = "FydVGBm1('" & lblBM.Caption & "','" & frmBxV.mtA.Value & "')"
                frmBxV.lblFw.Caption = lblBM.Caption
                frmBxV.lblFw.ToolTipText = ""
            Else
                tt = " FydVGywy('" & lblRen.Caption & "','" & frmBxV.mtA.Value & "')"
                frmBxV.lblFw.Caption = lblRen.Caption
                frmBxV.lblFw.ToolTipText = lblUid.Caption
            End If
        Else
            If lblBM.Caption <> "" Then
                tt = "FydVGBm2('" & lblBM.Caption & "','" & comQy.Text & "','" & frmBxV.mtA.Value & "')"
                frmBxV.lblFw.Caption = lblBM.Caption
                frmBxV.lblFw.ToolTipText = ""
            Else
                tt = " FydVGywy('" & lblRen.Caption & "','" & frmBxV.mtA.Value & "')"
                frmBxV.lblFw.Caption = lblRen.Caption
                frmBxV.lblFw.ToolTipText = lblUid.Caption
            End If
        End If
        frmBxV.adoBxV.Close
        frmBxV.adoBxV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        Set frmBxV.dtgBx.DataSource = frmBxV.adoBxV
        If frmBxV.adoBxV.RecordCount > 0 Then
            frmBxV.dtgBx.FixedRows = 0
            frmBxV.dtgBx.MergeCol(1) = True
            frmBxV.dtgBx.MergeCol(2) = True
            frmBxV.dtgBx.MergeCol(3) = True
            frmBxV.dtgBx.MergeCol(4) = True
            frmBxV.dtgBx.MergeCol(5) = True
            frmBxV.dtgBx.MergeCol(7) = True
            frmBxV.dtgBx.MergeCells = 3
            frmBxV.dtgBx.FixedRows = 1
        End If


    Case "fyBB"
            If lblBM.Caption <> "" Then
                fyBB.lblFw.Caption = lblBM.Caption
                fyBB.lblFw.ToolTipText = ""
                fyBB.chkLb(0).Enabled = False
            Else

                fyBB.lblFw.Caption = lblRen.Caption
                fyBB.lblFw.ToolTipText = lblUid.Caption
                fyBB.chkLb(0).Enabled = True
            End If
            fyBB.ZOrder 0
    Case "frmBB"
            If lblBM.Caption <> "" Then
                frmBB.lblFw.Caption = lblBM.Caption
                frmBB.lblFw.ToolTipText = ""

            Else
                
                frmBB.lblFw.Caption = lblRen.Caption
                frmBB.lblFw.ToolTipText = lblUid.Caption

            End If
            fyBB.ZOrder 0
    Case "HLB"
        HLB.txtH(3).Text = lblRen.Caption
        HLB.txtH(3).ToolTipText = lblUid.Caption
    Case "Dialog"
        Set adoZ = CreateObject("adodb.recordset")
        tt = "update newfuwu set ywy='" & lblRen.Caption & "',uid='" & lblUid.Caption & "' where fwid=" & Dialog.Fwid
        adoZ.Close
        adoZ.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        adoZ.Close
        Set adoZ = Nothing
        Dialog.Fwid = 0
        Call mod1.refEnvent(1)
    Case "frmGGL"
        Dim TG As Boolean '是否添加,重复姓名将不添加进列表
        On Error Resume Next
        If lblUid.Caption <> "" Then
            'If mod1.comJZ = False Then Exit Sub
            frmGGL.adoRen.Recordset.MoveFirst
            TG = True
            Do While Not frmGGL.adoRen.Recordset.EOF
                If frmGGL.adoRen.Recordset.Fields("username").Value = lblRen.Caption Then
                    TG = False
                    Exit Do
                End If
                frmGGL.adoRen.Recordset.MoveNext
            Loop
            If TG = True Then
                frmGGL.adoRen.Recordset.AddNew "username", lblRen.Caption
            End If

        Else
                tt = "SELECT dbo.worker.UserName, dbo.worker.UserId, dbo.RLA.bm1" & _
                        " FROM         dbo.BM AS BM_2 RIGHT OUTER JOIN" & _
                      " dbo.RLA INNER JOIN dbo.worker ON dbo.RLA.Auid = dbo.worker.UserId INNER JOIN dbo.BM INNER JOIN" & _
                     " dbo.BM AS BM_1 ON dbo.BM.gglid = BM_1.BMID ON dbo.worker.BM = dbo.BM.BM ON BM_2.BMID = BM_1.gglid" & _
                " WHERE     (dbo.worker.BM = '" & lblBM.Caption & "' OR dbo.RLA.bm1 = '" & lblBM.Caption & "' OR dbo.RLA.bm2 = '" & lblBM.Caption & "' OR dbo.RLA.bm3 = '" & lblBM.Caption & "' OR" & _
                  " BM_1.BM = '" & lblBM.Caption & "' OR BM_2.BM = '" & lblBM.Caption & "') and (dbo.worker.zzF = 1)" & _
               " ORDER BY dbo.worker.qy, dbo.worker.BM"
            'tt = "select username from renyuan where bm='" & lblBm.Caption & "'"
'''''            If lblBm.Caption = "行政人事" Then
'''''                tt = "select username from renyuan where bm='" & lblBm.Caption & "' or bm='商务部'"
'''''            End If
            Set mod1.HTP = CreateObject("adodb.recordset")
            mod1.HTP.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
            mod1.HTP.MoveFirst
            Do While Not mod1.HTP.EOF

                If frmGGL.adoRen.Recordset.RecordCount > 0 Then
                    frmGGL.adoRen.Recordset.MoveFirst
                    TG = True
                    Do While Not frmGGL.adoRen.Recordset.EOF
                        If frmGGL.adoRen.Recordset.Fields("username").Value = mod1.HTP.Fields("username").Value Then
                            TG = False
                            Exit Do
                        End If
                        If TG = True Then
                            frmGGL.adoRen.Recordset.AddNew "username", mod1.HTP.Fields("username").Value
                            frmGGL.adoRen.Recordset.Update "userid", mod1.HTP.Fields("userid").Value
                        End If
                        frmGGL.adoRen.Recordset.MoveNext
                    Loop
                Else
                    frmGGL.adoRen.Recordset.AddNew "username", mod1.HTP.Fields("username").Value
                          frmGGL.adoRen.Recordset.Update "userid", mod1.HTP.Fields("userid").Value
                End If

                mod1.HTP.MoveNext
            Loop
        End If
        Set frmGGL.dtgRen.DataSource = frmGGL.adoRen
    End Select
    Select Case Trim(Ren.OpenForm)
'''''    Case "bView"
'''''        bView.Enabled = True
'''''        bView.ZOrder 0
'''''    Case "b1"
'''''        b1.Enabled = True
'''''        b1.ZOrder 0
    Case "frmRen"
        frmRen.Enabled = True
        frmRen.ZOrder 0
    Case "frmFYBX"
        frmFYBX.Enabled = True
        frmFYBX.ZOrder 0
    Case "frmKhbrG"
        frmKhbrG.Enabled = True
        frmKhbrG.ZOrder 0
    Case "htBrowG"
        htBrowG.Enabled = True
        htBrowG.ZOrder 0
    Case "frmGzBG"
        frmGzBG.Enabled = True
        frmGzBG.ZOrder 0
    Case "frmBxV"
        frmBxV.Enabled = True
        frmBxV.ZOrder 0


    Case "fyBB"
        fyBB.Enabled = True
        fyBB.WindowState = 2
        fyBB.ZOrder 0
    Case "frmBB"
        frmBB.Enabled = True
        frmBB.WindowState = 2
        frmBB.ZOrder 0
    Case "HLB"
        HLB.Enabled = True
        HLB.ZOrder 0
    Case "Dialog"
        Dialog.Enabled = True
        Dialog.ZOrder 0
    Case "frmGGL"
        frmGGL.Enabled = True
        frmGGL.ZOrder 0
    Case "frmZu"
        frmZu.Enabled = True
        frmOL.Show
        frmOL.ZOrder 0
    Case "frmBxBrow"
        frmZu.Enabled = True
        frmBxBrow.ZOrder 0
        frmBxBrow.Enabled = True
    Case "frmGzbN"
        frmGzbN.Enabled = True
        frmGzbN.ZOrder 0
    Case "Dialog"
        Dialog.Enabled = True
        Dialog.ZOrder 0
    End Select
    
    Ren.Visible = False
    frmQy.Visible = False
End If
End Sub

Private Sub comQy_Click()
Static oo As Integer
Dim tt As String
Dim Tbm As String
On Error Resume Next
oo = oo + 2
''    Unload txt1(oo)
''    Load txt1(oo)
''    txt1(oo).Visible = True
'txt1(oo).se
'txt1(oo).Nodes.Remove txt1(oo).SelectedItem.Index
'For oo = -200 To 300
'    txt1(oo).Nodes.Remove oo
'Next
Load txt1(oo)
txt1(oo).Visible = True
txt1(oo - 2).Visible = False
txt1(1).Visible = False
            If comQy.Text = "全公司" Then
                tt = "renOpen"
                Ren.frmQy.Visible = True
            Else
                    tt = "renOpenQy('" & comQy.Text & "')"
            End If

    mod1.adoRen.Close
    mod1.adoRen.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    mod1.adoRen.MoveFirst

    Ren.txt1(oo).Nodes.Add , , mod1.adoRen.Fields("bm").Value, mod1.adoRen.Fields("bm").Value
    Ren.txt1(oo).Nodes.Add mod1.adoRen.Fields("bm").Value, tvwChild, mod1.adoRen.Fields("userid").Value, mod1.adoRen.Fields("username").Value
    Tbm = mod1.adoRen.Fields("bm").Value
    'Tuid = mod1.adoRen.Fields("userid").Value
    mod1.adoRen.MoveNext
    Do While Not mod1.adoRen.EOF
        If mod1.adoRen.Fields("bm").Value = Tbm Then
            Ren.txt1(oo).Nodes.Add Tbm, tvwChild, adoRen.Fields("userid").Value, adoRen.Fields("username").Value
        Else
            Ren.txt1(oo).Nodes.Add Tbm, tvwNext, adoRen.Fields("bm").Value, mod1.adoRen.Fields("bm").Value
            Tbm = mod1.adoRen.Fields("bm").Value
            Ren.txt1(oo).Nodes.Add Tbm, tvwChild, adoRen.Fields("userid").Value, adoRen.Fields("username").Value
        End If
        mod1.adoRen.MoveNext
    Loop
    Ren.lblRen.Caption = ""
    Ren.lblUid.Caption = ""
'If Ren.Visible = False Then Exit Sub
'Call mod1.RenShow(comQy.Text)
End Sub


Private Sub dtgRen_Click()
Dim Ren As String: Dim Uid As String: Dim Bm As String
Dim ji As Integer
dtgN.Row = dtgRen.Row
dtgN.Col = 3
ji = Val(dtgN.Text)
Bm = dtgN.Text
dtgN.Col = 1
Ren = dtgN.Text
dtgN.Col = 2
Uid = dtgN.Text
If ji = 0 Then
    lblRen.Caption = Ren
    lblUid.Caption = Uid
    lblBM.Caption = Bm
Else
    lblRen.Caption = ""
    lblUid.Caption = ""
    lblBM.Caption = Ren
    'lblBm.Caption = ""
End If
dtgN.Col = 1
End Sub

Private Sub dtgRen_DblClick()
Dim tt As String
On Error Resume Next
Dim GGlId As Single
Dim Bm As String
Dim Bmid As Single
Dim ii As Integer: Dim oo As Integer
Dim Ra: Dim ua: Dim Rb: Dim ub


Dim ji As Integer
dtgRen.Col = 3
ji = Val(dtgRen.Text)
If ji = 0 Then Exit Sub
dtgRen.Col = 2
Bmid = Val(dtgRen.Text)
dtgRen.Col = 4
GGlId = Val(dtgRen.Text)
If ji > 1 Then
    frmBJ.ToolTipText = GGlId
Else
    frmBJ.ToolTipText = Bmid
End If
dtgRen.Col = 1
Bm = dtgRen.Text

tt = "select bm,bmid,ji,gglid from bm where gglid=" & Bmid & " and zzf=1;" & _
    "SELECT dbo.worker.UserName, dbo.worker.UserId,'" & Bm & "' FROM dbo.worker left outer JOIN dbo.RLA ON dbo.worker.UserId = dbo.RLA.Auid WHERE" & _
    " ((dbo.RLA.bm1 = '" & Bm & "') OR (dbo.worker.BM = '" & Bm & "') or (dbo.rla.bm2='" & Bm & "') or (dbo.rla.bm3='" & Bm & "')) and worker.zzf=1 order by worker.bmjl desc,worker.qy,worker.userid"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    Set mod1.HTP = mod1.HTP.NextRecordset
    Rb = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    ua = UBound(Ra, 2)
    ub = UBound(Rb, 2)
    dtgRen.Clear
    If ua >= 0 Then
        For oo = 0 To ua
            dtgRen.Row = oo
            For ii = 1 To 4
                dtgRen.Col = ii
                dtgRen.Text = Ra(ii - 1, oo)
                dtgRen.CellForeColor = &H8000000D
            Next
        Next
    End If
    If ub >= 0 Then
        For oo = 0 To ub
            dtgRen.Row = oo + ua + 1
            For ii = 1 To 4
                dtgRen.Col = ii
                dtgRen.Text = Rb(ii - 1, oo)
            Next
        Next
    End If
    Call NB
    frmBJ.Visible = True

    lblBj.Caption = Bm
    lblBj.ToolTipText = Bmid
    lblBj.Tag = ji + 1
    OBm = lblBj.Caption
End Sub


Private Sub Form_Load()
Dim oo As Integer
Dim tt As String
On Error Resume Next
Ren.Width = 4605
Ren.Height = 6090
Load txt1(1)
txt1(1).Visible = True

'Call mod1.ResizeForm(Me) '确保窗体改变时控件随之改变
'设置区域下拉框
tt = "Select qy from YzQy"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
For oo = comQy.ListCount - 1 To 0 Step -1
    comQy.RemoveItem oo
Next
comQy.AddItem "全公司"
For oo = 1 To mod1.HTP.RecordCount
    comQy.AddItem mod1.HTP.Fields("Qy"), oo
    mod1.HTP.MoveNext
Next
    
    comQy.Text = "全公司"
    dtgRen.Left = 0
    dtgRen.Top = 0
dtgRen.ColWidth(1) = 2000
dtgRen.Row = 0
dtgRen.ColWidth(2) = 0
dtgRen.ColWidth(0) = 0
dtgRen.ColWidth(3) = 0
End Sub



Private Sub Form_Unload(Cancel As Integer)
'If MDI.Cq = False Then
Select Case Trim(Ren.OpenForm)
Case "frmPeiView"
    frmPeiView.Enabled = True
    frmPeiView.ZOrder 0
Case "frmZu"
    frmZu.Enabled = True
    frmZu.ZOrder 0
Case "frmRen"
    frmRen.Enabled = True
    frmRen.ZOrder 0
Case "frmFYBX"
    frmFYBX.Enabled = True
    frmFYBX.ZOrder 0
Case "frmKhbrG"
    frmKhbrG.Enabled = True
    frmKhbrG.ZOrder 0
Case "htBrowG"
    htBrowG.Enabled = True
    htBrowG.ZOrder 0
Case "frmGzBG"
    frmGzBG.Enabled = True
    frmGzBG.ZOrder 0
Case "frmBxV"
    frmBxV.Enabled = True
    frmBxV.ZOrder 0


Case "fyBB"
    fyBB.Enabled = True
    fyBB.ZOrder 0
Case "frmBB"
    frmBB.Enabled = True
    frmBB.ZOrder 0
Case "frmGGL"
    frmGGL.Enabled = True
    frmGGL.ZOrder 0
Case "Dialog"
    Dialog.Enabled = True
    Dialog.ZOrder 0
Case "frmBxBrow"
        frmZu.Enabled = True
        frmBxBrow.ZOrder 0
        frmBxBrow.Enabled = True
Case "frmGzbN"
    frmGzbN.ZOrder
    frmGzbN.Enabled = True
End Select
frmQy.Visible = False
'End If
End Sub

Private Sub txt1_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
If Left(txt1(Index).SelectedItem.Key, 2) = "HM" Then
    lblRen.Caption = txt1(Index).SelectedItem.Text
    lblUid.Caption = txt1(Index).SelectedItem.Key
    lblBM.Caption = ""
Else
    lblRen.Caption = ""
    lblUid.Caption = ""
    lblBM.Caption = txt1(Index).SelectedItem.Text
End If
End Sub

Public Sub NB()
Dim ii As Integer: Dim oo As Integer
dtgN.Clear
dtgN.Cols = dtgRen.Cols
dtgN.Rows = dtgRen.Rows
For oo = 0 To dtgRen.Rows - 1
    dtgN.Row = oo: dtgRen.Row = oo
    For ii = 0 To 4
        dtgN.Col = ii: dtgRen.Col = ii
        dtgN.Text = dtgRen.Text
    Next
Next
End Sub
