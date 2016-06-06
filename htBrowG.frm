VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form htBrowG 
   Caption         =   "查询框"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.Frame frmXDC 
      Caption         =   "询价单查询"
      Height          =   1695
      Left            =   12810
      TabIndex        =   16
      Top             =   1380
      Width           =   2385
      Begin NiceFormControl.NiceButton cmdXJView 
         Height          =   345
         Left            =   330
         TabIndex        =   17
         Top             =   1200
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   609
         BTYPE           =   3
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "htBrowG.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Style           =   15
         Caption         =   "询价单查询"
      End
      Begin MSComCtl2.DTPicker dd1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dddddd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   390
         TabIndex        =   18
         Top             =   540
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         CalendarBackColor=   12648384
         CalendarTitleBackColor=   16448
         Format          =   138936321
         CurrentDate     =   38100
      End
   End
   Begin VB.CommandButton cmdCG 
      Caption         =   "采购报表"
      Height          =   345
      Left            =   12930
      TabIndex        =   15
      Top             =   7980
      Width           =   2295
   End
   Begin VB.CommandButton cmdX 
      Caption         =   "销售报表"
      Height          =   375
      Left            =   12900
      TabIndex        =   14
      Top             =   7410
      Width           =   2385
   End
   Begin NiceFormControl.NiceButton cmdZF 
      Height          =   435
      Left            =   12870
      TabIndex        =   13
      Top             =   6840
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   767
      BTYPE           =   1
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   8454143
      BCOLO           =   8454143
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421631
      MPTR            =   1
      MICON           =   "htBrowG.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Style           =   21
      Caption         =   "作废合同"
   End
   Begin NiceFormControl.NiceButton cmdWZX 
      Height          =   465
      Left            =   12870
      TabIndex        =   12
      Top             =   6300
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   820
      BTYPE           =   1
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "htBrowG.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Style           =   0
      Caption         =   "未执行合同"
   End
   Begin VB.CommandButton cmdRef 
      Caption         =   "查    询"
      Height          =   375
      Left            =   12840
      TabIndex        =   11
      Top             =   5190
      Width           =   2385
   End
   Begin VB.CommandButton cmdVall 
      Caption         =   "显示全部"
      Height          =   375
      Left            =   12870
      TabIndex        =   7
      Top             =   5790
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.TextBox txtZ 
      Height          =   315
      Left            =   13890
      TabIndex        =   6
      Top             =   4680
      Width           =   1305
   End
   Begin VB.ComboBox comLx 
      Height          =   300
      ItemData        =   "htBrowG.frx":0054
      Left            =   13890
      List            =   "htBrowG.frx":0064
      TabIndex        =   5
      Text            =   "合同编号"
      Top             =   4140
      Width           =   1395
   End
   Begin VB.CommandButton cmdFw 
      Caption         =   "查询范围"
      Height          =   315
      Left            =   12810
      TabIndex        =   4
      Top             =   3510
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00800000&
      Caption         =   "详 情"
      Height          =   345
      Left            =   12900
      TabIndex        =   2
      Top             =   360
      Width           =   2265
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "退 出"
      Height          =   345
      Left            =   14010
      TabIndex        =   1
      Top             =   8670
      Width           =   1005
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8565
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   15108
      _Version        =   393216
      BackColor       =   -2147483634
      Rows            =   100
      Cols            =   15
      BackColorSel    =   -2147483641
      BackColorBkg    =   -2147483636
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   15
   End
   Begin VB.Label Label2 
      Caption         =   "值："
      Height          =   405
      Left            =   13410
      TabIndex        =   10
      Top             =   4710
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "查询条件："
      Height          =   225
      Left            =   12870
      TabIndex        =   9
      Top             =   4200
      Width           =   945
   End
   Begin VB.Label lblFw 
      Height          =   285
      Left            =   14010
      TabIndex        =   8
      Top             =   3540
      Width           =   1155
   End
   Begin VB.Label Label5 
      Caption         =   "  双击列表记录可打开"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12780
      TabIndex        =   3
      Top             =   420
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "htBrowG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim ZT As String
Public Sub dtgbrFF(Ra, ua)
On Error Resume Next
Dim oo As Integer
Dim ii As Integer: Dim LR As Double
    htBrowG.dtgBr.Visible = False
    
    If ua < 32 Then
        dtgBr.Rows = 32
    Else
        dtgBr.Rows = ua + 2
    End If
    htBrowG.dtgBr.Clear
    htBrowG.dtgBr.Row = 0:
    htBrowG.dtgBr.Col = 1: htBrowG.dtgBr.Text = "询价单编号"
    htBrowG.dtgBr.Col = 2: htBrowG.dtgBr.Text = "项目名称":
    htBrowG.dtgBr.Col = 3: htBrowG.dtgBr.Text = "操作日期":
    htBrowG.dtgBr.Col = 4: htBrowG.dtgBr.Text = "采购员"
    htBrowG.dtgBr.Col = 5: htBrowG.dtgBr.Text = "删除否":

    For ii = 1 To 5
        dtgBr.Col = ii
        dtgBr.CellFontBold = True
    Next
    For oo = 1 To ua
'        dtgBr.Col = 10
'        lr = Val(dtgBr.Text)
        dtgBr.Row = oo
        For ii = 1 To 5
            htBrowG.dtgBr.Col = ii
            htBrowG.dtgBr.Text = Trim(Ra(ii - 1, oo - 1))
            If Ra(5, oo - 1) > 0 Then
                dtgBr.CellForeColor = &H0&
            Else
                dtgBr.CellForeColor = &HFF&
            End If
'''''            If Ra(10, oo - 1) > 0 Then
'''''                dtgBr.CellForeColor = &H800000
'''''            End If
        Next
    Next
    dtgBr.Visible = True

End Sub
Private Sub CancelButton_Click()
htBrowG.Visible = False
frmZu.Enabled = True
End Sub



Private Sub cmdDel_Click()
Dim tt As String
Dim Hid As Long
On Error Resume Next
dtgBr.Col = 8
Hid = dtgBr.Text
tt = "update htping set delf=0 where hid=" & Hid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
mod1.HTP.Close
adoBr.Requery
Set dtgBr.DataSource = adoBr
End Sub

Private Sub cmdCg_Click()
FmxcNewCG.Show
Call FmxcNewCG.dtgCFF
End Sub

Private Sub cmdFw_Click()
Set Ren.XForm = New htBrowG
Call mod1.RenXz("htBrowG", Me, 0)
End Sub

Private Sub cmdRef_Click()
Dim tt As String
Dim oo As Integer: Dim ii As Integer
Dim Ra, ua
On Error Resume Next
If (mod1.DName = "张砚纯" Or mod1.DName = "" Or mod1.DName = "吴金荣" Or mod1.DName = "徐瑛") And frmZu.NC.Caption = "产品事务" Then
    If mod1.DName <> "张砚纯" And mod1.DName <> "徐瑛" Then
        Select Case comLx.Text
            Case "合同金额"
                tt = "select 项目名称,询价日期,人工费,最低销售价,类型,编号,业务员,bid,uid from xunjiaZC where 人工费=" & Val(txtZ.Text) & " or 最低销售价=" & Val(txtZ.Text) & " order by bid desc"
            Case "项目名称"
                tt = "select 项目名称,询价日期,人工费,最低销售价,类型,编号,业务员,bid,uid from xunjiaZC where 项目名称 like '%" & Trim(txtZ.Text) & "%' order by bid desc"
            Case "合同编号"
                MsgBox "没这个功能"
                Exit Sub
        End Select
    Else
        Select Case comLx.Text
            Case "合同金额"
                tt = "select 项目名称,询价日期,人工费,最低销售价,类型,编号,业务员,bid,uid from xunjiaZ where 人工费=" & Val(txtZ.Text) & " or 最低销售价=" & Val(txtZ.Text) & " order by bid desc"
            Case "项目名称"
                tt = "select 项目名称,询价日期,人工费,最低销售价,类型,编号,业务员,bid,uid from xunjiaZ where 项目名称 like '%" & Trim(txtZ.Text) & "%' order by bid desc"
            Case "合同编号"
                MsgBox "没这个功能"
                Exit Sub
        End Select

    End If
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2)
    htBrowG.dtgBr.Clear
    htBrowG.dtgBr.Row = 0: htBrowG.dtgBr.Col = 1: htBrowG.dtgBr.Text = "项目名称": htBrowG.dtgBr.Col = 2: htBrowG.dtgBr.Text = "询价日期"
    htBrowG.dtgBr.Col = 3: htBrowG.dtgBr.Text = "人工费": htBrowG.dtgBr.Col = 4: htBrowG.dtgBr.Text = "最低销售价": htBrowG.dtgBr.Col = 5: htBrowG.dtgBr.Text = "类型"
    htBrowG.dtgBr.Col = 6: htBrowG.dtgBr.Text = "编号": htBrowG.dtgBr.Col = 7: htBrowG.dtgBr.Text = "业务员": htBrowG.dtgBr.ColWidth(8) = 0: htBrowG.dtgBr.ColWidth(9) = 0
    htBrowG.dtgBr.ColWidth(1) = 2500: htBrowG.dtgBr.ColWidth(2) = 1000
    For oo = 1 To La + 1
        htBrowG.dtgBr.Row = oo
        For ii = 1 To 9
            htBrowG.dtgBr.Col = ii
            htBrowG.dtgBr.Text = Ra(ii - 1, oo - 1)
        Next
    Next
    Exit Sub
End If
Select Case comLx.Text

    Case "合同金额"

            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where ggl='" & mod1.DHid & "' and 合同金额=" & Val(txtZ) & " and 合同编号<>'HMNEW' order by 合同日期 desc"
            If mod1.DName = "初永友" Then
                tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 部门='" & mod1.Bm & "' and 合同金额=" & Val(txtZ) & " and 合同编号<>'HMNEW' order by 合同日期 desc"
            ElseIf mod1.DName = "颜继明" Then
                tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where (区域='杭州' or 区域='南京' or 区域='烟台') and 合同金额=" & Val(txtZ) & " and 合同编号<>'HMNEW' order by 合同日期 desc"
            ElseIf mod1.DName = "胡文婷" Then
                tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where (区域='南京' or 区域='烟台') and 合同金额=" & Val(txtZ) & " and 合同编号<>'HMNEW' order by 合同日期 desc"
            End If
        If mod1.Qy <> "上海" And mod1.Bm <> "武汉" And mod1.Bq2 = True Then '外地办文员
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 区域='" & mod1.Qy & "' and 合同金额=" & Val(txtZ) & " and 合同编号<>'HMNEW'  order by 合同日期 desc"
        ElseIf mod1.DName = "高芳" Or mod1.DName = "吴芳" Or mod1.DName = "王国君" Then    '外地办文员
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 区域='" & mod1.Qy & "' and 合同金额=" & Val(txtZ) & " and 合同编号<>'HMNEW'  order by 合同日期 desc"
        ElseIf mod1.DName = "霍艳" Or mod1.DName = "王全红" Or mod1.DName = "张婉秋" Or mod1.DName = "曾弋津" Or mod1.DName = "李建" Or mod1.DName = "杨燕" Then

            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 区域='北京'  and 合同金额=" & Val(txtZ) & " and 合同编号<>'HMNEW'order by 合同日期 desc  "
        
'''        ElseIf mod1.DName = "邹晨" Then
'''            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr  from htView1 where 合同金额=" & Val(txtZ) & _
'''            " and 公司代号=" & mod1.comId & "  and not(部门='维销部3' or 部门='产品部1' or 部门='产品部2') and 合同编号<>'HMNEW'   order by 合同日期 desc"
        ElseIf mod1.DName = "徐瑛" Or mod1.DName = "邹晨" Or mod1.DName = "陆军" Or mod1.DName = "张戬贇" Or mod1.DName = "乔继敏" Or mod1.DName = "陈文超" Or mod1.DName = "倪东海" Or mod1.DName = "朱婷婷" Or mod1.DName = "王绣霞" Or mod1.DName = "徐瑛" Or mod1.DName = "宋晓丹" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 合同金额=" & Val(txtZ) & " and 合同编号<>'HMNEW'   order by 合同日期 desc"
        ElseIf mod1.KhK = 3 Or mod1.DName = "周春云" Or mod1.DName = "马晓聪" Or mod1.DName = "倪东海" Or mod1.Bm = "商务部" Or mod1.DName = "乔继敏" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 合同金额=" & Val(txtZ) & " and 合同编号<>'HMNEW'  order by 合同日期 desc"
'''        ElseIf mod1.DName = "徐瑛" Then
'''            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr  from htView1 where 合同金额=" & Val(txtZ) & _
'''            " and ((合同性质<>'E. 产品合同'  and newf<2) or ((合同性质='维保' or 合同性质='大修') and newf=2))  and 合同编号<>'HMNEW'   order by 合同日期 desc"
'        ElseIf mod1.DName = "郑刚" Then
'            tt = "Select * from htView1 where 合同金额=" & Val(txtZ) & " and (合同性质='大修' or 合同性质='维保' or 合同性质='C. 维保合同' or 合同性质='D. 维修合同' or 合同性质='工程分包')  and 区域<>'上海' and 公司代号=0  order by 合同日期 desc"

        ElseIf mod1.DName = "" Or mod1.DName = "徐瑛" Or mod1.DName = "宋晓丹" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 合同金额=" & Val(txtZ) & "  order by 合同日期 desc"
        ElseIf mod1.DName = "郑刚" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 合同金额=" & Val(txtZ) & " and 合同编号<>'HMNEW'   and 项目归属人='王红'    order by 合同日期 desc"
        
        End If
        If mod1.DName = "陆军" Or mod1.DName = "张戬贇" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 合同金额=" & Val(txtZ) & " and 合同编号<>'HMNEW' and (ggl='" & mod1.DHid & "' or ywy='张平')  order by 合同日期 desc"
        
        End If
'''''        If mod1.BM = "工程二部" Then
'''''            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr  from htViewP where 合同金额=" & Val(txtZ) & " and 合同编号<>'HMNEW'  order by 部门,合同日期 desc"
'''''        End If
    Case "项目名称"
        tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where  ggl='" & mod1.DHid & "'  and 项目名称 like '%" & Trim(txtZ) & "%' and 合同编号<>'HMNEW'  order by 合同日期 desc"
            If mod1.DName = "初永友" Then
                tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where  部门='" & mod1.Bm & "'  and 项目名称 like '%" & Trim(txtZ) & "%' and 合同编号<>'HMNEW'  order by 合同日期 desc"
            ElseIf mod1.DName = "颜继明" Then
                tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where  (区域='杭州' or 区域='南京' or 区域='烟台')  and 项目名称 like '%" & Trim(txtZ) & "%' and 合同编号<>'HMNEW'  order by 合同日期 desc"
            ElseIf mod1.DName = "胡文婷" Then
                tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where  (区域='南京' or 区域='烟台')  and 项目名称 like '%" & Trim(txtZ) & "%' and 合同编号<>'HMNEW'  order by 合同日期 desc"
            End If
       If mod1.Qy <> "上海" And mod1.Bm <> "武汉" And mod1.Bq2 = True Then '外地办文员
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 区域='" & mod1.Qy & "' and 项目名称 like '%" & Trim(txtZ) & "%'  and 合同编号<>'HMNEW'   order by 合同日期 desc"
       ElseIf mod1.DName = "高芳" Or mod1.DName = "吴芳" Or mod1.DName = "王国君" Then    '外地办文员
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where  区域='" & mod1.Qy & "' and 项目名称 like '%" & Trim(txtZ) & "%'  and 合同编号<>'HMNEW'   order by 合同日期 desc"
        ElseIf mod1.DName = "霍艳" Or mod1.DName = "王全红" Or mod1.DName = "张婉秋" Or mod1.DName = "曾弋津" Or mod1.DName = "李建" Or mod1.DName = "杨燕" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where (区域='北京')  and 项目名称 like '%" & Trim(txtZ) & "%'  and 合同编号<>'HMNEW'   order by 合同日期 desc"
'''        ElseIf mod1.DName = "邹晨" Then
'''            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr  from htView1 where 项目名称 like '%" & Trim(txtZ) & _
'''            "%' and 公司代号=" & mod1.comId & "  and not(部门='维销部3' or 部门='产品部1' or 部门='产品部2')  and 合同编号<>'HMNEW'    order by 合同日期 desc"
        ElseIf mod1.DName = "徐瑛" Or mod1.DName = "邹晨" Or mod1.DName = "陆军" Or mod1.DName = "张戬贇" Or mod1.DName = "乔继敏" Or mod1.DName = "陈文超" Or mod1.DName = "朱婷婷" Or mod1.DName = "王绣霞" Or mod1.DName = "徐瑛" Or mod1.DName = "宋晓丹" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 项目名称 like '%" & Trim(txtZ) & "%' and 合同编号<>'HMNEW'   order by 合同日期 desc"
        ElseIf mod1.KhK = 3 Or mod1.DName = "周春云" Or mod1.DName = "马晓聪" Or mod1.DName = "倪东海" Or mod1.Bm = "商务部" Or mod1.DName = "乔继敏" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr  from htView1 where 项目名称 like '%" & Trim(txtZ) & "%'     order by 合同日期 desc"
''''        ElseIf mod1.DName = "徐瑛" Then
''''            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr from htView1 where 项目名称 like '%" & Trim(txtZ) & _
''''            "%' and ((合同性质<>'E. 产品合同'  and newf<2) or ((合同性质='维保' or 合同性质='大修') and newf=2))  and 合同编号<>'HMNEW'   order by 合同日期 desc"
        ElseIf mod1.DName = "彭海翔" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 项目名称 like '%" & Trim(txtZ) & _
            "%' and (合同性质='大修' or 合同性质='维保' or 合同性质='C. 维保合同' or 合同性质='D. 维修合同' or 合同性质='工程分包')   and 公司代号=1  and 合同编号<>'HMNEW'    order by 合同日期 desc"
        ElseIf mod1.DName = "" Or mod1.DName = "徐瑛" Or mod1.DName = "宋晓丹" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 项目名称 like '%" & Trim(txtZ) & _
            "%'   order by 合同日期 desc"
        ElseIf mod1.DName = "郑刚" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 项目名称 like '%" & Trim(txtZ) & "%' and 合同编号<>'HMNEW'  and 项目归属人='王红'     order by 合同日期 desc"
        ElseIf mod1.DName = "汪燕明" Or mod1.DName = "朱婷婷" Or mod1.DName = "吴金荣" Or mod1.DName = "吴金荣" Then
            tt = "Select 项目归属人,项目名称,合同日期,合同性质,'不知道',合同编号,状态,hid,newF,lr,fid  from htView1 where 项目名称 like '%" & Trim(txtZ) & "%'    order by 合同日期 desc"

        End If
        If mod1.Bm = "工程二部" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr  from htViewP where  项目名称 like '%" & Trim(txtZ) & "%'  and 合同编号<>'HMNEW'   order by 部门,合同日期 desc"
        End If
        If mod1.DName = "陆军" Or mod1.DName = "张戬贇" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 项目名称 like '%" & txtZ.Text & "%' and 合同编号<>'HMNEW' and (ggl='" & mod1.DHid & "' or ywy='张平')  order by 合同日期 desc"
        
        End If
    Case "合同编号"

        tt = "Select  top 1 项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 部门='" & mod1.Bm & "' and 合同编号 like '%" & Trim(txtZ) & "%' order by 合同日期 desc"

            If mod1.DName = "初永友" Then
                tt = "Select top 1  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 部门='" & mod1.Bm & "' and 合同编号 like '%" & Trim(txtZ) & "%' order by 合同日期 desc"
            ElseIf mod1.DName = "颜继明" Then
                        tt = "Select  top 1 项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where ( 区域='杭州' or  区域='南京' or  区域='烟台') and 合同编号 like '%" & Trim(txtZ) & "%' order by 合同日期 desc"
            ElseIf mod1.DName = "胡文婷" Then
                        tt = "Select  top 1 项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where (区域='南京' or 区域='烟台') and 合同编号 like '%" & Trim(txtZ) & "%' order by 合同日期 desc"
            End If
        If mod1.Qy <> "上海" And mod1.DName <> "武汉" And mod1.Bq2 = True Then  '外地办文员
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 区域='" & mod1.Qy & "' and 合同编号 like '%" & Trim(txtZ) & "%' order by 合同日期 desc"
        ElseIf mod1.DName = "高芳" Or mod1.DName = "吴芳" Or mod1.DName = "王国君" Then  '外地办文员
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 区域='" & mod1.Qy & "' and 合同编号 like '%" & Trim(txtZ) & "%' order by 合同日期 desc"
        ElseIf mod1.DName = "霍艳" Or mod1.DName = "王全红" Or mod1.DName = "张婉秋" Or mod1.DName = "曾弋津" Or mod1.DName = "李建" Or mod1.DName = "杨燕" Then
            tt = "Select   项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where (区域='北京')  and 合同编号 like '%" & Trim(txtZ) & "%' order by 合同日期 desc"
        ElseIf mod1.DName = "徐瑛" Or mod1.DName = "邹晨" Or mod1.DName = "陆军" Or mod1.DName = "张戬贇" Or mod1.DName = "乔继敏" Or mod1.DName = "陈文超" Or mod1.DName = "倪东海" Or mod1.DName = "朱婷婷" Or mod1.DName = "王绣霞" Or mod1.DName = "徐瑛" Or mod1.DName = "宋晓丹" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where hid= " & Trim(txtZ) & "  order by 合同日期 desc"
        ElseIf mod1.KhK = 3 Or mod1.DName = "周春云" Or mod1.DName = "马晓聪" Or mod1.DName = "倪东海" Or mod1.Bm = "商务部" Or mod1.DName = "乔继敏" Then
            tt = "Select 项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 合同编号 like '%" & Trim(txtZ) & "%'  order by 合同日期 desc"
'''        ElseIf mod1.DName = "徐瑛" Then
'''            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr  from htView1 where 合同编号 like '%" & Trim(txtZ) & _
'''            "%' and ((合同性质<>'E. 产品合同'  and newf<2) or ((合同性质='维保' or 合同性质='大修') and newf=2))  order by 合同日期 desc"
        ElseIf mod1.DName = "彭海翔" Then
            tt = "Select top 1 项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 合同编号 like '%" & Trim(txtZ) & "%' and (合同性质='大修' or 合同性质='维保' or 合同性质='C. 维保合同' or 合同性质='D. 维修合同' or 合同性质='工程分包')   and 公司代号=1   order by 合同日期 desc"
        ElseIf mod1.DName = "" Or mod1.DName = "徐瑛" Or mod1.DName = "宋晓丹" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 合同编号 like '%" & _
            Trim(txtZ) & "%'   order by 合同日期 desc"
        ElseIf mod1.DName = "郑刚" Then
            tt = "Select 项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 合同编号 like '%" & Trim(txtZ) & "%'  and 项目归属人='王红'      order by 合同日期 desc"
        ElseIf mod1.DName = "汪燕明" Or mod1.DName = "朱婷婷" Or mod1.DName = "吴金荣" Or mod1.DName = "吴金荣" Then
            tt = "Select 项目归属人,项目名称,合同日期,合同性质,'不知道',合同编号,状态,hid,newF,lr,fid  from htView1 where 合同编号 like '%" & Trim(txtZ) & "%'    order by 合同日期 desc"
        End If
'''''        If mod1.BM = "工程二部" Then
'''''            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr  from htViewP where  合同编号 like '%" & Trim(txtZ) & "%' order by 部门,合同日期 desc"
'''''        End If
        If mod1.DName = "陆军" Or mod1.DName = "张戬贇" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 合同编号 like '%" & txtZ.Text & "%' and 合同编号<>'HMNEW' and (ggl='" & mod1.DHid & "' or ywy='张平')  order by 合同日期 desc"
        
        End If
    Case "合同执行编号"
        tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 部门='" & mod1.Bm & "' and 合同编号 like '%" & Trim(txtZ) & "%' order by 合同日期 desc"


        If mod1.Qy <> "上海" And mod1.Bq2 = True Then '外地办文员
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where 区域='" & mod1.Qy & "' and zbh like '%" & Trim(txtZ) & "%' order by 合同日期 desc"
'''        ElseIf mod1.DName = "邹晨" Then
'''            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr  from htView1 where zbh like '%" & Trim(txtZ) & "%' and 公司代号=" & mod1.comId & _
'''            "  and not(部门='维销部3' or 部门='产品部1' or 部门='产品部2')  order by 合同日期 desc"
        ElseIf mod1.DName = "徐瑛" Or mod1.DName = "邹晨" Or mod1.DName = "陆军" Or mod1.DName = "张戬贇" Or mod1.DName = "乔继敏" Or mod1.DName = "倪东海" Or mod1.DName = "孟智峰" Or mod1.DName = "朱婷婷" Or mod1.DName = "王绣霞" Or mod1.DName = "徐瑛" Or mod1.DName = "宋晓丹" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where zbh like '%" & Trim(txtZ) & "%' order by 合同日期 desc"
        ElseIf mod1.KhK = 3 Or mod1.DName = "周春云" Or mod1.DName = "马晓聪" Or mod1.DName = "倪东海" Or mod1.Bm = "商务部" Or mod1.DName = "乔继敏" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where zbh like '%" & Trim(txtZ) & "%'  order by 合同日期 desc"
'''        ElseIf mod1.DName = "徐瑛" Then
'''            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr  from htView1 where 合同编号 like '%" & Trim(txtZ) & _
'''            "%' and ((合同性质<>'E. 产品合同'  and newf<2) or ((合同性质='维保' or 合同性质='大修') and newf=2))  order by 合同日期 desc"
        ElseIf mod1.DName = "彭海翔" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where zbh like '%" & Trim(txtZ) & "%' and (合同性质='大修' or 合同性质='维保' or 合同性质='C. 维保合同' or 合同性质='D. 维修合同' or 合同性质='工程分包')   and 公司代号=1   order by 合同日期 desc"
        ElseIf mod1.DName = "" Or mod1.DName = "徐瑛" Or mod1.DName = "宋晓丹" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where zbh like '%" & _
            Trim(txtZ) & "%'   order by 合同日期 desc"
        ElseIf mod1.DName = "霍艳" Or mod1.DName = "王全红" Or mod1.DName = "张婉秋" Or mod1.DName = "曾弋津" Or mod1.DName = "李建" Or mod1.DName = "杨燕" Then
            tt = "Select  项目归属人,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htView1 where zbh like '%" & Trim(txtZ) & "%' and 区域='北京' order by 合同日期 desc"
        End If
End Select

        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        ua = UBound(Ra, 2)
        Call Bref(Ra, ua + 1)
End Sub

Private Sub cmdVall_Click()
'''''Dim tt As String
'''''On Error Resume Next
'''''    If mod1.KhK = 1 And mod1.BM <> "行政人事" And mod1.BM = "维销部2" Then
'''''        tt = "Select * from htView1 where 部门='" & mod1.BM & "' order by 部门,合同日期 desc"
'''''    ElseIf mod1.KhK = 1 And mod1.BM <> "行政人事" And mod1.BM <> "维销部2" Then
'''''        tt = "Select * from htView2 where 部门='" & mod1.BM & "' order by 部门,合同日期 desc"
'''''    ElseIf mod1.KhK = 1 And mod1.BM = "行政人事" And Qy <> "上海" Then
'''''        tt = "select * FROM HTVIEW WHERE 区域='" & mod1.Qy & "' order by 部门,合同日期 desc"
'''''    ElseIf (mod1.KhK = 2 Or mod1.DName = "徐瑛") And Not (mod1.DName = "张寅" Or mod1.DName = "彭海翔") Then
'''''        If mod1.Qy = "广州" Then
'''''            tt = "Select * from htView1 where 区域='广州'  order by 合同日期 desc"
'''''        Else
'''''            tt = "Select * from htView1 where 区域<>'广州' and not(部门='维销部3' or 部门='产品部1' or 部门='产品部2')   order by 合同日期 desc"
'''''        End If
'''''    ElseIf mod1.DName = "王蕾" Then
'''''            tt = "select * from htview1 where 公司代号=" & mod1.DName & " order by 合同日期 desc"
'''''    ElseIf mod1.BM = "商务部" Then
'''''            tt = "select * from htview1 order by 合同日期 desc"
'''''    ElseIf mod1.Bq2 = True And mod1.Qy <> "上海" Then
'''''            tt = "select * from htview1 and 区域='" & mod1.Qy & "' order by 合同日期 desc"
'''''    ElseIf mod1.KhK = 3 Then
'''''            tt = "select * from htview1 where 公司代号=" & mod1.comId & " order by 合同日期 desc"
'''''    ElseIf mod1.DName = "张寅" Then
'''''        tt = "Select * from htView1 where (合同性质='大修' or 合同性质='维保' or 合同性质='C. 维保合同' or 合同性质='D. 维修合同') and not(状态='评审' or 状态='盖章')  and 公司代号=0   order by 合同日期 desc"
''''''    ElseIf mod1.DName = "郑刚" Then
''''''        tt = "Select * from htView1 where (合同性质='大修' or 合同性质='维保' or 合同性质='C. 维保合同' or 合同性质='D. 维修合同') and not(状态='评审' or 状态='盖章') and 区域<>'上海' and 公司代号=0  order by 合同日期 desc"
'''''    ElseIf mod1.DName = "彭海翔" Then
'''''        tt = "Select * from htView1 where (合同性质='大修' or 合同性质='维保' or 合同性质='C. 维保合同' or 合同性质='D. 维修合同') and not(状态='评审' or 状态='盖章')  and 公司代号=1   order by 合同日期 desc"
'''''    End If
'''''    If mod1.BM = "工程二部" Then
'''''        tt = "Select * from htViewP order by 部门,合同日期 desc"
'''''    End If
'''''    htBrowG.adoBr.Close
'''''    htBrowG.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''    Set htBrowG.dtgBr.DataSource = htBrowG.adoBr
'''''    If htBrowG.adoBr.RecordCount > 0 Then
'''''        htBrowG.dtgBr.FixedRows = 0
'''''        htBrowG.dtgBr.MergeCol(1) = True
'''''        htBrowG.dtgBr.MergeCol(2) = True
'''''        htBrowG.dtgBr.MergeCol(3) = True
'''''        htBrowG.dtgBr.MergeCol(4) = True
'''''        htBrowG.dtgBr.MergeCol(7) = True
'''''        htBrowG.dtgBr.MergeCol(13) = True
'''''        htBrowG.dtgBr.MergeCells = 3
'''''        htBrowG.dtgBr.FixedRows = 1
'''''    End If
End Sub

Private Sub cmdWZX_Click()
Dim tt As String
Dim oo As Integer: Dim ii As Integer
Dim Ra, ua
On Error Resume Next
            tt = "Select  项目归属人,项目名称,qrq,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htview6 where ggl='" & mod1.DHid & "' order by qrq desc"
            If mod1.DName = "初永友" Then
                tt = "Select  项目归属人,项目名称,qrq,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htview6 where 部门='" & mod1.Bm & "' order by qrq desc"
            ElseIf mod1.DName = "颜继明" Then
                tt = "Select  项目归属人,项目名称,qrq,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htview6 where (区域='杭州' or 区域='南京' or 部门='烟台') order by qrq desc"
            ElseIf mod1.DName = "胡文婷" Then
                tt = "Select  项目归属人,项目名称,qrq,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htview6 where (区域='南京') order by qrq desc"
            End If
        If mod1.Qy <> "上海" And mod1.DName <> "武汉" And mod1.Bq2 = True Then '外地办文员
            tt = "Select  项目归属人,项目名称,qrq,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htview6 where 区域='" & mod1.Qy & "' order by qrq desc"
        ElseIf mod1.DName = "高芳" Or mod1.DName = "吴芳" Then  '外地办文员
            tt = "Select  项目归属人,项目名称,qrq,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htview6 where (区域='武汉' or 区域='广州') order by qrq desc"
        ElseIf mod1.DName = "霍艳" Then

            tt = "Select  项目归属人,项目名称,qrq,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htview6 where (区域='北京' )  order by qrq desc"
        

        ElseIf mod1.DName = "徐瑛" Or mod1.DName = "徐瑛" Or mod1.DName = "沈维" Or mod1.DName = "乔继敏" Or mod1.DName = "陈文超" Or mod1.DName = "孟智峰" Or mod1.DName = "朱婷婷" Then
            tt = "Select  项目归属人,项目名称,qrq,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htview6 order by qrq desc"
        ElseIf mod1.KhK = 3 Or mod1.DName = "周春云" Or mod1.DName = "马晓聪" Or mod1.Bm = "商务部" Then
            tt = "Select  项目归属人,项目名称,qrq,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htview6 order by qrq desc"

        ElseIf mod1.DName = "" Or mod1.DName = "吴金荣" Or mod1.DName = "徐瑛" Then
            tt = "Select  项目归属人,项目名称,qrq,合同性质,合同金额,合同编号,状态,hid,newF,lr,fid  from htview6 order by qrq desc"
            Exit Sub
        End If
        
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        ua = UBound(Ra, 2)
        Call Bref(Ra, ua + 1)
End Sub

Private Sub cmdX_Click()
FmxcXB.Show

End Sub

Private Sub cmdXJView_Click()
Dim Ra
Dim La As Integer
Me.Enabled = True
    mod1.BTZ = 36
    frmGxBiao.Visible = False
    tt = "SELECT  dbo.XunJiaD.bid,dbo.XunJiaD.xmmc, HMText.dbo.xunC.trq, " & _
      " HMText.dbo.xunC.ywy , dbo.XunJiaD.DelF " & _
    " FROM dbo.XunJiaD RIGHT OUTER JOIN" & _
    "  HMText.dbo.xunC ON dbo.XunJiaD.bid = HMText.dbo.xunC.bh" & _
    " where year(HMText.dbo.xunC.Trq)=" & Year(dd1.Value) & _
    " and month(HMText.dbo.xunC.trq)=" & Month(dd1.Value) & _
    " and day(HMText.dbo.xunC.trq)=" & Day(dd1.Value) & _
    " and (hmtext.dbo.xunc.ywy='朱婷婷' or hmtext.dbo.xunc.ywy='吴金荣') order by dbo.xunjiad.bid desc"
    
    Set mod1.HTP = CreateObject("adodb.recordset")
    On Error Resume Next
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2)
    Call dtgbrFF(Ra, La)
End Sub

Private Sub cmdZF_Click()
Dim tt As String
Dim oo As Integer: Dim ii As Integer
Dim Ra, ua
On Error Resume Next
            tt = "Select  业务员,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,0,fid  from htviewdel where ggl='" & mod1.DHid & "' order by 合同日期 desc"
            If mod1.DName = "颜继明" Then
                tt = "Select  业务员,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,0,fid  from htviewdel where (区域='杭州' or 区域='南京' or 部门='烟台') order by 合同日期 desc"
            ElseIf mod1.DName = "胡文婷" Then
                tt = "Select  业务员,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,0,fid  from htviewdel where (区域='南京') order by 合同日期 desc"
            End If
        If mod1.Qy <> "上海" And mod1.DName <> "武汉" And mod1.Bq2 = True Then '外地办文员
            tt = "Select  业务员,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,0,fid  from htviewdel where 区域='" & mod1.Qy & "' order by 合同日期 desc"
        ElseIf mod1.DName = "高芳" Or mod1.DName = "吴芳" Then  '外地办文员
            tt = "Select  业务员,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,0,fid  from htviewdel where (区域='武汉' or 区域='广州') order by 合同日期 desc"
        ElseIf mod1.DName = "霍艳" Then

            tt = "Select  业务员,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,0,fid  from htviewdel where (区域='北京' )  order by 合同日期 desc"
        

        ElseIf mod1.DName = "徐瑛" Or mod1.DName = "倪旭" Or mod1.DName = "沈维" Or mod1.DName = "乔继敏" Or mod1.DName = "乔继敏" Or mod1.DName = "孟智峰" Then
            tt = "Select  业务员,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,0,fid  from htviewdel order by 合同日期 desc"
        ElseIf mod1.KhK = 3 Or mod1.DName = "周春云" Or mod1.DName = "马晓聪" Or mod1.Bm = "商务部" Then
            tt = "Select  业务员,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,0,fid  from htviewdel order by 合同日期 desc"

        ElseIf mod1.DName = "" Or mod1.DName = "吴金荣" Or mod1.DName = "徐瑛" Then
            tt = "Select  业务员,项目名称,合同日期,合同性质,合同金额,合同编号,状态,hid,newF,0,fid  from htviewdel order by 合同日期 desc"
            Exit Sub
        End If
        
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        ua = UBound(Ra, 2)
        Call Bref(Ra, ua + 1)
End Sub

Private Sub dtgBr_DblClick()
Static Px As Boolean

If dtgBr.Row = 1 Then
    If Px = True Then
        dtgBr.Sort = 2
        Px = False
    Else
        dtgBr.Sort = 1
        Px = True
    End If
'Else
'    MsgBox MGa.ColData(1)
End If
End Sub


Private Sub dtgBr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static ZF As Boolean
If Button <> 2 Then Exit Sub
If ZF = False Then
        htBrowG.dtgBr.FixedRows = 0
        htBrowG.dtgBr.MergeCol(1) = True
        htBrowG.dtgBr.MergeCol(2) = True
        htBrowG.dtgBr.MergeCol(3) = True
        htBrowG.dtgBr.MergeCol(4) = True
        htBrowG.dtgBr.MergeCol(7) = True
        htBrowG.dtgBr.MergeCol(13) = True
        htBrowG.dtgBr.MergeCells = 0
        htBrowG.dtgBr.FixedRows = 1
        ZF = True
Else
        htBrowG.dtgBr.FixedRows = 0
        htBrowG.dtgBr.MergeCol(1) = True
        htBrowG.dtgBr.MergeCol(2) = True
        htBrowG.dtgBr.MergeCol(3) = True
        htBrowG.dtgBr.MergeCol(4) = True
        htBrowG.dtgBr.MergeCol(7) = True
        htBrowG.dtgBr.MergeCol(13) = True
        htBrowG.dtgBr.MergeCells = 3
        htBrowG.dtgBr.FixedRows = 1
        ZF = False
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 68 And mod1.DName = "马晓聪" Then
    If cmdDel.Visible = True Then
        cmdDel.Visible = False
    Else
        cmdDel.Visible = True
    End If
End If

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight

Set adoBr = CreateObject("adodb.recordset")
dtgBr.ColWidth(0) = 300
dtgBr.ColWidth(2) = 4000
dtgBr.ColWidth(3) = 1900
dtgBr.ColWidth(4) = 1200
dtgBr.ColWidth(6) = 2000
dtgBr.ColWidth(7) = 1000
dtgBr.ColWidth(8) = 0 'hid
dtgBr.ColWidth(9) = 0 'newF
dtgBr.ColWidth(10) = 0 'lr
comLx.Text = "合同编号"
''''''nc.LoadSkin 5
If mod1.DName <> "徐瑛" And mod1.DName <> "张文琴" And mod1.DName <> "乔继敏" And mod1.DName <> "于晓静" Then
    cmdX.Visible = False
End If
If mod1.DName <> "陈文超" And mod1.DName <> "张文琴" And mod1.DName <> "乔继敏" And mod1.DName <> "顾珣" Then
    cmdCG.Visible = False
End If
If mod1.DName <> "宋晓丹" And mod1.DName <> "马晓聪" And mod1.DName <> "乔继敏" Then
    cmdXJView.Visible = False

    frmXDC.Visible = True
Else
    frmXDC.Visible = True
        dd1.Value = mod1.DQda
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
htBrowG.Visible = False
frmZu.Enabled = True
Cancel = True
End If
End Sub

Private Sub OKButton_Click()
Dim tt As String
Dim xZ As String
Dim NewF As Integer
Dim Hid As Long
Dim Bid As Long
Dim ZL As String
Dim Ra
'Dim Lid As String
On Error Resume Next

''''''If mod1.DName = "张砚纯" Or ((mod1.DName = "周春云" Or mod1.DName = "" Or mod1.DName = "倪旭") And frmZu.NC.Caption = "产品事务") Then
''''''    dtgBr.Col = 5: ZL = Trim(dtgBr.Text)
''''''    dtgBr.Col = 8: Bid = Val(dtgBr.Text)
''''''
''''''        If ZL <> "配件" And ZL <> "产品" Then
''''''            Call frmWBXX.Qing
''''''            Call frmWBXX.Bound(Bid)
''''''            frmWBXX.Show
''''''            frmWBXX.ZOrder 0
''''''            Exit Sub
''''''        End If
''''''
''''''        If Bid > 8113 And ZL <> "配件" And ZL <> "产品" Then
''''''            Call frmWBXNew.Qing
''''''            Call frmWBXNew.Bound(Bid)
''''''            frmWBXNew.Show
''''''            frmWBXNew.ZOrder 0
''''''            Exit Sub
''''''        End If
''''''    Call modBJD.BJDWBQing
''''''    Call modBJD.BJDGXQing
''''''    Call modBJD.BJDBound(Bid, ZL)
''''''    'frmWBXJ.Show
''''''    If ZL = "产品" Or ZL = "零配件" Or ZL = "配件" Then
''''''    frmGXBj.Show
''''''    Else
''''''    frmWBXJ.Show
''''''    End If
''''''    Exit Sub
''''''End If

' 判断是打开询价单还是合同评审。
dtgBr.Col = 1
If Val(dtgBr.Text) > 0 Then
    Bid = Val(dtgBr.Text)
    If Bid = 0 Then Exit Sub
    Call FmxcXJ.Bound(Bid)
    FmxcXJ.Show
    FmxcXJ.ZOrder 0
Exit Sub
End If

'打开合同评审
dtgBr.Col = 4
xZ = dtgBr.Text
dtgBr.Col = 8
Hid = dtgBr.Text
dtgBr.Col = 9
NewF = dtgBr.Text
'Lid = Str(Lid)

If (mod1.Bm = "技术中心" Or mod1.Bm = "维保中心") And mod1.DName <> "朱婷婷" And mod1.DName <> "周春云" Then
    tt = "select htbh from htping where hid=" & Hid
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    tt = "select fid from ht where htbh='" & Ra(0, 0) & "' and xz=1"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workHT, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    
    Dim bt() As Byte
    On Error Resume Next
    Kill "c:\work\*.xls": Kill "c:\work\*.doc"
    tt = "select fnr,fsize,fname from ht where fid=" & Val(Ra(0, 0)) & " and xz=1"
    frmGGL.adoFile.Recordset.Close
    frmGGL.adoFile.Recordset.Open tt, mod1.workHT, adOpenKeyset, adLockReadOnly, adCmdText
    ReDim bt(frmGGL.adoFile.Recordset.Fields("Fsize").Value) As Byte
    bt() = frmGGL.adoFile.Recordset.Fields("FNR").GetChunk(frmGGL.adoFile.Recordset.Fields("Fsize").Value + 1)
    
    Open ("c:\work\" & frmGGL.adoFile.Recordset.Fields("fname").Value) For Binary As #2
    Put #2, , bt()
    Close #2
    
        frmGGL.OLE2.SourceDoc = "c:\work\" & frmGGL.adoFile.Recordset.Fields("fname").Value
        frmGGL.OLE2.Action = 1
        frmGGL.OLE2.DoVerb (-2)
    
    Exit Sub
End If

mod1.BTZ = 6
If mod1.DKZ(Hid, 1) = True Then
        MsgBox "这份表单正由" & mod1.DKRen & "打开,请稍候再试,或与马晓聪联系."
        Exit Sub
End If

frmWait.Visible = True
frmWait.ZOrder 0
frmWait.Refresh
'htBrowg.MousePointer = 11
htBrowG.Enabled = False
'mod1.MPld = False '初始化,不生成配料单
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

FmxcNew.Width = mod1.FWidth + 500
FmxcNew.Height = mod1.FHeight
FmxcNew.frmNewLx.Left = 5070
FmxcNew.frmNewLx.Top = 0
''''    If mod1.DName = "朱婷婷" Or mod1.DName = "汪燕明" Or mod1.DName = "吴金荣" Or mod1.DName = "吴金荣" Then
''''        Call FmxcNew.Xian
''''    End If
End Sub


Public Sub Bref(Ra, ua)
On Error Resume Next
Dim oo As Integer
Dim ii As Integer: Dim LR As Double
    htBrowG.dtgBr.Visible = False
    htBrowG.dtgBr.Clear
    htBrowG.dtgBr.Row = 0: htBrowG.dtgBr.Col = 1: htBrowG.dtgBr.Text = "项目归属人"
    htBrowG.dtgBr.Col = 2: htBrowG.dtgBr.Text = "项目名称": htBrowG.dtgBr.Col = 3: htBrowG.dtgBr.Text = "日期": htBrowG.dtgBr.Col = 4: htBrowG.dtgBr.Text = "合同性质"
    htBrowG.dtgBr.Col = 5: htBrowG.dtgBr.Text = "合同金额": htBrowG.dtgBr.Col = 6: htBrowG.dtgBr.Text = "合同编号": htBrowG.dtgBr.Col = 7: htBrowG.dtgBr.Text = "状态"
    For ii = 1 To 12
        dtgBr.Col = ii
        dtgBr.CellFontBold = True
    Next
    For oo = 1 To ua
'        dtgBr.Col = 10
'        lr = Val(dtgBr.Text)
        dtgBr.Row = oo
        For ii = 1 To 13
            htBrowG.dtgBr.Col = ii
            htBrowG.dtgBr.Text = Trim(Ra(ii - 1, oo - 1))
            If Ra(9, oo - 1) > 0 Then
                dtgBr.CellForeColor = &H0&
            Else
                dtgBr.CellForeColor = &HFF&
            End If
            If Ra(10, oo - 1) > 0 Then
                dtgBr.CellForeColor = &H800000
            End If
        Next
    Next
    dtgBr.Visible = True
    If ua < 32 Then
        dtgBr.Rows = 32
    Else
        dtgBr.Rows = ua + 2
    End If
End Sub

Private Sub txtZ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmdRef_Click
End If
End Sub


