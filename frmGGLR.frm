VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGGLKan 
   Caption         =   "消息浏览"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   8415
   Begin VB.CommandButton cmdRef 
      Caption         =   "查   询"
      Height          =   315
      Left            =   6720
      TabIndex        =   6
      Top             =   8640
      Width           =   1185
   End
   Begin VB.ComboBox comBj 
      Height          =   300
      Left            =   3270
      TabIndex        =   4
      Text            =   "comBj"
      Top             =   8670
      Width           =   825
   End
   Begin VB.ComboBox comLx 
      Height          =   300
      ItemData        =   "frmGGLR.frx":0000
      Left            =   1140
      List            =   "frmGGLR.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   8670
      Width           =   1485
   End
   Begin MSComCtl2.DTPicker dtpZ 
      Height          =   255
      Left            =   4650
      TabIndex        =   5
      Top             =   8670
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   450
      _Version        =   393216
      CalendarBackColor=   8454016
      CalendarTitleBackColor=   16711808
      CalendarTrailingForeColor=   -2147483635
      Format          =   106758145
      CurrentDate     =   38797
   End
   Begin VB.ComboBox txtZ 
      Height          =   300
      ItemData        =   "frmGGLR.frx":003D
      Left            =   4650
      List            =   "frmGGLR.frx":003F
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   8640
      Width           =   1965
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgJl 
      Height          =   8565
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   15108
      _Version        =   393216
      BackColor       =   -2147483634
      BackColorBkg    =   -2147483636
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label4 
      Caption         =   "比较:"
      Height          =   225
      Left            =   2760
      TabIndex        =   3
      Top             =   8700
      Width           =   585
   End
   Begin VB.Label Label3 
      Caption         =   "值:"
      Height          =   255
      Left            =   4260
      TabIndex        =   2
      Top             =   8700
      Width           =   315
   End
   Begin VB.Label Label1 
      Caption         =   "查询类别:"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   8700
      Width           =   885
   End
End
Attribute VB_Name = "frmGGLKan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AdoJl As Object

Public GGGCCC As Boolean '如果单击了滚动条,则出错将消失
Dim CCC As Boolean '网格点击出错否








Private Sub cmdRef_Click()
Dim tt As String
On Error Resume Next
Select Case comLx.Text
Case "类别"
    If txtZ.Text <> "其它类" Then
        tt = "select left(gnr,10)+'...' as 内容提要,zz as 发送人,rq as 发送日期,gid,lb as 类别," & mod1.DName & _
        " AS 看过否 from ggl where not(" & mod1.DName & " is null) and lb='" & txtZ.Text & "' and left(zz,1)<>'n' order by " & mod1.DName & ",gid desc"
    Else
        tt = "select left(gnr,10)+'...' as 内容提要,zz as 发送人,rq as 发送日期,gid,lb as 类别," & mod1.DName & _
        " AS 看过否 from ggl where not(" & mod1.DName & " is null) and lb is null and left(zz,1)<>'n'  order by " & mod1.DName & ",gid desc"
    End If
Case "发送人"
        tt = "select left(gnr,10)+'...' as 内容提要,zz as 发送人,rq as 发送日期,gid,lb as 类别," & mod1.DName & _
        " AS 看过否 from ggl where not(" & mod1.DName & " is null) and zz like '%" & txtZ.Text & "%'  and left(zz,1)<>'n'  order by " & mod1.DName & ",gid desc"
Case "发送日期"
    Select Case comBj.Text
    Case "="
            tt = "select left(gnr,10)+'...' as 内容提要,zz as 发送人,rq as 发送日期,gid,lb as 类别," & mod1.DName & _
        " AS 看过否 from ggl where not(" & mod1.DName & " is null) and year(rq)=" & Year(dtpZ.Value) & " and month(rq)=" & _
        Month(dtpZ.Value) & " and day(rq)=" & Day(dtpZ.Value) & " and left(zz,1)<>'n'  order by " & mod1.DName & ",gid desc"
    Case ">"
            tt = "select left(gnr,10)+'...' as 内容提要,zz as 发送人,rq as 发送日期,gid,lb as 类别," & mod1.DName & _
        " AS 看过否 from ggl where not(" & mod1.DName & " is null) and rq>='" & dtpZ.Value & "' and left(zz,1)<>'n'  order by " & mod1.DName & ",gid desc"
    Case "<"
            tt = "select left(gnr,10)+'...' as 内容提要,zz as 发送人,rq as 发送日期,gid,lb as 类别," & mod1.DName & _
        " AS 看过否 from ggl where not(" & mod1.DName & " is null) and rq<='" & dtpZ.Value & "' and left(zz,1)<>'n'  order by " & mod1.DName & ",gid desc"
    End Select
Case "内容"
            tt = "select left(gnr,10)+'...' as 内容提要,zz as 发送人,rq as 发送日期,gid,lb as 类别," & mod1.DName & _
        " AS 看过否 from ggl where not(" & mod1.DName & " is null) and gnr like '%" & txtZ.Text & "%' and left(zz,1)<>'n'  order by " & mod1.DName & ",gid desc"
Case "匿名类"
            tt = "select left(gnr,10)+'...' as 内容提要,'匿名者' as 发送人,rq as 发送日期,gid,lb as 类别," & mod1.DName & _
        " AS 看过否 from ggl where  left(zz,1)='n'  order by " & mod1.DName & ",gid desc"
End Select
Set frmGGLKan.AdoJl = CreateObject("adodb.recordset")
frmGGLKan.AdoJl.Close
frmGGLKan.AdoJl.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGGLKan.dtgJl.DataSource = frmGGLKan.AdoJl
'Call dtgJl.ScrollTrack
dtgJl.Row = frmGGLKan.AdoJl.RecordCount - 1
End Sub


Private Sub comLx_Click()
Dim oo As Integer
On Error Resume Next
For oo = 4 To 0 Step -1
    comBj.RemoveItem oo
Next
For oo = 9 To 0 Step -1
    txtZ.RemoveItem oo
Next
dtpZ.Visible = False
txtZ.Text = ""
Select Case comLx.Text
Case "类别"
    comBj.AddItem "="
    comBj.Text = "="
    txtZ.AddItem "公告类"
    txtZ.AddItem "一般类"
    txtZ.AddItem "通知类"
    txtZ.AddItem "派工类"
    txtZ.AddItem "到帐类"
    txtZ.AddItem "胡萝卜"
    txtZ.AddItem "晨会类"
    txtZ.AddItem "其它类"
    txtZ.AddItem "评审修改"
    txtZ.Text = "公告类"
Case "发送人"
    comBj.AddItem "包含"
    comBj.Text = "包含"
Case "发送日期"
    comBj.AddItem "="
    comBj.AddItem ">"
    comBj.AddItem "<"
    comBj.Text = "="
    dtpZ.Visible = True
Case "内容"
    comBj.AddItem "包含"
    comBj.Text = "包含"
End Select
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dtgJl_Click()
Dim tt As String
On Error Resume Next
'If frmGGLKan.GGGCCC = False Then
'    If CCC = False Then
'        dtgJl.Row = dtgJl.Row + 1
'    Else
'        dtgJl.Row = 1
'    End If
'End If
Zou:
On Error Resume Next
dtgJl.Col = 4
modGGL.Oid = dtgJl.Text
If Not (modGGL.Oid > 0) Then Exit Sub


    frmGGL.frmCa.Visible = False
    frmGGL.frmCb.Visible = False
    frmGGL.cmdSave.Enabled = False
    Set frmGGL.adoGGl = CreateObject("adodb.recordset")
        tt = "Select top 1 gnr,zz,rq,gid,fdx,wzid, " & mod1.DName & ",lb,fid from ggl where  gid=" & modGGL.Oid

    'frmGGL.adoGGl.Close
    frmGGL.adoGGl.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
    If frmGGL.adoGGl.RecordCount = 1 Then
        Oid = frmGGL.adoGGl.Fields("gid").Value
        frmGGL.rihNr.Text = frmGGL.adoGGl.Fields("Gnr").Value
        If Left(frmGGL.adoGGl.Fields("zz").Value, 1) = "n" Then
            frmGGL.lblZZ.Caption = "匿名者"
        Else
            frmGGL.lblZZ.Caption = frmGGL.adoGGl.Fields("zz").Value
        End If
        frmGGL.lblDate.Caption = frmGGL.adoGGl.Fields("rq").Value
        If IsNull(frmGGL.adoGGl.Fields("lb").Value) = True Then
            frmGGL.comLb.Visible = False
            frmGGL.lblLb.Visible = False
        Else
            frmGGL.comLb.Text = frmGGL.adoGGl.Fields("lb").Value
            frmGGL.comLb.Visible = True
            frmGGL.lblLb.Visible = True
            frmGGL.comLb.Locked = True
        End If

        frmGGL.Show
        frmGGL.ZOrder 0
        frmZu.Enabled = False
        
        '判断字颜色
        frmGGL.rihNr.SelStart = 0
        frmGGL.rihNr.SelLength = Len(frmGGL.rihNr.Text)
    
        If frmGGL.adoGGl.Fields(mod1.DName).Value = False Then
            frmGGL.rihNr.SelColor = &HFF0000
        Else
            frmGGL.rihNr.SelColor = &H80000012
        End If
    
        frmGGL.rihNr.SelFontSize = frmGGL.adoGGl.Fields("Fdx").Value
        frmGGL.rihNr.SelStart = 0
        frmGGL.rihNr.SelLength = 0
    'End If
    
        If frmGGL.lblZZ.Caption = mod1.DName Or mod1.DName = "马晓聪" Then
        frmGGL.cmdDel.Enabled = True
        Else
        frmGGL.cmdDel.Enabled = False
        End If
        

        frmGGL.cmdYjb.Visible = False
        
'        If IsNull(frmGGL.adoGG.Recordset.Fields("wzid").Value) = False Then
'
'
'            If Left(frmGGL.rihNr.Text, 3) = "请注意" Then
'                frmGGL.cmdYjb.Visible = True
'            Else
'                frmGGL.cmdXq.Visible = True
'            End If
'        End If
        frmGGL.cmdZx.Enabled = True
        frmGGL.cmdReply.Enabled = True
        frmGGL.frmRen.Visible = False
        frmGGL.cmdPre.Enabled = True
    Else
        frmGGL.cmdNext.Enabled = False
    End If
    frmGGL.cmdYjb.Visible = False

If frmGGL.adoGGl.RecordCount = 1 Then
    If IsNull(frmGGL.adoGGl.Fields("wzid").Value) = False Then
    
    
        If Left(frmGGL.rihNr.Text, 3) = "请注意" Then
            frmGGL.cmdYjb.Visible = True
        Else

        End If
    End If
End If

Exit Sub
'frmGGL.ZOrder 0

''取得公告栏的最大记录Id，以便和新的公告记录区别
'tt = "Select max(Gid) from ggl"
'adoMaxG.Recordset.Close
'adoMaxG.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'MaxGid = adoMaxG.Recordset.Fields(0).Value
End Sub

Private Sub dtgJl_DblClick()
Static Px As Boolean
On Error Resume Next
If dtgJl.Row = 1 Then
    If Px = True Then
        dtgJl.Sort = 2
        Px = False
    Else
        dtgJl.Sort = 1
        Px = True
    End If
'Else
'    MsgBox MGa.ColData(1)
End If
End Sub

Private Sub dtgJl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y > 270 And Y < 525 Then
    CCC = True
Else
    CCC = False
End If
End Sub

Private Sub dtgJl_Scroll()
frmGGLKan.GGGCCC = True
End Sub

Private Sub dtpZ_CloseUp()
txtZ.Text = dtpZ.Value
End Sub

Private Sub Form_Load()
frmGGLKan.Height = 9600
frmGGLKan.Width = 8550
Me.Top = 0
dtgJl.ColWidth(0) = 300
dtgJl.ColWidth(1) = 2500
dtgJl.ColWidth(3) = 2000
dtgJl.ColWidth(4) = 0
comLx.Text = "类别"
comBj.AddItem "="
'txtZ.AddItem "公告类"
'txtZ.AddItem "一般类"
'txtZ.AddItem "通知类"
'txtZ.AddItem "派工类"
'txtZ.AddItem "到帐类"
txtZ.Text = "公告类"
'txtZ.AddItem "其它类"
GGGCCC = False
CCC = False
dtpZ.Value = Date
txtZ.Text = dtpZ.Value
End Sub

