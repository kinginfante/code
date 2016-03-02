VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmMeetDetail 
   BackColor       =   &H00C0FFC0&
   Caption         =   "会议详情"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15210
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3825
      Left            =   6210
      TabIndex        =   12
      Top             =   3450
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印"
      Height          =   765
      Left            =   11820
      Picture         =   "frmMeetDetail.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8250
      Width           =   675
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H00C0FFC0&
      Caption         =   "修改"
      Height          =   765
      Left            =   12510
      Picture         =   "frmMeetDetail.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "修改"
      Top             =   8250
      Width           =   675
   End
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "审核"
      Height          =   765
      Left            =   10410
      Picture         =   "frmMeetDetail.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8250
      Width           =   675
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "提交"
      Height          =   765
      Left            =   13200
      Picture         =   "frmMeetDetail.frx":0DB6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8250
      Width           =   675
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "返回"
      Height          =   765
      Left            =   14550
      Picture         =   "frmMeetDetail.frx":1420
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8250
      Width           =   585
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "作废"
      Enabled         =   0   'False
      Height          =   765
      Left            =   13890
      Picture         =   "frmMeetDetail.frx":1522
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8250
      Width           =   645
   End
   Begin VB.CommandButton cmdBr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "浏览"
      Height          =   765
      Left            =   11100
      Picture         =   "frmMeetDetail.frx":16AC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8250
      Width           =   705
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   4965
      Left            =   0
      TabIndex        =   0
      Top             =   4110
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   8758
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
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMeet 
      Height          =   7305
      Left            =   6180
      TabIndex        =   2
      Top             =   0
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   12885
      _Version        =   393216
      BackColor       =   16777152
      Rows            =   30
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   16777152
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.OLE OLE1 
      Height          =   495
      Left            =   7530
      TabIndex        =   11
      Top             =   7980
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblMc 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   780
      TabIndex        =   9
      Top             =   30
      Width           =   4095
   End
   Begin VB.Label lblNR 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   90
      TabIndex        =   1
      Top             =   570
      Width           =   5925
   End
End
Attribute VB_Name = "frmMeetDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bh As Long
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
lblMc.Caption = ""
lblNR.Caption = ""
Call dtgMeetFF
Call dtgPFF
End Sub

Public Sub Bound(Mid As Long)
Dim tt As String
Dim Ra, Rb
Dim Lb As Integer
Call Qing
tt = "select mc,lx,rq,bm,mtime,mtzu,mtji,bz,cren from meet where mid=" & Mid & _
    ";select xz+': '+nr,wcf,wt,fa,wrq,zren,did from meetdetail where mid=" & Mid & " order by zren"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
lblMc.Caption = Ra(0, 0) & " 会议记录"
lblNR.Caption = "会议性质: " & Ra(1, 0) & "     部门:" & Ra(3, 0) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                "日期:" & Ra(2, 0) & "     时间:" & Ra(4, 0) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                "参会者:" & Ra(8, 0) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                "会议主持:" & Ra(5, 0) & "     会议记录:" & Ra(6, 0)
Call Me.MeetBound(Rb)
Bh = Mid
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
End Sub

Private Sub cmdBr_Click()
Me.Visible = False
frmMeet.Show
frmMeet.ZOrder 0
End Sub


Private Sub cmdPrint_Click()
Dim tt As String
Dim Ra
tt = "declare @lid int;" & _
    "insert into openfile (lx,bh,ren,uid,rq) values ('会议记录'," & Bh & ",'" & mod1.DName & "','" & mod1.DHid & "',getdate());" & _
    "select @@identity"

Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.wzcc, adOpenForwardOnly, adLockReadOnly, adCmdText
Set mod1.HTP = mod1.HTP.NextRecordset
Ra = mod1.HTP.GetRows
Set mod1.HTP = Nothing
Open "c:\work\tempId.txt" For Output As #5
Write #5, Ra(0, 0)
Close #5


On Error Resume Next
    OLE1.SourceDoc = "c:\work\会议记录.xls"
    OLE1.Action = 1
    OLE1.DoVerb (-2)
End Sub


Private Sub Form_Load()
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
Call Me.dtgMeetFF
End Sub

Public Sub dtgMeetFF()
dtgMeet.Clear
dtgMeet.Cols = 7
dtgMeet.Rows = 30
dtgMeet.Row = 0
dtgMeet.Col = 0: dtgMeet.Text = "内容": dtgMeet.CellFontBold = True
dtgMeet.Col = 1: dtgMeet.Text = "完成情况": dtgMeet.CellFontBold = True
dtgMeet.Col = 2: dtgMeet.Text = "存在问题": dtgMeet.CellFontBold = True
dtgMeet.Col = 3: dtgMeet.Text = "解决方案": dtgMeet.CellFontBold = True
dtgMeet.Col = 4: dtgMeet.Text = "完成时间": dtgMeet.CellFontBold = True
dtgMeet.Col = 5: dtgMeet.Text = "责任人": dtgMeet.CellFontBold = True
dtgMeet.ColWidth(0) = 2610
dtgMeet.ColWidth(1) = 1065
dtgMeet.ColWidth(2) = 1380
dtgMeet.ColWidth(4) = 1005
dtgMeet.ColWidth(6) = 0
End Sub

Public Sub MeetBound(Rb)
Dim Lb As Integer
Dim oo As Integer
Call Me.dtgMeetFF
Lb = UBound(Rb, 2) + 1
On Error Resume Next
For oo = 1 To Lb
    dtgMeet.Row = oo
    dtgMeet.Col = 0: dtgMeet.Text = Rb(0, oo - 1)
    dtgMeet.Col = 1: dtgMeet.Text = Rb(1, oo - 1)
    dtgMeet.Col = 2: dtgMeet.Text = Rb(2, oo - 1)
    dtgMeet.Col = 3: dtgMeet.Text = Rb(3, oo - 1)

    dtgMeet.Col = 4: dtgMeet.Text = Rb(4, oo - 1)
    If dtgMeet.Text = "1900/1/1" Then dtgMeet.Text = ""
    dtgMeet.Col = 5: dtgMeet.Text = Rb(5, oo - 1)
Next

End Sub
