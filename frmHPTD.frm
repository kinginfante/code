VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmHPTD 
   BackColor       =   &H00C0FFC0&
   Caption         =   "��Ʒ���ϲ�ѯ"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15210
   Begin VB.Frame frmZB 
      BackColor       =   &H00FFFFC0&
      Caption         =   "����ܱ�"
      Height          =   5295
      Left            =   8610
      TabIndex        =   34
      Top             =   1260
      Visible         =   0   'False
      Width           =   15255
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3405
         Left            =   0
         TabIndex        =   35
         Top             =   240
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   6006
         _Version        =   393216
         BackColor       =   16777152
         FixedCols       =   0
         BackColorFixed  =   16777152
         BackColorBkg    =   16777152
         WordWrap        =   -1  'True
         TextStyleFixed  =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdLP 
      Caption         =   "Command1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   5280
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   405
      Left            =   8100
      TabIndex        =   28
      Top             =   5640
      Width           =   5115
      Begin VB.CommandButton cmdC 
         Caption         =   "��ѯ"
         Height          =   285
         Left            =   3240
         TabIndex        =   31
         Top             =   30
         Width           =   945
      End
      Begin VB.ComboBox comLx 
         Height          =   300
         ItemData        =   "frmHPTD.frx":0000
         Left            =   810
         List            =   "frmHPTD.frx":0022
         TabIndex        =   30
         Text            =   "��������"
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox txtZ 
         Height          =   285
         Left            =   1920
         TabIndex        =   29
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ��ʽ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   30
         Width           =   735
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgXN 
      Height          =   1005
      Left            =   10860
      TabIndex        =   27
      Top             =   6840
      Visible         =   0   'False
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   1773
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   885
      Left            =   8670
      TabIndex        =   16
      Top             =   6870
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1561
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer timQuit 
      Interval        =   1000
      Left            =   6780
      Top             =   480
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�޸�"
      Height          =   765
      Left            =   13830
      Picture         =   "frmHPTD.frx":0088
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "�޸�"
      Top             =   8280
      Width           =   675
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "��   ��"
      Height          =   1185
      Left            =   2850
      TabIndex        =   14
      Top             =   2430
      Width           =   345
   End
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "���"
      Height          =   765
      Left            =   13080
      Picture         =   "frmHPTD.frx":0392
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8280
      Width           =   675
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����"
      Height          =   765
      Left            =   14580
      Picture         =   "frmHPTD.frx":07D4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8280
      Width           =   585
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "������"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   30
      TabIndex        =   6
      Top             =   7230
      Visible         =   0   'False
      Width           =   6315
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "����"
         Height          =   285
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton OptT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�ܾ�"
         Height          =   195
         Left            =   5220
         TabIndex        =   9
         Top             =   870
         Width           =   675
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ͬ��"
         Height          =   225
         Left            =   5220
         TabIndex        =   8
         Top             =   480
         Width           =   705
      End
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1365
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   300
         Width           =   4965
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLP 
      Height          =   3405
      Left            =   60
      TabIndex        =   4
      Top             =   1740
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   6006
      _Version        =   393216
      BackColor       =   16777152
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      TextStyleFixed  =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgX 
      Height          =   3405
      Left            =   3720
      TabIndex        =   5
      Top             =   1740
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   6006
      _Version        =   393216
      BackColor       =   16777152
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      TextStyleFixed  =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   3405
      Left            =   60
      TabIndex        =   11
      Top             =   5580
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   6006
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      Height          =   285
      Left            =   3720
      TabIndex        =   26
      Top             =   1380
      Width           =   1305
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "�����ϵ��(˫�����������ʾ��һ�������"
      Height          =   285
      Left            =   180
      TabIndex        =   25
      Top             =   1350
      Width           =   3555
   End
   Begin VB.Label lblJz 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      Height          =   375
      Left            =   8640
      TabIndex        =   24
      Top             =   780
      Width           =   5745
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "���û���"
      Height          =   345
      Left            =   7410
      TabIndex        =   23
      Top             =   780
      Width           =   1215
   End
   Begin VB.Label lblXN 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   345
      Left            =   4650
      TabIndex        =   22
      Top             =   780
      Width           =   2355
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "��Ʒ�ͺ�"
      Height          =   345
      Left            =   3630
      TabIndex        =   21
      Top             =   780
      Width           =   1005
   End
   Begin VB.Label lblOname 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   345
      Left            =   900
      TabIndex        =   20
      Top             =   780
      Width           =   1965
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ԭ�����"
      Height          =   285
      Left            =   60
      TabIndex        =   19
      Top             =   780
      Width           =   795
   End
   Begin VB.Label lblYpb 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   315
      Left            =   8640
      TabIndex        =   18
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ԭ��Ʒ��"
      Height          =   315
      Left            =   7440
      TabIndex        =   17
      Top             =   120
      Width           =   945
   End
   Begin VB.Label lblPartName 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label5"
      Height          =   255
      Left            =   4650
      TabIndex        =   3
      Top             =   120
      Width           =   2385
   End
   Begin VB.Label lblbh 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label4"
      Height          =   315
      Left            =   900
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "��Ʒ����"
      Height          =   285
      Left            =   3630
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "���"
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "frmHPTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timZm As Integer
Dim Lbh As String '�༭ʱ����ʱ���
Dim Tid As String

Dim GXC As Integer '��ϵ������Ĭ��Ϊ0
Public Sub dtgPFF()
Dim oo As Integer
For oo = 1 To dtgP.Rows - 1
    dtgP.RowHeight(oo) = dtgP.RowHeight(0) * 2
Next
dtgP.Clear
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "����": dtgP.Col = 1: dtgP.Text = "����": dtgP.Col = 2: dtgP.Text = "ְ��": dtgP.Col = 3: dtgP.Text = "������": dtgP.Col = 4: dtgP.Text = "���":
dtgP.ColWidth(0) = 1665
dtgP.ColWidth(1) = 1005
dtgP.ColWidth(2) = 0
 dtgP.ColWidth(3) = 3570: dtgP.ColWidth(4) = 1035
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub
Private Sub cmdBack_Click()
Me.Visible = False
End Sub

Private Sub cmdC_Click()
Dim tt As String
Dim LT1 As String
Dim LT2 As String
Dim LT3 As String
Dim JT As String
Dim DelF As Integer
DelF = 1
'''''If chkDel.Value = 1 Then
'''''    DelF = 0
'''''End If
JT = ",oname,gg,xn,pb,jz,ypb,bm1,bm2,bm3,l1,l2,l3,jyf,bz"
Select Case comLx.Text
Case "��������"
    tt = "select bh,partname,pid" & JT & " from nlpmxc where (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bh='" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "' or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " bz like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%')  and delf=" & DelF & " "
Case "���"
    If Len(Replace(txtZ.Text, vbCrLf, "", 1)) = 1 And Val(Replace(txtZ.Text, vbCrLf, "", 1)) > 0 Then
        tt = "select bh,partname,pid" & JT & " from nlpmxc where left(bh,1)='" & Replace(txtZ.Text, vbCrLf, "", 1) & "' and delf=" & DelF & " "
    ElseIf Len(Replace(txtZ.Text, vbCrLf, "", 1)) = 2 And Val(Replace(txtZ.Text, vbCrLf, "", 1)) > 0 Then
        tt = "select bh,partname,pid" & JT & " from nlpmxc where left(bh,2)='" & Replace(txtZ.Text, vbCrLf, "", 1) & "' and delf=" & DelF & " "
    ElseIf Len(Replace(txtZ.Text, vbCrLf, "", 1)) = 3 And Val(Replace(txtZ.Text, vbCrLf, "", 1)) > 0 Then
        tt = "select bh,partname,pid" & JT & " from nlpmxc where left(bh,3)='" & Replace(txtZ.Text, vbCrLf, "", 1) & "' and delf=" & DelF & " "
    ElseIf Len(Replace(txtZ.Text, vbCrLf, "", 1)) = 5 And Val(Replace(txtZ.Text, vbCrLf, "", 1)) > 0 Then
        tt = "select bh,partname,pid" & JT & ",delf from nlpmxc where bh='" & Replace(txtZ.Text, vbCrLf, "", 1) & "'"
    End If
Case "���"
    tt = "select bh,partname,pid" & JT & " from nlpmxc where (l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%')  and delf=" & DelF
Case "����"
    tt = "select bh,partname,pid" & JT & " from nlpmxc where (bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%')  and delf=" & DelF
Case "ԭ�����"
    tt = "select bh,partname,pid" & JT & " from nlpmxc where oname='" & Replace(txtZ.Text, vbCrLf, "", 1) & "' and delf=" & DelF
Case "����Ʒ��"
    tt = "select bh,partname,pid" & JT & " from nlpmxc where pb='" & Replace(txtZ.Text, vbCrLf, "", 1) & "' and delf=" & DelF & " "
Case "���û���"
    tt = "select bh,partname,pid" & JT & " from nlpmxc where jz like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' and delf=" & DelF & " "
Case "����"
    tt = "select bh,partname,'ԭ�����:'+oname+' '+gg+' '+xn+' ',pid from nlpmxc where (lb1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or lb2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%')  and delf=" & DelF & " "
Case "������Ʒ"
    tt = "select bh,partname,pid" & JT & " from nlpmxc where left(bh,1)='H' and (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%'  or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%') and delf=" & DelF & " "
Case "���������"
    tt = "select bh,partname,pid" & JT & " from nlpmxc where left(bh,1)='B' and (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%'  or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%') and delf=" & DelF & "  "
Case "�º����̼��׺�"
    tt = "select bh,partname,pid" & JT & " from nlpmxc where left(bh,1)='A' and (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%'  or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%') and delf=" & DelF & "  "
Case "ԭ�����"
    tt = "select bh,partname,pid" & JT & " from nlpmxc where left(bh,1)='9' and (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%'  or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%') and delf=" & DelF & "  "
Case "��Ʒ��"
    tt = "select bh,partname,pid" & JT & " from nlpmxc where left(bh,1)='8' and (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%'  or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%') and delf=" & DelF & "  "
Case "��ʱ��"
    tt = "select bh,partname,pid" & JT & " from nlpmxc where left(bh,1)='3' and (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%'  or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%') "
Case "�����"
    tt = "select bh,partname,pid" & JT & " from nlpmxc where (left(bh,1)='1'" & _
        "or left(bh,1)='2' or left(bh,1)='4' or left(bh,1)='5' or left(bh,1)='6' or left(bh,1)='7') and (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%'  or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%') and delf=" & DelF & "  "
End Select
If tt = "" Then Exit Sub
If mod1.Bm <> "��������" And mod1.Bm <> "ά������" Then
    If comLx.Text <> "���" And Left(txtZ.Text, 1) <> "3" Then
        tt = tt & " and lc=100"
    End If
End If
tt = tt & " order by bh"
Call dtgXBound(tt)

End Sub

Private Sub cmdEdit_Click()
Dim tt As String
Dim oo As Integer
Dim Rb, RC
Dim Lb As Integer
Dim Pb As String
Dim JZ As String
If cmdEdit.Caption <> "���" And cmdEdit.Caption <> "ɾ��" Then Exit Sub
If cmdEdit.Caption = "���" And Lbh = "" Or cmdEdit.Caption = "ɾ��" And Tid = 0 Then Exit Sub

'

    timZm = 1 '�༭

     '���¼������Ʒ������Ʒ�������û���
     If cmdEdit.Caption = "ɾ��" Then
        tt = "select pb,jz from nlpmxctd where ybh='" & Lbh & "' and bh<>'" & lblbh.Caption & "' order by pb"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        On Error Resume Next
        Rb = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        Lb = UBound(Rb, 2) + 1
        For oo = 0 To Lb - 1
            If Not (InStr(1, Pb, Rb(0, oo)) > 0) Then
            
                Pb = Pb & Rb(0, oo) & " "
            End If
            JZ = JZ & "(" & Rb(0, oo) & ")" & Rb(1, oo) & " "

        Next
     ElseIf cmdEdit.Caption = "���" Then
        '������ǰ�������ϵ���γɵ�Ʒ�������
        tt = "select pb,jz from nlpmxctd where ybh='" & Lbh & "' order by pb;" & _
            "select ypb,jz from nlpmxc where bh='" & Lbh & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        On Error Resume Next
        Rb = mod1.HTP.GetRows
        Set mod1.HTP = mod1.HTP.NextRecordset
        RC = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        Lb = UBound(Rb, 2) + 1
        Pb = RC(0, 0)
        JZ = RC(1, 0)
        For oo = 0 To Lb - 1
            If Not (InStr(1, Pb, Rb(0, oo)) > 0) Then
            
                Pb = Pb & Rb(0, oo) & " "
            End If
            If Not (InStr(1, JZ, Rb(1, oo)) > 0) Then
            JZ = JZ & "(" & Rb(0, oo) & ")" & Rb(1, oo) & " "
            End If
        Next
        Pb = Trim(Pb):
        JZ = Trim(JZ)
        '������ڵ�Ʒ�ƺͻ���
        If Not (InStr(1, Pb, frmHPZL.txtPb.Text)) > 0 Then
            Pb = Pb & " " & frmHPZL.txtPb.Text
        End If
        If Not (InStr(1, JZ, frmHPZL.txtJz.Text) > 0) Then
            JZ = JZ & " " & frmHPZL.txtJz.Text
        End If
   End If
        Set mod1.HTP = Nothing
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "�»�Ʒ���"
        mod1.cmd.Parameters("@NBLX") = "�༭"
        mod1.cmd.Parameters("@bh") = lblbh.Caption
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = cmdEdit.Caption
        mod1.cmd.Parameters("@mt2") = Tid
        mod1.cmd.Parameters("@mt3") = Lbh
        mod1.cmd.Parameters("@mt18") = Pb
        mod1.cmd.Parameters("@mt19") = JZ
        mod1.cmd.Parameters("@mlt1") = ""
        mod1.cmd.Parameters("@mm1") = 0
        mod1.cmd.Parameters("@mb1") = 0
        mod1.cmd.Parameters("@md1") = Null
        Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
        mod1.cmd.Execute
        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
            MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
                cmdSave.Enabled = False
            Exit Sub
        Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
            Me.Enabled = False
            frmWaitA.Visible = True
            frmWaitA.Timer2.Enabled = False
    
            frmWaitA.ZOrder 0
            frmWaitA.Timer2.Enabled = True
            timWait.Enabled = True
            'cmdSave.Enabled = False
        End If
    Set mod1.cmd = Nothing

End Sub

Private Sub cmdLP_Click(Index As Integer)
Dim oo As Integer
On Error Resume Next
If Index < GXC Then
      Call Me.dtpLPBound(cmdLP(Index).Caption)
      
      For oo = 30 To Index + 1 Step -1
        Unload cmdLP(oo)
      Next
      GXC = Index
End If

End Sub

Private Sub cmdMod_Click()

cmdEdit.Visible = True
End Sub

Private Sub dtgLP_Click()
cmdEdit.Caption = "ɾ��"
dtgN.Row = dtgLP.Row
dtgN.Col = 2
Tid = dtgN.Text
dtgN.Col = 0
Lbh = dtgN.Text

End Sub

Private Sub dtgXZ_Click()
cmdEdit.Caption = "���"
End Sub

Private Sub dtgLP_DblClick()
Dim tt As String
Dim JT As String

    JT = ",oname,gg,xn,pb,jz,ypb,bm1,bm2,bm3,l1,l2,l3,jyf,bz"
    tt = "select bh,partname,pid" & JT & " from nlpmxc where bh='" & Lbh & "'"
  Call Me.dtgXBound(tt)
  
  Call Me.dtpLPBound(Lbh)
  
  
End Sub

Private Sub dtgX_Click()
cmdEdit.Caption = "���"
dtgXN.Row = dtgX.Row
dtgXN.Col = 2
Tid = 0
dtgXN.Col = 0
Lbh = dtgXN.Text
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
frmZB.Left = 0
frmZB.Top = 0
End Sub

Public Sub dtgXBound(tt As String)
Dim Ra
Dim La
Dim LNR As String
Dim zz As Integer
Dim oo As Long
dtgX.Visible = False
Call dtgXFF

Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgX.Rows = La + 30
dtgXN.Rows = La + 30
    If comLx.Text = "���" And mod1.Bm <> "��������" Then
        If Ra(15, 0) = False Then
        MsgBox "�޴˻�Ʒ��Ϣ���밴�����������в�ѯ����ѯ������!"
        Exit Sub
        End If
    End If
For oo = 1 To La
    LNR = ""
    dtgX.Row = oo
    dtgX.Col = 0: dtgX.Text = Ra(0, oo - 1)
    dtgX.Col = 1: dtgX.Text = Ra(1, oo - 1)
    dtgX.Col = 2: dtgX.Text = Ra(2, oo - 1)
    If Ra(3, oo - 1) <> "" Then
        LNR = "ԭ�����: " & Ra(3, oo - 1) & " " & Chr(13) & Chr(10)
    End If
    If Ra(4, oo - 1) <> "" Then
        LNR = LNR + "��װ���: " & Ra(4, oo - 1) & " " & Chr(13) & Chr(10)
    End If
    If Ra(5, oo - 1) <> "" Then
        LNR = LNR + "��Ʒ�ͺ�: " & Ra(5, oo - 1) & " " & Chr(13) & Chr(10)
    End If
    If Ra(6, oo - 1) <> "" Then
        LNR = LNR + "����Ʒ��: " & Ra(6, oo - 1) & " " & Chr(13) & Chr(10)
    End If
    If Ra(7, oo - 1) <> "" Then
        LNR = LNR + "���û���: " & Ra(7, oo - 1) & " " & Chr(13) & Chr(10)
    End If
    If Ra(8, oo - 1) <> "" Then
        LNR = LNR + "ԭ��Ʒ��: " & Ra(8, oo - 1) & " " & Chr(13) & Chr(10)
    End If
'''''    If Ra(9, oo - 1) <> "" Then
'''''        LNR = LNR + "����1:" & Ra(9, oo - 1) & " "
'''''    End If
'''''    If Ra(10, oo - 1) <> "" Then
'''''        LNR = LNR + "����2:" & Ra(10, oo - 1) & " "
'''''    End If
'''''    If Ra(11, oo - 1) <> "" Then
'''''        LNR = LNR + "����3:" & Ra(11, oo - 1) & " "
'''''    End If
'''''    If Ra(12, oo - 1) <> "" Then
'''''        LNR = LNR + "���1:" & Ra(12, oo - 1) & " "
'''''    End If
'''''    If Ra(13, oo - 1) <> "" Then
'''''        LNR = LNR + "���2:" & Ra(13, oo - 1) & " "
'''''    End If
'''''    If Ra(14, oo - 1) <> "" Then
'''''        LNR = LNR + "���3:" & Ra(14, oo - 1) & " "
'''''    End If
    If Ra(16, oo - 1) <> "" Then           '��ע
        LNR = LNR + "��ע: " & Ra(16, oo - 1) & " " & Chr(13) & Chr(10)
    End If
    frmZu.lblDtg.Caption = LNR
    dtgX.RowHeight(oo) = frmZu.lblDtg.Height
    dtgX.Col = 3: dtgX.Text = LNR
    dtgXN.Row = oo
    dtgXN.Col = 0: dtgXN.Text = Ra(0, oo - 1)
    dtgXN.Col = 1: dtgXN.Text = Ra(1, oo - 1)
    dtgXN.Col = 2: dtgXN.Text = Ra(2, oo - 1)
    dtgXN.Col = 3: dtgXN.Text = LNR
    If oo = La Then
        Jpid = Ra(2, oo - 1)
    End If
    If Jpid < 10 Then
        Jpid = 0
    End If
    '������ʾ��ɫ
    If Ra(15, oo - 1) = True Then
        For zz = 0 To 4
            dtgX.Col = zz: dtgX.CellForeColor = &H80000012
        Next
    Else
        For zz = 0 To 4
            dtgX.Col = zz: dtgX.CellForeColor = &HFF&
        Next
    End If
Next
dtgX.Visible = True
End Sub
Public Sub dtgXFF()
Dim oo As Long
dtgX.Clear
dtgX.Rows = 300
dtgX.Cols = 5
dtgX.Row = 0
dtgX.Col = 0: dtgX.Text = "���": dtgX.CellFontBold = True
dtgX.Col = 1: dtgX.Text = "��Ʒ����": dtgX.CellFontBold = True
dtgX.Col = 3: dtgX.Text = "����": dtgX.CellFontBold = True
dtgX.Col = 2: dtgX.Text = Pid: dtgX.CellFontBold = True
dtgX.Col = 4: dtgX.Text = "���÷�"
dtgXN.Clear
dtgXN.Rows = 300
dtgXN.Cols = 5
dtgX.ColWidth(3) = 7815
dtgX.ColWidth(1) = 1860
dtgX.ColWidth(2) = 0
dtgX.ColWidth(5) = 0
For oo = 1 To 299
    dtgX.RowHeight(oo) = dtgX.RowHeight(0) * 2
Next
End Sub
Public Sub Bound()
Dim Ra
Dim La As Integer
Dim oo As Integer
Dim tt As String
Call Qing

Call Me.dtgPFF
Call Me.dtgXFF

lblbh.Caption = frmHPZL.txtBh.Text
lblPartName.Caption = frmHPZL.txtPartName.Text
Me.lblJz.Caption = frmHPZL.txtJz.Text
Me.lblOname.Caption = frmHPZL.txtOname.Text
Me.lblXN.Caption = frmHPZL.txtXN.Text
Me.lblYpb.Caption = frmHPZL.txtYpb.Text
Me.cmdLP(0).Caption = lblbh.Caption
Call dtpLPBound(lblbh.Caption)

End Sub

Public Sub dtgLPFF()
Dim oo As Long
dtgLP.Clear
dtgLP.Rows = 300
dtgLP.Cols = 4
dtgLP.Row = 0
dtgLP.Col = 0: dtgLP.Text = "���": dtgLP.CellFontBold = True
dtgLP.Col = 1: dtgLP.Text = "ԭ�����": dtgLP.CellFontBold = True

dtgN.Clear
dtgN.Rows = 300
dtgN.Cols = 4
dtgLP.ColWidth(0) = 2160
dtgLP.ColWidth(1) = 0
dtgLP.ColWidth(2) = 0
dtgLP.ColWidth(3) = 0
End Sub

Public Sub Qing()
Dim oo As Integer
On Error Resume Next
For oo = 30 To 1 Step -1
    Unload cmdLP(oo)
Next
GXC = 0
lblbh.Caption = ""
lblPartName.Caption = ""
Me.lblJz.Caption = ""
Me.lblOname.Caption = ""
Me.lblXN.Caption = ""
Me.lblYpb.Caption = ""

dtgLP.Clear
dtgP.Clear
End Sub

Private Sub timQuit_Timer()
Dim tt As String
Dim Rb, RC, RD, RE
Dim Lb As Integer
On Error Resume Next
Dim oo As Integer
Dim jj As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then
'''''    tt = "select ljbh,detail,mj,dj,jdj,sl,jhg,drq,zbq,delf,lid,ljmc,gyid1,gyid2,gyid3,gdj1,gdj2,gdj3,mc1,mc2,mc3,gyid  from XJDetail where bid=" & Val(FmxcXJ.lblBid.ToolTipText) & " order by delf desc,lid desc"
'''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''    Rb = mod1.HTP.GetRows
'''''    mod1.HTP.Close
'''''    Set mod1.HTP = Nothing
'''''    Lb = UBound(Rb, 2)
'''''    Call FmxcXJ.dtgBrBound(Rb, Lb)

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
If mod1.WP.Fields("cf").Value = 1 Then '�ύ�ɹ�
    mod1.Ti = 5
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    timWait.Enabled = False
    If timZm = 1 Then
''''''''        Pid = mod1.WP.Fields("mm1").Value
''''''''        If LCUid = "" Then
''''''''            LCUid = mod1.DHid
''''''''            LCRen = mod1.DName
''''''''        End If
''''''''
        Call dtpLPBound(lblbh.Caption)

    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then 'δ���

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '����ʧ��
    ii = MsgBox("���������ڴ�����������ʱ,�������´���:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "��������!")
    Unload frmWaitA
    Me.Enabled = True
'''''    If timZm = 1 Then
'''''        NiceButton1.Enabled = False
'''''    End If
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("���������ڴ�����������ʱ,��ʱ!", vbExclamation + vbOKOnly, "��������!")
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

Public Sub dtpLPBound(Bh As String)
Dim tt As String
Dim Ra
Dim La As Integer

tt = "select tdbh,oname,tid,zt from nlptdOname where bh='" & Bh & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1

If La = 0 Then Exit Sub
Call Me.dtgLPFF
dtgLP.Rows = La + 20
dtgN.Rows = dtgLP.Rows
For oo = 1 To La
    dtgLP.Row = oo
    dtgLP.Col = 0: dtgLP.Text = Ra(0, oo - 1)
    If Ra(3, oo - 1) = 2 Then 'ɾ��ȷ��
        dtgLP.CellForeColor = &HFF&: dtgLP.Col = 3: dtgLP.Text = "ɾ��ȷ��"
    ElseIf Ra(3, oo - 1) = 3 Then '���ȷ��
        dtgLP.CellForeColor = &HC00000: dtgLP.Col = 3: dtgLP.Text = "���ȷ��"
    Else
        dtgLP.CellForeColor = &H0&
    End If
    dtgLP.Col = 1: dtgLP.Text = Ra(1, oo - 1)
    dtgLP.Col = 2: dtgLP.Text = Ra(2, oo - 1)
    dtgN.Row = oo
    dtgN.Col = 0: dtgN.Text = Ra(0, oo - 1)
    dtgN.Col = 1: dtgN.Text = Ra(1, oo - 1)
    dtgN.Col = 2: dtgN.Text = Ra(2, oo - 1)
Next

If Bh <> Me.lblbh.Caption Then '�����˫���������ɹ�ϵ��ť
    GXC = GXC + 1
    Load cmdLP(GXC)
    cmdLP(GXC).Left = cmdLP(GXC - 1).Left + 100 + cmdLP(GXC - 1).Width
    cmdLP(GXC).Caption = Bh
    cmdLP(GXC).Visible = True
End If

End Sub
