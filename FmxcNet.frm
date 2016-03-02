VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FmxcNet 
   BackColor       =   &H00C0FFFF&
   Caption         =   "网上订单-----------杰升在线订购"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdBack 
      Height          =   435
      Left            =   14640
      Picture         =   "FmxcNet.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   8670
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMn 
      Height          =   795
      Left            =   14250
      TabIndex        =   46
      Top             =   2370
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   1402
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMa 
      Height          =   9105
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   16060
      _Version        =   393216
      BackColor       =   12648384
      BackColorFixed  =   12648384
      BackColorBkg    =   12648447
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgDetail 
      Height          =   3945
      Left            =   7140
      TabIndex        =   41
      Top             =   4680
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   6959
      _Version        =   393216
      BackColor       =   12648384
      BackColorFixed  =   12648384
      BackColorBkg    =   12648447
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblAddTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   8730
      TabIndex        =   45
      Top             =   4320
      Width           =   2235
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "下单日期"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7530
      TabIndex        =   44
      Top             =   4320
      Width           =   1245
   End
   Begin VB.Label lblSum 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   13650
      TabIndex        =   43
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "总金额"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   12720
      TabIndex        =   42
      Top             =   4320
      Width           =   645
   End
   Begin VB.Label lblCompany_taxid 
      BackStyle       =   0  'Transparent
      Caption         =   "Label39"
      Height          =   225
      Left            =   8730
      TabIndex        =   40
      Top             =   3990
      Width           =   5385
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "税号"
      Height          =   225
      Left            =   7530
      TabIndex        =   39
      Top             =   3990
      Width           =   1215
   End
   Begin VB.Label lblBank_username 
      BackStyle       =   0  'Transparent
      Caption         =   "Label37"
      Height          =   225
      Left            =   12660
      TabIndex        =   38
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label lblBank_account 
      BackStyle       =   0  'Transparent
      Caption         =   "Label36"
      Height          =   225
      Left            =   8730
      TabIndex        =   37
      Top             =   3645
      Width           =   2295
   End
   Begin VB.Label lblBank_name 
      BackStyle       =   0  'Transparent
      Caption         =   "Label35"
      Height          =   225
      Left            =   12660
      TabIndex        =   36
      Top             =   3240
      Width           =   1875
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "开户名"
      Height          =   225
      Left            =   11460
      TabIndex        =   35
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "开户帐号"
      Height          =   225
      Left            =   7530
      TabIndex        =   34
      Top             =   3645
      Width           =   1275
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "开户行"
      Height          =   225
      Left            =   11460
      TabIndex        =   33
      Top             =   3300
      Width           =   1155
   End
   Begin VB.Label lblBank 
      BackStyle       =   0  'Transparent
      Caption         =   "Label31"
      Height          =   225
      Left            =   8730
      TabIndex        =   32
      Top             =   3285
      Width           =   2445
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "开户银行"
      Height          =   225
      Left            =   7530
      TabIndex        =   31
      Top             =   3285
      Width           =   975
   End
   Begin VB.Label lblCompany_fax 
      BackStyle       =   0  'Transparent
      Caption         =   "Label29"
      Height          =   225
      Left            =   12660
      TabIndex        =   30
      Top             =   2940
      Width           =   2205
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "传真"
      Height          =   225
      Left            =   11460
      TabIndex        =   29
      Top             =   2940
      Width           =   585
   End
   Begin VB.Label lblCompany_phone 
      BackStyle       =   0  'Transparent
      Caption         =   "Label27"
      Height          =   225
      Left            =   8730
      TabIndex        =   28
      Top             =   2940
      Width           =   2445
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "电话"
      Height          =   225
      Left            =   7530
      TabIndex        =   27
      Top             =   2940
      Width           =   1095
   End
   Begin VB.Label lblCompany_areacode 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      Height          =   225
      Left            =   8730
      TabIndex        =   26
      Top             =   2235
      Width           =   1545
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "邮编"
      Height          =   225
      Left            =   7530
      TabIndex        =   25
      Top             =   2235
      Width           =   1035
   End
   Begin VB.Label lblCompany_addr 
      BackStyle       =   0  'Transparent
      Caption         =   "Label24"
      Height          =   225
      Left            =   8730
      TabIndex        =   24
      Top             =   2595
      Width           =   6435
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "公司地址"
      Height          =   225
      Left            =   7530
      TabIndex        =   23
      Top             =   2595
      Width           =   1125
   End
   Begin VB.Label lblCompany_type 
      BackStyle       =   0  'Transparent
      Caption         =   "Label22"
      Height          =   225
      Left            =   12660
      TabIndex        =   22
      Top             =   2220
      Width           =   1995
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "公司类型"
      Height          =   225
      Left            =   11460
      TabIndex        =   21
      Top             =   2220
      Width           =   915
   End
   Begin VB.Label lblCompany_Name 
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      Height          =   225
      Left            =   8730
      TabIndex        =   20
      Top             =   1890
      Width           =   1785
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "公司名称"
      Height          =   225
      Left            =   7530
      TabIndex        =   19
      Top             =   1890
      Width           =   1065
   End
   Begin VB.Label lblUserType 
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
      Height          =   225
      Left            =   12660
      TabIndex        =   18
      Top             =   510
      Width           =   1485
   End
   Begin VB.Label label17 
      BackStyle       =   0  'Transparent
      Caption         =   "用户类型"
      Height          =   225
      Left            =   11460
      TabIndex        =   17
      Top             =   510
      Width           =   1065
   End
   Begin VB.Label lblWork 
      BackStyle       =   0  'Transparent
      Caption         =   "Label16"
      Height          =   225
      Left            =   12660
      TabIndex        =   16
      Top             =   1185
      Width           =   2355
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "职务"
      Height          =   225
      Left            =   11460
      TabIndex        =   15
      Top             =   1185
      Width           =   1035
   End
   Begin VB.Label lblPerson_Addr 
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      Height          =   225
      Left            =   8730
      TabIndex        =   14
      Top             =   1530
      Width           =   6375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "地址"
      Height          =   225
      Left            =   7530
      TabIndex        =   13
      Top             =   1530
      Width           =   1095
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      Height          =   225
      Left            =   8730
      TabIndex        =   12
      Top             =   1185
      Width           =   2595
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "EMail"
      Height          =   225
      Left            =   7530
      TabIndex        =   11
      Top             =   1185
      Width           =   675
   End
   Begin VB.Label lblMobile 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      Height          =   225
      Left            =   12660
      TabIndex        =   10
      Top             =   855
      Width           =   2355
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "手机"
      Height          =   225
      Left            =   11460
      TabIndex        =   9
      Top             =   855
      Width           =   945
   End
   Begin VB.Label lblTName 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   225
      Left            =   8730
      TabIndex        =   8
      Top             =   855
      Width           =   2565
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      Height          =   225
      Left            =   7530
      TabIndex        =   7
      Top             =   855
      Width           =   555
   End
   Begin VB.Label lblGroup 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      Height          =   225
      Left            =   8730
      TabIndex        =   6
      Top             =   510
      Width           =   2565
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "等级"
      Height          =   225
      Left            =   7530
      TabIndex        =   5
      Top             =   510
      Width           =   495
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   225
      Left            =   12660
      TabIndex        =   4
      Top             =   180
      Width           =   2385
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "客户编码"
      Height          =   225
      Left            =   11460
      TabIndex        =   3
      Top             =   180
      Width           =   1185
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   8730
      TabIndex        =   2
      Top             =   180
      Width           =   2265
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "用户"
      Height          =   225
      Left            =   7530
      TabIndex        =   1
      Top             =   180
      Width           =   615
   End
End
Attribute VB_Name = "FmxcNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tt As String

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
End Sub

Private Sub dtgMa_Click()
Dim tt As String
Dim Uid As String
Dim UserType As String
dtgMn.Row = dtgMa.Row
dtgMn.Col = 2
UserType = dtgMn.Text
dtgMn.Col = 7
Uid = dtgMn.Text
If Val(Uid) = 0 Then Exit Sub
'''''''tt = "declare @groupid int;" & _
'''''''    "select uid,code,groupid,username,tname,userpwd,mobile,email,person_addr,work,user_type,company_name,company_type,company_addr,company_areacode,company_phone,company_fax," & _
'''''''   "bank,bank_name,bank_account,bank_username,company_taxid from tb_member where uid='" & Uid & "';" & _
'''''''   "select @groupid=groupid,@uid=uid from tb_member where uid='" & Uid & "';" & _
'''''''   "select name from tb_usergroup where groupid=@groupid"
tt = "select uid,code,groupid,username,tname,userpwd,mobile,email,person_addr,work,user_type,company_name,company_type,company_addr,company_areacode,company_phone,company_fax," & _
   "bank,bank_name,bank_account,bank_username,company_taxid from tb_member where uid='" & Uid & "'"
Call DetailBound(tt, UserType)
End Sub


Private Sub dtgMa_DblClick()
Dim tt As String
Dim Id As Long
Dim Ra
Dim La As Integer
dtgMn.Row = dtgMa.Row
dtgMn.Col = 5
Id = dtgMn.Text
tt = "select pid,num,sum from tb_order where id=" & Id
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
mod1.HTP.Close
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0

End Sub

Public Sub Bound(tt As String)
Dim Ra
Dim La As Integer
Call Qing
Dim oo As Integer
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workJs, adOpenForwardOnly, adLockReadOnly, adCmdText
'mod1.HTP.Open tt, mod1.workJs, adOpenForwardOnly, adLockBatchOptimistic, adCmdText
'Adodc1.Recordset.Open tt, mod1.workJs, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
For oo = 1 To La
    dtgMa.Row = oo
    dtgMa.Col = 0: dtgMa.Text = Ra(0, oo - 1)
    dtgMa.Col = 1: dtgMa.Text = Ra(1, oo - 1)
    dtgMa.Col = 2: dtgMa.Text = Ra(2, oo - 1)
    dtgMa.Col = 3: dtgMa.Text = Ra(3, oo - 1)
    dtgMa.Col = 4: dtgMa.Text = Ra(4, oo - 1)
    dtgMa.Col = 5: dtgMa.Text = Ra(5, oo - 1)
    dtgMa.Col = 6: dtgMa.Text = Ra(6, oo - 1)
    dtgMa.Col = 7: dtgMa.Text = Ra(7, oo - 1)
    dtgMn.Row = oo
    dtgMn.Col = 0: dtgMn.Text = Ra(0, oo - 1)
    dtgMn.Col = 1: dtgMn.Text = Ra(1, oo - 1)
    dtgMn.Col = 2: dtgMn.Text = Ra(2, oo - 1)
    dtgMn.Col = 3: dtgMn.Text = Ra(3, oo - 1)
    dtgMn.Col = 4: dtgMn.Text = Ra(4, oo - 1)
    dtgMn.Col = 5: dtgMn.Text = Ra(5, oo - 1)
    dtgMn.Col = 6: dtgMn.Text = Ra(6, oo - 1)
    dtgMn.Col = 7: dtgMn.Text = Ra(7, oo - 1)
Next

End Sub

Public Sub Qing()
Call DetailQing
dtgMa.Clear
dtgMa.Cols = 8
dtgMa.Rows = 100
dtgMa.Row = 0
dtgMa.Col = 0: dtgMa.Text = "日期": dtgMa.CellFontBold = True
dtgMa.Col = 1: dtgMa.Text = "客户": dtgMa.CellFontBold = True
dtgMa.Col = 2: dtgMa.Text = "等级": dtgMa.CellFontBold = True
dtgMa.Col = 3: dtgMa.Text = "金额": dtgMa.CellFontBold = True
dtgMa.Col = 4: dtgMa.Text = "状态": dtgMa.CellFontBold = True
dtgMa.Col = 5: dtgMa.Text = "id": dtgMa.CellFontBold = True
dtgMa.Col = 6: dtgMa.Text = "groupid": dtgMa.CellFontBold = True
dtgMa.Col = 7: dtgMa.Text = "uid": dtgMa.CellFontBold = True
dtgMa.ColWidth(0) = 2010
dtgMa.ColWidth(1) = 1715
dtgMa.ColWidth(2) = 1000
dtgMa.ColWidth(4) = 1000
dtgMa.ColWidth(5) = 1000
dtgMa.ColWidth(6) = 1000
dtgMa.ColWidth(7) = 1000
dtgMn.Clear
dtgMn.Cols = 8
dtgMn.Rows = 100

End Sub

Private Sub Label31_Click()

End Sub

Public Sub DetailQing()
Me.lblUserName.ToolTipText = ""
Me.lblCode.Caption = ""
Me.lblGroup.ToolTipText = "": Me.lblGroup.Caption = ""
Me.lblUserName.Caption = ""
Me.lblTName.Caption = ""
Me.lblMobile.Caption = ""
Me.lblEmail.Caption = ""
Me.lblPerson_Addr.Caption = ""
Me.lblWork.Caption = ""
Me.lblUserType.Caption = ""
Me.lblCompany_Name.Caption = ""
Me.lblCompany_type.Caption = ""
Me.lblCompany_addr.Caption = ""
Me.lblCompany_areacode.Caption = ""
Me.lblCompany_phone.Caption = ""
Me.lblCompany_fax.Caption = ""
Me.lblBank.Caption = ""
Me.lblBank_name.Caption = ""
Me.lblBank_account.Caption = ""
Me.lblBank_username.Caption = ""
Me.lblCompany_taxid.Caption = ""
Me.lblSum.Caption = ""
Me.lblAddTime.Caption = ""

dtgDetail.Clear
dtgDetail.Cols = 5
dtgDetail.Rows = 30
dtgDetail.Row = 0
dtgDetail.Col = 0: dtgDetail.Text = "编码": dtgDetail.CellFontBold = True
dtgDetail.Col = 1: dtgDetail.Text = "货品": dtgDetail.CellFontBold = True
dtgDetail.Col = 2: dtgDetail.Text = "单价": dtgDetail.CellFontBold = True
dtgDetail.Col = 3: dtgDetail.Text = "数量": dtgDetail.CellFontBold = True
dtgDetail.Col = 4: dtgDetail.Text = "Pid": dtgDetail.CellFontBold = True
dtgDetail.ColWidth(1) = 4635
dtgDetail.ColWidth(4) = 0
End Sub

Public Sub DetailBound(tt As String, UserType As String)
Dim Ra
Dim Rb
Dim RC
Call DetailQing
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workJs, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
'''''Set mod1.HTP = mod1.HTP.NextRecordset
'''''Rb = mod1.HTP.GetRows
'''''Set mod1.HTP = mod1.HTP.NextRecordset
'''''RC = mod1.HTP.GetRows
''''''mod1.HTP.Close
Set mod1.HTP = Nothing

  
Me.lblUserName.ToolTipText = Ra(0, 0)
Me.lblCode.Caption = Ra(1, 0)
Me.lblGroup.ToolTipText = Ra(2, 0)
Me.lblGroup.Caption = UserType
Me.lblUserName.Caption = Ra(3, 0)
Me.lblTName.Caption = Ra(4, 0)
Me.lblMobile.Caption = Ra(5, 0)
Me.lblEmail.Caption = Ra(6, 0)
Me.lblPerson_Addr.Caption = Ra(7, 0)
Me.lblWork.Caption = Ra(8, 0)
Me.lblUserType.Caption = Ra(9, 0)
Me.lblCompany_Name.Caption = Ra(10, 0)
Me.lblCompany_type.Caption = Ra(11, 0)
Me.lblCompany_addr.Caption = Ra(12, 0)
Me.lblCompany_areacode.Caption = Ra(13, 0)
Me.lblCompany_phone.Caption = Ra(14, 0)
Me.lblCompany_fax.Caption = Ra(15, 0)
Me.lblBank.Caption = Ra(16, 0)
Me.lblBank_name.Caption = Ra(17, 0)
Me.lblBank_account.Caption = Ra(18, 0)
Me.lblBank_username.Caption = Ra(19, 0)
Me.lblCompany_taxid.Caption = Ra(20, 0)
Me.lblSum.Caption = ""
Me.lblAddTime.Caption = ""

dtgDetail.Clear
dtgDetail.Cols = 5
dtgDetail.Rows = 30
dtgDetail.Row = 0
dtgDetail.Col = 0: dtgDetail.Text = "编码": dtgDetail.CellFontBold = True
dtgDetail.Col = 1: dtgDetail.Text = "货品": dtgDetail.CellFontBold = True
dtgDetail.Col = 2: dtgDetail.Text = "单价": dtgDetail.CellFontBold = True
dtgDetail.Col = 3: dtgDetail.Text = "数量": dtgDetail.CellFontBold = True
dtgDetail.Col = 4: dtgDetail.Text = "Pid": dtgDetail.CellFontBold = True
dtgDetail.ColWidth(1) = 4635
dtgDetail.ColWidth(4) = 0
End Sub
