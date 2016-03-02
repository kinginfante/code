VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FmxcBJ 
   BackColor       =   &H00FFFFC0&
   Caption         =   "询价单"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11505
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   11505
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   8916
      _Version        =   393216
      BackColor       =   16777152
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
Attribute VB_Name = "FmxcBJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Width = mod1.FWidth + 500
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
dtgBr.Left = 0
dtgBr.Top = 0
dtgBr.Width = Me.Width
dtgBr.Height = Me.Height - 1000
dtgBr.Rows = 200
End Sub

Public Sub dtgbrFF()
dtgBr.Clear
dtgBr.Cols = 6
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "豪曼编号": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "名称": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "数量": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "单价": dtgBr.CellFontBold = True
dtgBr.Col = 4: dtgBr.Text = "小计": dtgBr.CellFontBold = True
dtgBr.Col = 5: dtgBr.Text = "优惠价": dtgBr.CellFontBold = True
'dtgBr.Col = 5: dtgBr.Text = "询价单编号": dtgBr.CellFontBold = True
dtgBr.ColWidth(1) = 10000
dtgBr.ColWidth(5) = 1100
dtgBr.ColWidth(4) = -1

End Sub

Public Sub Bound(Hid As Long)
Dim tt As String
Dim oo As Long
Dim Ra
Dim La As Long
Me.Caption = "报价清单 合同编号：" & FmxcNew.txtHtbh.Text
tt = "select ljbh,ljmc,sl,sddj,sdxg,zbq,detail,sdyh  from XJbao where htbh='" & Trim(Str(Hid)) & "' and delf=1 order by ywlx,lid"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgBr.Rows = La + 100
For oo = 1 To La
    dtgBr.Row = oo
    If Ra(0, oo - 1) = "" Then
        dtgBr.Col = 1: dtgBr.Text = Ra(6, oo - 1)
        dtgBr.Col = 2: dtgBr.Text = Ra(2, oo - 1)
        dtgBr.Col = 3: dtgBr.Text = Ra(3, oo - 1)
        dtgBr.Col = 4: dtgBr.Text = Ra(4, oo - 1)
        'dtgBr.Col = 5: dtgBr.Text = "XJD" & Ra(7, oo - 1)
        dtgBr.Col = 5: dtgBr.Text = Ra(7, oo - 1)
    Else
        dtgBr.Col = 0: dtgBr.Text = Ra(0, oo - 1)
        dtgBr.Col = 1: dtgBr.Text = Ra(1, oo - 1)
        dtgBr.Col = 2: dtgBr.Text = Ra(2, oo - 1)
        dtgBr.Col = 3: dtgBr.Text = Ra(3, oo - 1)
        dtgBr.Col = 4: dtgBr.Text = Ra(4, oo - 1)
        'dtgBr.Col = 5: dtgBr.Text = "XJD" & Ra(7, oo - 1)
        dtgBr.Col = 5: dtgBr.Text = Ra(7, oo - 1)
    End If
Next
End Sub
