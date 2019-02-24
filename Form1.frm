VERSION 5.00
Begin VB.Form frm 
   AutoRedraw      =   -1  'True
   Caption         =   "SPR-BBCodeTools"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   8415
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fm1 
      Caption         =   "咕咕咕"
      Height          =   1215
      Left            =   5880
      TabIndex        =   30
      Top             =   600
      Width           =   2415
      Begin VB.Label Label1 
         Caption         =   "这里我打算弄点小尾巴设置，这就是为什么下面要Copy"
         Height          =   855
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton cColors 
      Caption         =   ">"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "前景色"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cColor 
      Caption         =   "色"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "前景色/背景色"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cBCs 
      Caption         =   "<"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "背景色"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cList 
      Caption         =   "序"
      Height          =   375
      Left            =   4200
      TabIndex        =   26
      ToolTipText     =   "无序列表/有序列表"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cAlign_r 
      Caption         =   "右"
      Height          =   375
      Left            =   2400
      TabIndex        =   24
      ToolTipText     =   "右对齐"
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cAlign_c 
      Caption         =   "中"
      Height          =   375
      Left            =   2040
      TabIndex        =   23
      ToolTipText     =   "居中对齐"
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cAlign_l 
      Caption         =   "左"
      Height          =   375
      Left            =   1680
      TabIndex        =   22
      ToolTipText     =   "左对齐"
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cP 
      Caption         =   "自"
      Height          =   375
      Left            =   1320
      TabIndex        =   25
      ToolTipText     =   "自动排版"
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cSize 
      Caption         =   "字号:2号"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      ToolTipText     =   "字号"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdGaoji 
      Caption         =   "高级模式"
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer tmr1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5880
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6360
      Top             =   1560
   End
   Begin VB.CommandButton cCode 
      Caption         =   ">_"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      ToolTipText     =   "代码"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cHr 
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      ToolTipText     =   "分割线"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cQuote 
      Caption         =   "“"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      ToolTipText     =   "引用"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cUrl 
      Caption         =   "链"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "超链接"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "+"
      Height          =   375
      Left            =   7920
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "-"
      Height          =   375
      Left            =   5880
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtPage 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Page"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cImg 
      Caption         =   "图"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "插入图片"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cS 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      ToolTipText     =   "删除线"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cU 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   5
      ToolTipText     =   "下划线"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cI 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      ToolTipText     =   "斜体"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cB 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "粗体"
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox chkTop 
      Caption         =   "窗口置顶"
      Height          =   255
      Left            =   5880
      TabIndex        =   18
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CheckBox chk3 
      Caption         =   "手游版块签名档"
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CheckBox chk2 
      Caption         =   "新人指引"
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   960
      Width           =   2415
   End
   Begin VB.CheckBox chk1 
      Caption         =   "答疑图"
      Height          =   255
      Left            =   5880
      TabIndex        =   15
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox txtMsg 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label labColor 
      BackColor       =   &H00000000&
      Height          =   68
      Left            =   1680
      TabIndex        =   27
      Top             =   480
      Width           =   375
   End
   Begin VB.Menu m1 
      Caption         =   "编辑"
      Begin VB.Menu m1_All 
         Caption         =   "全选"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu m3 
      Caption         =   "视图"
      Begin VB.Menu m3_FontSize 
         Caption         =   "调整显示字号"
      End
   End
   Begin VB.Menu m5 
      Caption         =   "代码"
      Begin VB.Menu m5_hide 
         Caption         =   "隐藏内容[hide]"
      End
      Begin VB.Menu m5_1 
         Caption         =   "常用自定义编辑器代码"
         Begin VB.Menu m5_1_fly 
            Caption         =   "滚动文字[fly]"
         End
         Begin VB.Menu m5_1_sub 
            Caption         =   "下标[sub]"
         End
         Begin VB.Menu m5_1_sup 
            Caption         =   "上标[sup]"
         End
         Begin VB.Menu m5_1_hr1 
            Caption         =   "-"
         End
         Begin VB.Menu m5_1_ruby 
            Caption         =   "注释[ruby]"
         End
      End
   End
   Begin VB.Menu m10 
      Caption         =   "模板"
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RButton As Boolean

Dim GaojiMode As Boolean

Dim MouseSave As String
Dim MouseStart As String
Dim MouseStop As String

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub cAlign_c_Click()
    BBCode "align", , "center"
End Sub

Private Sub cAlign_l_Click()
    BBCode "align", , "left"
End Sub

Private Sub cAlign_r_Click()
    BBCode "align", , "right"
End Sub

Private Sub cB_Click()
    BBCode "b"
End Sub

Private Sub cCode_Click()
    BBCode "code"
End Sub

Private Sub cColor_Click()
    frmColor.Show
    If GaojiMode = False Then
        frmColor.Move Me.Left + 1000, Me.Top + 150
    Else
        frmColor.Move Me.Left + cColor.Left + 200, Me.Top + cColor.Top + cColor.Height + 820
    End If
End Sub

Private Sub cColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then RButton = True Else RButton = False
    MouseStart = "color"
    cColors.Visible = True
    cBCs.Visible = True
    cImg.Visible = False
    cS.Visible = False
End Sub

Private Sub cColor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then RButton = True Else RButton = False
    Timer1.Enabled = True
End Sub

Private Sub cColors_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseStop = "colors"
    cColors.Visible = False
End Sub

Private Sub chkTop_Click()
    If chkTop.Value = 1 Then SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    If chkTop.Value = 0 Then SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub cI_Click()
    BBCode "i"
End Sub

Private Sub cImg_Click()
    If RButton = False Then
        BBCode "img"
    ElseIf RButton = True Then
        GetBBCode "img", "网络图片", "网络图片地址", "宽(可选)", "高(可选)"
        RButton = False
        frmInput.Move Me.Left + cImg.Left + 200, Me.Top + cImg.Top + cImg.Height + 820
    End If
End Sub

Private Sub cImg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then RButton = True Else RButton = False
    MouseStart = "img"
    tmr1.Enabled = True
End Sub

Private Sub cImg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseStop = "img"
    If MouseStart <> "img" Then ColorReset
End Sub

Private Sub cImg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then RButton = True Else RButton = False
    Timer1.Enabled = True
End Sub

Private Sub cList_Click()
    If RButton = False Then
        BBCodeList
    Else
        BBCodeList , 1
    End If
End Sub

Private Sub cList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then RButton = True Else RButton = False
End Sub

Private Sub cList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then RButton = True Else RButton = False
End Sub

Private Sub cmdClear_Click()
    If PageMode = False Then
        txtMsg.Text = ""
        txtMsg.SetFocus
    Else
        
    End If
End Sub

Private Sub cmdCopy_Click()
    If PageMode = False Then
        Dim X As String
        X = txtMsg.Text
        If chk1.Value = 1 Then X = X & vbCrLf & Form2.txtImg
        If chk2.Value = 1 Then X = X & vbCrLf & Form2.txtNew
        If chk3.Value = 1 Then X = X & vbCrLf & Form2.txtMobile
        Clipboard.Clear
        Clipboard.SetText X
    Else
        Select Case Page
        Case "code:hide"
            Dim Y As String
            Y = txtMsg.Text
            BackPage
            GetBBCode "hide", "隐藏内容", "隐藏内容", "需要积分", , Y
        End Select
    End If
End Sub

Private Sub cmdDown_Click()
    If PageMode = False Then
        SetPage Page - 1
        If Page < 1 Then cmdDown.Enabled = False Else cmdDown.Enabled = True
        If Page > 29 Then cmdUp.Enabled = False Else cmdUp.Enabled = True
    End If
End Sub

Private Sub cHr_Click()
    Dim X As Long
    X = frm.txtMsg.SelStart
    frm.txtMsg.SelText = vbCrLf & "[hr]"
    frm.txtMsg.SetFocus
    frm.txtMsg.SelStart = X + 6
End Sub

Private Sub cmdGaoji_Click()
    If GaojiMode = False Then
        GaojiMode = True
        cmdGaoji.Caption = "普通模式"
        Me.WindowState = 2
    Else
        GaojiMode = False
        cmdGaoji.Caption = "高级模式"
        Me.WindowState = 0
    End If
End Sub

Private Sub cQuote_Click()
    BBCode "quote"
End Sub

Private Sub cmdUp_Click()
    If IsNumeric(Page) = True Then PageMode = False
    If PageMode = False Then
        SetPage Page + 1
        If Page < 1 Then cmdDown.Enabled = False Else cmdDown.Enabled = True
        If Page > 29 Then cmdUp.Enabled = False: cmdDown.SetFocus Else cmdUp.Enabled = True
    End If
End Sub

Private Sub cS_Click()
    BBCode "s"
End Sub

Private Sub cU_Click()
    BBCode "u"
End Sub

Private Sub cUrl_Click()
    GetBBCode "url", "超链接", "链接文本", "URL地址"
    frmInput.Move Me.Left + cUrl.Left + 200, Me.Top + cUrl.Top + cUrl.Height + 820
End Sub

Private Sub cUrl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseStart = "url"
    tmr1.Enabled = True
End Sub

Private Sub cUrl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseStop = "url"
    If MouseStart <> "url" Then ColorReset
End Sub

Private Sub cUrl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
End Sub

Private Sub Form_Click()
    Form_GotFocus
End Sub

Private Sub Form_GotFocus()
    Unload Form2
    Unload frmColor
    Unload frmInput
End Sub

Private Sub Form_Load()
    Form2.Show
    Form2.Visible = False
    Page = 1
    txtPage.Text = 1
    PageMode = False
    RButton = False
    GaojiMode = False
    
    TextColor(0) = 0
    TextColor(1) = 0
    TextColor(2) = 0
    
    ColorReset
End Sub

Private Sub Form_Paint()
On Error Resume Next
    If Me.Height < 4320 Then Me.Height = 4320
    If Me.Width < 8655 Then Me.Width = 8655
    
    If Me.WindowState = 2 And GaojiMode = False Then cmdGaoji_Click
    If Me.WindowState = 0 And GaojiMode = True Then cmdGaoji_Click
    
    If GaojiMode = False Then
        txtMsg.Top = 600
        cSize.Visible = False
        'cColor.Visible = False
        cAlign_l.Visible = False
        cAlign_c.Visible = False
        cAlign_r.Visible = False
        cP.Visible = False
        If TopMode = True Then TopMode = False: frm.chkTop.Value = 1
        frm.chkTop.Enabled = True
    Else
        txtMsg.Top = 1080
        cSize.Visible = True
        'cColor.Visible = True
        cAlign_l.Visible = True
        cAlign_c.Visible = True
        cAlign_r.Visible = True
        cP.Visible = True
        If frm.chkTop = 1 Then TopMode = True: frm.chkTop.Value = 0
        frm.chkTop.Enabled = False
    End If
    txtMsg.Height = Me.Height - txtMsg.Top - 945
    txtMsg.Width = Me.Width - txtMsg.Left - 2880
    
    fm1.Left = Me.Width - 2775
    cmdCopy.Left = Me.Width - 2775
    cmdCopy.Top = Me.Height - 1440
    cmdClear.Left = Me.Width - 2775
    cmdClear.Top = Me.Height - 2040
    
    chkTop.Left = Me.Width - 2775
    chkTop.Top = Me.Height - 2400
    
    chk1.Left = Me.Width - 2775
    chk2.Left = Me.Width - 2775
    chk3.Left = Me.Width - 2775
    
    cmdDown.Left = Me.Width - 2775
    txtPage.Left = Me.Width - 2415
    cmdUp.Left = Me.Width - 735
    
    cmdGaoji.Left = Me.Width - 3855
End Sub

Private Sub Form_Resize()
    Form_Paint
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub m1_All_Click()
    txtMsg.SetFocus
    txtMsg.SelStart = 0
    txtMsg.SelLength = Len(txtMsg.Text)
End Sub

Private Sub m3_FontSize_Click()
On Error Resume Next
    txtMsg.FontSize = InputBox("请输入字号(默认为9)", "调整显示字号", txtMsg.FontSize)
End Sub

Private Sub m5_1_fly_Click()
    BBCode "fly"
End Sub

Private Sub m5_1_ruby_Click()
    GetBBCode "ruby", "注释", "正文", "注释"
End Sub

Private Sub m5_1_sub_Click()
    BBCode "sub"
End Sub

Private Sub m5_1_sup_Click()
    BBCode "sup"
End Sub

Private Sub m5_hide_Click()
    SetPage "code:hide"
End Sub

Private Sub Timer1_Timer()
    If (MouseStart = "img" And MouseStop = "url") Or (MouseStart = "url" And MouseStop = "img") Then
        If MsgBox("是否插入带图片链接？这将替换您选择的内容。", vbYesNo, "带图片链接") = vbYes Then
            txtMsg.SelText = "[url=网页链接][img]图片链接[/img][/url]"
        End If
    End If
    If MouseStart = "color" And MouseStop = "colors" Then
        BBCode "color", , "RGB(" & TextColor(0) & "," & TextColor(1) & "," & TextColor(2) & ")"
    End If
    MouseSave = ""
    MouseStart = ""
    MouseStop = ""
    cColors.Visible = False
    cBCs.Visible = False
    cImg.Visible = True
    cS.Visible = True
    ColorReset
    Timer1.Enabled = False
End Sub

Private Sub tmr1_Timer()
    If MouseStart = "img" Then cUrl.BackColor = RGB(128, 256, 128)
    If MouseStart = "url" Then cImg.BackColor = RGB(128, 256, 128)
    If MouseStart = "color" Then cColors.Visible = True
    tmr1.Enabled = False
End Sub

Private Sub txtMsg_GotFocus()
    Form_GotFocus
End Sub
