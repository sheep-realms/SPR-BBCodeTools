VERSION 5.00
Begin VB.Form frm 
   Caption         =   "Edit"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   8415
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cCode 
      Caption         =   ">_"
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cHr 
      Caption         =   "--"
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cQuote 
      Caption         =   """"
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cUrl 
      Caption         =   "链"
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "+"
      Height          =   375
      Left            =   7920
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "-"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
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
      TabIndex        =   12
      Text            =   "Page"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cImg 
      Caption         =   "图"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox chkTop 
      Caption         =   "窗口置顶"
      Height          =   255
      Left            =   5880
      TabIndex        =   6
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CheckBox chk3 
      Caption         =   "手游版块签名档"
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CheckBox chk2 
      Caption         =   "新人指引"
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.CheckBox chk1 
      Caption         =   "答疑图"
      Height          =   255
      Left            =   5880
      TabIndex        =   2
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
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   5655
   End
   Begin VB.Menu m1 
      Caption         =   "编辑"
      Begin VB.Menu m1_All 
         Caption         =   "全选"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RButton As Boolean

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub cB_Click()
    BBCode "b"
End Sub

Private Sub cCode_Click()
    BBCode "code"
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
        Dim X As Integer
        X = txtMsg.SelStart
        txtMsg.SelText = "[img][/img]"
        txtMsg.SetFocus
        txtMsg.SelStart = X + 5
    ElseIf RButton = True Then
        txtMsg.SelText = "[img]" & InputBox("网络图片地址：", "插入图片") & "[/img]"
        txtMsg.SetFocus
        txtMsg.SelStart = Len(txtMsg.Text)
        RButton = False
    End If
End Sub

Private Sub cImg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then RButton = True
End Sub

Private Sub cImg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then RButton = True
End Sub

Private Sub cmdClear_Click()
    txtMsg.Text = ""
    txtMsg.SetFocus
End Sub

Private Sub cmdCopy_Click()
    Dim X As String
    X = txtMsg.Text
    If chk1.Value = 1 Then X = X & vbCrLf & Form2.txtImg
    If chk2.Value = 1 Then X = X & vbCrLf & Form2.txtNew
    If chk3.Value = 1 Then X = X & vbCrLf & Form2.txtMobile
    Clipboard.Clear
    Clipboard.SetText X
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

Private Sub cQuote_Click()
    BBCode "quote"
End Sub

Private Sub cmdUp_Click()
    If PageMode = False Then
        SetPage Page + 1
        If Page < 1 Then cmdDown.Enabled = False Else cmdDown.Enabled = True
        If Page > 29 Then cmdUp.Enabled = False Else cmdUp.Enabled = True
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
End Sub

Private Sub Form_Load()
    Form2.Show
    Form2.Visible = False
    Page = 1
    txtPage.Text = 1
    PageMode = False
    RButton = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub m1_All_Click()
    txtMsg.SetFocus
    txtMsg.SelStart = 0
    txtMsg.SelLength = Len(txtMsg.Text)
End Sub
