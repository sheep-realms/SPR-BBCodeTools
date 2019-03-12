VERSION 5.00
Begin VB.UserControl UC_MsgBar 
   BackColor       =   &H00F0F0F0&
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ScaleHeight     =   855
   ScaleWidth      =   4815
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   240
   End
   Begin VB.Label labText 
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label labTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Shape sp 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "UC_MsgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Cilck()
Public Event DblCilck()
'Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Timer()
Public Event TickEnd()

Dim Ticks As Integer

Public Property Let Interval(ByVal Value As Integer)
    tmr.Interval = Value
End Property

Public Property Get Interval() As Integer
    Interval = tmr.Interval
End Property

Public Property Let Tick(ByVal Value As Integer)
    Ticks = Value
    TickChange
End Property

Public Property Get Tick() As Integer
    Tick = Ticks
End Property

Public Sub Output(ByVal Title As String, Text As String)
    labTitle.Caption = Title
    labText.Caption = Text
End Sub

Public Sub Color(ByVal r As Integer, g As Integer, b As Integer)
    Dim rs As String, gs As String, bs As String
    rs = r + (255 - r) * 0.8
    gs = g + (255 - g) * 0.8
    bs = b + (255 - b) * 0.8
    UserControl.BackColor = RGB(rs, gs, bs)
    rs = r + (255 - r) * 0.5
    gs = g + (255 - g) * 0.5
    bs = b + (255 - b) * 0.5
    rs = rs - r * 0.1
    gs = gs - g * 0.1
    bs = bs - b * 0.1
    sp.BackColor = RGB(rs, gs, bs)
End Sub

Private Sub TickChange()
    tmr.Enabled = True
End Sub

Private Sub labText_Click()
    UserControl_Click
End Sub

Private Sub labText_DblClick()
    UserControl_DblClick
End Sub

Private Sub labTitle_Click()
    UserControl_Click
End Sub

Private Sub labTitle_DblClick()
    UserControl_DblClick
End Sub

Private Sub tmr_Timer()
    RaiseEvent Timer
    If Ticks = -1 Then Exit Sub
    If Ticks > 0 Then
        Ticks = Ticks - 1
    Else
        tmr.Enabled = False
        RaiseEvent TickEnd
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Cilck
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblCilck
End Sub


'Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent MouseDown(Button, Shift, X, Y)
'End Sub
