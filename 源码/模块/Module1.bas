Attribute VB_Name = "Module1"
Option Explicit

Public Page As String
Public PageData As String
Public PageUp As String
Public PageSelStart As Long
Public PageSelLength As Long
Public PageMode As Boolean
Public TopMode As Boolean

Public TextColor(2) As Integer

Public BBInputMode As Boolean
Public BBInputV(2) As String
Public BBInputL(2) As String
Public BBInputC As String

Public txtSave(30) As String

Public Function BBCode(ByVal Code As String, Optional ByVal V1 As String, Optional ByVal V2 As String, Optional ByVal V3 As String, Optional ByVal Mode As Boolean)
    Dim X As Long
    Dim i As Integer
    Dim j As Long
    
    If V1 = "" Then V1 = frm.txtMsg.SelText
    i = Len(Code) + 2
    j = Len(V1)
    X = frm.txtMsg.SelStart
    
    If V2 = "" And V3 = "" And Mode = False Then
        If V1 = "" Then
            frm.txtMsg.SelText = "[" & Code & "][/" & Code & "]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i
        Else
            frm.txtMsg.SelText = "[" & Code & "]" & V1 & "[/" & Code & "]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i
            frm.txtMsg.SelLength = j
        End If
    ElseIf (Mode = True) Or (V2 <> "" And V3 <> "") Then
        i = i + 2 + Len(V2) + Len(V3)
        If V1 = "" Then
            frm.txtMsg.SelText = "[" & Code & "=" & V2 & "," & V3 & "][/" & Code & "]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i
        Else
            frm.txtMsg.SelText = "[" & Code & "=" & V2 & "," & V3 & "]" & V1 & "[/" & Code & "]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i
            frm.txtMsg.SelLength = j
        End If
    ElseIf V3 = "" Then
        i = i + 1 + Len(V2)
        If V1 = "" Then
            frm.txtMsg.SelText = "[" & Code & "=" & V2 & "][/" & Code & "]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i
        Else
            frm.txtMsg.SelText = "[" & Code & "=" & V2 & "]" & V1 & "[/" & Code & "]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i
            frm.txtMsg.SelLength = j
        End If
    End If
End Function

Public Function BBCodeList(Optional ByVal V1 As String, Optional ByVal V2 As String)
    Dim X As Long
    Dim i As Integer
    Dim j As Long
    
    If V1 = "" Then V1 = frm.txtMsg.SelText
    i = 6
    j = Len(V1)
    X = frm.txtMsg.SelStart
    
    If Mid(frm.txtMsg.Text, frm.txtMsg.SelStart + 3, 7) = "[/list]" Then
        frm.txtMsg.SelText = vbCrLf & "[*]"
        frm.txtMsg.SetFocus
        frm.txtMsg.SelStart = X + 5
    ElseIf V2 = "" Then
        If V1 = "" Then
            frm.txtMsg.SelText = "[list]" & vbCrLf & "[*]" & vbCrLf & "[/list]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i + 5
        Else
            V1 = Replace(V1, vbCrLf, vbCrLf & "[*]")
            X = Len(V1)
            frm.txtMsg.SelText = "[list]" & vbCrLf & "[*]" & V1 & vbCrLf & "[/list]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i + 5
            'frm.txtMsg.SelLength = j
        End If
    Else
        i = i + 1 + Len(V2)
        If V1 = "" Then
            frm.txtMsg.SelText = "[list=" & V2 & "]" & vbCrLf & "[*]" & vbCrLf & "[/list]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i + 5
        Else
            V1 = Replace(V1, vbCrLf, vbCrLf & "[*]")
            X = Len(V1)
            frm.txtMsg.SelText = "[list=" & V2 & "]" & vbCrLf & "[*]" & V1 & vbCrLf & "[/list]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i + 5
            'frm.txtMsg.SelLength = j
        End If
    End If

End Function

Public Function GetBBCode(ByVal Code As String, Optional ByVal Title As String, Optional ByVal L1 As String, Optional ByVal L2 As String, Optional ByVal L3 As String, Optional ByVal V1 As String, Optional ByVal V2 As String, Optional ByVal V3 As String)
    BBInputC = Code
    BBInputL(0) = L1
    BBInputL(1) = L2
    BBInputL(2) = L3
    If V1 = "" And frm.txtMsg.SelText <> "" Then V1 = frm.txtMsg.SelText
    BBInputV(0) = V1
    BBInputV(1) = V2
    BBInputV(2) = V3
    If frm.chkTop = 1 Then TopMode = True: frm.chkTop.Value = 0
    frm.chkTop.Enabled = False
    frmInput.Show
    If Title = "" Then frmInput.Caption = Code Else frmInput.Caption = Title
    ColorReset
End Function

Public Function SetPage(Value As String)
    Dim Values As String
    Values = Value
    If IsNumeric(Page) = True Then
        txtSave(Page) = frm.txtMsg.Text
        PageSelStart = frm.txtMsg.SelStart
        PageSelLength = frm.txtMsg.SelLength
    Else
    
    End If
    
    frm.txtMsg.Text = ""
    
    If IsNumeric(Values) = True Then
        PageMode = False
        frm.cmdUp.Enabled = True
        frm.cmdDown.Enabled = True
        frm.txtMsg.Text = txtSave(Values)
        frm.txtMsg.SelStart = PageSelStart
        frm.txtMsg.SelLength = PageSelLength
        frm.cmdCopy.Caption = "&Copy"
    Else
        PageMode = True
        frm.cmdUp.Enabled = False
        frm.cmdDown.Enabled = False
        Select Case Values
        Case "code:hide"
            frm.cmdCopy.Caption = "完成"
        End Select
    End If
    
    PageUp = Page
    'If Value <> Values Then MsgBox "数据异常！"
    '你他喵的，这玩意自己会变？逗我呢？
    frm.txtPage.Text = Values
    Page = Values
End Function

Public Function BackPage()
    SetPage PageUp
End Function

Public Function ColorReset()
    frm.cUrl.BackColor = &H8000000F
    frm.cImg.BackColor = &H8000000F
    
    frm.labColor.BackColor = RGB(TextColor(0), TextColor(1), TextColor(2))
    frm.cColors.BackColor = RGB(TextColor(0), TextColor(1), TextColor(2))
    'If (TextColor(0) + TextColor(1) + TextColor(2)) / 2 < 128 Then
End Function
