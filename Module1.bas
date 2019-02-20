Attribute VB_Name = "Module1"
Option Explicit

Public Page As String
Public PageMode As Boolean
Public TopMode As Boolean

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
End Function

Public Function SetPage(Value As String)
    If IsNumeric(Page) = True Then
        txtSave(Page) = frm.txtMsg.Text
    Else
    
    End If
    
    frm.txtMsg.Text = ""
    
    If IsNumeric(Value) = True Then
        frm.txtMsg.Text = txtSave(Value)
    Else
    
    End If
    
    frm.txtPage.Text = Value
    Page = Value
End Function
