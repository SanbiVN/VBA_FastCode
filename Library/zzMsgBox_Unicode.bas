Attribute VB_Name = "zzMsgBox_Unicode"

' __   _____   _ ®
' \ \ / / _ | / \
'  \ \ /| _ \/ / \
'   \_/ |___/_/ \_\
'

Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
  X As Long
  Y As Long
End Type
#If VBA7 Then
Private Declare PtrSafe Function GetWindowRect Lib "USER32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function SetWindowsHookEx Lib "USER32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As Long
Private Declare PtrSafe Function CallNextHookEx Lib "USER32" (ByVal hHook As LongPtr, ByVal CodeNo As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "USER32" (ByVal hHook As LongPtr) As Long
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Function FindWindowEx Lib "USER32" Alias "FindWindowExA" (ByVal ParenthWnd As LongPtr, ByVal ChildHwnd As LongPtr, ByVal classname As String, ByVal Caption As String) As LongPtr
Private Declare PtrSafe Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As Long
Private Declare PtrSafe Function SetWindowPos Lib "USER32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function CreateWindowEx Lib "USER32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal cp As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare PtrSafe Function SetWindowTextW Lib "USER32" (ByVal hwnd As LongPtr, ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function MsgBoxTimeoutW Lib "USER32" Alias "MessageBoxTimeoutW" (ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As VbMsgBoxStyle, ByVal wlange As Long, ByVal dwTimeout As Long) As Long
Private Declare PtrSafe Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function ClientToScreen Lib "USER32" (ByVal hwnd As LongPtr, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
#Else
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "USER32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MsgBoxTimeoutW Lib "user32" Alias "MessageBoxTimeoutW" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As VbMsgBoxStyle, ByVal wlange As Long, ByVal dwTimeout As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" ( ByVal hHook As Long, ByVal CodeNo As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" ( ByVal hHook As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( ByVal ParenthWnd As Long, ByVal ChildhWnd As Long, ByVal className As String, ByVal Caption As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" ( ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" ( ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" ( ByVal hObject As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" ( ByVal h As Long, ByVal W As Long, ByVal e As Long, ByVal o As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function SetWindowTextW Lib "user32" ( ByVal hwnd As Long, ByVal lpString As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
#End If
#If VBA7 And Win64 Then
Private hDlgHook^, hDlgHWnd^
#ElseIf VBA7 Then
Private hDlgHook As LongPtr, hDlgHWnd As LongPtr
#Else
Private hDlgHook&, hDlgHWnd&
#End If

Private hFont&, newRECT As RECT, newPoint As POINTAPI, iShowUnderCursor As Boolean
Private Sub Alert_test()
  Alert "Xin ch" & ChrW(224) & "o, b" & ChrW(7841) & "n mu" & ChrW(7889) & "n bao nhi" & ChrW(234) & "u gi" & ChrW(226) & "y t" & ChrW(7921) & " " & ChrW(273) & ChrW(7897) & "ng " & ChrW(273) & ChrW(243) & "ng th" & ChrW(244) & "ng b" & ChrW(225) & "o?", vbOKCancel, Timeout:=5
End Sub
Private Sub Alert_test2()
  'Return Value:
  ' End Timeout = 32000 (Het thoi gian chon)
  ' OK = 1 (Xac Nhan)
  ' Cancel = 2 (Huy 1)
  ' Abort = 3 (Huy 2)
  ' Retry = 4 (Thu Lai)
  ' Ignore = 5 (Bo Qua)
  ' Yes = 6 (Co)
  ' No = 7 (Khong)
  
  'Debug.Print Alert("OK?", vbOKCancel, Timeout:=5)
  'Debug.Print Alert("OK?", vbAbortRetryIgnore, Timeout:=5)
  'Debug.Print Alert("OK?", vbYesNoCancel, Timeout:=5)

End Sub
' Last Edit: 09/03/2020 17:01
#If VBA7 Then
Public Function Alert(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As String = "Thông báo", Optional ByVal hwnd As LongPtr = &H0, Optional ByVal Timeout& = 2, Optional ByVal ShowUnderCursor As Boolean = True) As VbMsgBoxResult
#Else
Public Function Alert(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As String = "Thông báo", Optional ByVal hwnd& = &H0, Optional ByVal Timeout& = 2, Optional ByVal ShowUnderCursor As Boolean = True) As VbMsgBoxResult
#End If
  iShowUnderCursor = ShowUnderCursor
  If Timeout <= 0 Then Timeout = 3600
  #If Win64 Then
    hDlgHook = SetWindowsHookEx(5, AddressOf HookProcMsgBox, Application.HinstancePtr, GetCurrentThreadId())
  #Else
    hDlgHook = SetWindowsHookEx(5, AddressOf HookProcMsgBox, Application.hInstance, GetCurrentThreadId())
  #End If
  Call SetWindowPos(hDlgHWnd, -1, 0, 0, 0, 0, &H2 Or &H1)
  Alert = MsgBoxTimeoutW(hwnd, VBA.StrConv(Prompt, 64), VBA.StrConv(title, 64), Buttons Or &H2000&, 0&, Timeout * 1000)
  DeleteObject hFont
End Function

#If VBA7 And Win64 Then
Private Function HookProcMsgBox&(ByVal nCode&, ByVal wParam^, ByVal lParam^)
  Dim hStatic1^, hStatic2^, hButton^, nCaption$, lCaption$
#ElseIf VBA7 Then
Private Function HookProcMsgBox&(ByVal nCode&, ByVal wParam As LongPtr, ByVal lParam As LongPtr)
  Dim hStatic1 As LongPtr, hStatic2 As LongPtr, hButton As LongPtr, nCaption$, lCaption$
#Else
Private Function HookProcMsgBox&(ByVal nCode&, ByVal wParam&, ByVal lParam&)
  Dim hStatic1&, hStatic2&, hButton&, nCaption$, lCaption$
#End If
  HookProcMsgBox = CallNextHookEx(hDlgHook, nCode, wParam, lParam)
  If nCode = 5 Then
    hFont = CreateFont(13, 0, 0, 0, 500, 0, 0, 0, 0, 0, 0, 0, 0, "Tahoma")
    hStatic1 = FindWindowEx(wParam, 0&, "Static", VBA.vbNullString)
    hStatic2 = FindWindowEx(wParam, hStatic1, "Static", VBA.vbNullString)
    hDlgHWnd = wParam
    Call SetWindowPos(hDlgHWnd, -3, 0, 0, 0, 0, &H2 Or &H1)
    If hStatic2 = 0 Then hStatic2 = hStatic1
    SendMessage hStatic2, &H30, hFont, ByVal 1&
    '--------------------------------------
    nCaption = "&X" & VBA.ChrW(225) & "c nh" & VBA.ChrW(7853) & "n"
    lCaption = "OK":      GoSub Send
    nCaption = "&C" & VBA.ChrW(243)
    lCaption = "&Yes":    GoSub Send
    nCaption = "&Kh" & VBA.ChrW(244) & "ng"
    lCaption = "&No":     GoSub Send
    nCaption = "&H" & VBA.ChrW(7911) & "y"
    lCaption = "Cancel":  GoSub Send
    nCaption = "&Th" & VBA.ChrW(7917) & " l" & VBA.ChrW(7841) & "i"
    lCaption = "&Retry":  GoSub Send
    nCaption = "&B" & VBA.ChrW(7887) & " qua"
    lCaption = "&Ignore": GoSub Send
    nCaption = "H" & VBA.ChrW(7911) & "&y b" & VBA.ChrW(7887)
    lCaption = "&Abort":  GoSub Send
    nCaption = "Tr" & VBA.ChrW(7907) & " &gi" & VBA.ChrW(250) & "p"
    lCaption = "Help":    GoSub Send
    '--------------------------------------
    If iShowUnderCursor Then
      GetCursorPos newPoint
      GetWindowRect wParam, newRECT
      MoveWindow wParam, newPoint.X, newPoint.Y, (newRECT.Right - newRECT.Left - 1), (newRECT.Bottom - newRECT.Top - 1), False
    End If
    UnhookWindowsHookEx hDlgHook
  End If
Exit Function
Send:
  hButton = FindWindowEx(wParam, 0&, "Button", lCaption)
  SendMessage hButton, &H30, hFont, 0
  SetWindowTextW hButton, StrPtr(nCaption)
Return
End Function


Private Sub Text2CodeVBA_test()
  Dim p$
  p = Application.InputBox("Input")
  ' Dán vãn baÒn bãÌng phím tãìt Ctrl+V
  If p = vbNullString Then
    Exit Sub
  End If
  With VBA.GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    .SetText Text2CodeVBA(p), 1
    .PutInClipboard
    VBA.CreateObject("WScript.Shell").SendKeys "^v"
  End With
End Sub

Function Text2CodeVBA(ByVal Text As String, Optional ByVal procedureName$, Optional ByVal limitRows% = 300, Optional ByVal limitColumns% = 950)
  Dim L&
  L = Len(Text)
  If L < 1 Then Exit Function
  Dim i&, s, s1, s2, s3$, s4$, t$, lt$, t1$, t2$, k&, kk&, v&
  t1 = "Dim s$"
  If procedureName <> "" Then
    t2 = "s = s & """
  Else
    t2 = """"
  End If
  s3 = t2
  For i = 1 To L
    t = Mid(Text, i, 1)
    v = 0
    Select Case t
    Case """": s3 = s3 & """"""
    Case vbCr:
    Case vbLf:
      k = k + 1
      If k > limitRows Then
        GoSub join
      Else
        s3 = s3 & """ & vbLF" & vbLf & IIf(i = L, "", "s = s & """)
      End If
    Case Else
      'StrConv(t, 64) Like "[! ][!" & VBA.vbNullChar & "]" Or
      v = AscW(t)
      If v > 127 Then
        s3 = s3 & """ & ChrW(" & CStr(v) & ") " & IIf(i = L, "", "& """)
      Else
        s3 = s3 & t
      End If
    End Select
    If Len(Split(s3, vbLf)(UBound(Split(s3, vbLf)))) >= limitColumns Then
      s3 = s3 & """ & vbLF" & vbNewLine & IIf(i = L, "", "s = s & """)
    End If
    lt = t
  Next i
  GoSub join
  If kk > 0 Then
    s = s2
  End If
  Text2CodeVBA = s

Exit Function
join:
  If s3 <> t2 Then
    kk = kk + 1
    If procedureName <> "" Then
      s1 = s1 & "s = s & " & procedureName & kk & " & n" & vbNewLine
      s2 = s2 & "Function " & procedureName & kk & "()" & vbNewLine & _
              t1 & vbNewLine & s3 & IIf(s3 Like "*& vbLF" & vbNewLine, "", """") & vbNewLine & _
              procedureName & kk & " = s" & vbNewLine & _
              "End Function" & vbNewLine
    Else
      s2 = s3 & IIf(v > 127 Or s3 Like "*& vbLF" & vbNewLine, "", """")
    End If
  End If
  k = 0: s3 = t2
Return
End Function

