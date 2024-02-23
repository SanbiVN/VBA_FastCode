Attribute VB_Name = "zzzzzSheetControler"
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias "GetTickCount" () As Long
#Else
Private Declare Function getTickCount Lib "kernel32" Alias "GetTickCount" () As Long
#End If
Private Const DelayDoubleClick As Double = 0.95
Private Const GroupCenter = "GroupCenter"
Private CallerAutoScroll$
Private TimeAutoScroll%
Private ScrollDown%, ScrollToRight%
Private AllowAutoScroll As Boolean
Private AllowTurnSheet As Boolean
Private Function CenterButtons()
  CenterButtons = Array("btnMoveCenter", _
            "btnMoveRight", "btnMoveBottom", "btnMoveLeft", "btnMoveTop", _
            "btnPreviousSheet", "btnNextSheet")
End Function

'/////////////////////////////////////////////////
Sub btnMoveCenter()
  Call Controler(CenterButtons(0))
End Sub
'/////////////////////////////////////////////////

Sub btnMoveRight()
  Call Controler(CenterButtons(1))
End Sub

Sub btnMoveBottom()
  Call Controler(CenterButtons(2))
End Sub
Sub btnMoveLeft()
  Call Controler(CenterButtons(3))
End Sub

Sub btnMoveTop()
  Call Controler(CenterButtons(4))
End Sub

Sub btnPreviousSheet()
  Call Controler(CenterButtons(5))
End Sub
Sub btnNextSheet()
  Call Controler(CenterButtons(6))
End Sub
Sub Center_Control()
  On Error Resume Next
  ThisWorkbook.ActiveSheet.Shapes("GroupCenter").Placement = 3
  Application.ActiveWindow.ScrollRow = 1
  Application.ActiveWindow.ScrollColumn = 1
  On Error GoTo 0
End Sub

' Move
Private Sub Controler(ByVal Button$)
  Const NumberMoveRow = 10, NumberMoveCol = 5
  On Error Resume Next
  If typename(Application.Caller(1)) <> "String" Then Exit Sub
  Static MoveWS%, DMoveWS As Date, Actived As Object
  Dim Row&, Col%, sRow&, sCol%, a, aw As Object, home As Object
  a = CenterButtons
  Set aw = ActiveWorkbook.ActiveSheet
  If CallerAutoScroll <> Button Then GoSub UnMoveNumber: AllowTurnSheet = False
  GoSub GetClick
  SetUsedRangeLimit Row, Col
  sRow = Application.ActiveWindow.ScrollColumn
  sCol = Application.ActiveWindow.ScrollRow
  Select Case Button
    Case a(5), a(6)
      'If MoveWS >= 1 Then AllowTurnSheet = True
      GoSub AutoSheet
    Case a(0)
      If MoveWS >= 1 Then
        Set a = Actived
        For Each home In ActiveWorkbook.Worksheets
          If home.visible Then
            Exit For
          End If
        Next
        If Not aw Is home Then
          home.Activate
        Else
          If Not a Is home Then
            a.Activate
          End If
        End If
        Set Actived = aw
      End If
      If AllowAutoScroll Then
        AllowAutoScroll = False
      Else
        Center_Control
      End If
    Case Else
      If CallerAutoScroll <> Button Then AllowAutoScroll = False
      If MoveWS >= 1 Then AllowAutoScroll = True
      Select Case Button
        Case a(1): ScrollDown = 0: ScrollToRight = NumberMoveCol
        Case a(2): ScrollDown = NumberMoveRow: ScrollToRight = 0
        Case a(3): ScrollDown = 0: ScrollToRight = -NumberMoveCol
        Case a(4): ScrollDown = -NumberMoveRow: ScrollToRight = 0
      End Select
      GoSub DoScroll
  End Select
  If MoveWS >= 1 Then GoSub UnMoveNumber
  CallerAutoScroll = Button
Ends:
  Set aw = Nothing
  Set home = Nothing
Exit Sub
DoScroll:
  If (ScrollToRight > 0 And sCol > Col) Or (ScrollDown > 0 And sRow > Row) Then GoSub UnMoveNumber: GoTo Ends
  If TimeAutoScroll < 150 Then TimeAutoScroll = 150
  If AllowAutoScroll Then
    Do Until Not AllowAutoScroll
      Delay TimeAutoScroll: GoSub scroll:
    Loop
  Else
    GoSub scroll
  End If
Return
scroll:
  With Application
    If (ScrollDown > 0 And .ActiveWindow.ScrollRow > Row) Or (ScrollDown < 0 And .ActiveWindow.ScrollRow < 12) Or (ScrollToRight > 0 And .ActiveWindow.ScrollColumn > Col) Or (ScrollToRight < 0 And .ActiveWindow.ScrollColumn < 7) Then AllowAutoScroll = False
    .ActiveWindow.SmallScroll Down:=ScrollDown, ToRight:=ScrollToRight
  End With
Return
AutoSheet:
  On Error Resume Next
  If AllowTurnSheet Then
    Do Until Not AllowTurnSheet
      GoSub turnSheet
      AllowTurnSheet = Err.Number <> 91
      Delay 400:
    Loop
  Else
    GoSub turnSheet
  End If
  On Error GoTo 0
Return
GetClick:
  If DMoveWS = 0 Then
    MoveWS = 0: DMoveWS = VBA.Now + 1 / 24 / 60 / 60 * DelayDoubleClick
  Else: MoveWS = MoveWS + 1:
    If VBA.Now > DMoveWS Then GoSub UnMoveNumber
  End If
Return
turnSheet:
  If a(6) = Button Then
    ThisWorkbook.ActiveSheet.Next.Select
  Else
    ThisWorkbook.ActiveSheet.Previous.Select
  End If
Return
UnMoveNumber: MoveWS = 0: DMoveWS = 0: Return
End Sub

Private Sub CenterButtonsCreate_test()
  Call CenterButtonsCreate(ActiveSheet)
End Sub
Sub CenterButtonsCreate(ByVal ws As Worksheet)
Attribute CenterButtonsCreate.VB_ProcData.VB_Invoke_Func = " \n14"
  On Error Resume Next
  Dim s As Object, sn$
  Set s = ws.Shapes(GroupCenter)
'  If Not s Is Nothing Then
'    GoTo Ends
'  End If
  CenterControlDelete ws
  Dim rt&, F&, L!, t!, W!, H!, a, Y&
  a = CenterButtons
  Y = 52
  W = 10.52551
  H = 11.08622
  sn = a(1): L = 27: t = 14: rt = 1: F = 0: GoSub M
  sn = a(2): L = 14: t = 27: rt = 1: F = 90:  GoSub M
  sn = a(3): L = 1: t = 14: rt = 0: F = 0:  GoSub M
  sn = a(4): L = 14: t = 1: rt = 0: F = 90: GoSub M
  sn = a(5): L = 1: t = 29: W = 10: H = 7: rt = 0: F = 0: GoSub M
  sn = a(6): L = 27: t = 29: W = 10: H = 7: rt = 1: F = 0: GoSub M
  Y = 1
  sn = a(0): L = 14.5: t = 14.5: W = 10: H = 10: F = 45: rt = 1: GoSub M

  With ws.Shapes.Range(Array(a(0), a(1), a(2), a(3), a(4), a(5), a(6)))
    .Group
    .Name = GroupCenter
    ws.Shapes(GroupCenter).Placement = 3
  End With
    On Error GoTo 0
Ends:
  Set s = Nothing
Exit Sub
M:
  Set s = Nothing
  Set s = ws.Shapes(sn)
  If s Is Nothing Then
    Set s = ws.Shapes.AddShape(Y, L, t, W, H)
    s.Name = sn
  End If
  With s
    .LEFT = L
    .TOP = t
    .Width = W
    .Height = H
    .Flip rt
    .Rotation = F
    .Line.visible = 0
    .OnAction = "'" & ws.Parent.Name & "'!" & sn
    .Placement = 3
    With .Fill
      .ForeColor.RGB = 57548
      .Solid
    End With
  End With
Return
End Sub

Sub CreateControlCenterAll()
  Call CreateControlCenterWB(ActiveWorkbook)
End Sub
Public Sub CreateControlCenterWB(WB As Workbook)

  On Error Resume Next
  Dim s As Object, ws As Worksheet, aws As Worksheet, item
  Dim r%, c%, W!, H!, Rng As Range
  Set aws = WB.ActiveSheet
  For Each ws In WB.Worksheets
    ws.Activate
    CenterButtonsCreate ws
    Set s = ws.Shapes(GroupCenter)
    For Each item In s.GroupItems
      If item.OnAction Like "*!*" Then
        item.OnAction = Split(item.OnAction, "!")(1)
      End If
    Next
    For r = 1 To 10
      For c = 1 To 10
        Set Rng = ws.Range("A1")(r, c)
        If Rng.LEFT >= s.Width And Rng.TOP >= s.Height Then
          If r <> ActiveWindow.SplitRow Or ActiveWindow.SplitColumn <> c Then
            If ActiveWindow.SplitRow > 0 Or ActiveWindow.SplitColumn > 0 Then
              'ActiveWindow.FreezePanes = False
            Else
              Rng.Select
              ActiveWindow.FreezePanes = True
            End If
          End If
          GoTo N
        End If
      Next
    Next
N:
  Next
  aws.Activate
  On Error GoTo 0
End Sub


Sub CenterControlDeleteAll()
  On Error Resume Next
  Dim a
  For Each a In ThisWorkbook.Worksheets
    a.Shapes(GroupCenter).Delete
    CenterControlDelete a
  Next
  On Error GoTo 0
End Sub

Private Sub CenterControlDelete(ByVal ws As Worksheet)
  On Error Resume Next
  Dim a
  For Each a In CenterButtons
    ws.Shapes(a).Delete
  Next
  On Error GoTo 0
End Sub


Sub DeleteShapes(ByVal ws As Worksheet, ParamArray Shapes())
  Dim s
  On Error Resume Next
  For Each s In Shapes
    ws.Shapes(s).Delete:
  Next s
  On Error GoTo 0
End Sub

Sub SetUsedRangeLimit(ByRef LastRow&, ByRef LastCol%, ParamArray Args())
  Dim Arg
  On Error Resume Next
  With ThisWorkbook.ActiveSheet
    For Each Arg In Args
      If LCase$(.Name) = LCase$(Arg) Then Exit Sub
    Next
    Err.clear
    LastRow = .Cells.find("*", After:=.Cells(1), LookIn:=xlFormulas, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    LastCol = .Cells.find("*", After:=.Cells(1), LookIn:=xlFormulas, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column
  End With
  On Error GoTo 0
End Sub
Private Sub Delay(Optional ByVal MiliSecond% = 1000)
  Dim Start&, check&
  Start = getTickCount&()
  Do Until check >= Start + MiliSecond
    DoEvents
    check = getTickCount&()
  Loop
End Sub
