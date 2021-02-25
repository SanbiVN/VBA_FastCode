VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pppQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private head As pppQueueItem
Private tail As pppQueueItem
Private countQ As Long
Private Sub Class_Initialize()
    countQ = 0
End Sub
Private Sub Class_Terminate()
    countQ = 0
    Set head = Nothing
    Set tail = Nothing
End Sub
Public Property Get IsEmpty() As Boolean
    IsEmpty = ((head Is Nothing) And (tail Is Nothing))
End Property
Public Property Get Count() As Long
    Count = countQ
End Property
Public Property Get Peek() As Variant
    Peek = head.value
End Property
Public Function EnQueue(v As Variant)
  Dim pppQueueItem As New pppQueueItem
  If VBA.IsObject(v) Then
    Set pppQueueItem.value = v
  Else
    pppQueueItem.value = v
  End If
  If Me.IsEmpty = True Then
    Set head = pppQueueItem
    Set tail = head
  Else
    Set tail.NextItem = pppQueueItem
    Set tail = pppQueueItem
  End If
  countQ = countQ + 1
  Set pppQueueItem = Nothing
End Function

Public Function DeQueue() As Variant
  If Me.IsEmpty = True Then
    DeQueue = Null
  Else
    If VBA.IsObject(head.value) Then
      Set DeQueue = head.value
    Else
      DeQueue = head.value
    End If
    If head Is tail Then
      Set head = Nothing
      Set tail = Nothing
      countQ = 0
    Else
      Set head = head.NextItem
      countQ = countQ - 1
    End If
  End If
End Function

Public Function clear()
    countQ = 0
    Set head = Nothing
    Set tail = Nothing
End Function

Public Function ToArray() As Variant
  Dim sizeQ As Long
  Dim result() As Variant
  Dim index As Long
  Dim tmp As pppQueueItem
  sizeQ = Me.Count - 1
  If sizeQ > -1 Then
    ReDim result(sizeQ)
    Set tmp = head
    For index = 0 To sizeQ
      If VBA.IsObject(tmp.value) Then
        Set result(index) = tmp.value
      Else
        result(index) = tmp.value
      End If
      Set tmp = tmp.NextItem
    Next index
    ToArray = result
  Else
    Erase result
  End If
  Set tmp = Nothing
End Function
