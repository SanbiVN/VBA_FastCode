VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pppHandleEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event complete(obj As pppThreadHandle, CCW As pppTaskHTTP)
Public Event Opened(threadID$)
Public Event Closed(threadID$)
Sub NotifyComplete(obj As pppThreadHandle, Optional CCW As pppTaskHTTP)
RaiseEvent complete(obj, CCW)
End Sub
Sub NotifyThreadOpened(threadID$)
RaiseEvent Opened(threadID)
End Sub
Sub NotifyThreadClosed(threadID$)
RaiseEvent Closed(threadID)
End Sub
