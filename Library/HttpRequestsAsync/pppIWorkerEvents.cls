VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pppIWorkerEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event complete(CCW As pppTaskHTTP)
Public Event statusChange(statusVal As Variant)
Public Event Started()
Sub complete(Optional CCW As pppTaskHTTP)
RaiseEvent complete(CCW)
End Sub
Sub statusChange(statusVal As Variant)
RaiseEvent statusChange(statusVal)
End Sub
Sub Started()
RaiseEvent Started
End Sub
