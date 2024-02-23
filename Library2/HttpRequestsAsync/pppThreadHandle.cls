VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pppThreadHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents workerEvents As pppIWorkerEvents
Attribute workerEvents.VB_VarHelpID = -1
Private eventHandle As pppHandleEvents
Private workerObject As pppIWorker
Private taskID$, handleID$, handleArgs As Variant
Private Sub workerEvents_Complete(CCW As pppTaskHTTP)
eventHandle.NotifyComplete Me, CCW
End Sub
Public Property Set Worker(ByVal workObj As pppIWorker)
Set workerObject = workObj.CreateNew
Set workerEvents = New pppIWorkerEvents
Set workerObject.Events = workerEvents
End Property
Sub OpenHandle(multiThreadGroup As pppMultiThread, delegate As pppHandleEvents)
Set eventHandle = delegate
handleID = multiThreadGroup.NewHandleID
eventHandle.NotifyThreadOpened Name
End Sub
Sub Execute(Optional Args As pppTaskHTTP)
workerObject.Execute Args
End Sub
Public Property Get Task$()
Task = taskID
End Property
Public Property Let Task(val$)
taskID = val
End Property
Public Property Get Name$()
Name = handleID
End Property
Private Sub Class_Terminate()
eventHandle.NotifyThreadClosed Me.Name
Set eventHandle = Nothing
Set workerObject = Nothing
End Sub
