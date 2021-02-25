VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pppMultiThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private threadGroup As New VBA.Collection
Private maxThreads As Long
Private minThreads As Long
Private iterableQueue As pppQueue
Private iterableSize As Long
Private WithEvents threadEvents As pppHandleEvents
Attribute threadEvents.VB_VarHelpID = -1
Private workerClass As pppIWorker
Private autoQuitEnabled As Boolean
Private openThreadCount As Long
Private openTaskCount As Long
Private closedTaskCount As Long
Private successfulTaskCount As Long
Private newThreadIndex As Long
Private newTaskIndex As Long
Private taskIDset As VBA.Collection
Private freeThreads As VBA.Collection
Private StartTime As Date
Public Event TaskComplete(CCW As pppTaskHTTP, taskID$, threadID$, timeTaken As Date)
Public Event ThreadOpened(threadCount As Long, threadID$)
Public Event ThreadClosed(threadCount As Long, threadID$)
Public Event complete(timeTaken As Date)
Public Event Closed(timeTaken As Date)
Public Event Opened(StartTime As Date)
Private Type Instruction
    threadID As String
    instructionBody As InstructionType
End Type
Private Enum InstructionType
    mltCloseThread
    mltOpenThread
    mltSetTask
    mltDoNothing
    mltQuit
End Enum
Private Sub Class_Initialize()
  maxThreads = Application.MultiThreadedCalculation.threadCount
  If maxThreads > 4 Then maxThreads = CInt(maxThreads * 2 / 3)
  minThreads = 1
  newThreadIndex = 1
  newTaskIndex = 1
  autoQuitEnabled = True
  Set threadEvents = New pppHandleEvents
  Set taskIDset = New VBA.Collection
  Set freeThreads = New VBA.Collection
  StartTime = VBA.Now
  RaiseEvent Opened(StartTime)
End Sub
Sub Execute()
  If iterableSize = 0 Then
    VBA.Err.Raise 5, Description:="You must set size argument to a non-zero value, or a non-empty iterable first"
  ElseIf workerClass Is Nothing Then
    VBA.Err.Raise 5, Description:="You must set the async class argument first"
  Else
    doInstructions
  End If
End Sub
Sub Quit()
  RaiseEvent complete(VBA.Now - StartTime)
  Set threadEvents = Nothing
End Sub
Private Sub doInstructions(Optional freeThreadID$, Optional loopcount As Long = 1)
  Dim instructionVal As Instruction
  If freeThreadID <> vbNullString Then freeThread = freeThreadID
  instructionVal = getInstruction()
  Select Case instructionVal.instructionBody
  Case InstructionType.mltCloseThread
    closeThread instructionVal.threadID
  Case InstructionType.mltOpenThread
    openThread
  Case InstructionType.mltSetTask
    Dim taskThread As pppThreadHandle
    Set taskThread = threadGroup(instructionVal.threadID)
    assignTaskID (taskThread.Name)
    BusyThread = taskThread.Name
    openTaskCount = openTaskCount + 1
    If Not iterableQueue Is Nothing Then
      taskThread.Execute iterableQueue.DeQueue
    Else
      taskThread.Execute
    End If
  Case InstructionType.mltQuit
    Me.Quit
    instructionVal.instructionBody = mltDoNothing
  Case InstructionType.mltDoNothing
  Case Else
    VBA.Err.Raise 5, , "Error 5: doInstructions"
  End Select
  If instructionVal.instructionBody <> mltDoNothing Then
    Debug.Assert loopcount < maxThreads * 3 + 5
    doInstructions loopcount:=loopcount + 1
  End If
End Sub
Private Function getInstruction() As Instruction
  Dim results As Instruction
  If openThreadCount > openTaskCount Then
    If newTaskIndex > iterableSize Then
      results.instructionBody = mltCloseThread
      results.threadID = freeThread
    ElseIf openThreadCount <= maxThreads Then
      results.instructionBody = mltSetTask
      results.threadID = freeThread
    Else
      results.instructionBody = mltCloseThread
      results.threadID = freeThread
    End If
  Else
    If openThreadCount < maxThreads And newTaskIndex <= iterableSize Then
      results.instructionBody = mltOpenThread
    ElseIf openThreadCount = 0 And autoQuitEnabled Then
      results.instructionBody = mltQuit
    Else
      results.instructionBody = mltDoNothing
    End If
  End If
  getInstruction = results
End Function
Private Sub openThread()
  Dim newThread As New pppThreadHandle
  newThread.OpenHandle Me, threadEvents
  threadGroup.Add newThread, newThread.Name
  openThreadCount = openThreadCount + 1
  freeThread = newThread.Name
End Sub
Private Property Let freeThread(threadID$)
    freeThreads.Add threadID, threadID
End Property
Private Property Get freeThread$()
    freeThread = freeThreads(1)
    freeThreads.Remove (1)
End Property
Private Property Let BusyThread(threadID$)
    On Error Resume Next
    ' Hay biò lôÞi viò trí này: seÞ sõìm khãìc phuòc lôÞi
    freeThreads.Remove threadID
    On Error GoTo 0
End Property
Private Sub assignTaskID(threadID$)
  Dim newThread As pppThreadHandle
  Set newThread = threadGroup(threadID)
  newThread.Task = NewTaskID
  Set newThread.Worker = AsyncClass
End Sub
Private Sub closeThread(threadID$, Optional Failed As Boolean = False)
  Dim oldThread As pppThreadHandle
  Set oldThread = threadGroup(threadID)
  threadGroup.Remove oldThread.Name
  BusyThread = oldThread.Name
  Set oldThread = Nothing
  openThreadCount = openThreadCount - 1
End Sub
Public Property Get Size() As Variant
    Size = iterableSize
End Property
Sub IncreaseSize(ByVal QueueK, Optional ByVal bExecute As Boolean = False)
  Set Me.AsyncClass = New pppHTTPWorker
  If threadEvents Is Nothing Then Set threadEvents = New pppHandleEvents
  If iterableQueue Is Nothing Then Set iterableQueue = New pppQueue
  If TypeOf QueueK Is pppQueue Then
    If QueueK.Count Then
      iterableSize = iterableSize + QueueK.Count
      While QueueK.Count > 0
        iterableQueue.EnQueue QueueK.DeQueue()
      Wend
    End If
  Else
    iterableSize = iterableSize + 1
    iterableQueue.EnQueue QueueK
  End If
  If bExecute Then Me.Execute
End Sub
Public Property Set AsyncClass(ByVal workObj As pppIWorker)
    Set workerClass = workObj
End Property
Public Property Get AsyncClass() As pppIWorker
    Set AsyncClass = workerClass
End Property
Public Property Let autoQuit(ByVal value As Boolean)
    autoQuitEnabled = value
End Property
Public Property Get NewHandleID$()
    NewHandleID = "Handle " & newThreadIndex
    newThreadIndex = newThreadIndex + 1
End Property
Private Property Get NewTaskID$()
    NewTaskID = "Task " & newTaskIndex
    taskIDset.Add newTaskIndex, NewTaskID
    newTaskIndex = newTaskIndex + 1
End Property
Private Sub Class_Terminate()
    RaiseEvent Closed(VBA.Now - StartTime)
End Sub
Private Sub threadEvents_Closed(threadID$)
    RaiseEvent ThreadClosed(openThreadCount, threadID)
End Sub
Private Sub threadEvents_Opened(threadID$)
    RaiseEvent ThreadOpened(openThreadCount, threadID)
End Sub
Private Sub threadEvents_Complete(obj As pppThreadHandle, CCW As pppTaskHTTP)
    RaiseEvent TaskComplete(CCW, obj.Task, obj.Name, VBA.Now - StartTime)
    openTaskCount = openTaskCount - 1
    closedTaskCount = closedTaskCount + 1
    successfulTaskCount = successfulTaskCount + 1
    doInstructions obj.Name
End Sub
