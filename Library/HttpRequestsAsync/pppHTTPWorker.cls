VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pppHTTPWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private XMLHTTP As Object
Private WithEvents XMLHttpRequest As WinHttp.WinHttpRequest
Attribute XMLHttpRequest.VB_VarHelpID = -1
Private CW As pppTaskHTTP
Private FailedRetries%, iRetry%
Implements pppIWorker
Private Type TWorker
    Events As pppIWorkerEvents
End Type
Private this As TWorker
Private Function pppIWorker_CreateNew() As pppIWorker
    Set pppIWorker_CreateNew = New pppHTTPWorker
End Function
Private Property Set pppIWorker_Events(RHS As pppIWorkerEvents)
    Set this.Events = RHS
End Property
Private Sub pppIWorker_Execute(Optional Args As pppTaskHTTP)
  On Error Resume Next
  Started
  Set CW = Args
  FailedRetries = CW.ErrorRuns
  If FailedRetries = 0 Then FailedRetries = 1
  Call SetHTTP
End Sub
Private Sub Started()
  If Not this.Events Is Nothing Then
    this.Events.Started
  End If
End Sub
Private Sub statusChange(ByVal statusText$)
  If Not this.Events Is Nothing Then
    this.Events.statusChange statusText
  End If
End Sub
Private Sub complete(Optional ByVal work As pppTaskHTTP)
  Set XMLHTTP = Nothing
  Set XMLHttpRequest = Nothing
  If Not this.Events Is Nothing Then
    this.Events.complete work
  End If
End Sub
Private Sub SetHTTP()
  Select Case VBA.typename(CW.http)
  Case "IServerXMLHTTPRequest2"
    Set XMLHTTP = CW.http
    XMLHTTP.OnReadyStateChange = Me
    XMLHTTP.Send CW.Api
  Case "WinHttpRequest"
    Set XMLHttpRequest = CW.http
    XMLHttpRequest.Send CW.Api
  Case Else
    
  End Select
End Sub

Sub OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0
'ri:Attribute OnReadyStateChange.VB_UserMemId = 0
  With XMLHTTP
    statusChange .statusText
    If .readyState = 4 Then
      CW.Status = .Status: complete CW
    End If
  End With
End Sub
'/////////////////////////////////////////////////////////////////
'Private Sub XMLHttpRequest_OnResponseStart(ByVal Status As Long, ByVal ContentType$)
'    CW.Status = XMLHttpRequest.Status: complete CW
'End Sub

'Private Sub XMLHttpRequest_OnResponseDataAvailable(data() As Byte)
'    CW.Status = XMLHttpRequest.Status: complete CW
'End Sub

Private Sub XMLHttpRequest_OnResponseFinished()
  CW.Status = XMLHttpRequest.Status: complete CW
End Sub
Private Sub XMLHttpRequest_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription$)
  CW.Status = XMLHttpRequest.Status: complete CW
End Sub
'/////////////////////////////////////////////////////////////////


Private Sub Class_Terminate()
  Set XMLHTTP = Nothing
  Set XMLHttpRequest = Nothing
  Set CW = Nothing
End Sub
