VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeneralInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public HTTPWorks As VBA.Collection
Public Workname As String
Public httpIndex  As Integer
Public HttpWorkCount As Integer
Public Running As Long
Public ErrorRuns As Integer
Public finally As Boolean
Public Failed As Boolean
Public timeTaken As Date

Public items As VBA.Collection
' WorkHTTP required in new class:
Public Property Get WorkHTTP(ByVal Key$) As pppTaskHTTP
  If Key <> vbNullString Then
    On Error Resume Next
    Dim o As Object
    If HTTPWorks Is Nothing Then
      Set HTTPWorks = New VBA.Collection
    Else
      Set o = HTTPWorks(Key)
      VBA.Err.clear
    End If
    If o Is Nothing Then
      HTTPWorks.Add New pppTaskHTTP, Key
    Else
      Set o = Nothing
    End If
    Set WorkHTTP = HTTPWorks(Key)
    WorkHTTP.Name = Key
  End If
End Property

Public Sub Remove(ByVal Key$)
  On Error Resume Next
  HTTPWorks.Remove Key
  On Error GoTo 0
End Sub
Public Sub clearWorks()
  Set HTTPWorks = Nothing
End Sub


Private Sub Class_Initialize()
  ErrorRuns = 2
  Set items = New VBA.Collection
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  Set items = Nothing
  Set HTTPWorks = Nothing
  On Error GoTo 0
End Sub
