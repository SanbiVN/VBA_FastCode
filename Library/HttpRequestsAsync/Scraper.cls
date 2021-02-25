VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Scraper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private WithEvents HttpAsync As pppMultiThread
Attribute HttpAsync.VB_VarHelpID = -1
Private Works As VBA.Collection
Private RE As Object, HF As Object, RRR As Range

Private Sub Class_Initialize()
  Set RE = glbRegex
  Set HF = glbHtmlfile
End Sub
' Kêìt thúc lõìp
Private Sub Class_Terminate()
  Set RE = Nothing
  Set HttpAsync = Nothing
  Set Works = Nothing
'
End Sub

Private Sub Init()
  If HttpAsync Is Nothing Then
    Set HttpAsync = New pppMultiThread
  End If
  If Works Is Nothing Then
    Set Works = New VBA.Collection
  End If
End Sub

Private Sub HttpAsync_TaskComplete(CCW As pppTaskHTTP, taskID$, threadID$, timeTaken As Date)
  
  On Error Resume Next
  Dim work As GeneralInfo
  Set work = Works(CCW.Workname)
  RE.Pattern = "^(\d+)\."
  Dim M, res$, cls, i%, r&
  res = CCW.http.ResponseText
  HF.body.innerhtml = res
  Set cls = GetClass("entry-content", "div", HF)
  If IsObject(cls) Then
    Set cls = cls.childNodes
    For i = 0 To cls.Length - 1
      res = ""
      If Not cls(i).getElementsByTagName("strong") Is Nothing Then
        res = cls(i).innerText
        VBA.Err.clear
        Set M = RE.Execute(res)
        If M.Count And VBA.Err.Number = 0 Then
          VBA.Err.clear
          r = CLng(M(0).submatches(0))
          If VBA.Err.Number = 0 Then
            If Not cls(i + 1).getElementsByTagName("img") Is Nothing Then
              RRR(r, 1).value = r
              RRR(r, 2).value = cls(i + 1).getElementsByTagName("img")(0).src
            End If
          End If
        End If
      End If
    Next
  End If
  work.httpIndex = work.httpIndex + 1
  Debug.Print "Task:"; work.httpIndex, work.HttpWorkCount
  If work.httpIndex >= work.HttpWorkCount Then
    'Call Transpose(total, wsNghiaCumTu.Range("G4"))
    'wsNghiaCumTu.Range("G4").Resize(UBound(total, 2)) = Application.Transpose(total)
  End If
  work.Remove CCW.Name
End Sub

Public Sub scrap(ByVal url$, ByVal taskname$, Optional ByVal Workname$ = "Download Picture")
  Dim work As GeneralInfo
  Call Init
  Set work = newWork(Workname, Works)
  With work
    With .WorkHTTP(Workname)
      .url = url
      Set .work = Me
      .Name = url
      .Workname = Workname
      .ErrorRuns = 2
      Set .http = XMLHTTP
      .Method = "GET"
      .http.Open .Method, .url, True
      .http.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    End With
    Call HttpAsync.IncreaseSize(.WorkHTTP(Workname), True)
  End With
End Sub

Sub getByList()
  Dim r As Range, i%, Workname$, taskname$
  Dim work As GeneralInfo
  Call Init
  Workname = "getByList"
  Set work = newWork(Workname, Works)
  Set r = wsNghiaCumTu.Range("E4")
  Set RRR = r(1, 3)
  work.HttpWorkCount = 26
  work.httpIndex = 0

  For i = 1 To work.HttpWorkCount
    taskname = r(i, 1).value
    With work.WorkHTTP(taskname)
      .url = r(i, 2).value
      Set .work = Me
      .Name = taskname
      .Workname = Workname
      .ErrorRuns = 2
      Set .http = XMLHTTP
      .Method = "GET"
      .http.Open .Method, .url, True
      .http.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    End With
    Call HttpAsync.IncreaseSize(work.WorkHTTP(taskname), True)
  Next
  Set r = Nothing
  Set work = Nothing
End Sub

