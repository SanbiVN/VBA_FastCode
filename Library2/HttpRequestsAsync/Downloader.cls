VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Downloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private WithEvents HttpAsync As pppMultiThread
Attribute HttpAsync.VB_VarHelpID = -1
Private Works As VBA.Collection
Private RE As Object
Private Sub Class_Initialize()
  Set RE = glbRegex
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
  Dim i As Integer
  Dim j As Long
  Dim File As String
  Dim Ext As String
  Dim FileName As String
  Dim arrBytes() As Byte
  Dim bytes() As Byte
  CreateFolder CCW.Path
  Ext = Split(CCW.http.GetResponseHeader("Content-Type"), "/")(1)

  Select Case Ext
  Case "jpeg":  CCW.Extend = ".jpg"
  Case "mpeg": CCW.Extend = ".mp3"
  Case "svg+xml", "svg+xml": CCW.Extend = ".svg"
  End Select
  File = CCW.Path & CCW.FileName & CCW.Extend
  i = FreeFile
  Open File For Binary As #i: Put #i, , CCW.http.ResponseBody
  ReDim bytes(LOF(i) - 1)
  Close i
  Open File For Binary As #i: Get #i, 13, bytes: Close i
  Open File For Binary As #i: Put #i, 1, bytes
  Close i
  If CCW.SheetShowImage <> "" Then
    Select Case CCW.Extend
    Case ".jpg", ".jpeg", ".png"
      ImageShowAfterDownload = File
      TextFontImageAfterDownload = CCW.TextFontImage
      TitleFontImageAfterDownload = CCW.TitleFontImage
      showImage
    End Select
  End If
End Sub

Private Sub showImage()
  Application.OnTime VBA.Now(), "'" & ThisWorkbook.Name & "'!" & ShowImageAfterDownload
End Sub


Public Sub Save( _
         ByVal url As String, _
         ByVal Path As String, _
Optional ByVal DefaultFileName As String = "", _
Optional ByVal DefaultExtend As String = "", _
Optional ByVal WriteExists As Boolean, _
Optional ByVal SheetShowImage As String, _
Optional ByVal TextFontImage As String, _
Optional ByVal TitleFontImage As String, _
Optional ByVal Workname$ = "Download Picture")
          
  Call Init
  
  Path = Path & IIf(Path Like "*[\/]", "", Application.PathSeparator)
  Dim FileName As String, Extend As String, Matchs
  Dim k As Long, work As GeneralInfo
  
  RE.Pattern = ".+\/{2}.+\/{1}(.+)(\.\w+)\?*.*"

  Set Matchs = RE.Execute(url)
  If Matchs.Count Then
    FileName = DecodeURL(Matchs(0).submatches(0))
    Extend = Matchs(0).submatches(1)
  End If

  If DefaultFileName <> "" Then
    FileName = DefaultFileName
  End If
  If DefaultExtend <> "" Then
    Extend = DefaultExtend
  End If
  
  If Len(Extend) > 0 Then
    If LEFT(Extend, 1) <> "." Then
      Extend = "." & Extend
    End If
  End If

'      Do While FileName = ""
'        FileName = DefaultFileName: k = k + 1
'        FileName = VBA.Dir(path & FileName & Extend, vbSystem)
'        If k > 10 Then
'          Exit Sub
'        End If
'      Loop

  If VBA.Dir(Path & FileName & Extend, vbSystem) <> "" Then
    If Not WriteExists Then
      If SheetShowImage <> "" Then
        Select Case Extend
        Case ".jpg", ".jpeg", ".png"
          ImageShowAfterDownload = Path & FileName & Extend
          TextFontImageAfterDownload = TextFontImage
          TitleFontImageAfterDownload = TitleFontImage
          showImage
        End Select
      End If
      Exit Sub
    End If
  End If


  Set work = newWork(Workname, Works)
  With work
    With .WorkHTTP(Workname)
        .url = url
      Set .work = Me
      .Name = Path & FileName & Extend
      .Workname = Workname
      .ErrorRuns = 2
      
      .Path = Path
      .FileName = FileName
      .Extend = Extend
      
      .SheetShowImage = SheetShowImage
      .TextFontImage = TextFontImage
      .TitleFontImage = TitleFontImage
      Set .http = XMLHTTP
      .Method = "GET"
      .http.Open .Method, .url, True
      .http.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    End With
    Call HttpAsync.IncreaseSize(.WorkHTTP(Workname), True)
  End With
End Sub

Private Function CreateFolder(ByVal FolderPath As String) As Boolean
  Dim FolderArray, tmp$, i As Integer, UB As Integer, tFolder$
  tFolder = FolderPath
  If VBA.RIGHT(tFolder, 1) = "\" Then tFolder = LEFT(tFolder, VBA.Len(tFolder) - 1)
  If tFolder Like "\\*\*" Then tFolder = Replace(tFolder, "\", "@", 1, 3)
  FolderArray = Split(tFolder, "\")
  FolderArray(0) = Replace(FolderArray(0), "@", "\", 1, 3)
  On Error GoTo Ends
  UB = UBound(FolderArray)
  For i = 0 To UB
    tmp = tmp & FolderArray(i) & "\"
    If VBA.Dir(tmp, vbDirectory) = "" Then
      VBA.MkDir (tmp)
    End If
    CreateFolder = (i = UB) And Len(FolderArray(i)) > 0 And FolderArray(i) <> " "
  Next
Ends:
End Function


Public Function BlobToFile(strFile As String, ByRef Field As Object) As Long
  On Error GoTo BlobToFileError
  Dim nFileNum As Integer
  Dim abytData() As Byte
  BlobToFile = 0
  nFileNum = FreeFile
  Open strFile For Binary Access Write As nFileNum
  abytData = Field
  Put #nFileNum, , abytData
  BlobToFile = LOF(nFileNum)
BlobToFileExit:
  If nFileNum > 0 Then Close nFileNum
  Exit Function
BlobToFileError:
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, _
  "Error writing file in BlobToFile"
  BlobToFile = 0
  Resume BlobToFileExit
End Function
'Function:  FileToBlob - Loads a file into a binary field.
'Parameter: strFile - Full path and filename of the source file.
'Parameter: Field - The binary field into which the file is to be loaded.
Public Function FileToBlob(strFile As String, ByRef Field As Object)
  On Error GoTo FileToBlobError
  If Len(Dir(strFile)) > 0 Then
    Dim nFileNum As Integer
    Dim byteData() As Byte
    nFileNum = FreeFile()
    Open strFile For Binary Access Read As nFileNum
    If LOF(nFileNum) > 0 Then
      ReDim byteData(1 To LOF(nFileNum))
      Get #nFileNum, , byteData
      Field = byteData
    End If
  Else
    MsgBox "Error: File not found", vbCritical, _
    "Error reading file in FileToBlob"
  End If
FileToBlobExit:
  If nFileNum > 0 Then Close nFileNum
  Exit Function
FileToBlobError:
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, _
  "Error reading file in FileToBlob"
  Resume FileToBlobExit
End Function

Private Function DecodeURL(text As Variant, Optional Encode = True)
  Static o As Object
  If o Is Nothing Then
    Set o = VBA.CreateObject("htmlfile")
    o.parentWindow.execScript "function decodeURI(s) {return decodeURIComponent(s)}", "javascript"
  End If
  If Encode Then
    DecodeURL = o.parentWindow.decodeURI(text)
  End If
End Function
