Attribute VB_Name = "z_Chrome"
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
#Else
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
#End If
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10





Public Const SW_FORCEMINIMIZE = 11
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Private Type STARTUPINFO
   CB As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Const SYNCHRONIZE = 1048576
Const NORMAL_PRIORITY_CLASS = &H20&
Const STARTF_USEPOSITION = &H4&
Const STARTF_USESIZE = &H2&
  
Private Const S_OK = &H0
Private Const S_FALSE = &H1
Private Const E_INVALIDARG = &H80070057
Private Const SHGFP_TYPE_CURRENT = 0
Private Const SHGFP_TYPE_DEFAULT = 1
Private Const GW_HWNDNEXT = 2
Private Const GA_PARENT = 1


Public retainedChromeHwnd As Long, ChildHwnd As Long, ChildFound As Boolean, origChildFound As Boolean
Public NextHandle As Boolean, GotNextParent As Boolean


#If VBA7 Then
  Public Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Boolean
  Public Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As LongPtr) As Long
   Public Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
  ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
  Public Declare PtrSafe Function GetWindowTextW Lib "user32.dll" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal cch As LongPtr) As Long
    Public Declare PtrSafe Function GetWindowPlacement Lib _
            "user32" (ByVal hWnd As LongPtr, _
            ByRef lpwndpl As WINDOWPLACEMENT) As Integer
    Public Declare PtrSafe Function SetWindowPlacement Lib "user32" _
           (ByVal hWnd As LongPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Integer
    
  Public Declare PtrSafe Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal lpszPath As String) As Long
  Public Declare PtrSafe Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As LongPtr) As Long
  Public Declare PtrSafe Function ShowWindow Lib "user32.dll" (ByVal hWnd As LongPtr, ByVal lCmdShow As Long) As Boolean
  Public Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hWnd As LongPtr, ByVal flags As Long) As Long
  Public Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
  Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
  Public Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal wCmd As Long) As Long
  Public Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, lpdwProcessId As Long) As Long
  Public Declare PtrSafe Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As LongPtr, ByVal wFlag As Long) As Long
  Public Declare PtrSafe Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
  Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long
  Public Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As LongPtr) As Long
  Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Public Declare PtrSafe Function MoveWindow Lib "user32.dll" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
  Public Declare PtrSafe Function CreateProcess Lib "kernel32" _
            Alias "CreateProcessA" ( _
             ByVal lpApplicationName As String, _
             ByVal lpCommandLine As String, _
                   lpProcessAttributes As Any, _
                   lpThreadAttributes As Any, _
             ByVal bInheritHandles As Long, _
             ByVal dwCreationFlags As Long, _
                   lpEnvironment As Any, _
             ByVal lpCurrentDriectory As String, _
                   lpStartupInfo As STARTUPINFO, _
                   lpProcessInformation As PROCESS_INFORMATION) As Long

  Public Declare PtrSafe Function OpenProcess Lib "kernel32.dll" _
     (ByVal dwAccess As Long, _
      ByVal fInherit As Integer, _
      ByVal hObject As Long) As Long

  Public Declare PtrSafe Function TerminateProcess Lib "kernel32" _
     (ByVal hProcess As Long, _
     ByVal uExitCode As Long) As Long

  Public Declare PtrSafe Function CloseHandle Lib "kernel32" _
     (ByVal hObject As Long) As Long
  Public Declare PtrSafe Function SendMessage _
                           Lib "user32" _
                             Alias "SendMessageA" ( _
                               ByVal hWnd As LongPtr, _
                               ByVal wMsg As Long, _
                               ByVal wParam As Long, _
                               ByRef lParam As Any) As Long
  Public Declare PtrSafe Function getTickCount Lib "kernel32" Alias "GetTickCount" () As Long
  Public Declare PtrSafe Function AlertTimeOut Lib "user32" Alias "MessageBoxTimeoutA" ( _
    ByVal hWnd As LongPtr, ByVal lpText$, ByVal lpCaption$, _
    ByVal wType As VbMsgBoxStyle, ByVal wlange As Long, ByVal dwTimeout As Long) As Long
  
#Else
  Public Declare Function AlertTimeOut Lib "user32" Alias "MessageBoxTimeoutA" ( _
    ByVal hwnd As Long, ByVal lpText$, ByVal lpCaption$, _
    ByVal wType As VbMsgBoxStyle, ByVal wlange As Long, ByVal dwTimeout As Long) As Long
  
  Public Declare Function GetTickCount Lib "kernel32" () As Long
  Public Declare Function SendMessage _
            Lib "user32" _
              Alias "SendMessageA" ( _
                ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long
Public Declare Function GetWindowTextW Lib "user32.dll" (ByVal hwnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
  
    Public  Declare Function GetWindowPlacement Lib _
            "user32" (ByVal hwnd As Long, _
            ByRef lpwndpl As WINDOWPLACEMENT) As Integer
    Public  Declare Function SetWindowPlacement Lib "user32" _
           (ByVal hwnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Integer
    
  Public  Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal lpszPath As String) As Long
  Public  Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
  Public  Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal lCmdShow As Long) As Boolean
  Public  Declare Function GetAncestor Lib "user32" (ByVal hwnd As Long, ByVal flags As Long) As Long
  Public  Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
  Public  Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
  Public  Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
  Public  Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
  Public  Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
  Public  Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
  Public  Declare Function GetForegroundWindow Lib "user32" () As Long
  Public  Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
  Public  Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Public  Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public  Declare Function CreateProcess Lib "kernel32" _
     Alias "CreateProcessA" _
     (ByVal lpApplicationName As String, _
     ByVal lpCommandLine As String, _
     lpProcessAttributes As Any, _
     lpThreadAttributes As Any, _
     ByVal bInheritHandles As Long, _
     ByVal dwCreationFlags As Long, _
     lpEnvironment As Any, _
     ByVal lpCurrentDriectory As String, _
     lpStartupInfo As STARTUPINFO, _
     lpProcessInformation As PROCESS_INFORMATION) As Long
  Public  Declare Function OpenProcess Lib "kernel32.dll" _
     (ByVal dwAccess As Long, _
     ByVal fInherit As Integer, _
     ByVal hObject As Long) As Long

  Public  Declare Function TerminateProcess Lib "kernel32" _
     (ByVal hProcess As Long, _
     ByVal uExitCode As Long) As Long

  Public  Declare Function CloseHandle Lib "kernel32" _
     (ByVal hObject As Long) As Long
  Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Boolean
  Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
   Public  Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
  ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

#If VBA7 Then
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As Long) As Long
Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As LongPtr) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare PtrSafe Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#Else
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If



Const HWND_BROADCAST = &HFFFF
Const SC_MONITORPOWER = &HF170
Const WM_SYSCOMMAND = &H1125

#If Win64 Then
  Public hwnd_chrome  As LongPtr
#Else
  Public hwnd_chrome As Long
#End If

Private Type POINTAPI
   X As Integer
   Y As Integer
End Type

Public Type WINDOWPLACEMENT
     Length As Integer
     flags As Integer
     showCmd As Integer
     ptMinPosition As POINTAPI
     ptMaxPosition As POINTAPI
     rcNormalPosition As RECT
End Type

Public oChromeDriver As Object
Public SeleDriver As Object 'Selenium.WebDriver
Public SeleKeys 'As New Selenium.Keys
Public SeleBy 'As New Selenium.By

Public Const Si_Chrome$ = "https://www.google.com/chrome/"
Public Const Si_SeleniumBasic$ = "https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0"
Public Const Si_ChromeDriver$ = "http://chromedriver.chromium.org/downloads"

Private GoNextParent As Boolean

Private Sub test_SEConnectChrome()
  
  Dim ti#, hWnd&: ti = Timer
  
  Set SeleDriver = Nothing
  Debug.Print SEConnectChrome(SeleDriver, _
  url:="https://translate.google.com/?hl=vi#view=home&op=translate&sl=en&tl=vi", _
  maximize:=True, checkUrl:=True, startIfExists:=True, _
  visible:=5)
  SeleDriver.Window.maximize
  SEChromeClose SeleDriver
  'wVisible_App SeleDriver.Title & "", , 9
End Sub


Function UserDataDirForGame()
  UserDataDirForGame = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp")) & "\remote-profile-cr"
End Function


Function SEConnectChrome( _
              Optional ByRef Driver As Object, _
              Optional ByRef boolStart As Boolean, _
              Optional ByVal url As String, _
              Optional ByVal indexPage As Integer = 1, _
              Optional ByRef boolWait As Boolean, _
              Optional ByRef checkUrl As Boolean, _
              Optional ByVal startIfExists As Boolean, _
              Optional ByVal refreshIfExists As Boolean, _
              Optional ByVal startNewTab As Boolean, _
              Optional ByVal boolApp As Boolean, _
              Optional ByVal popupBlocking As Boolean, _
              Optional ByVal visible = vbNormalFocus, _
              Optional ByVal maximize As Boolean, _
              Optional ByVal position As String = "0,0", _
              Optional ByVal screen As String = "1900,1053", _
              Optional ByVal disableGPU As Boolean = False, _
              Optional ByVal boolClose As Boolean, _
              Optional ByVal chromePath As String, _
              Optional ByVal Port As Long = 9222, _
              Optional ByVal UserDataDir As String, _
              Optional ByVal PrivateMode As Boolean) As Boolean
  DoEvents 'Open
  If chromePath = "" Then
    chromePath = getChromePath
  End If
  Dim Win, Process, isBrowserOpen As Boolean, isOpen As Boolean, k%, i%
  If Port <= 0 Then Port = 9222
  If UserDataDir = vbNullString Then
    UserDataDir = UserDataDirForGame
  End If
  GoSub CheckCR: isOpen = isBrowserOpen
  If Not isOpen Then
    If Not boolStart And Not startIfExists Then GoTo Ends
    Dim CmdLn$
    CmdLn = _
      IIf(Port > 0, " --remote-debugging-port=" & Port, "") & _
      IIf(UserDataDir <> vbNullString, " --user-data-dir=""" & UserDataDir & """", "") & _
      " --lang=en" & _
      IIf(url = "", "", IIf(boolApp, " --app=", " ")) & url & _
      IIf(maximize And visible <> 0, " --start-maximized", "") & _
      IIf(position <> vbNullString And Not maximize, " --window-position=" & position, "") & _
      IIf(screen <> vbNullString And Not maximize, " --window-size=" & screen, "") & _
      IIf(disableGPU, " --disable-gpu", "") & _
      IIf(PrivateMode, " --incognito", "") & _
      IIf(popupBlocking, "", " --disable-popup-blocking") & " --disable-sync"
    Shell chromePath & "" & CmdLn, vbNormalFocus
    Do Until isBrowserOpen:
      GoSub CheckCR:
      Delay 500
      k = k + 1: If k > 12 Then GoTo Ends
    Loop
  End If
  If Driver Is Nothing Then
    'Set Driver = New selenium.ChromeDriver
    Set Driver = VBA.CreateObject("selenium.ChromeDriver")
    Driver.SetCapability "debuggerAddress", "127.0.0.1:" & Port
    Driver.Start "chrome"
'    Driver.Timeouts.ImplicitWait = 5000
'    Driver.Timeouts.PageLoad = 5000
'    Driver.Timeouts.Server = 10000
    GoSub FindChrDrv
  End If
  GoSub checkUrl
Ends:
SEConnectChrome = isBrowserOpen
Set Process = Nothing: Set Win = Nothing
Exit Function
checkUrl:
  If Not checkUrl Then Return
  k = 0
  On Error Resume Next
  checkUrl = False
  GoSub compareUrl
  For Each Win In Driver.Windows
    If Err.Number <> 0 Then Return
    Win.Activate
    GoSub compareUrl
  Next
nextCheckUrl:
  If Not checkUrl Or (startIfExists And checkUrl) Then
    If startNewTab Then
      Driver.ExecuteScript "window.open(arguments[0], '_blank');", url
    Else
      Driver.Get url
    End If
  End If
  On Error GoTo 0
Return
compareUrl:
  If LCase$(Driver.url) Like "*" & LCase$(url) & "*" Then
    k = k + 1: If indexPage = k Then checkUrl = True: GoTo nextCheckUrl
  End If
Return
FindChrDrv:
  For Each Process In VBA.GetObject("winmgmts:\\.\root\CIMV2") _
    .ExecQuery("SELECT * FROM Win32_Process WHERE Name = ""chromedriver.exe""", , 48)
    Set oChromeDriver = Process: Exit For
  Next
Return
CheckCR:
  For Each Process In VBA.GetObject("winmgmts:\\.\root\CIMV2") _
    .ExecQuery("SELECT * FROM Win32_Process WHERE Name = ""chrome.exe""", , 48)
    If LCase$(Process.commandLine) Like _
       LCase$("*chrome*--remote-debugging-port=" & _
        Port & "*--user-data-dir=*" & LCase$(UserDataDir) & "*") Then
      isBrowserOpen = True: If boolClose Then Process.Terminate: Set Driver = Nothing: GoTo Ends
      Exit For
    End If
  Next
Return
End Function

Sub btn_UpdateChromedriver()
  Debug.Print UpdateChromedriver
  'Shell "explorer.exe """ & Environ$("USERPROFILE") & "\AppData\Local\SeleniumBasic" & """", vbNormalFocus
End Sub
Function UpdateChromedriver(Optional ByVal chromePath As String = "") As Boolean
  If chromePath = "" Then
    chromePath = getChromePath
  End If
  On Error Resume Next
  Dim LastedUpdate As String
  Dim FSO As Object, SEPath$
  Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
  Dim XMLHTTP As Object
  Dim a, Tmp1$, tmp$, eURL$, temp$, info$
  Const LATEST_RELEASE = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
  Const url$ = "https://chromedriver.storage.googleapis.com/"

  Const EXE$ = "\chromedriver.exe"
  Const ZIP$ = "\chromedriver_win32.zip"
  SEPath$ = Environ$("USERPROFILE") & "\AppData\Local\SeleniumBasic"
  temp = Environ("TEMP"): GoSub DelTemp
  If Not FSO.FileExists(chromePath) Then Exit Function
  info = FSO.GetFileVersion(chromePath)
  eURL = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & Split(info, ".")(0)
  GoSub http
  LastedUpdate = VBA.GetSetting("Chromedriver", "Update", "Last")
  If LastedUpdate < tmp Then
    GoSub Download
  Else
    UpdateChromedriver = True
  End If
Ends: Set FSO = Nothing
Exit Function
Download:
On Error Resume Next
  eURL = url$ & tmp & "/chromedriver_win32.zip"
  If URLDownloadToFile(0, eURL, temp & ZIP, 0, 0) = 0 Then
    GoSub Extract
    Call VBA.SaveSetting("Chromedriver", "Update", "Last", tmp)
  End If
On Error GoTo 0
Return
Extract:
On Error Resume Next
With VBA.CreateObject("Shell.Application")
  .Namespace(temp & "\").CopyHere .Namespace(temp & ZIP).items
End With
With FSO
  If .FileExists(temp & EXE) Then
    If .FolderExists(SEPath) Then FSO.CopyFile temp & EXE, SEPath & EXE, True
  End If
  UpdateChromedriver = Err.Number = 0
End With
On Error GoTo 0
GoSub DelTemp
Return

DelTemp:
On Error Resume Next
  FSO.DeleteFile temp & ZIP
  FSO.DeleteFile temp & EXE
On Error GoTo 0
Return
http:
Set XMLHTTP = VBA.CreateObject("MSXML2.XMLHTTP.6.0")
With XMLHTTP
  .Open "GET", eURL, False
  .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36"
  .setRequestHeader "Content-type", "application/x-www-form-urlencoded"
  .Send
  tmp = VBA.Trim(Application.Clean(.ResponseText))
End With
Return
End Function



Function SwitchIFrame(ByVal IFrame, _
                      ByVal Driver As Object, _
             Optional ByVal Timeout% = -1, _
             Optional ByVal hRaise As Boolean = True) As Boolean
  On Error Resume Next
  Driver.SwitchToFrame IFrame, Timeout, hRaise
  SwitchIFrame = Err.Number = 0
  On Error GoTo 0
End Function


Sub SEChromeClose(ByRef Driver As Object)
  Dim i, p
  On Error Resume Next
  Debug.Print Driver.Windows.Count
  For i = Driver.Windows.Count To 1 Step -1
    Driver.Windows(i).Close
  Next:
  Set Driver = Nothing
  For Each p In VBA.GetObject("winmgmts:\\.\root\CIMV2") _
    .ExecQuery("SELECT * FROM Win32_Process WHERE Name = ""chromedriver.exe""", , 48)
    p.Terminate
  Next
  On Error GoTo 0
End Sub


Sub SECloseTabByURL(ByRef Driver As Object, ByVal url$, Optional ByVal index% = 1)
  If Driver Is Nothing Then Exit Sub
  Dim Win, i%, k%
  On Error Resume Next
  For i = 1 To Driver.Windows.Count
    Set Win = Driver.Windows(i)
    If LCase(Driver.url) Like LCase(url) Then
      k = k + 1
      If index = k Then
        Win.Close: Exit For
      End If
    End If
  Next
  focusChrome
  Set Win = Nothing
  Set Driver = Nothing
  On Error GoTo 0
End Sub




Function HTMLNode(ByVal obj As Object, _
              ParamArray ArrayNodes()) As Object
  DoEvents 'Open
  Dim aNode
  On Error Resume Next
  For Each aNode In ArrayNodes
    If obj.childNodes(aNode) Is Nothing Then Exit Function
    Set obj = obj.childNodes(aNode)
    If Err.Number <> 0 Then Exit Function
  Next
  Set HTMLNode = obj
End Function
  



Public Function getNameFromHwnd$(hWnd&)
  Dim Title As String * 255
  Dim tLen&
  tLen = GetWindowTextLength(hWnd)
  GetWindowTextW hWnd, Title, 255
  getNameFromHwnd = LEFT(Title, tLen)
End Function



Private Sub test_FindWindow()

  Debug.Print GetForegroundWindow()


End Sub

 Private Sub FormsdsClick()
     Dim pInfo As PROCESS_INFORMATION
     Dim sInfo As STARTUPINFO
     Dim sNull As String
     Dim lSuccess As Long
     Dim lRetValue As Long

     With sInfo
        .CB = Len(sInfo)
        'set flag to tell that size and position info _
         has been set in structure
        .dwFlags = STARTF_USESIZE + STARTF_USEPOSITION
        ' set position and size
        .dwX = 200
        .dwY = 100
        .dwXSize = 800
        .dwYSize = 400
    End With
     lSuccess = CreateProcess(sNull, _
                             "C:\Program Files\7-Zip\7zFM.exe", _
                             ByVal 0&, _
                             ByVal 0&, _
                             1&, _
                             NORMAL_PRIORITY_CLASS, _
                             ByVal 0&, _
                             sNull, _
                             sInfo, _
                             pInfo)
     If lSuccess = False Then MsgBox "launch failed"
     lRetValue = CloseHandle(pInfo.hThread)
     lRetValue = CloseHandle(pInfo.hProcess)

  End Sub



#If VBA7 Then
Public Function GetHWND_App(Optional ByRef Title$, _
                            Optional ByRef Class$) As LongPtr
#Else
Public Function GetHWND_App(Optional ByRef Title$, _
                            Optional ByRef Class$) As Long
#End If
  #If Win64 Then
    Dim hWndThis As LongPtr
  #Else
    Dim hWndThis As Long
  #End If
  Title = LCase$(Title)
  Class = LCase$(Class)

  Dim sTitle As String, sClass As String
  hWndThis = FindWindow(vbNullString, vbNullString)
  
  While hWndThis
    If Title <> vbNullString Then
      sTitle = Space$(260)
      sTitle = LEFT$(sTitle, GetWindowTextW(hWndThis, StrPtr(sTitle), 260))
      sTitle = LCase$(sTitle)
    End If
    If Class <> vbNullString Then
      sClass = Space$(100)
      sClass = LEFT$(sClass, GetClassName(hWndThis, sClass, 100))
      sClass = LCase$(sClass)
    End If
    
    If Title <> vbNullString And Class <> vbNullString Then
      If InStr(1, sTitle, Title, vbTextCompare) And _
         sClass Like Class Then
        GetHWND_App = hWndThis: Exit Function
      End If
    ElseIf Title <> vbNullString And Class = vbNullString Then
      
      If InStr(1, sTitle, Title, vbTextCompare) Then
        GetHWND_App = hWndThis: Exit Function
      End If
    ElseIf Title = vbNullString And Class <> vbNullString Then
      If sClass Like Class Then
        GetHWND_App = hWndThis: Exit Function
      End If
    End If
    hWndThis = GetWindow(hWndThis, GW_HWNDNEXT)
  Wend
End Function
Sub wVisible_App(Optional Title$, Optional Class$, Optional ByVal wVisible = 9)
  #If Win64 Then
    Dim hWndThis As LongPtr
  #Else
    Dim hWndThis As Long
  #End If
  hWndThis = GetHWND_App(Title, Class)
  If hWndThis <> 0 Then
      ShowWindow hWndThis, wVisible
  End If
End Sub
Sub btnImmediateShow()
  On Error Resume Next
  Application.VBE.Windows("Immediate").visible = Not Application.VBE.Windows("Immediate").visible
End Sub
'=======================================
'       Clear Immediate Window
'=======================================
Sub btnImmediateWindowClear_call()
  On Error Resume Next
  VBA.CreateObject("WScript.Shell").SendKeys "^g^a{BS}", True
  Application.VBE.ActiveCodePane.Window.SetFocus
End Sub
Sub btnImmediateWindowClear()
  On Error GoTo Ends
  'VBA.CreateObject ("WScript.Shell")
  VBA.CreateObject("WScript.Shell").SendKeys "^g^a{BS}", True
Ends:
  'Application.OnTime VBA.Now(), "'" & Application.ThisWorkbook.Name & "'!LinkedImmediateWindow"
End Sub
Sub LinkedImmediateWindow()
  On Error Resume Next
  Dim Immediate
  Set Immediate = Application.VBE.Windows("Immediate")
  Immediate.visible = True
  With Application.VBE
    .MainWindow.LinkedWindows.Remove Immediate
    With Immediate
      .LEFT = 305
      .TOP = 850
      .Width = 1000
      .Height = 220
    End With
    '.MainWindow.LinkedWindows.add Immediate
    'Immediate.Close
  End With
End Sub

Sub toggleFullScreen()
  If Application.DisplayFormulaBar Then
    Application.DisplayFormulaBar = False
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"", False)"
    ActiveWindow.DisplayHeadings = False
  Else
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"", True)"
  End If
End Sub

Sub SetXY()
  'LinkedImmediateWindow

  Call SetXY_Excel(1, 1, 690, 780)
  Call SetXY_Chrome
  'toggleFullScreen
End Sub
Sub SetXY_Excel(Optional ByVal L% = 0, _
                  Optional ByVal t% = 0, _
                  Optional ByVal W% = 840, _
                  Optional ByVal H% = 787)
  DoEvents
  Dim Win
  If Application.Version > 14 Then
    Set Win = Application.Windows(1)
  Else
    Set Win = Application
  End If
  Win.WindowState = xlNormal
  GoSub p
  Set Win = Application.VBE.MainWindow
  If Win.visible Then
    Win.WindowState = 0
    L = 0: t = 40: W = 1310: H = 1010
    GoSub p
  End If
  Set Win = Nothing
Exit Sub
p:
  With Win
'    If .WindowState = xlMaximized Then
'      .WindowState = xlNormal
'    End If
    If .LEFT <> L Then .LEFT = L
    If .TOP <> t Then .TOP = t
   
    If .Width <> W Then .Width = W
    If .Height <> H Then .Height = H
  End With
Return
End Sub

Private Sub SetXY_Chrome_test()
  Static b As Boolean
  Call SetXY_Chrome(HideChrome:=b)
End Sub
Sub SetXY_Chrome(Optional L% = 1408, _
                  Optional t% = 0, _
                  Optional W% = 500, _
                  Optional H% = 1053, _
                  Optional ByRef HideChrome As Boolean)
  DoEvents
  'chromeConnect
  On Error Resume Next
  If SeleDriver Is Nothing Then Exit Sub

  If HideChrome Then
    SeleDriver.Window.SetPosition 1921, 0
    SeleDriver.Window.SetSize W, H

    HideChrome = False
  Else
    HideChrome = True
    SeleDriver.Window.SetPosition L, t
    SeleDriver.Window.SetSize W, H
  End If
End Sub





Sub TurnOffScreen()
    Call SendMessage(HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, 2)
End Sub

Sub TurnOnScreen()
    Call SendMessage(HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, -1)
End Sub
    

Sub Delay(Optional ByVal MiliSecond% = 1000)
  Dim Start&, check&
  Start = getTickCount&()
  Do Until check >= Start + MiliSecond
    DoEvents
    check = getTickCount&()
  Loop
End Sub

Function ParentNode(ByVal Element As Object, Optional ByVal Timeout& = -1, Optional ByVal returnError As Boolean = 0) As Object
  Set ParentNode = Element.FindElementByXPath("./..", Timeout, returnError)
End Function

Function childNodes(ByVal Element As Object, Optional ByVal Timeout& = 0) As Object
  Set childNodes = Element.FindElementsByXPath("./child::*", , Timeout)
End Function
Function childNode(ByVal Element As Object, Optional ByVal Timeout& = -1, Optional ByVal returnError As Boolean = 0) As Object
  On Error Resume Next
  Set childNode = Element.FindElementByXPath("./child::*", Timeout, returnError)
  On Error GoTo 0
End Function

Function ChildInChilds( _
                 ByVal ElementObject As Object, _
            ParamArray childs()) As Object
'  On Error Resume Next
'  Dim index
'  For Each index In childs
'    Set ElementObject = childNodes(ElementObject)(index)
'  Next
'  Set ChildInChilds = ElementObject
'  On Error GoTo 0
End Function

Function FindClass(ByVal Element As Object, ByVal classname As String, Optional ByVal Timeout& = -1, Optional ByVal returnError As Boolean = 0) As Object
  Set FindClass = Element.FindElementByClass(classname, Timeout, returnError)
End Function

Function getTextByClass( _
           ByVal classname As String, _
           ByVal Element As Object, _
  Optional ByVal Default = "")
  On Error Resume Next
  Set Element = Element.FindElementByClass(classname, 0, 0)
  If Not Element Is Nothing Then
    getTextByClass = Element.Attribute("textContent")
  Else
    getTextByClass = Default
  End If
  On Error GoTo 0
End Function

Function compTextByClass( _
           ByVal text$, _
           ByVal classname As String, _
           ByVal Element As Object) As Boolean
  On Error Resume Next
  Set Element = Element.FindElementByClass(classname, 0, 0)
  If Not Element Is Nothing Then
    compTextByClass = LCase(Trim(Element.Attribute("textContent"))) Like LCase(text)
  End If
  On Error GoTo 0
End Function

#If Win64 Then
Function InstanceToWnd(ByVal target_pid As Long) As LongPtr
#Else
Function InstanceToWnd(ByVal target_pid As Long) As Long
#End If
  Dim hWnd As var64, pid As Long, thread_id As Long
  hWnd.Long = FindWindow("Chrome_WidgetWin_1", vbNullString)
  Do While hWnd.Long <> 0
    If GetParent(hWnd.Long) = 0 Then
      thread_id = GetWindowThreadProcessId(hWnd.Long, pid)
      If pid = target_pid Then
        InstanceToWnd = hWnd.Long
        Exit Do
      End If
    End If
    hWnd.Long = GetWindow(hWnd.Long, 2)
  Loop
End Function
'#If Win64 Then
'Function InstanceToWnd(ByVal target_pid As Long) As LongPtr
'#Else
'Function InstanceToWnd(ByVal target_pid As Long) As Long
'#End If
'  Dim hwnd As var64, pid As Long, thread_id As Long
'  hwnd.Long = FindWindow(vbNullString, vbNullString)
'  Do While hwnd.Long <> 0
'    If GetParent(hwnd.Long) = 0 Then
'      thread_id = GetWindowThreadProcessId(hwnd.Long, pid)
'      If pid = target_pid Then
'        InstanceToWnd = hwnd.Long
'        Exit Do
'      End If
'    End If
'    hwnd.Long = GetWindow(hwnd.Long, 2)
'  Loop
'End Function
#If Win64 Then
Function BringWindowToFront(ByVal hWnd As LongPtr) As Boolean
#Else
Function BringWindowToFront(ByVal hWnd As Long) As Boolean
#End If

  Dim ThreadID1 As Long, ThreadID2 As Long, nRet As Long
  On Error Resume Next
  If hWnd = GetForegroundWindow() Then
    BringWindowToFront = True
  Else
    ThreadID1 = GetWindowThreadProcessId(GetForegroundWindow, ByVal 0&)
    ThreadID2 = GetWindowThreadProcessId(hWnd, ByVal 0&)
    Call AttachThreadInput(ThreadID1, ThreadID2, True)
    nRet = SetForegroundWindow(hWnd)
    If IsIconic(hWnd) Then
      Call ShowWindow(hWnd, 9) ' SW_RESTORE)
      'Call ShowWindow(hwnd, 5) 'SW_SHOW)
    Else
      Call ShowWindow(hWnd, SW_SHOWNORMAL) 'SW_SHOW 5)
    End If
    BringWindowToFront = CBool(nRet)
    Call AttachThreadInput(ThreadID1, ThreadID2, False)
  End If
End Function

Public Sub focusChrome()
  Dim H As var64
  H.Long = GetChromeHandleByProcessID(MAINPORT, "")
  'Putfocus h.Long
  'PostMessage h.Long, &H201, 0&, ByVal MakeDWord(5, 5) 'WM_LBUTTONDOWN
  'PostMessage h.Long, &H202, 0&, ByVal MakeDWord(5, 5) 'WM_LBUTTONUP
  BringWindowToFront H.Long
End Sub
#If Win64 Then
Public Function GetChromeHandleByProcessID( _
   Optional ByVal Port$, _
   Optional ByVal UserDataDir$) As LongPtr
#Else
Public Function GetChromeHandleByProcessID( _
  Optional ByVal Port$, _
  Optional ByVal UserDataDir$) As Long
#End If
  Dim p
  For Each p In VBA.GetObject("winmgmts:\\.\root\CIMV2") _
    .ExecQuery("SELECT * FROM Win32_Process WHERE Name = ""chrome.exe""", , 48)
    If LCase$(p.commandLine) Like _
       LCase$("*chrome*--remote-debugging-port=" & _
        Port & "*--user-data-dir=*" & LCase$(UserDataDir) & "*") Then
      GetChromeHandleByProcessID = InstanceToWnd(p.Processid)
      Exit For
    End If
  Next
End Function


Public Sub closeChromeGames()
  Dim p
  For Each p In VBA.GetObject("winmgmts:\\.\root\CIMV2") _
    .ExecQuery("SELECT * FROM Win32_Process WHERE (Name = ""chrome.exe"" and commandLine like '%chrome.exe"" --remote-debugging-port=" & MAINPORT & "%')", , 48)
    p.Terminate
  Next
  Set p = Nothing
  Set SeleDriver = Nothing
End Sub

Private Function MakeDWord(LoWord As Integer, HiWord As Integer) As Long
    MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function
 


Private Sub ChromeMove_test()
  Call ChromeMove(0, 0, 600, 800)
End Sub

Sub ChromeMove(ByVal LEFT&, ByVal TOP&, ByVal Width&, ByVal Height&)
  Dim H As Long
  H = ChromeHandle
  If H > 0 Then
    MoveWindow H, LEFT, TOP, Width, Height, True
  End If
End Sub

Function ChromeHandle() As Long
  Dim s$, L&, p As Object
  For Each p In VBA.GetObject("winmgmts:") _
     .ExecQuery("SELECT * FROM win32_process Where name = 'chrome.exe'", , 48)
    'find the Handle for the Chrome Browser
    ChromeChilds 0&, 0
    s = String$(100, Chr$(0))
    L = GetClassName(ChildHwnd, s, 100)
    'loop incase of more siblings
    While Not VBA.LEFT$(s, L) = "Chrome_WidgetWin_1"
      ChildHwnd = GetAncestor(ChildHwnd, 1) ' GA_PARENT)
      s = String$(100, Chr$(0))
      L = GetClassName(ChildHwnd, s, 100)
      'Duplicate of classname but WidgetWin_0
      If ChildHwnd = 0 Then
        origChildFound = True
        ChromeChilds retainedChromeHwnd, 0
      End If
    Wend
    ChromeHandle = ChildHwnd
    Exit For
  Next
End Function

Public Sub ChromeChilds(hparent&, xcount&)
  Dim ChromeID&, strtext$, ChromeClassName$, ChromeHwnd&
  Dim lngret&
  ChromeHwnd = FindWindowEx(hparent, 0&, vbNullString, vbNullString)
  If origChildFound = True Then
    ChromeHwnd = retainedChromeHwnd
    origChildFound = False
  End If
  If ChildFound = True And GoNextParent = True Then
    Exit Sub
  ElseIf ChildFound = True Then
    NextHandle = True
    ChildFound = False
  End If
  While ChromeHwnd <> 0
    strtext = String$(100, Chr$(0))
    lngret = GetClassName(ChromeHwnd, strtext, 100)
    ChromeClassName = LEFT$(strtext, lngret)
    If ChromeClassName = "Chrome_RenderWidgetHostHWND" Then
      ChildFound = True
      ChildHwnd = ChromeHwnd
    End If
    xcount = xcount + 1
    ChromeChilds ChromeHwnd, xcount 'loop through next level of child windows
    If ChildFound = True Then Exit Sub
    ChromeHwnd = FindWindowEx(hparent, ChromeHwnd, vbNullString, vbNullString)
    If hparent = 0 And NextHandle = True Then
      retainedChromeHwnd = ChromeHwnd
      ChildFound = True
      GoNextParent = True
    End If
  Wend
End Sub


Function getChromePath()
  getChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
  If Len(Dir(getChromePath)) = 0 Then
    getChromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    If Len(Dir(getChromePath)) = 0 Then
      getChromePath = GetFolder(&H1C) & "\Google\Chrome\Application\chrome.exe"
      If Len(Dir(getChromePath)) = 0 Then
        getChromePath = vbNullString
      End If
    End If
  End If
End Function

Function GetFolder(ByVal lngFolder&)
 Dim strPath$, strBuffer As String * 1000
 If SHGetFolderPath(0&, lngFolder, 0&, 0, strBuffer) = 0 Then
   strPath = LEFT$(strBuffer, InStr(strBuffer, Chr$(0)) - 1)
 Else
   strPath = vbNullString
 End If
 GetFolder = strPath
End Function


Sub ChromeWindow_Click()
  SetCursorPos 1900, 32
  mouse_event &H2, 0, 0, 0, 0
  mouse_event &H4, 0, 0, 0, 0
End Sub

