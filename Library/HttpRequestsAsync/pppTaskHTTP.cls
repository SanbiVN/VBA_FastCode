VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pppTaskHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Name$, ID$, Method$
Public Workname$
Public work As Object
Public http As Object
Public url$
Public Api$
Public Cookie$
Public Status
Public WorkIndex%
Public httpIndex%
Public ErrorRuns%
Public timeTaken As Date
Public Directive As EnumDirective
Public Path$, FileName$, Extend$, DefaultExtend$
Public SheetShowImage$
Public TextFontImage$
Public TitleFontImage$
Private Sub Class_Initialize()
  ErrorRuns = 2
End Sub
Private Sub Class_Terminate()
  Set http = Nothing
  Set work = Nothing
End Sub

