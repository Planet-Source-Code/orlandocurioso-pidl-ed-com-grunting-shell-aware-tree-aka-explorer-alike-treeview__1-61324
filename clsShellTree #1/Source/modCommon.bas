Attribute VB_Name = "modCommon"
'---------------------------------------------------------------------------------------
' Module    : modCommon
' Author    : OrlandoCurioso 16.05.2005
' Purpose   :
'
'---------------------------------------------------------------------------------------
Option Explicit

Public Const MySEH_ERROR = 12345&

Public Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)

Public Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
   Private Const EXCEPTION_CONTINUE_EXECUTION = -1
   Private Const EXCEPTION_CONTINUE_SEARCH = 0
   Private Const EXCEPTION_EXECUTE_HANDLER = 1
'Private Declare Function UnhandledExceptionFilter Lib "kernel32" (ByRef ExceptionInfo As EXCEPTION_POINTERS) As Long

' Visual C++ Debugger: press F10 twice to return to VB
' SetUnhandledExceptionFilter ByVal 0&  ' -> remove hook prior breaking
Public Declare Sub DebugBreak Lib "kernel32" ()

'' >= WinXP
'Public Enum EFaultRepRetVal
'    frrvOk = 0
'    frrvOkManifest = 1
'    frrvOkQueued = 2
'    frrvErr = 3
'    frrvErrNoDW = 4
'    frrvErrTimeout = 5
'    frrvLaunchDebugger = 6
'    frrvOkHeadless = 7
'End Enum
''Public Declare Function ReportFault Lib "Faultrep" (pep As EXCEPTION_POINTERS, ByVal dwMode As Long) As EFaultRepRetVal
'Public Declare Function ReportFaultPtr Lib "Faultrep" Alias "ReportFault" (pep As Long, ByVal dwMode As Long) As EFaultRepRetVal

' EscapePressed ()
Private Type MSGTYPE
   hWnd     As Long
   message  As Long
   wParam   As Long
   lParam   As Long
   time     As Long
   X        As Long      'pt As POINTAPI
   Y        As Long
End Type

Private Declare Function GetInputState Lib "user32" () As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSGTYPE, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
   Private Const WM_KEYFIRST = &H100
   Private Const WM_KEYLAST = &H108
   Private Const PM_REMOVE = &H1

'Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Const KEYEVENTF_KEYUP = &H2


Private m_bInIDE As Boolean
'

Public Property Get InIDE() As Boolean
   Debug.Assert (pIsInIDE)
   InIDE = m_bInIDE
End Property
Private Property Get pIsInIDE() As Boolean
   m_bInIDE = True
   pIsInIDE = True
End Property

' call as Debug.Assert DbgPrt (condition,"text1","text2",..)
Public Function DbgPrt(ByVal bNoiseLevel As Boolean, ParamArray va()) As Boolean
   On Error GoTo err_h
   Dim i As Long
   If bNoiseLevel Then
      For i = LBound(va) To UBound(va) - 1
         Debug.Print va(i); Tab;
      Next
      Debug.Print va(i)
   End If
   DbgPrt = True
err_h:
End Function

Public Sub LogEvent(ByVal sLog As String, Optional ByVal bNewFile As Boolean = False)
   On Error Resume Next
   Dim sFIle As String
   Dim iFile As Long
   sFIle = App.Path & "\Log.txt"
   If bNewFile Then Kill sFIle
   iFile = FreeFile
   Open sFIle For Append As #iFile
   Print #iFile, sLog
   Close #iFile
End Sub

' http://www.devx.com/vb2themax/Tip/18962
Public Function EscapePressed(Optional msgText As String) As Boolean
   Dim mess As MSGTYPE
   
   If GetInputState() Then
      PeekMessage mess, 0, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE
      If mess.wParam = vbKeyEscape Then
         If Len(msgText) = 0 Then
            ' Escape was pressed, return True without showing a msgbox
            EscapePressed = True
         Else
            ' Escape was pressed, ask user to confirm
            EscapePressed = (MsgBox(msgText, vbQuestion + vbYesNo) = vbYes)
         End If
      End If
   End If
   
End Function

Public Function MyExceptionHandler(lpEP As Long) As Long
   Dim lRes As VbMsgBoxResult
   
   lRes = MsgBox("EXCEPTION ???" & vbCrLf & "Continue in VB, End, or invoke Debugger ?", _
                  vbYesNoCancel Or vbCritical, App.Title & "MyExceptionHandler")
                  
   Select Case lRes
   
      Case vbYes
         ' Continue execution in VB
         If InIDE Then
            Stop
            MyExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
            ' raise VB error :
            ' if proc where exception occured has no err handler, this proc is called again!
            On Error GoTo 0
            Err.Raise MySEH_ERROR, "MyExceptionHandler", "EXCEPTION ???"
         Else
            ' # only allow in IDE, unless MySEH_ERROR handled in all procs #
            ' default dlg : stop or invoke default Debugger
            MyExceptionHandler = EXCEPTION_CONTINUE_SEARCH
         End If
         
       Case vbNo
         ' End
         MyExceptionHandler = EXCEPTION_EXECUTE_HANDLER
           
       Case vbCancel
         ' default dlg : stop or invoke default Debugger
         MyExceptionHandler = EXCEPTION_CONTINUE_SEARCH
   
   End Select
   
End Function
