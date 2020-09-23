VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShellNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsShellNotify
' Author    : OrlandoCurioso 28.04.2005  / Brad Martinez / Paul Caton
' Purpose   : Creates a window and registers it to receive Shell change notifications.
'             On notification raises ShellNotification event.
' CanDo     : Use window to detect system wide changes.(ie WM_SETTINGCHANGE,WM_WININICHANGE)
'---------------------------------------------------------------------------------------
Option Explicit

' Project conditional compilation constants
'#Const UNICODE = x
'#Const WIN32_IE = 1024  ' 1024 == &H400 >= IE4 >= Shell32 4.71 == Win98,(NT4)
'                        ' 1280 == &H500 >= IE5 >= Shell32 5.00 >= Win2k

' Private conditional compilation constants:
#Const HEXORG = 0      ' original hex pairs to bytes conversion in Subclass_Start()

'========================================================================================
' Declarations
'========================================================================================

Public Event ShellNotification(ByVal EventID As SHCN_EventIDs, ByVal dwItem1 As Long, ByVal dwItem2 As Long, ByVal bDuplicate As Boolean)

' Shell notification event IDs (MSDN: SHChangeNotify function)
Public Enum SHCN_EventIDs
   SHCNE_RENAMEITEM = &H1           '(D) A non-folder item has been renamed.
   SHCNE_CREATE = &H2               '(D) A non-folder item has been created.
   SHCNE_DELETE = &H4               '(D) A non-folder item has been deleted.
   SHCNE_MKDIR = &H8                '(D) A folder item has been created.
   SHCNE_RMDIR = &H10               '(D) A folder item has been removed.
   SHCNE_MEDIAINSERTED = &H20       '(G) Storage media has been inserted into a drive.
   SHCNE_MEDIAREMOVED = &H40        '(G) Storage media has been removed from a drive.
   SHCNE_DRIVEREMOVED = &H80        '(G) A drive has been removed.
   SHCNE_DRIVEADD = &H100           '(G) A drive has been added.
   SHCNE_NETSHARE = &H200           'A folder on the local computer is being shared via the network.
   SHCNE_NETUNSHARE = &H400         'A folder on the local computer is no longer being shared via the network.
   SHCNE_ATTRIBUTES = &H800         '(D) The attributes of an item or folder have changed.
   SHCNE_UPDATEDIR = &H1000         '(D) The contents of an existing folder have changed, but the folder still exists and has not been renamed.
   SHCNE_UPDATEITEM = &H2000        '(D) An existing non-folder item has changed, but the item still exists and has not been renamed.
   SHCNE_SERVERDISCONNECT = &H4000  'The computer has disconnected from a server.
   SHCNE_UPDATEIMAGE = &H8000&      '(G) An image in the system image list has changed.
   SHCNE_DRIVEADDGUI = &H10000      '(G) A drive has been added and the shell should create a new window for the drive.
   SHCNE_RENAMEFOLDER = &H20000     '(D) The name of a folder has changed.
   SHCNE_FREESPACE = &H40000        '(G) The amount of free space on a drive has changed.
   SHCNE_ASSOCCHANGED = &H8000000   '(G) A file type association has changed.
   SHCNE_DISKEVENTS = &H2381F       '(D) Specifies a combination of all of the disk event identifiers.
   SHCNE_GLOBALEVENTS = &HC0581E0   '(G) Specifies a combination of all of the global event identifiers.
   SHCNE_ALLEVENTS = &H7FFFFFFF
   SHCNE_INTERRUPT = &H80000000     'The specified event occurred as a result of a system interrupt. It is stripped out before the clients of SHCNNotify_ see it.
#If (WIN32_IE >= &H400) Then
   SHCNE_EXTENDED_EVENT = &H4000000 ' see SHCNEE_flags
#End If
End Enum

' SHCNE_EXTENDED_EVENT extended events.  These events are ordinals. This is not a bitfield.
#If (WIN32_IE >= &H400) Then
Public Enum SHCNEE_flags
   SHCNEE_ORDERCHANGED = 2         ' dwItem2 is the pidl of the changed folder
   SHCNEE_MSI_CHANGE = 4           ' dwItem2 is the product code
   SHCNEE_MSI_UNINSTALL = 5        ' dwItem2 is the product code
End Enum
#End If

'---------------------------------------------------------------------------------------
' Private declares
'---------------------------------------------------------------------------------------

Private Const WM_DESTROY            As Long = &H2
Private Const WM_USER               As Long = &H400

Private Const WS_POPUP              As Long = &H80000000

'// Shell notifications

' self defined message for the shell to send to window
Private Const WM_SHELL_NOTIFY       As Long = WM_USER + 1

Private Type SHNOTIFYSTRUCT         ' VB workaround for 64bit Long
   dwItem1     As Long
   dwItem2     As Long
End Type

Private Type SHChangeNotifyEntry
   pidl        As Long              ' absolute pidl of folder to watch.
   fRecursive  As Long              ' flag to indicate watching of subfolders.
End Type

' Notification flags SHChangeNotify
' uFlags & SHCNF_TYPE is an ID which indicates what dwItem1 and dwItem2 mean (documented for SHChangeNotify)
Public Enum SHCN_ItemFlags
   SHCNF_IDLIST = &H0         'LPITEMIDLIST
   SHCNF_PATHA = &H1          'path name
   SHCNF_PRINTERA = &H2       'printer friendly name
   SHCNF_DWORD = &H3          'DWORD
   SHCNF_PATHW = &H5          'path name
   SHCNF_PRINTERW = &H6       'printer friendly name
   SHCNF_TYPE = &HFF
   SHCNF_FLUSH = &H1000       'Flushes the system event buffer. The function does not return until the system is finished processing the given event.
   SHCNF_FLUSHNOWAIT = &H2000 'Flushes the system event buffer. The function returns immediately regardless of whether the system is finished processing the given event.
#If UNICODE Then
   SHCNF_PATH = SHCNF_PATHW
   SHCNF_PRINTER = SHCNF_PRINTERW
#Else
   SHCNF_PATH = SHCNF_PATHA
   SHCNF_PRINTER = SHCNF_PRINTERA
#End If
End Enum

#If WIN32_IE >= &H500 Then

' documented >= Win2K
Private Enum SHCNRF_ItemFlags
   SHCNRF_InterruptLevel = &H1         ' Interrupt level notifications from the file system.
   SHCNRF_ShellLevel = &H2             ' Shell-level notifications from the shell.
   SHCNRF_RecursiveInterrupt = &H1000  ' Interrupt events on the whole subtree. This flag must be combined with the SHCNRF_InterruptLevel flag. When using this flag, notifications must also be made recursive by setting the fRecursive member of the corresponding SHChangeNotifyEntry structure referenced by pfsne to TRUE.
   SHCNRF_NewDelivery = &H8000         ' Messages received use shared memory. Call SHChangeNotification_Lock to access the actual data. Call SHChangeNotification_Unlock to release the memory when done.
End Enum

Private Declare Function SHChangeNotifyRegister Lib "shell32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long, ByVal wEventsMask As Long, ByVal wMsg As Long, ByVal cItems As Long, lpItems As SHChangeNotifyEntry) As Long
Private Declare Function SHChangeNotifyDeregister Lib "shell32.dll" (ByVal ulID As Long) As Long
'Private Declare Function SHChangeNotification_Lock Lib "shell32.dll" (xxx) As Long
'Private Declare Function SHChangeNotification_Unlock Lib "shell32.dll" (xxx) As Long

Private Declare Function ILIsEqual Lib "shell32" (ByVal pidl1 As Long, ByVal pidl2 As Long) As Boolean 'long
Private Declare Function ILClone Lib "shell32" (ByVal pidl As Long) As Long
Private Declare Sub ILFree Lib "shell32" (ByVal pidl As Long)

#Else

' undocumented, yet exist by ordinal >= Shell32 4.0
Private Declare Function SHChangeNotifyRegister Lib "shell32.dll" Alias "#2" (ByVal hWnd As Long, ByVal dwFlags As Long, ByVal wEventsMask As Long, ByVal wMsg As Long, ByVal cItems As Long, lpItems As SHChangeNotifyEntry) As Long
Private Declare Function SHChangeNotifyDeregister Lib "shell32.dll" Alias "#4" (ByVal ulID As Long) As Long
'Private Declare Function SHChangeNotification_Lock Lib "shell32.dll" Alias "#644" (xxx) As Long
'Private Declare Function SHChangeNotification_Unlock Lib "shell32.dll" Alias "#645" (xxx) As Long

Private Declare Function ILIsEqual Lib "shell32" Alias "#21" (ByVal pidl1 As Long, ByVal pidl2 As Long) As Boolean
Private Declare Function ILClone Lib "shell32" Alias "#18" (ByVal pidl As Long) As Long
Private Declare Sub ILFree Lib "shell32" Alias "#155" (ByVal pidl As Long)

#End If

Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As SHCN_EventIDs, ByVal uFlags As SHCN_ItemFlags, ByVal dwItem1 As Long, ByVal dwItem2 As Long)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
   Private Const HWND_MESSAGE As Long = -3   ' >= Win2K
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Type SHNOTIFY_LAST
   eSHID       As SHCN_EventIDs
   dwItem1     As Long
   dwItem2     As Long
   dwCopy1     As Long
End Type

Private m_hwnd                As Long
Private m_hSHNotify           As Long
Private m_eEventID            As SHCN_EventIDs
Private m_tSHLast             As SHNOTIFY_LAST

'========================================================================================
' Subclasser declarations
'========================================================================================

Private Enum eMsgWhen
    [MSG_AFTER] = 1                                  'Message calls back after the original (previous) WndProc
    [MSG_BEFORE] = 2                                 'Message calls back before the original (previous) WndProc
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE 'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES     As Long = -1          'All messages added or deleted
Private Const CODE_LEN         As Long = 200         'Length of the machine code in bytes
Private Const GWL_WNDPROC      As Long = -4          'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04         As Long = 88          'Table B (before) address patch offset
Private Const PATCH_05         As Long = 93          'Table B (before) entry count patch offset
Private Const PATCH_08         As Long = 132         'Table A (after) address patch offset
Private Const PATCH_09         As Long = 137         'Table A (after) entry count patch offset

Private Type tSubData                                'Subclass data type
    hWnd                       As Long               'Handle of the window being subclassed
    nAddrSub                   As Long               'The address of our new WndProc (allocated memory).
    nAddrOrig                  As Long               'The address of the pre-existing WndProc
    nMsgCntA                   As Long               'Msg after table entry count
    nMsgCntB                   As Long               'Msg before table entry count
    aMsgTblA()                 As Long               'Msg after table array
    aMsgTblB()                 As Long               'Msg Before table array
End Type

Private sc_aSubData()          As tSubData           'Subclass data array
#If HEXORG = 1 Then
Private sc_aBuf(1 To CODE_LEN) As Byte               'Code buffer byte array
#Else
Private sc_aBuf(1 To 50)       As Long               'Code buffer Long array
#End If
Private sc_pCWP                As Long               'Address of the CallWindowsProc
Private sc_pEbMode             As Long               'Address of the EbMode IDE break/stop/running function
Private sc_pSWL                As Long               'Address of the SetWindowsLong function
  
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
'

'========================================================================================
' Subclass handler - MUST be the first Public routine in this file. That includes public properties also
'========================================================================================

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
   
   Dim tNotify       As SHNOTIFYSTRUCT
   Dim bDuplicateMsg As Boolean
   
   On Error GoTo Proc_Error

   Debug.Assert lng_hWnd = m_hwnd
   
   If uMsg = WM_DESTROY Then
      Subclass_StopAll
      UnregisterNotify
      m_hwnd = 0
      Exit Sub
   Else
      Debug.Assert uMsg = WM_SHELL_NOTIFY
      bHandled = True
   End If
   
   ' -- Received a shell change notification! --
   
   ' wParam is a pointer to a SHNOTIFYSTRUCT, lParam is the ID of the notication event.
   ' SHNOTIFYSTRUCT contains fully qualified pidl(s) relative to desktop, freed by shell!
   CopyMemory tNotify, ByVal wParam, LenB(tNotify)
   
   ' -- detect duplicate messages:
   ' dwItemX pointers are reused for same item and different EventID's.
   ' If item resides in ie filesystem and virtual folder, for every location a
   ' notify is sent(easily mistaken as duplicate notify).
   ' Further pointers are reused for new items with indices,
   ' ie "New Folder" and "New Folder (1)" or "Item.lnk" and "Item (1).lnk".
   ' So equal pidls(as pointer values) don't necessarily imply same pidls(as ItemIDLists).
   
   ' Why does shell send duplicate notifications ???
   ' Sending an item from MyDocuments to recycle bin results in 4 x SHCNE_DELETE
   ' + 2 x SHCNE_DELETE for same item in MyComputer. ???
   
   ' -- Chosen detection strategy (after several failures)
   ' For the first unique notif the EventID and pointer values are stored.
   ' Additionally the dwItem1 is copied and stored. Prior the last copy is freed.
   ' dwItem2 is discarded #? is dwItem1 sufficient? #.
   ' For following notifs, if EventID and pointer values match, dwItem1 is compared
   ' to stored copy.For duplicates storage remains unaltered, for unique cycle rebegins.
   ' On termination last copy is freed in UnregisterNotify().
   
   With m_tSHLast
      
      If lParam = .eSHID Then
         If tNotify.dwItem1 = .dwItem1 Then
            If tNotify.dwItem2 = .dwItem2 Then
            
               bDuplicateMsg = ILIsEqual(tNotify.dwItem1, .dwCopy1)
            End If
         End If
      End If
      
      If Not bDuplicateMsg Then
         ' store current notify
         .eSHID = lParam
         .dwItem1 = tNotify.dwItem1
         .dwItem2 = tNotify.dwItem2
         
         ' ILFree & ILClone handle passed zero values
         ILFree .dwCopy1
         .dwCopy1 = ILClone(tNotify.dwItem1)
         
      End If
      
   End With
   
   Debug.Print vbCrLf & IIf(bDuplicateMsg, "DUPLICATE ", "") & "SHELL NOTIFICATION: " & dbgSHNotify_GetEventStr(lParam)
      
   RaiseEvent ShellNotification(lParam, tNotify.dwItem1, tNotify.dwItem2, bDuplicateMsg)

   Exit Sub

Proc_Error:
   Debug.Print "Error: " & Err.Number & ". " & Err.Description, vbOKOnly Or vbCritical, App.Title & ".clsShellNotify: Sub zSubclass_Proc"
   If InIDE Then Stop: Resume
End Sub

'========================================================================================
' Class Code
'========================================================================================

' set notifications, we want (MSDN: SHChangeNotify function)
'Public Property Get EventID() As SHCN_EventIDs
'   EventID = m_eEventID
'End Property
Public Property Let EventID(ByVal eEventID As SHCN_EventIDs)
   m_eEventID = eEventID
End Property

Public Function Initialize(ByVal hOwner As Long) As Boolean
   
   If (m_hwnd <> 0) Then
      Call Class_Terminate
   End If
   
#If WIN32_IE >= &H500 Then
   ' message-only window
   m_hwnd = CreateWindowEx(0&, "MESSAGE", vbNullString, 0&, 0&, 0&, 0&, 0&, HWND_MESSAGE, 0&, 0&, ByVal 0&)
#Else
   m_hwnd = CreateWindowEx(0&, "STATIC", vbNullString, WS_POPUP, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
#End If
   
   If (m_hwnd <> 0) Then
      
      Subclass_Start m_hwnd
      Subclass_AddMsg m_hwnd, WM_SHELL_NOTIFY, MSG_BEFORE
      Subclass_AddMsg m_hwnd, WM_DESTROY, MSG_BEFORE
      
      Initialize = True
   End If
   
End Function

Private Sub Class_Terminate()
   
   Subclass_StopAll
   UnregisterNotify
   
   If m_hwnd > 0 Then
      If DestroyWindow(m_hwnd) = 0 Then
         Debug.Assert False
      End If
   End If
   m_hwnd = 0
End Sub

' # call sucessive for multiple folders , don't use array of tNotify's #
Public Function RegisterNotify(ByVal pidl As Long) As Boolean

   Dim tNotify As SHChangeNotifyEntry
   
   Debug.Assert m_hwnd
   Debug.Assert m_eEventID
   
   UnregisterNotify
   
   With tNotify
      .pidl = pidl
      .fRecursive = True
   End With
   
#If WIN32_IE >= &H500 Then
   
   ' SHCNRF_ItemFlags
   m_hSHNotify = SHChangeNotifyRegister(m_hwnd, SHCNRF_ShellLevel, _
                                        m_eEventID, WM_SHELL_NOTIFY, 1, tNotify)
#Else

   ' SHCN_ItemFlags ???
   m_hSHNotify = SHChangeNotifyRegister(m_hwnd, SHCNF_TYPE Or SHCNF_IDLIST, _
                                        m_eEventID, WM_SHELL_NOTIFY, 1, tNotify)
#End If

   RegisterNotify = (m_hSHNotify <> 0)

End Function

Public Sub UnregisterNotify()
      
   ILFree m_tSHLast.dwCopy1:  m_tSHLast.dwCopy1 = 0
      
   If m_hSHNotify Then
      ' Unregister our handle.
      If SHChangeNotifyDeregister(m_hSHNotify) <> 0 Then
         m_hSHNotify = 0
      Else
         Debug.Assert False
      End If
   End If
End Sub

Public Sub SendNotify(ByVal wEventId As SHCN_EventIDs, ByVal uFlags As SHCN_ItemFlags, ByVal dwItem1 As Long, ByVal dwItem2 As Long)
   
   Debug.Print "SENT SHChangeNotify: " & dbgSHNotify_GetEventStr(wEventId)
   SHChangeNotify wEventId, uFlags, dwItem1, dwItem2
   Debug.Print "RTN  SHChangeNotify: " & dbgSHNotify_GetEventStr(wEventId)
End Sub

Private Function dbgSHNotify_GetEventStr(dwEventID As Long) As String
   Dim sEvent As String
  ' Returns the event string associated with the specified event ID value.
   
   Select Case dwEventID
      Case SHCNE_RENAMEITEM:       sEvent = "SHCNE_RENAMEITEM"       '&H1
      Case SHCNE_CREATE:           sEvent = "SHCNE_CREATE"           '&H2
      Case SHCNE_DELETE:           sEvent = "SHCNE_DELETE"           '&H4
      Case SHCNE_MKDIR:            sEvent = "SHCNE_MKDIR"            '&H8
      Case SHCNE_RMDIR:            sEvent = "SHCNE_RMDIR"            '&H10
      Case SHCNE_MEDIAINSERTED:    sEvent = "SHCNE_MEDIAINSERTED"    '&H20
      Case SHCNE_MEDIAREMOVED:     sEvent = "SHCNE_MEDIAREMOVED"     '&H40
      Case SHCNE_DRIVEREMOVED:     sEvent = "SHCNE_DRIVEREMOVED"     '&H80
      Case SHCNE_DRIVEADD:         sEvent = "SHCNE_DRIVEADD"         '&H100
      Case SHCNE_NETSHARE:         sEvent = "SHCNE_NETSHARE"         '&H200
      Case SHCNE_NETUNSHARE:       sEvent = "SHCNE_NETUNSHARE"       '&H400
      Case SHCNE_ATTRIBUTES:       sEvent = "SHCNE_ATTRIBUTES"       '&H800
      Case SHCNE_UPDATEDIR:        sEvent = "SHCNE_UPDATEDIR"        '&H1000
      Case SHCNE_UPDATEITEM:       sEvent = "SHCNE_UPDATEITEM"       '&H2000
      Case SHCNE_SERVERDISCONNECT: sEvent = "SHCNE_SERVERDISCONNECT" '&H4000
      Case SHCNE_UPDATEIMAGE:      sEvent = "SHCNE_UPDATEIMAGE"      '&H8000&
      Case SHCNE_DRIVEADDGUI:      sEvent = "SHCNE_DRIVEADDGUI"      '&H10000
      Case SHCNE_RENAMEFOLDER:     sEvent = "SHCNE_RENAMEFOLDER"     '&H20000
      Case SHCNE_FREESPACE:        sEvent = "SHCNE_FREESPACE"        '&H40000
    
#If (WIN32_IE >= &H400) Then
      Case SHCNE_EXTENDED_EVENT:   sEvent = "SHCNE_EXTENDED_EVENT"   '&H4000000
#End If
    
      Case SHCNE_ASSOCCHANGED:     sEvent = "SHCNE_ASSOCCHANGED"     '&H8000000
    
      Case SHCNE_DISKEVENTS:       sEvent = "SHCNE_DISKEVENTS"       '&H2381F
      Case SHCNE_GLOBALEVENTS:     sEvent = "SHCNE_GLOBALEVENTS"     '&HC0581E0
      Case SHCNE_ALLEVENTS:        sEvent = "SHCNE_ALLEVENTS"        '&H7FFFFFFF
      Case SHCNE_INTERRUPT:        sEvent = "SHCNE_INTERRUPT"        '&H80000000
      Case Else:                   sEvent = "SHCNE_ unknown event":  Debug.Assert False
   End Select
  
   dbgSHNotify_GetEventStr = sEvent

End Function


'========================================================================================
' Subclass code - The programmer may call any of the following Subclass_??? routines
'========================================================================================

Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
'Parameters:
'   lng_hWnd - The handle of the window for which the uMsg is to be added to the callback table
'   uMsg     - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
'   When     - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  
    With sc_aSubData(zIdx(lng_hWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
''Delete a message from the table of those that will invoke a callback.
''Parameters:
''   lng_hWnd - The handle of the window for which the uMsg is to be removed from the callback table
''   uMsg     - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
''   When     - Whether the msg is to be removed from the before, after or both callback tables
'
'    With sc_aSubData(zIdx(lng_hWnd))
'        If (When And eMsgWhen.MSG_BEFORE) Then
'            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
'        End If
'        If (When And eMsgWhen.MSG_AFTER) Then
'            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
'        End If
'    End With
'End Sub

Private Function Subclass_InIDE() As Boolean
'Return whether we're running in the IDE.
    
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

' # OpCode taken from WinSubHook2.cSubclass v1.02, padded 6 zero bytes for same CODE_LEN #
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Start subclassing the passed window handle
'Parameters:
'   lng_hWnd - The handle of the window to be subclassed
'Returns;
'   The sc_aSubData() index

  Dim i                        As Long                       'Loop index
  Dim nSubIdx                  As Long                       'Subclass data index

  Const GMEM_FIXED             As Long = 0                   'Fixed memory GlobalAlloc flag
  Const PAGE_EXECUTE_READWRITE As Long = &H40&               'Allow memory to execute without violating XP SP2 Data Execution Prevention
  Const PATCH_01               As Long = 18                  'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02               As Long = 68                  'Address of the previous WndProc
  Const PATCH_03               As Long = 78                  'Relative address of SetWindowsLong
  Const PATCH_06               As Long = 116                 'Address of the previous WndProc
  Const PATCH_07               As Long = 121                 'Relative address of CallWindowProc
  Const PATCH_0A               As Long = 186                 'Address of the owner object
  Const FUNC_CWP               As String = "CallWindowProcA" 'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM               As String = "EbMode"          'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL               As String = "SetWindowLongA"  'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER               As String = "user32"          'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5               As String = "vba5"            'Location of the EbMode function if running VB5
  Const MOD_VBA6               As String = "vba6"            'Location of the EbMode function if running VB6

    'If it's the first time through here..
    If (sc_aBuf(1) = 0) Then

#If HEXORG = 1 Then

  Dim sSubCode                 As String                     'Subclass code string
  Dim j                        As Long                       'Loop index

        'Build the hex pair subclass string
        sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
                   "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
                   "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E32F09" & _
                   "C978078B450CF2AF75248D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF501CC3" & _
                   "000000"

        'Convert the string from hex pairs to bytes and store in the machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            sc_aBuf(j) = CByte("&H" & Mid$(sSubCode, i, 2))                       'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                      'Next pair of hex characters

        'Get API function addresses
        If (Subclass_InIDE) Then                                                  'If we're running in the VB IDE
            sc_aBuf(16) = &H90                                                    'Patch the code buffer to enable the IDE state code
            sc_aBuf(17) = &H90                                                    'Patch the code buffer to enable the IDE state code
            sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                            'Get the address of EbMode in vba6.dll
            If (sc_pEbMode = 0) Then                                              'Found?
                sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                        'VB5 perhaps
            End If
        End If

#Else
       ' Load OpCode from Long array (10 x faster)
              
        sc_aBuf(1) = &H83E58955:       sc_aBuf(26) = &HFFC30000
        sc_aBuf(2) = &H3157F8C4:       sc_aBuf(27) = &H75FF1475
        sc_aBuf(3) = &HFC4589C0:       sc_aBuf(28) = &HC75FF10
        sc_aBuf(4) = &HEBF84589:       sc_aBuf(29) = &H680875FF
        sc_aBuf(5) = &HE80E&:          sc_aBuf(30) = &H0&
        sc_aBuf(6) = &HF8830000:       sc_aBuf(31) = &HE8&
        sc_aBuf(7) = &H85217402:       sc_aBuf(32) = &HFC458900
        sc_aBuf(8) = &HE82474C0:       sc_aBuf(33) = &HBFD231C3
        sc_aBuf(9) = &H30&:            sc_aBuf(34) = &H0&
        sc_aBuf(10) = &HF87D83:        sc_aBuf(35) = &HB9&
        sc_aBuf(11) = &H38E80A75:      sc_aBuf(36) = &H1E800
        sc_aBuf(12) = &HE8000000:      sc_aBuf(37) = &HE3C30000
        sc_aBuf(13) = &H4D&:           sc_aBuf(38) = &H78C9092F
        sc_aBuf(14) = &HFC458B5F:      sc_aBuf(39) = &HC458B07
        sc_aBuf(15) = &H10C2C9:        sc_aBuf(40) = &H2475AFF2
        sc_aBuf(16) = &H26E8&:         sc_aBuf(41) = &H5014458D
        sc_aBuf(17) = &H68F1EB00:      sc_aBuf(42) = &H5010458D
        sc_aBuf(18) = &H0&:            sc_aBuf(43) = &H500C458D
        sc_aBuf(19) = &H75FFFC6A:      sc_aBuf(44) = &H5008458D
        sc_aBuf(20) = &HE808&:         sc_aBuf(45) = &H50FC458D
        sc_aBuf(21) = &HE0EB0000:      sc_aBuf(46) = &H50F8458D
        sc_aBuf(22) = &HBF4AD231:      sc_aBuf(47) = &HB852&
        sc_aBuf(23) = &H0&:            sc_aBuf(48) = &H8B500000
        sc_aBuf(24) = &HB9&:           sc_aBuf(49) = &H1C50FF00
        sc_aBuf(25) = &H2DE800:        sc_aBuf(50) = &HC3&

        'Get API function addresses
        If (Subclass_InIDE) Then                                                  'If we're running in the VB IDE
            sc_aBuf(4) = &H90F84589                                               'Patch the code buffer to enable the IDE state code
            sc_aBuf(5) = &HE890                                                   'Patch the code buffer to enable the IDE state code
            sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                            'Get the address of EbMode in vba6.dll
            If (sc_pEbMode = 0) Then                                              'Found?
                sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                        'VB5 perhaps
            End If
        End If

        i = 2 * CODE_LEN + 1   ' for calling VirtualProtect
#End If

        Call zPatchVal(VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me))                  'Patch the address of this object instance into the static machine code buffer

        sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                   'Get the address of the CallWindowsProc function
        sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                   'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                     'Create the first sc_aSubData element

    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If (nSubIdx = -1) Then                                                    'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                                   'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                  'Create a new sc_aSubData element
        End If

        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)

        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                             'Allocate memory for the machine code WndProc
        Call VirtualProtect(ByVal .nAddrSub, CODE_LEN, PAGE_EXECUTE_READWRITE, i)  'Mark memory as executable
        Call RtlMoveMemory(ByVal .nAddrSub, sc_aBuf(1), CODE_LEN)                 'Copy the machine code from the static byte array to the code array in sc_aSubData
        .hWnd = lng_hWnd                                                          'Store the hWnd
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                'Set our WndProc in place

        Call zPatchRel(.nAddrSub, PATCH_01, sc_pEbMode)                           'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                           'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, sc_pSWL)                              'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                           'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, sc_pCWP)                              'Patch the relative address of the CallWindowProc api function
    End With
End Function

Private Sub Subclass_StopAll()
'Stop all subclassing
  
  Dim i As Long
  
    i = UBound(sc_aSubData())                                                     'Get the upper bound of the subclass data array
    Do While i >= 0                                                               'Iterate through each element
        With sc_aSubData(i)
            If (.hWnd <> 0) Then                                                  'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hWnd)                                         'Subclass_Stop
            End If
        End With
    
        i = i - 1                                                                 'Next element
    Loop
End Sub

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Stop subclassing the passed window handle
'Parameters:
'   lng_hWnd - The handle of the window to stop being subclassed
  
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                       'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                                    'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                                    'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                                                'Release the machine code memory
        .hWnd = 0                                                                 'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                             'Clear the before table
        .nMsgCntA = 0                                                             'Clear the after table
        Erase .aMsgTblB                                                           'Erase the before table
        Erase .aMsgTblA                                                           'Erase the after table
    End With
End Sub

'----------------------------------------------------------------------------------------
'These z??? routines are exclusively called by the Subclass_??? routines.
'----------------------------------------------------------------------------------------

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
'Worker sub for Subclass_AddMsg
  
  Dim nEntry  As Long                                                             'Message table entry index
  Dim nOff1   As Long                                                             'Machine code buffer offset 1
  Dim nOff2   As Long                                                             'Machine code buffer offset 2
  
    If (uMsg = ALL_MESSAGES) Then                                                 'If all messages
        nMsgCnt = ALL_MESSAGES                                                    'Indicates that all messages will callback
      Else                                                                        'Else a specific message number
        Do While nEntry < nMsgCnt                                                 'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1
        
            If (aMsgTbl(nEntry) = 0) Then                                         'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                            'Re-use this entry
                Exit Sub                                                          'Bail
            ElseIf (aMsgTbl(nEntry) = uMsg) Then                                  'The msg is already in the table!
                Exit Sub                                                          'Bail
            End If
        Loop                                                                      'Next entry

        nMsgCnt = nMsgCnt + 1                                                     'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                              'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                   'Store the message number in the table
    End If

    If (When = eMsgWhen.MSG_BEFORE) Then                                          'If before
        nOff1 = PATCH_04                                                          'Offset to the Before table
        nOff2 = PATCH_05                                                          'Offset to the Before table entry count
      Else                                                                        'Else after
        nOff1 = PATCH_08                                                          'Offset to the After table
        nOff2 = PATCH_09                                                          'Offset to the After table entry count
    End If

    If (uMsg <> ALL_MESSAGES) Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                          'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                                         'Patch the appropriate table entry count
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
'Return the memory address of the passed function in the passed dll
    
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                                        'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
''Worker sub for Subclass_DelMsg
'
'  Dim nEntry As Long
'
'    If (uMsg = ALL_MESSAGES) Then                                                 'If deleting all messages
'        nMsgCnt = 0                                                               'Message count is now zero
'        If When = eMsgWhen.MSG_BEFORE Then                                        'If before
'            nEntry = PATCH_05                                                     'Patch the before table message count location
'          Else                                                                    'Else after
'            nEntry = PATCH_09                                                     'Patch the after table message count location
'        End If
'        Call zPatchVal(nAddr, nEntry, 0)                                          'Patch the table message count to zero
'      Else                                                                        'Else deleteting a specific message
'        Do While nEntry < nMsgCnt                                                 'For each table entry
'            nEntry = nEntry + 1
'            If (aMsgTbl(nEntry) = uMsg) Then                                      'If this entry is the message we wish to delete
'                aMsgTbl(nEntry) = 0                                               'Mark the table slot as available
'                Exit Do                                                           'Bail
'            End If
'        Loop                                                                      'Next entry
'    End If
'End Sub

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the sc_aSubData() array index of the passed hWnd
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                            'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If (.hWnd = lng_hWnd) Then                                            'If the hWnd of this element is the one we're looking for
                If (Not bAdd) Then                                                'If we're searching not adding
                    Exit Function                                                 'Found
                End If
            ElseIf (.hWnd = 0) Then                                               'If this an element marked for reuse.
                If (bAdd) Then                                                    'If we're adding
                    Exit Function                                                 'Re-use it
                End If
            End If
        End With
        zIdx = zIdx - 1                                                           'Decrement the index
    Loop
  
    If (Not bAdd) Then
        Debug.Assert False                                                        'hWnd not found, programmer error
    End If

'If we exit here, we're returning -1, no freed elements were found
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
'Patch the machine code buffer at the indicated offset with the relative address to the target address.
    
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
'Patch the machine code buffer at the indicated offset with the passed value
    
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
'Worker function for Subclass_InIDE
    
    zSetTrue = True
    bValue = True
End Function

'========================================================================================
'========================================================================================

#If False Then
   Public SHCNE_RENAMEITEM
   Public SHCNE_CREATE
   Public SHCNE_DELETE
   Public SHCNE_MKDIR
   Public SHCNE_RMDIR
   Public SHCNE_MEDIAINSERTED
   Public SHCNE_MEDIAREMOVED
   Public SHCNE_DRIVEREMOVED
   Public SHCNE_DRIVEADD
   Public SHCNE_NETSHARE
   Public SHCNE_NETUNSHARE
   Public SHCNE_ATTRIBUTES
   Public SHCNE_UPDATEDIR
   Public SHCNE_UPDATEITEM
   Public SHCNE_SERVERDISCONNECT
   Public SHCNE_UPDATEIMAGE
   Public SHCNE_DRIVEADDGUI
   Public SHCNE_RENAMEFOLDER
   Public SHCNE_FREESPACE
   Public SHCNE_ASSOCCHANGED
   Public SHCNE_DISKEVENTS
   Public SHCNE_GLOBALEVENTS
   Public SHCNE_ALLEVENTS
   Public SHCNE_INTERRUPT
   
   Public SHCNE_EXTENDED_EVENT
   
   Public SHCNEE_ORDERCHANGED
   Public SHCNEE_MSI_CHANGE
   Public SHCNEE_MSI_UNINSTALL
   
   
   Public SHCNF_IDLIST
   Public SHCNF_PATHA
   Public SHCNF_PRINTERA
   Public SHCNF_DWORD
   Public SHCNF_PATHW
   Public SHCNF_PRINTERW
   Public SHCNF_TYPE
   Public SHCNF_FLUSH
   Public SHCNF_FLUSHNOWAIT
   Public SHCNF_PATH
   Public SHCNF_PRINTER
   Public SHCNF_PATH
   Public SHCNF_PRINTER
   
#End If

'========================================================================================
'========================================================================================

