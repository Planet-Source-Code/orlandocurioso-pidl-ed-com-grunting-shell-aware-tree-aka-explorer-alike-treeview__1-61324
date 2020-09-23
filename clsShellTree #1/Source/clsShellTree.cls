VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShellTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsShellTree
' Author    : OrlandoCurioso 08.04.2005
' Credits   : "Brought to you by Brad Martinez"                         http://btmtz.mvps.org/index.html
'             EnumDeskVB,VBExplorer,SHChangeNotify,SHContextMenu demos, ISHF_Ex.tlb #1.2
'             Nikolaos Achilleos Bozinis (Shell Explorer's Cookbook)    http://netez.com/2xExplorer/shellFAQ/shmain.html
'             Eduardo Morcillo (IDataObject weak reference)             http://www.mvps.org/emorcillo/en/code/vb6/getidataobject.shtml
'             Roman Engels (IContextMenu3 handling)                     http://codetools.com/shell/shellcontextmenu.asp
'
' Purpose   : Using ucTreeView as FolderBrowser
'
'             Drag & Drop Image:
'             - IDragHelper degrades gracefully for <Win2K.
'
' Info      : Shell Notifications                  http://www.geocities.com/SiliconValley/4942/notify.html
'             Shell Functions (version info!)      http://members.ozemail.com.au/~geoffch/samples/win32/shell/shell32/functions/
'             Shell Explorer's Cookbook            http://netez.com/2xExplorer/shellFAQ/shmain.html
'             MSDN: Shell Programmer's Guide, Shell Reference, OLE and Data Transfer
'
' GUID's    : Desktop            ""
'             My Computer        "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
'             My Documents       "::{450d8fba-ad25-11d0-98a8-0800361b1103}"
'             Recycle Bin        "::{645FF040-5081-101B-9F08-00AA002F954E}"
'
' ToDo      :
'             Implement Shell Change Notifications (SCN) for networks.
'             Implement code reacting to file operations, when used without SCN.
'             MySEH_ERROR
'             Q235630 PRB: SHGetFileInfo Caches Drive Information (<= Win2K)
'             Same drop cursors as explorer: # which resource are they loaded from ??? #
'             OCX package
'
' CanDo     : SHRestricted api >= Win2K
'             SHGetAttributesFromDataObject   >= WinXP
'             IDropTarget by hwnd : "OleDropTargetInterface" window property, "OleDropTargetMarshalHwnd" ?
'             VB Shell object: TypeLib 'Microsoft Internet Controls' SHDocVw.dll
'
' Requires  : modSortTreeCB
'             IShellFolder Extended Type Library v1.3 OC (ISHF_Ex.tlb)
'             clsShellNotify (SHCN = 1)
'             clsDropSource,modDropSource (OLEDD = 1)
' Test      :
'             Unicode
'             Win2K
'             SHContextMenuCmdIDs const for all OS ?
'             Network :
'             - DisplayName(SHGDN_NORMAL and SHGDN_INFOLDER) interchanged for items ???
'             - DisplayName(SHGDN_FORPARSING) non localized name of 'Entire Network' ?
' Bugs      :
'             ISF.CreateViewObject on Win98 crashes (== 'New Folder' button not possible!)
'---------------------------------------------------------------------------------------
' Update #1 :
'             IDropSource: clsDropSource as private class, no dll needed
'             Get icon indices with SHMapPIDLToSystemImageListIndex

Option Explicit

' Project conditional compilation constants
'#Const UNICODE = 1
'#Const WIN32_IE = 1024  ' 1024 == &H400 >= IE4 >= Shell32 4.71 >= Win98,(NT4)
'                        ' 1280 == &H500 >= IE5 >= Shell32 5.00 >= Win2k
                        
' Private conditional compilation constants (use same settings as ucTreeView,FLDBR = 1 & DDIMG = 0)
#Const CUSTDRAW = 1     ' OwnerDraw
#Const MULSEL = 1       ' Multiple Selection
#Const OLEDD = 1        ' OLE Drag & Drop             (requires clsDropSource,modDropSource)

#Const SHCN = 1         ' Shell Change Notifications  (requires clsShellNotify)

'// general

Private Type RECT2
   X1 As Long
   Y1 As Long
   X2 As Long
   Y2 As Long
End Type

Private Type POINTAPI
   X  As Long
   Y  As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT2, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
   Private Const SM_CXDRAG = 68
'   Private Const SM_CYDRAG = 69

'//

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Private Declare Function IsBadReadPtr Lib "kernel32" (ByRef lp As Any, ByVal ucb As Long) As Long
'Private Declare Function IsBadWritePtr Lib "kernel32" (ByRef lp As Any, ByVal ucb As Long) As Long
    
'Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

'Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
'Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
'Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

'// OLE D&D

Private Declare Function CoCreateInstance Lib "ole32" (rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As Any, ppv As Any) As Long
   Private Const CLSCTX_INPROC_SERVER As Long = 1&

Private Declare Function DoDragDrop Lib "ole32" (ByVal pDataObject As Long, ByVal pDropSource As Long, ByVal AllowedEffects As Long, Effects As Long) As Long
'Private Declare Function DoDragDrop Lib "ole32" (pDataObject As IDataObject, pDropSource As IDropSource, ByVal AllowedEffects As Long, ByVal  pEffects As Long) As Long
   Private Const DRAGDROP_S_DROP = &H40100
   Private Const DRAGDROP_S_CANCEL = &H40101


'   Private Const CFSTR_SHELLIDLIST = "Shell IDList Array"
'   Private Const CFSTR_SHELLIDLISTOFFSET = "Shell Object Offsets"
'   Private Const CFSTR_NETRESOURCES = "Net Resource"
'   Private Const CFSTR_FILEDESCRIPTORA = "FileGroupDescriptor"
'   Private Const CFSTR_FILEDESCRIPTORW = "FileGroupDescriptorW"
'   Private Const CFSTR_FILECONTENTS = "FileContents"
'   Private Const CFSTR_FILENAMEA = "FileName"
'   Private Const CFSTR_FILENAMEW = "FileNameW"
'   Private Const CFSTR_PRINTERGROUP = "PrinterFriendlyName"
'   Private Const CFSTR_FILENAMEMAPA = "FileNameMap"
'   Private Const CFSTR_FILENAMEMAPW = "FileNameMapW"
'   Private Const CFSTR_SHELLURL = "UniformResourceLocator"
'   Private Const CFSTR_INETURLA = CFSTR_SHELLURL
'   Private Const CFSTR_INETURLW = "UniformResourceLocatorW"
'   Private Const CFSTR_PREFERREDDROPEFFECT = "Preferred DropEffect"
'   Private Const CFSTR_PERFORMEDDROPEFFECT = "Performed DropEffect"
'   Private Const CFSTR_PASTESUCCEEDED = "Paste Succeeded"
   Private Const CFSTR_INDRAGLOOP = "InShellDragLoop"
'   Private Const CFSTR_MOUNTEDVOLUME = "MountedVolume"
'   Private Const CFSTR_PERSISTEDDATAOBJECT = "PersistedDataObject"
'   Private Const CFSTR_TARGETCLSID = "TargetCLSID"
'   Private Const CFSTR_LOGICALPERFORMEDDROPEFFECT = "Logical Performed DropEffect"
'   Private Const CFSTR_AUTOPLAY_SHELLIDLISTS = "Autoplay Enumerated IDList Array"

'   Private Const CFSTR_DRAGCONTEXT = "DragContext"       ' internal to the shell drag/drop helpers
'   Private Const CFSTR_DRAGBITS = "DragImageBits"        ' internal to the shell drag/drop helpers

'// shlwapi.dll Unicode + ANSI always exist

Private Declare Function PathFileExistsW Lib "Shlwapi" (ByVal lpszPath As Long) As Boolean
Private Declare Function PathIsDirectoryW Lib "Shlwapi" (ByVal lpszPath As Long) As Boolean

'//
'Private Declare Function SetErrorMode Lib "kernel32.dll" (ByVal wMode As Long) As Long
'
'Private Const SEM_FAILCRITICALERRORS As Long = &H1&
'Private Const SEM_NOGPFAULTERRORBOX As Long = &H2&
'Private Const SEM_NOALIGNMENTFAULTEXCEPT As Long = &H4&
'Private Const SEM_NOOPENFILEERRORBOX As Long = &H8000&

'// ShellExecute

Private Declare Function ShellExecuteEx Lib "shell32.dll" (lpExecInfo As SHELLEXECUTEINFO) As Long

Private Type SHELLEXECUTEINFO
  cbSize       As Long
  fMask        As SEE_MASK
  hWnd         As Long
  lpVerb       As Long
  lpFile       As Long             ' lp's changed to pointers to strings
  lpParameters As Long
  lpDirectory  As Long
  nShow        As SW_cmds
  hInstApp     As Long
  '  Optional fields
  lpIDList     As Long
  lpClass      As Long
  hkeyClass    As Long
  dwHotKey     As Long
  hIcon        As Long
  hProcess     As Long
End Type

Private Enum SEE_MASK
   SEE_MASK_CLASSNAME = &H1
   SEE_MASK_CLASSKEY = &H3      ' CLASSKEY overrides CLASSNAME
   SEE_MASK_IDLIST = &H4
   SEE_MASK_INVOKEIDLIST = &HC  ' INVOKEIDLIST overrides IDLIST
   SEE_MASK_ICON = &H10
   SEE_MASK_HOTKEY = &H20
   SEE_MASK_NOCLOSEPROCESS = &H40
   SEE_MASK_CONNECTNETDRV = &H80
   SEE_MASK_FLAG_DDEWAIT = &H100
   SEE_MASK_DOENVSUBST = &H200
   SEE_MASK_FLAG_NO_UI = &H400
End Enum

Private Enum SW_cmds
  SW_HIDE = 0
  SW_NORMAL = 1
  SW_SHOWNORMAL = 1
  SW_SHOWMINIMIZED = 2
  SW_MAXIMIZE = 3
  SW_SHOWMAXIMIZED = 3
  SW_SHOWNOACTIVATE = 4
  SW_SHOW = 5
  SW_MINIMIZE = 6
  SW_SHOWMINNOACTIVE = 7
  SW_SHOWNA = 8
  SW_RESTORE = 9
  SW_MAX = 10
  SW_SHOWDEFAULT = 10
End Enum

'// Shell

Public Enum SSF_Flags
   SSF_ALL = -1                     ' &HFFFF&/&HFFFFFFFF
   SSF_SHOWALLOBJECTS = &H1
   SSF_SHOWEXTENSIONS = &H2
   SSF_HIDDENFILEEXTS = &H4
'   SSF_SERVERADMINUI = &H4
   SSF_SHOWCOMPCOLOR = &H8
   SSF_SORTCOLUMNS = &H10           ' SHGetSetSettings
   SSF_SHOWSYSFILES = &H20
   ' &H40 unused
   SSF_DOUBLECLICKINWEBVIEW = &H80
   SSF_SHOWATTRIBCOL = &H100
   SSF_DESKTOPHTML = &H200
   SSF_WIN95CLASSIC = &H400
   SSF_DONTPRETTYPATH = &H800
   SSF_MAPNETDRVBUTTON = &H1000
   SSF_SHOWINFOTIP = &H2000
   SSF_HIDEICONS = &H4000&          ' unused
   SSF_NOCONFIRMRECYCLE = &H8000&   ' as unsigned Long !!!
   ' SHGetSetSettings
   SSF_FILTER = &H10000
   SSF_WEBVIEW = &H20000            ' "Show common tasks in folders" / "Use Windows classic folders"
   SSF_SHOWSUPERHIDDEN = &H40000
   SSF_SEPPROCESS = &H80000
   SSF_NONETCRAWLING = &H100000
   SSF_STARTPANELON = &H200000
   SSF_SHOWSTARTPAGE = &H400000
End Enum

' HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced
'Private Declare Sub SHGetSettings Lib "shell32" (lpsfs As SHELLFLAGSTATE, ByVal dwMask As Long)
Private Declare Sub SHGetSettingsAny Lib "shell32" Alias "SHGetSettings" (lpsfs As Any, ByVal dwMask As Long)
' SHGetSetSettings ordinal #68

#If WIN32_IE >= &H500 Then
Private Declare Function SHMapPIDLToSystemImageListIndex Lib "shell32" (ByVal psf As IShellFolder, ByVal pidl As Long, piIndexSel As Long) As Long
Private Declare Function ILFindLastID Lib "shell32" (ByVal pidl As Long) As Long
#End If

'// Implementation

Private Const vbDropEffectLink = 4&  ' DROPEFFECT_LINK

Public Enum eFTOptions  ' OR'ed
   ftFileOperations = 1
   ftContextMenu = 2
   ftSHNotify = 4
   ftCollapseReset = 8
End Enum

Private Enum eOpenNodeResult
   on_Valid = 0
   on_Sucess
   on_Partial
   on_Error
End Enum

Private Type SHNOTIFY_LAST
   hNode1      As Long
   hNode2      As Long
   hParent1    As Long
   hParent2    As Long
End Type

#If SHCN Then
Private WithEvents m_cSN      As clsShellNotify
Attribute m_cSN.VB_VarHelpID = -1
#End If

#If OLEDD Then
Private m_IDataObject         As IDataObject
Private m_IDropSource         As IDropSource       ' clsDropSource/modDropSource
Private m_IDropTarget         As IDropTarget
Private m_IDTHelper           As IDropTargetHelper

Private m_bInDrag             As OLE_TRISTATE
Private m_tPDrag              As POINTAPI
Private m_hNodeDrag           As Long
#End If

Private WithEvents m_ucTree   As ucTreeView
Attribute m_ucTree.VB_VarHelpID = -1
Private m_hOwner              As Long

Private m_sPath               As String
Private m_pidlFQ              As Long
'Private m_pidlRel             As Long

Private m_eIncludeItems       As ESHCONTF
Private m_eOptions            As eFTOptions
Private m_bShowCompressed     As Boolean

Private m_bInEdit             As OLE_TRISTATE
Private m_sEdit               As String

Private m_tSHLast             As SHNOTIFY_LAST

Private m_hNodeCut            As Long
Private m_hNodeAdded          As Long

Private m_idxNewFile          As Long
'

' OR'ed
Public Property Get Options() As eFTOptions
   Options = m_eOptions
End Property
Public Property Let Options(ByVal eOptions As eFTOptions)

   m_eOptions = eOptions
   
#If SHCN Then
   If (m_eOptions And ftSHNotify) Then
'         pRegisterNotify
   Else
      pUnregisterNotify
   End If
#End If

   ' Renaming items
   If Not (m_ucTree Is Nothing) Then
      m_ucTree.LabelEdit = CBool(m_eOptions And ftFileOperations)
   End If
End Property

Public Property Get IncludeItems() As ESHCONTF
   IncludeItems = m_eIncludeItems
End Property
Public Property Let IncludeItems(ByVal eIncludeItems As ESHCONTF)
   If eIncludeItems Then
      m_eIncludeItems = eIncludeItems
   Else
      ' as default
      m_eIncludeItems = SHCONTF_FOLDERS
   End If
End Property

Public Property Get DisplayName(Optional ByVal uFlags As ESHGNO = _
                                SHGDN_FORPARSING Or SHGDN_FORADDRESSBAR) As String
   If m_pidlFQ Then
      DisplayName = GetFolderDisplayName(isfDesktop, m_pidlFQ, uFlags)
   End If
End Property

' Expects and returns:
' - Physical path    ie 'C:\', 'C:' is OK
' - Paths containing GUID's  ie '::{208D2C60-3AEA-1069-A2D7-08002B30309D}'
' - UNC paths local ok, (remote) server ???
' - With nonexisting or empty passed path selects & expands root. (no error raised)
Public Property Get CurrentPath() As String
   ' set in m_ucTree_BeforeSelectionChange
   CurrentPath = m_sPath
End Property
Public Property Let CurrentPath(ByVal sCurrentPath As String)
   Dim pidlFQ  As Long
   
   If m_sPath <> sCurrentPath Then
      
      If LenB(sCurrentPath) Then
         
         If Len(sCurrentPath) = 2 Then
            If Right$(sCurrentPath, 1) = ":" Then
               ' add backslash for drive
               sCurrentPath = sCurrentPath & "\"
            End If
         End If

         pidlFQ = GetPIDLFromPath(m_hOwner, sCurrentPath)
      End If
      
      OpenNode pidlFQ, , bExpand:=True
      FreePIDL pidlFQ

   End If
End Property

' Don't free returned pidl
Public Property Get CurrentPidl() As Long
   ' set in m_ucTree_BeforeSelectionChange
   CurrentPidl = m_pidlFQ
End Property
' Expects fully qualified pidl freed by caller
Public Property Let CurrentPidl(ByVal pidlFQ As Long)
   
   If ValidatePidl(pidlFQ) Then
      OpenNode pidlFQ, , bExpand:=True
   End If
End Property

Public Property Get CurrentISF(Optional ByVal bRtnDesktop As Boolean) As IShellFolder

   If m_pidlFQ Then
      Set CurrentISF = GetIShellFolder(isfDesktop, m_pidlFQ, bRtnDesktop)
   End If
End Property

' hOwner : Owner of messageboxes. If zero, is Desktop.
Public Sub Initialize(oTree As Object, ByVal hOwner As Long)   ' oTree As ucTreeView
   
   m_hOwner = hOwner
   
   Set m_ucTree = oTree
   With m_ucTree
      .hImageList = GetSystemImagelist(SHGFI_SMALLICON)
      
      .LabelEdit = CBool(m_eOptions And ftFileOperations)   ' Renaming items
      
#If OLEDD Then
      .OLEDragAutoExpand = True
      .OLEDragInsertStyle = disDropHilite
      .OLEDragMode = drgManual
      .OLEDropMode = drpManual
#End If
      
#If CUSTDRAW Then
      If GetExplorerSettings(SSF_SHOWCOMPCOLOR) Then
         m_bShowCompressed = True
         .DoCustomDraw = .DoCustomDraw Or cdColor
      Else
         m_bShowCompressed = False
      End If
#End If

#If MULSEL Then
      .DoCustomDraw = .DoCustomDraw Or cdColor
      .MultiSelect = True
#End If

   End With
   
   ' Store uncounted reference for sorting callback's
   modSortTreeCB.FolderTree = m_ucTree
   
   ' Get icon index in system imagelist for new items once
#If WIN32_IE >= &H500 Then
   m_idxNewFile = GetDefaultFileIconIndex(vbNullString, vbNormal, SHGFI_SMALLICON)
   Debug.Assert m_idxNewFile = 3    ' # ? #
#Else
   m_idxNewFile = 0
#End If

End Sub

' ----------------------------------------------------------------------------------- '
'  Populate TreeView
' ----------------------------------------------------------------------------------- '

Public Sub Clear()

#If SHCN Then
   pUnregisterNotify
#End If
   pFreePidls
   m_ucTree.Clear
   m_sPath = vbNullString
End Sub

' pidlRoot freed by caller
Public Function AddRoot(Optional ByVal csidl As eCSIDL_VALUES = CSIDL_DESKTOP, _
                        Optional ByVal pidlRoot As Long = -1, _
                        Optional ByVal RootPath As String = vbNullString, _
                        Optional ByVal OnlySubDirs As Boolean = False) As Boolean
   Dim hRoot   As Long
   Dim bNoFree As Boolean
   
   Debug.Assert Not (m_ucTree Is Nothing)  ' Initialized ?
   
   Me.Clear
   
   Select Case True
   
      Case LenB(RootPath) <> 0
         If FolderExists(RootPath) Then
            pidlRoot = GetPIDLFromPath(m_hOwner, RootPath)
         Else
            Debug.Assert False
            pidlRoot = 0
            csidl = CSIDL_DESKTOP
         End If
         
      Case pidlRoot > 0
         ' ### test pidl ###
         bNoFree = True
         
      Case csidl >= CSIDL_DESKTOP
         '
         
      Case Else
         Debug.Assert False
         pidlRoot = 0
         csidl = CSIDL_DESKTOP
       
   End Select
   
   If pidlRoot <= 0 Then
      ' Get the pidl of the special folder
      pidlRoot = GetPIDLFromCSIDL(m_hOwner, csidl)
   End If
   
   If pidlRoot Then
      
      With m_ucTree
      
         If Not OnlySubDirs Then
         
            hRoot = pInsertRootFolder(pidlRoot)
            ' Expand the root, invoking pInsertFolderItems()
            .Expand hRoot
            
         Else
            
            ' ### check if pidlRoot has subfolders ###
            pInsertFolderItems pidlRoot, 0
            hRoot = .NodeRoot
         
         End If
         
         If hRoot Then
            .SelectedNode = hRoot
            AddRoot = True
         Else
            Debug.Assert .NodeCount = 0
            Debug.Assert False
         End If
      
      End With
      
#If SHCN Then
      If CBool(m_eOptions And ftSHNotify) Then
      
         If AddRoot Then
            ' Receive shell change notifications for tree contents
            pRegisterNotify pidlRoot
         End If
      End If
#End If
      
      If Not bNoFree Then
         ' Free our pidl. pInsertRootFolder/pInsertFolderItems() copied it...
         FreePIDL pidlRoot
      End If
      
   End If
End Function

Public Sub Refresh()
   Dim hNode   As Long
   
   On Error GoTo Proc_Error

   With m_ucTree
      
      .Redraw = False
      
      ' Loop through all nodes:
      ' - for non-expanded nodes remove any children (do first!)
      ' - for expanded nodes update folder contents
      
      hNode = .NodeRoot
      Do While hNode
         
         If Not .NodeExpanded(hNode) Then
            If .NodeChild(hNode) Then
               .Collapse hNode, , RemoveChildren:=True
            End If
         Else
            pUpdateFolder hNode
         End If
         
         hNode = .NodeNext(hNode)
      Loop
      
      .Redraw = True
   End With

   Exit Sub

Proc_Error:
   m_ucTree.Redraw = True
   Debug.Print "Error: " & Err.Number & ". " & Err.Description, vbOKOnly Or vbCritical, App.Title & ".clsShellTree: Sub Refresh"
   If InIDE Then Stop: Resume
End Sub

' Deletes stale or renamed items, adds new or renamed items.
Private Sub pUpdateFolder(hParent As Long)
   Dim isfParent     As IShellFolder
   Dim iEIDL         As IEnumIDList
   Dim pidlFQParent  As Long
   Dim pidlFQEn      As Long
   Dim pidlRelEn     As Long
   Dim pidlFQ        As Long
   Dim pidlRel       As Long
   Dim hChild        As Long
   Dim bInsert       As Boolean
   Dim bSort         As Boolean
   
   With m_ucTree
      Debug.Assert .Redraw = False
      
      ' Get the parent's IShellFolder from its fully qualified pidl
      pDecryptNodeKey .NodeKey(hParent), pidlFQParent
      Set isfParent = GetIShellFolder(isfDesktop, pidlFQParent)
      
      ' delete all nodes whose item doesn't exist anymore
      ' will also delete all renamed items
      hChild = .NodeChild(hParent)
      Do While hChild
         
         pDecryptNodeKey .NodeKey(hChild), , pidlRel
         
         If Not pValidatePidl(isfParent, pidlRel) Then
            Debug.Print "DELETED", .NodeText(hChild)
            .DeleteNode hChild
         End If
         
         hChild = .NodeNextSibling(hChild)
      Loop
      
      ' Create an enumeration object for the parent folder.
      ' Enumeration contains updated contents from shell (no need to get a fresh pidl)
      If S_OK = isfParent.EnumObjects(m_hOwner, m_eIncludeItems, iEIDL) Then
         
         ' Enumerate the contents of the parent folder
         Do While (iEIDL.Next(1, pidlRelEn, 0) = NOERROR)
            
            bInsert = True
            
            ' try finding node for this item
            hChild = .NodeChild(hParent)
            Do While hChild
               
               pDecryptNodeKey .NodeKey(hChild), pidlFQ, pidlRel
               
               If (S_OK = isfParent.CompareIDs(0&, pidlRelEn, pidlRel)) Then
                  ' node for this item exists
                  bInsert = False
                  Exit Do
               End If
               
               hChild = .NodeNextSibling(hChild)
            Loop
            
            If bInsert Then
               ' Create a fully qualified pidl for the current item. $SHGetRealIDL
               pidlFQEn = CombinePIDLs(pidlFQParent, pidlRelEn)
               If pidlFQEn Then
               
                  ' Insert the item under the parent folder.
                  hChild = pInsertItem(isfParent, pidlFQEn, pidlRelEn, hParent)
                  If hChild Then
                     ' inserted at least one item
                     bSort = True
                     Debug.Print "ADDED", .NodeText(hChild)
                  End If
                  
                  ' Free the current item's absolute pidl we created.
                  FreePIDL pidlFQEn
               End If
               
            Else
               Debug.Assert hChild
               Debug.Assert pidlFQ
               
                  ' pUpdateNodePIDL frees existing node pidls and assigns new !
                  ' any renamed items will be readded
                  pUpdateNodePIDL hParent, pidlFQParent, hChild, pidlFQ _
                                  , bSort:=False, bRename:=False
                                  
                  ' freed by pUpdateNodePIDL
                  pidlFQ = 0: pidlRel = 0
               
            End If
            
            ' Free the relative pidl the enumeration gave us.
            FreePIDL pidlRelEn
            
            If EscapePressed Then Exit Do
            
         Loop   ' ieidl.Next
       
      End If   ' SUCCEEDED(EnumObjects))
    
      If bSort Then
         ' Setup the callback and sort the parent folder
         m_ucTree.SortChildrenCB hParent, AddressOf modSortTreeCB.TreeViewCompareProc, _
                                 False, ObjPtr(isfParent)
      ElseIf .NodeChild(hParent) = 0 Then
         ' Remove its button, if shell notifies us about any changes
         If (m_eOptions And ftSHNotify) Then
            m_ucTree.NodePlusMinusButton(hParent) = False
         End If
      End If
      
   End With
   
End Sub

Private Function pInsertRootFolder(pidlFQ As Long) As Long
   Dim pidlRel As Long
  
   ' Get the specified root folder's relative pidl (free it, pInsertItem copies it).
   pidlRel = GetItemID(pidlFQ, GIID_LAST)
   If pidlRel Then
      ' Insert the root folder (both pidls are copied in the call)
      pInsertRootFolder = pInsertItem(GetIShellFolderParent(pidlFQ), pidlFQ, pidlRel, 0)
      FreePIDL pidlRel
   End If
   
End Function

Private Function pInsertItem(isfParent As IShellFolder, _
                             ByVal pidlFQ As Long, ByVal pidlRel As Long, _
                             ByVal hParent As Long) As Long
   Dim ulAttrs    As ESFGAO
   Dim hNode      As Long
   Dim idxImg     As Long
   Dim idxSelImg  As Long
   Dim sName      As String
   Dim sKey       As String      ' pidlFQ & pidlRel stored in NodeKey
   Dim bBtn       As Boolean
'   Dim sTag       As String
   
   On Error GoTo Proc_Error

   ' Get the item's attributes, specifiy what attributes we want.
   ulAttrs = SFGAO_HASSUBFOLDER Or SFGAO_HIDDEN Or _
             SFGAO_SHARE Or SFGAO_LINK ' Or SFGAO_GHOSTED
             
   If m_bShowCompressed Then ulAttrs = ulAttrs Or SFGAO_COMPRESSED
   
   If Not SUCCEEDED(isfParent.GetAttributesOf(1, pidlRel, ulAttrs)) Then
      ' error: ie &H80004005 Unknown error for empty drive
      If (m_eOptions And ftSHNotify) Then
         ' Shell will notify us about any change: no button
         ulAttrs = 0
      Else
         ' give node a button
         ulAttrs = SFGAO_HASSUBFOLDER
      End If
   End If
   
  ' Add the Node to the TreeView
   sName = GetFolderDisplayName(isfParent, pidlRel, SHGDN_INFOLDER)
   bBtn = CBool(ulAttrs And SFGAO_HASSUBFOLDER)
   
#If WIN32_IE >= &H500 Then
   idxImg = SHMapPIDLToSystemImageListIndex(isfParent, pidlRel, idxSelImg)
#Else
   idxImg = GetFileIconIndexPIDL(pidlFQ, SHGFI_SMALLICON)
   idxSelImg = GetFileIconIndexPIDL(pidlFQ, SHGFI_SMALLICON Or SHGFI_OPENICON)
#End If

   If idxSelImg = m_idxNewFile Then
      If idxImg <> idxSelImg Then
         ' selected icon should show as new/unknown item, but normal icon not ?
         ' don't believe it : bug with '.cpl' applets of Control Panel
         idxSelImg = idxImg
      End If
   End If

   ' copy the pidls and store them as NodeKey
   sKey = CStr(CopyPIDL(pidlFQ)) & ":" & CStr(CopyPIDL(pidlRel))
   
   With m_ucTree
   
      ' Add the (root) item to the treeview, returning it's hNode
      If (hParent = 0) Then
         hNode = .AddNode(, , sKey, sName, idxImg, idxSelImg, bBtn)
      Else
         hNode = .AddNode(hParent, rLast, sKey, sName, idxImg, idxSelImg, bBtn)
      End If
       
      ' ghosted image for hidden files
      If (ulAttrs And SFGAO_HIDDEN) Then .NodeGhosted(hNode) = True
      
      ' system imagelist overlay images:
      ' share overlay is the 1st, shortcut is 2nd, gray arrow is 3rd, no 4th image
      ' # §SHGetIconOverlayIndex #
      If (ulAttrs And SFGAO_SHARE) Then
         .NodeOverlayImage(hNode) = 1
      ElseIf (ulAttrs And SFGAO_LINK) Then
         .NodeOverlayImage(hNode) = 2
      End If
      
#If CUSTDRAW Then
      If (ulAttrs And SFGAO_COMPRESSED) Then
         ' # ? HowTo get color value for compressed files ? #
         .NodeForecolor(hNode) = vbBlue
      End If
#End If
      
   End With
   
   ' used with NewFolder command
   m_hNodeAdded = hNode
   
   pInsertItem = hNode
   
   Exit Function

Proc_Error:
   Debug.Print "Error: " & Err.Number & ". " & Err.Description
   If InIDE Then Stop: Resume
End Function

Private Sub pInsertFolderItems(pidlFQParent As Long, hParent As Long)
   Dim isfParent  As IShellFolder
   Dim iEIDL      As IEnumIDList
   Dim pidlRelEn  As Long
   Dim pidlFQ     As Long
   Dim hChild     As Long
   
   Debug.Assert m_ucTree.NodeChild(hParent) = 0
   
   ' Get the parent's IShellFolder from its fully qualified pidl
   Set isfParent = GetIShellFolder(isfDesktop, pidlFQParent)
   
   ' Create an enumeration object for the parent folder.
   ' .EnumObjects() complies with folder setting: Hide extension for registered files.
   ' .GetDisplayNameOf() does not: always returns with extension!
   If S_OK = isfParent.EnumObjects(m_hOwner, m_eIncludeItems, iEIDL) Then
                                                                
      ' Enumerate the contents of the parent folder
      Do While (iEIDL.Next(1, pidlRelEn, 0) = NOERROR)
         
         ' Create a fully qualified pidl for the current item. $SHGetRealIDL
         pidlFQ = CombinePIDLs(pidlFQParent, pidlRelEn)
         If pidlFQ Then
         
            ' Insert the item under the parent folder.
            hChild = pInsertItem(isfParent, pidlFQ, pidlRelEn, hParent)
            
            ' Free the current item's absolute pidl we created.
            FreePIDL pidlFQ
         End If
         
         ' Free the relative pidl the enumeration gave us.
         FreePIDL pidlRelEn
         
         If EscapePressed Then Exit Do
         
      Loop   ' ieidl.Next
    
   End If   ' SUCCEEDED(EnumObjects))
    
   If hChild Then
      ' Setup the callback and sort the parent folder
      m_ucTree.SortChildrenCB hParent, AddressOf modSortTreeCB.TreeViewCompareProc, _
                              False, ObjPtr(isfParent)
   ElseIf hParent Then
      ' The parent folder is expanding yet it has no subfolders. We'll assume
      ' that we've encountered a network folder which ISF::GetAttributesOf
      ' couldn't resolve when the folder was originally inserted.
      
      ' Remove its button, if shell notifies us about any changes
      If (m_eOptions And ftSHNotify) Then
         m_ucTree.NodePlusMinusButton(hParent) = False
      End If
   End If
   
End Sub

' ----------------------------------------------------------------------------------- '
'  The Monster
' ----------------------------------------------------------------------------------- '

' Expects fully qualified pidl freed by caller.
' pidlOpen = 0 is valid and corresponds to desktop or tree root.
' If sucessful, returns matching hNode. Selects and ensures visibility (bExpand = True).
' hStart: Search below this node instead of root.Fails if pidlOpen is a sibling of hStart.
' bExpand = false: searches only in current nodes.Does not select nor ensure visible.
' Returns eResult enum for sucess/failure.
' Returns hNext of last item it could resolve, for on_Partial (returns OpenNode = 0).
' - Selects hNext and ensures visibility (bExpand = True)
' bPartialUseText: for partial results, tries to resolve by TextCompare (ie renamed items).
Private Function OpenNode(ByVal pidlOpen As Long, _
                          Optional ByVal hStart As Long, _
                          Optional ByVal bExpand As Boolean, _
                          Optional ByRef eResult As eOpenNodeResult, _
                          Optional ByRef hNext As Long, _
                          Optional ByVal bPartialUseText As Boolean) As Long

   Dim isfParent     As IShellFolder
   Dim isfParentTree As IShellFolder
   Dim pidlCopyTree  As Long
   Dim pidlCopy      As Long
   Dim pidlFQ        As Long
   Dim pidlRel       As Long
   Dim pidl          As Long
   Dim hNode         As Long
   Dim hr            As Long
   Dim idxID         As Long
   Dim IDCountOpen   As Long
   Dim IDCountTree   As Long
   Dim bCompText     As Boolean
   Dim bRedraw       As Boolean
   Dim DBP           As Integer
   DBP = 0           ' 0 to 3
   
   On Error GoTo Proc_Error
   
   With m_ucTree
   
      eResult = on_Error
      hNext = 0
      
      If .NodeCount = 0 Then GoTo DONE
      
      If bExpand Then
         bRedraw = .Redraw
         .Redraw = False
         
         Screen.MousePointer = vbHourglass
      End If
      
      eResult = on_Valid
      
      If pidlOpen = 0 Then
         hNode = .NodeRoot
         GoTo DONE
      End If
      
      ' parent's IShellFolder of pidlOpen : start with Desktop
      Set isfParent = isfDesktop
      IDCountOpen = GetItemIDCount(pidlOpen)
      
      Debug.Assert DbgPrt(DBP >= 1, vbCrLf & "OpenNode" & String$(60, "-"))
      Debug.Assert DbgPrt(DBP >= 1, "OPEN PATH", GetFolderDisplayName(isfParent, pidlOpen, SHGDN_NORMAL), GetPathFromPIDL(pidlOpen), IDCountOpen, dbgWalkPIDL(pidlOpen))
      
      If hStart <> 0 Then
         hNode = hStart
      Else
         hNode = .NodeRoot
      End If
      
      ' try shortcut: is pidlOpen == NodeRoot/hStart
      pDecryptNodeKey .NodeKey(hNode), pidlFQ, pidlRel
      If isfParent.CompareIDs(0&, pidlOpen, pidlFQ) = 0 Then
         eResult = on_Sucess: GoTo DONE
      End If
      
      ' = 0: Desktop
      ' = 1: Virtual folder like MyComputer, MyDocuments etc. Child of Desktop.
      ' = 2: Drives or Virtual folder like System Props etc. Child of MyComputer.
      ' >  : descendants
      IDCountTree = GetItemIDCount(pidlFQ)
      
      ' -- with passed hStart skip this section --
      If hStart = 0 Then
      
         ' try shortcut: is pidlOpen == SelectedNode, or == parent of SelectedNode
         hNext = .SelectedNode
         hNode = hNext
         If hNode Then
            pDecryptNodeKey .NodeKey(hNode), pidlFQ, pidlRel
            If isfParent.CompareIDs(0&, pidlOpen, pidlFQ) = 0& Then
               eResult = on_Sucess: GoTo DONE
            End If
            
            hNode = .NodeParent(hNext)
            If hNode Then
               pDecryptNodeKey .NodeKey(hNode), pidlFQ, pidlRel
               If isfParent.CompareIDs(0&, pidlOpen, pidlFQ) = 0& Then
                  eResult = on_Sucess: GoTo DONE
               End If
            End If
         End If
         
         ' -- "Broadway the hard way": all shortcuts failed --
         
         hNode = .NodeRoot
         
         If (.NodeNextSibling(hNode) = 0) Then
            ' single root node:     start with it's first child
            If bExpand Then .Expand hNode
            hNode = .NodeChild(hNode)
            If hNode = 0 Then eResult = on_Error: GoTo DONE  ' no children ???
            IDCountTree = IDCountTree + 1
       ' Else
            ' multiple root nodes:  start with hNode = .NodeRoot
         End If
      
      Else
         ' -- instead of root start with hStart --
         ' in comments below read 'root' as hStart
         Debug.Assert hNode = hStart
      End If   ' hStart = 0

      ' -- Get isfParent & isfParentTree in sync --

      ' Get the tree root parent's IShellFolder from its fully qualified pidl
      pDecryptNodeKey .NodeKey(hNode), pidlFQ, pidlRel
      Set isfParentTree = GetIShellFolderParent(pidlFQ, fRtnDesktop:=True)
      
      Debug.Assert DbgPrt(DBP >= 2, "TREE ROOT", GetFolderDisplayName(isfParentTree, pidlRel, SHGDN_NORMAL), GetPathFromPIDL(pidlFQ), IDCountTree)
      
      pidl = CopyPIDL(pidlOpen)  ' !!! = 0 for desktop !!!
      
      For idxID = 1 To IDCountOpen

         ' compare items of pidlOpen with items of tree root: create simple pidls
         pidlCopy = GetItemID(pidlOpen, idxID)
         pidlCopyTree = GetItemID(pidlFQ, idxID)

         Debug.Assert DbgPrt(DBP >= 2, "COMPARE: " & idxID, GetFolderDisplayName(isfParent, pidlCopy, SHGDN_INFOLDER), "TREE: " & GetFolderDisplayName(isfParent, pidlCopyTree, SHGDN_FORPARSING))
         
         hr = isfParent.CompareIDs(0&, pidlCopy, pidlCopyTree)

         Select Case hr
            Case 0
               ' OK: common item in path, compare next item
               
               If idxID = IDCountTree Then
                  ' Start searching this item in tree
                  Debug.Assert DbgPrt(DBP >= 2, "START idxID: " & idxID, GetFolderDisplayName(isfParent, pidlCopy, SHGDN_NORMAL) & vbCrLf)
                  Debug.Assert (hStart <> 0) Or (IDCountOpen >= IDCountTree)
                  Exit For
               End If
               
            Case Is > 0
               ' Start searching this item in tree
               Debug.Assert DbgPrt(DBP >= 2, "START idxID: " & idxID, GetFolderDisplayName(isfParent, pidlCopy, SHGDN_NORMAL) & vbCrLf)
               
               Exit For
            
            Case Else
               ' Failed: no common item, bailout except if ...
               
               ' # Bug: CompareIDs fails for children (ID= 2) of MyDocuments.          #
               ' #      MS uses some 'SortOrderIndex hack' to add MyDocuments as       #
               ' #      first desktop child.                                           #
               ' # Fix: If pidlOpen refers to a MyDocuments child, start searching     #
               ' #      this item in tree from first child of current hNode.           #
               ' #      (for single root node, hNode already is first child of root)   #
               
               If idxID = 2 Then
                  
                  If IsPidlInMyDocuments(pidlOpen) Then
                     If bExpand Then .Expand hNode
                     hNext = .NodeChild(hNode)
                     
                     If hNext <> 0 Then
                        hNode = hNext
                        IDCountTree = IDCountTree + 1
                        
                        ' Start searching this item in tree
                        Debug.Assert DbgPrt(DBP >= 2, "START + FIX idxID: " & idxID, GetFolderDisplayName(isfParent, pidlCopy, SHGDN_NORMAL) & vbCrLf)
                        
                        Exit For
                     End If
                     
                  End If
               End If
               
               ' bailout
               FreePIDL pidlCopy
               FreePIDL pidlCopyTree
               
               eResult = on_Error: GoTo DONE
               
         End Select

         ' Bind to the subfolder
         Set isfParent = GetIShellFolder(isfParent, pidlCopy, fRtnDesktop:=False)

         ' Free the copy of the item identifier
         FreePIDL pidlCopy
         FreePIDL pidlCopyTree
         
         ' Get the next item ID
         pidl = GetNextItemID(pidl)

      Next
      
      FreePIDL pidlCopy
      FreePIDL pidlCopyTree
      
      ' -- Start searching nodes with current pidl in isfParent --
      Debug.Assert pidl
      Debug.Assert hNode
      hNext = 0&
      
      Do While pidl
         
         ' Copy the item identifier to a list by itself (simple pidl)
         pidlCopy = GetItemID(pidl, GIID_FIRST)
         
         ' reached terminator (err with GetFolderDisplayName)
         If GetItemIDSize(pidlCopy) = 0 Then
            FreePIDL pidlCopy
            Exit Do
         End If
         
         ' Display the name of the subfolder, which is searched for
         Debug.Assert DbgPrt(DBP >= 2, "SEARCHED FOLDER", GetFolderDisplayName(isfParent, pidlCopy, SHGDN_NORMAL) & vbCrLf)
         
         Do While hNode

            pDecryptNodeKey .NodeKey(hNode), pidlFQ, pidlRel
            Debug.Assert DbgPrt(DBP >= 2, "SEARCH IN", GetFolderDisplayName(isfParentTree, pidlRel, SHGDN_NORMAL), GetPathFromPIDL(pidlFQ), GetItemIDCount(pidlFQ), .NodeText(hNode))
            
            hr = isfParent.CompareIDs(0&, pidlCopy, pidlRel)
            
            If hr >= 0 Then
               
               If Not bCompText Then
                  ' low word of hr is valid
                  hr = LOWORD(hr)
                  
               Else
                  ' isfParent's ItemIDList content's refer to actual filesystem names.
                  ' Searching node with old name for renamed items will fail.
                  ' With bPartialUseText set, after a partial sucess, retries the branch
                  ' by comparing to NodeText(hNode).On match isfParent is reused for next
                  ' item, otherwise exits with on_Partial.
                  Debug.Assert DbgPrt(DBP >= 1, "TEXTCOMPARE", GetFolderDisplayName(isfParent, pidlCopy, SHGDN_NORMAL), .NodeText(hNode))
                  hr = StrComp(GetFolderDisplayName(isfParent, pidlCopy, SHGDN_NORMAL), .NodeText(hNode), vbTextCompare)
                ' Debug.Assert False
                  hr = Abs(hr)
               End If
               
               Select Case hr
   
                  Case Is > 0
                     ' first item should follow the second
                     hNext = .NodeNextSibling(hNode)
                     If hNext Then
                        ' continue with sibling
                        hNode = hNext
                     Else
                        ' Partial sucess: pidlOpen can be resolved in tree until hNode
                        eResult = on_Partial
                        
                        If bPartialUseText And Not bCompText Then
                           ' try again this branch by comparing to NodeText
                           bCompText = True
                           hNode = .NodeFirstSibling(hNode)
                         ' Debug.Assert False
                        Else
                           Exit Do
                        End If
                     End If
                     
                  Case 0
                     ' the two items match
                     Debug.Assert DbgPrt(DBP >= 3, "MATCH")
                     
                     If idxID < IDCountOpen Then
                        ' Continue with child (if any)
                        If bExpand Then .Expand hNode
                        hNext = .NodeChild(hNode)
                        
                        If hNext Then
                           ' Bind to current node's IShellfolder
                           Set isfParentTree = GetIShellFolder(isfParentTree, pidlRel)
                           hNode = hNext
                           
                           If bCompText Then
                              ' match achieved by NodeText compare (-> node has old name)
                              ' reset for next child item
                              bCompText = False
                              eResult = on_Valid
                           End If
                        Else
                           ' Partial sucess: pidlOpen can be resolved in tree until hNode
                           eResult = on_Partial
                        End If
                        
                     ElseIf idxID = IDCountOpen Then
                        ' SUCESS: pidlOpen fully resolved
                        ' bCompText is True, if match achieved by TextCompare
                        ' and hNode is first sibling of current branch.
                        eResult = on_Sucess
                     
                     Else
                        Debug.Print "FUBAR"
                        Debug.Assert False
                        eResult = on_Error
                     End If
                     
                     ' Done searching for hNode of current pidlCopy in tree
                     Exit Do
                     
   
                  Case Is < 0
                     ' first item should precede the second
                     hNext = .NodeNextSibling(hNode)
                     If hNext Then
                        ' still siblings in branch to search: continue
                        hNode = hNext
                     Else
                        Debug.Assert DbgPrt(DBP >= 2, "pidlOpen not in tree")
                        ' Parent (if any) as last resolved
                        hNext = .NodeParent(hNode)
                        If hNext Then hNode = hNext
                        eResult = on_Partial
                        Exit Do
                     End If
               End Select  ' LOWORD(hr)
            
            Else
               ' hr < 0 := isfParent.CompareIDs returned error
               Debug.Assert DbgPrt(DBP >= 1, "FUBAR", SUCCEEDED(hr))
'               Debug.Assert False
               eResult = on_Error
               
               Exit Do
            End If   ' hr >= 0
            
         Loop ' While hNode
            
         If (eResult <> on_Valid) Then
            FreePIDL pidlCopy
            Exit Do
         End If
         
         ' Bind to the subfolder
         Set isfParent = GetIShellFolder(isfParent, pidlCopy)
         
         ' Free the copy of the item identifier
         FreePIDL pidlCopy
         
         ' Get the next item ID
         pidl = GetNextItemID(pidl)
         idxID = idxID + 1
         
      Loop  ' While pidl
      
DONE:
      On Error Resume Next
      
      FreePIDL pidlCopy
      FreePIDL pidl
      
      Select Case eResult
      
         Case on_Sucess
            Debug.Assert (isfDesktop.CompareIDs(0&, pidlOpen, pidlFQ) = 0) Or bCompText
            Debug.Assert hNode
            hNext = 0
            
         Case on_Valid
            Debug.Assert (pidlOpen = 0)
            Debug.Assert hNode
            hNext = 0
            
         Case on_Partial
            Debug.Assert hNode
            Debug.Assert DbgPrt(DBP >= 2, "Resolved to: " & .NodeText(hNode))
            If hNode <> 0 And bExpand Then
               .SelectedNode = hNode
               .EnsureVisible hNode
            End If
            hNext = hNode
            hNode = 0
            
         Case on_Error
            Debug.Assert DbgPrt(DBP >= 2, "Error in: " & IIf(hNode, .NodeText(hNode), ""))
            hNext = hNode
            hNode = 0
            
      End Select
      
      If hNode <> 0 And bExpand Then
         .SelectedNode = hNode
         .EnsureVisible hNode
      End If
      
      If bExpand Then
         .Redraw = bRedraw
         Screen.MousePointer = vbDefault
      End If
      
   End With
   
   Debug.Assert DbgPrt(DBP >= 1, "OpenNode: " & IIf(hNode, "SUCESS", "FUBAR : ") & _
                                 Choose(eResult, "", "PARTIAL", "ABOVE", "BELOW", "ERROR"))

   OpenNode = hNode
   
   Exit Function

Proc_Error:
   Debug.Print "Error: " & Err.Number & ". " & Err.Description
   If InIDE Then Stop: Resume
   eResult = on_Error
   GoTo DONE
End Function

' ----------------------------------------------------------------------------------- '
'
' ----------------------------------------------------------------------------------- '

Private Sub Class_Initialize()
   
   m_eIncludeItems = SHCONTF_FOLDERS
'   SetUnhandledExceptionFilter AddressOf MyExceptionHandler
End Sub

Private Sub Class_Terminate()
   
#If SHCN Then
   pUnregisterNotify
#End If
   pFreePidls
   Set m_ucTree = Nothing
'   ' Remove the hook
'   SetUnhandledExceptionFilter ByVal 0&
End Sub

' keep identical to same proc in modSortTreeCB
Private Sub pDecryptNodeKey(sKey As String, Optional pidlFQ As Long, Optional pidlRel As Long)
   Dim saKey() As String
   saKey = Split(sKey, ":")
   
   pidlFQ = CLng(saKey(0))
   pidlRel = CLng(saKey(1))
End Sub

Private Sub pFreePidls(Optional ByVal hNode As Long, Optional ByVal idxNode As Long)
   Dim pidlFQ  As Long
   Dim pidlRel As Long
   
   With m_ucTree
      If hNode <> 0 Or idxNode <> 0 Then
         pDecryptNodeKey .NodeKey(hNode, idxNode), pidlFQ, pidlRel
         FreePIDL pidlFQ
         FreePIDL pidlRel
      Else
         For idxNode = 1 To .NodeCount
            pDecryptNodeKey .NodeKey(, idxNode), pidlFQ, pidlRel
            FreePIDL pidlFQ
            FreePIDL pidlRel
         Next
         m_pidlFQ = 0
      End If
   End With
   
End Sub

' ----------------------------------------------------------------------------------- '
'  Treeview events
' ----------------------------------------------------------------------------------- '

Private Sub m_ucTree_BeforeCollapse(ByVal hNode As Long, Cancel As Integer)

   If CBool(m_eOptions And ftCollapseReset) Then
      m_ucTree.Collapse hNode, CollapseChildren:=False, RemoveChildren:=True
   End If
End Sub

Private Sub m_ucTree_BeforeDelete(ByVal hNode As Long, ByVal idxNode As Long)
   ' ensures that pidls are freed, if client operates directly on ucTreeView
'   Debug.Print "BeforeDelete", m_ucTree.NodeText(, idxNode)
   pFreePidls , idxNode
End Sub

Private Sub m_ucTree_BeforeExpand(ByVal hNode As Long, ByVal ExpandedOnce As Boolean, Cancel As Integer)
   Dim pidlFQ     As Long
   Dim pidlRel    As Long
   Dim bRedraw    As Boolean
   Dim bInsert    As Boolean
   Dim eMP        As MousePointerConstants
   
   On Error GoTo Proc_Error
   
   ' Load on Demand
   If Not ExpandedOnce Then
      
      With m_ucTree
      
         pDecryptNodeKey .NodeKey(hNode), pidlFQ, pidlRel
      
         bInsert = True
         
#If OLEDD Then
         ' # !!! While dragging any popup of a (system) dialog causes a system hangup !!! #
         ' # -> don't autoexpand floppy, slow network paths or MyWebsites                 #

         If .OLEDragAutoExpand Then
            If .OLEGetDropInfo Then
            
               ' expanding as result of DragAutoExpand
               bInsert = pAllowAutoExpand(pidlFQ, pidlRel)
               
               
               If Not bInsert Then
                  Debug.Print "REJECTED AUTOEXPAND", .NodeText(hNode)
                  Debug.Print dbgShowAttributes(pidlFQ, pidlRel)
               End If
               
            End If
         End If
#End If  ' OLEDD
      
         If bInsert Then
            eMP = Screen.MousePointer
            Screen.MousePointer = vbHourglass
            bRedraw = .Redraw
            .Redraw = False
            
            pInsertFolderItems pidlFQ, hNode
            
            Screen.MousePointer = eMP
            .Redraw = bRedraw
         Else
            Cancel = 1
            
'            If (m_eOptions And ftSHNotify) Then
'               ' Shell will notify us about any change: remove button
'               .NodePlusMinusButton(hNode) = False
'            End If
         End If
         
      End With
   End If

   Exit Sub

Proc_Error:
   Debug.Print Err.Number, Err.Description
   If InIDE Then Stop: Resume
End Sub

Private Sub m_ucTree_AfterLabelEdit(ByVal hNode As Long, Cancel As Integer, NewString As String)
   Dim isfParent     As IShellFolder
   Dim pidlFQ        As Long
   Dim pidlRel       As Long
   Dim pidlRelNew    As Long
   
   ' as default return
   Cancel = True
   
   If LenB(NewString) Then
      
      If NewString = m_ucTree.NodeText(hNode) Then Exit Sub
      
      If LenB(m_sEdit) Then
         ' Only a part of NodeText may be edited
         If NewString = m_sEdit Then Exit Sub
      End If
      
      ' -- Rename item --
      pDecryptNodeKey m_ucTree.NodeKey(hNode), pidlFQ, pidlRel
      Set isfParent = GetIShellFolderParent(pidlFQ, fRtnDesktop:=True)
      
      If SUCCEEDED(isfParent.SetNameOf(m_hOwner, pidlRel, _
                   StrConv(NewString, vbUnicode), SHGDN_FOREDITING, pidlRelNew)) Then
                   
         If pidlRelNew <> 0 Then
         
            ' SUCCEEDED not reliable: at least with missing permissions, no error is raised
            If isfParent.CompareIDs(0&, pidlRel, pidlRelNew) <> 0 Then
               
               ' OK: Item was renamed
               ' With or without Shell change notifies, ucTree sets NewString as NodeText
               
               If (m_eOptions And ftSHNotify) Then
                  ' Strange,but true: cancel LabelEdit
                  ' otherwise pUpdateNodePIDL won't sort because of NewText = OldText
                  Debug.Assert Cancel = True
                  
               Else
                  ' without Shell change notifies must update node with new pidl here
                  Cancel = False
                  
                  Debug.Print "NOT YET IMPLEMENTED"
                  
                  ' Setup the callback and sort the parent folder
                  m_ucTree.SortChildrenCB m_ucTree.NodeParent(hNode), AddressOf modSortTreeCB.TreeViewCompareProc, _
                                          False, ObjPtr(isfParent)
                  
               End If
          
          ' Else: ' sanity
            End If
           
       ' Else: ' pidlRelNew is same item: renaming failed without error ( maybe shell showed dialog)
         End If
         
    ' Else: ' Failed with error: shell shows dialog (wrong tokens,permissions etc)
      End If
      
   End If

End Sub

Private Sub m_ucTree_BeforeLabelEdit(ByVal hNode As Long, Cancel As Integer, EditString As String, ByVal hEdit As Long)
   Dim isfParent     As IShellFolder
   Dim ulAttrs       As ESFGAO
   Dim pidlFQ        As Long
   Dim pidlRel       As Long
   
   ' Automatic LabelEdit on MouseDown : Cancel as default
   Cancel = True

   If m_bInEdit = Checked Then
   
      ' -- LabelEdit started manual on MouseUp
      With m_ucTree
         Debug.Assert .SelectedNode = hNode
         
         ' -- Can item be renamed ? --
         pDecryptNodeKey .NodeKey(hNode), pidlFQ, pidlRel
         Set isfParent = GetIShellFolderParent(pidlFQ, fRtnDesktop:=True)
         
         ulAttrs = SFGAO_CANRENAME
         
         If S_OK = isfParent.GetAttributesOf(1, pidlRel, ulAttrs) Then
         
            If (ulAttrs And SFGAO_CANRENAME) Then
            
               m_sEdit = GetFolderDisplayName(isfParent, pidlRel, SHGDN_FOREDITING)
               
               If LenB(m_sEdit) Then
               
                  If m_sEdit <> .NodeText(hNode) Then
                     ' Only a part of NodeText may be edited (ie drivename not driveletter):
                     EditString = m_sEdit
                  Else
                     ' OK editing NodeText
                     m_sEdit = vbNullString
                  End If
                  
                  ' OK start LabelEdit
                  Cancel = False
               End If
               
            End If
            
         End If   ' SUCCEEDED()
         
      End With
      
   End If   ' m_bInEdit = Checked
   
   m_bInEdit = Unchecked
End Sub

Private Sub m_ucTree_BeforeSelectionChange(ByVal hNodeNew As Long, ByVal hNodeOld As Long, Cancel As Integer)
   Dim pidlFQ     As Long
   Dim pidlRel    As Long
   
   If hNodeNew Then
      pDecryptNodeKey m_ucTree.NodeKey(hNodeNew), pidlFQ, pidlRel
      m_pidlFQ = pidlFQ
'      m_pidlRel = pidlRel
      
      If Not IsDesktopPIDL(pidlFQ) Then
         m_sPath = DisplayName(SHGDN_FORPARSING Or SHGDN_NORMAL)
         Debug.Assert LenB(m_sPath)
      Else
         m_sPath = vbNullString
      End If
      
'      Debug.Print GetFileIconIndexPIDL(pidlFQ, SHGFI_SMALLICON), GetFileIconIndexPIDL(pidlFQ, SHGFI_SMALLICON Or SHGFI_OPENICON)
'      Debug.Print "CurrentPath ", vbTab, vbTab, m_sPath
'
'      Debug.Print "SHGDN_NORMAL", vbTab, vbTab, DisplayName(SHGDN_NORMAL)
'      Debug.Print "SHGDN_INFOLDER", vbTab, DisplayName(SHGDN_INFOLDER)
'      Debug.Print "SHGDN_FORPARSING SHGDN_NORMAL", DisplayName(SHGDN_FORPARSING Or SHGDN_NORMAL)
'      Debug.Print "SHGDN_FORPARSING SHGDN_INFOLDER", DisplayName(SHGDN_FORPARSING Or SHGDN_INFOLDER)
'      Debug.Print "SHGDN_FORADDRESSBAR", vbTab, DisplayName(SHGDN_FORPARSING Or SHGDN_FORADDRESSBAR)
'      Debug.Print "SHGDN_FOREDITING", vbTab, DisplayName(SHGDN_FOREDITING)
'      Debug.Print ""
'
'      Debug.Print dbgShowAttributes(pidlFQ, pidlRel)
'      Debug.Print "dbgGetPIDLType", dbgGetPIDLType(pidlRel)
'      Debug.Print dbgDumpPIDL(pidlRel)
'      Debug.Print dbgEnumShellCommands(m_hOwner, CurrentISF, 1, pidlRel, , CMF_EXPLORE Or CMF_CANRENAME)
'      Debug.Print dbgEnumShellViewCommands(m_hOwner, CurrentISF, CMF_NORMAL)
   Else
      m_pidlFQ = 0
'      m_pidlRel = 0
      m_sPath = vbNullString
   End If
   
End Sub

Private Sub m_ucTree_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim hNode      As Long
   Dim cPidls     As Long
   Dim aPidls()   As Long
   Dim isfParent  As IShellFolder
   Dim cmd        As SHContextMenuCmdIDs

   With m_ucTree
   
      If KeyCode = vbKeyF5 Then
         ' -- update tree contents --
         Me.Refresh
         Exit Sub
      End If
      
      If (m_eOptions And ftFileOperations) = 0 Then Exit Sub

      hNode = .SelectedNode
      If (hNode = 0) Then Exit Sub

      If KeyCode = vbKeyReturn Then
      
         ' -- file/folder default action --
         
#If MULSEL Then
         Dim idx        As Long

         For idx = 1 To .SelectionCount
            m_ucTree_NodeDblClick .SelectionNode(idx), vbLeftButton, False
         Next
#Else
         m_ucTree_NodeDblClick hNode, vbLeftButton, False
         
#End If  ' MULSEL
         
         ' # Strange behaviour opening multiple files by contextmenu:              #
         ' # ie single txt file opens in notepad all right, but for multiple       #
         ' # before each file opens in notepad, the menus "Move items to folder"   #
         ' # and "Copy items to folder" are invoked (WinXP)???                     #
         ' # Same behaviour if pShellOperation is called for each file seperately. #
      '  pGetSelectedItemsPidl isfParent, cPidls, aPidls()
      '  cmd = FindShellCommand(m_hOwnerisfParent, cPidls, aPidls(1), "open")
      '  pShellOperation isfParent, cPidls, aPidls(1), cmd
         
         
      ElseIf KeyCode = vbKeyDelete Then
      
         ' Won't allow deleting single tree root
         If (hNode = .NodeRoot And (.NodeNextSibling(hNode) = 0)) Then Exit Sub
      
         ' invoke delete command on item's contextmenu (considers Shift key)
         cmd = cmd_Delete

      ElseIf GetAsyncKeyState(vbKeyControl) < 0 Then
      
         ' -- Cut, Copy, Paste --
         ' The easy way of Shell Data transfer: invoke command on item's contextmenu
         Select Case KeyCode
            Case vbKeyX:   cmd = cmd_Cut
            Case vbKeyC:   cmd = cmd_Copy
            Case vbKeyV:   cmd = cmd_Paste
            Case Else:     Exit Sub
         End Select

      End If   ' KeyCode
      
      If cmd Then
      
         pGetSelectedItemsPidl isfParent, cPidls, aPidls()
         
         If pShellOperation(isfParent, cPidls, aPidls(1), cmd) Then
            
            pResetCutItem cmd, hNode

         End If
      End If
      
   End With

End Sub

Private Sub m_ucTree_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
   Dim hNode      As Long
   Dim lfHit      As tvHitTestConstants
   
   With m_ucTree
      
      m_bInEdit = Unchecked
      
      hNode = .HitTest(X, Y, FullRowHit:=True, lfHit:=lfHit)
      If hNode Then
         
         If (lfHit And TVHT_ONITEMBUTTON) = 0 Then   ' == not expanding/collapsing
         
            If (Button = vbLeftButton) Then
               ' If clicked on previously selected node: start LabelEdit on MouseUp
               If (lfHit And TVHT_ONITEMLABEL) Then
                  If .SelectedNode = hNode Then
                     m_bInEdit = Gray
                  End If
               End If
            Else
               ' select node for possible following drag/ contextmenu operations
               If .SelectedNode <> hNode Then
                  .SelectedNode = hNode
               End If
            End If
            
#If OLEDD Then
            ' activate DragDetect on MouseMove
            m_tPDrag.X = X:   m_tPDrag.Y = Y
            m_bInDrag = Gray
#End If
         End If
         
      End If
   End With
End Sub

#If OLEDD Then

Private Sub m_ucTree_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
   Dim tR         As RECT2
   Dim dxy        As Long
   
   If m_bInDrag = Gray Then
      ' DragDetect rect: center m_tPDrag set in MouseDown
      SetRect tR, m_tPDrag.X, m_tPDrag.Y, m_tPDrag.X + 1, m_tPDrag.Y + 1
      dxy = GetSystemMetrics(SM_CXDRAG) \ 2
      InflateRect tR, dxy, dxy
      
      If PtInRect(tR, X, Y) = 0 Then
         ' mouse moved outside DragDetect rect since last MouseDown
         m_bInEdit = Unchecked
         
         m_hNodeDrag = m_ucTree.SelectedNode
         
         If m_hNodeDrag Then
         
            m_bInDrag = Checked
            Debug.Print "STARTDRAG"

            Call pOLEStartDrag
         
         End If
      End If
   End If
End Sub

#End If  ' OLEDD

Private Sub m_ucTree_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
   Dim isfParent  As IShellFolder
   Dim ulAttrs    As ESFGAO
   Dim hNode      As Long
   Dim pidlFQ     As Long
   Dim pidlRel    As Long
   Dim cPidls     As Long
   Dim aPidls()   As Long
   Dim CMF        As eCMF
   Dim cmd        As SHContextMenuCmdIDs
   
   With m_ucTree
   
#If OLEDD Then
      
      m_tPDrag.X = 0:   m_tPDrag.Y = 0
      
      If m_bInDrag = Checked Then
         ' After drag
         m_bInDrag = Unchecked
         m_bInEdit = Unchecked
         Exit Sub
      End If
      
      m_bInDrag = Unchecked
      
#End If  ' OLEDD
      
      hNode = .HitTest(X, Y, FullRowHit:=True)
      If hNode Then
      
         If (Button = vbLeftButton) Then
         
            ' -- Rename item, if left clicked on previously selected node
            '    (Explorer: renaming only by ContextMenu)
            If m_bInEdit = Gray Then
               m_bInEdit = Checked
               .StartLabelEdit hNode
            Else
               m_bInEdit = Unchecked
            End If
         
         ElseIf CBool(m_eOptions And ftContextMenu) Then
         
            ' # Show view's background context menu, if no node hit #
            ' # just testing, menu doesn't belong here              #
            If .HitTest(X, Y, FullRowHit:=False) = False Then
            
               pDecryptNodeKey .NodeKey(hNode), pidlFQ, pidlRel
               
            '  ' cmdv_NewFolder: Folder created as sibling of selected
            '  Set isfParent = GetIShellFolderParent(pidlFQ, fRtnDesktop:=False)
            
               ' cmdv_NewFolder: Folder created as child of selected
               Set isfParent = GetIShellFolder(isfDesktop, pidlFQ, fRtnDesktop:=False)
               
               If Not (isfParent Is Nothing) Then
                  cmd = ShowSHCtxMenuViewBK(.hWnd, isfParent, X, Y)
               End If
               Debug.Print "USERCHOICE", cmd
               Exit Sub
            End If
         
            ' -- Shell ContextMenu
            
            pGetSelectedItemsPidl isfParent, cPidls, aPidls()
            
            CMF = CMF_EXPLORE
            ulAttrs = SFGAO_CANRENAME
            
            If S_OK = isfParent.GetAttributesOf(cPidls, aPidls(1), ulAttrs) Then
               If (ulAttrs And SFGAO_CANRENAME) Then
                  CMF = CMF_EXPLORE Or CMF_CANRENAME
               End If
            End If
            
            cmd = ShowShellContextMenu(.hWnd, isfParent, cPidls, aPidls(1), X, Y, CMF)
            Debug.Print "USERCHOICE", cmd
            
            pResetCutItem cmd, hNode
            
            Select Case cmd
               
               Case cmd_Rename
                  Debug.Assert CMF = CMF_EXPLORE Or CMF_CANRENAME
                  m_bInEdit = Checked
                  .StartLabelEdit hNode
               
'               Case Else
'                  Debug.Print "USERCHOICE", cmd
            End Select
            
         End If
         
      End If
   End With
End Sub

Private Sub m_ucTree_NodeDblClick(ByVal hNode As Long, ByVal Button As Integer, bCancelExpansion As Boolean)
   Dim isfParent  As IShellFolder
   Dim ulAttrs    As ESFGAO
   Dim pidlFQ     As Long
   Dim pidlRel    As Long
   
   With m_ucTree
      
      ' for folder items, expand/collapse (comtl default)
      ' for non-folder items, trigger the default shell action & cancel node expansion
      If (m_eOptions And ftFileOperations) And (Button = vbLeftButton) Then
      
         If (m_eIncludeItems And SHCONTF_NONFOLDERS) Then
         
            If .NodePlusMinusButton(hNode) = False Then
               
               ' folder or file ?
               pDecryptNodeKey .NodeKey(hNode), pidlFQ, pidlRel
               
               Set isfParent = GetIShellFolderParent(pidlFQ, fRtnDesktop:=True)
               
               ulAttrs = SFGAO_FOLDER
               If S_OK = isfParent.GetAttributesOf(1, pidlRel, ulAttrs) Then
                  
                  If (ulAttrs And SFGAO_FOLDER) = 0 Then
                     ' item is a file: trigger default action
                     bCancelExpansion = True
                     
                     pShellExecute pidlFQ
                  End If
                  
               End If
               
          ' Else
               ' must be folder
            End If
            
         End If   ' (m_eIncludeItems And SHCONTF_NONFOLDERS)
         
      End If
      
   End With
End Sub

#If OLEDD Then

' ----------------------------------------------------------------------------------- '
' ----------------------------------------------------------------------------------- '
' OLE Drag & Drop
' ----------------------------------------------------------------------------------- '
' ----------------------------------------------------------------------------------- '


' ----------------------------------------------------------------------------------- '
' OLE Drag & Drop Interfaces
' ----------------------------------------------------------------------------------- '

' cPidls >1 : pidlRel as immediate children of isfParent. ByRef pidlRel !!!
Private Function pIDataObject(ByVal hOwner As Long, isfParent As IShellFolder, _
                              ByVal cPidls As Long, pidlRel As Long _
                              ) As IDataObject
   Dim hr      As Long
   Dim ido     As IDataObject
   
   ' Get a reference to the item's IDataObject interface
   hr = isfParent.GetUIObjectOf(hOwner, cPidls, pidlRel, GetRIID(rIID_IDataObject), 0, ido)
   If SUCCEEDED(hr) Then
      Set pIDataObject = ido
   End If
   
End Function

Private Function pIDropTarget(ByVal hOwner As Long, isfParent As IShellFolder, _
                              ByVal cPidls As Long, pidlRel As Long _
                              ) As IDropTarget
   Dim hr      As Long
   Dim idt     As IDropTarget
   
   Debug.Assert cPidls = 1
   
   ' Get a reference to the item's IDropTarget interface
   hr = isfParent.GetUIObjectOf(hOwner, cPidls, pidlRel, GetRIID(rIID_IDropTarget), 0, idt)
   If hr = S_OK Then
      Set pIDropTarget = idt
   End If
   
End Function

#If WIN32_IE >= &H500 Then

Private Function pIDropTargetHelper() As IDropTargetHelper
   Dim hr      As Long
   Dim idth    As IDropTargetHelper
   
   ' Create an in-process server drag-image manager object with IDropTargetHelper interface
   hr = CoCreateInstance(GetRIID(rCLSID_DragDropHelper), 0&, CLSCTX_INPROC_SERVER, _
                         GetRIID(rIID_IDropTargetHelper), idth)
   If hr = S_OK Then
      Set pIDropTargetHelper = idth
   End If
        
End Function

Private Function pIDragSourceHelper() As IDragSourceHelper
   Dim hr      As Long
   Dim idsh    As IDragSourceHelper
   
   ' Create an in-process server drag-image manager object with IDragSourceHelper interface
   hr = CoCreateInstance(GetRIID(rCLSID_DragDropHelper), 0&, CLSCTX_INPROC_SERVER, _
                         GetRIID(rIID_IDragSourceHelper), idsh)
   If hr = S_OK Then
      Set pIDragSourceHelper = idsh
   End If
        
End Function

#End If  ' WIN32_IE >= &H500

' ----------------------------------------------------------------------------------- '
' OLE Drag Source with IDropSource implementation + DoDragDrop API
' ----------------------------------------------------------------------------------- '

Private Sub pOLEStartDrag()
   Dim isfParent     As IShellFolder
   Dim ulAttrs       As ESFGAO
   Dim cPidls        As Long
   Dim aPidls()      As Long
   Dim AllowedEffects As Long
   Dim Effect        As Long
   Dim hr            As Long
   
   On Error GoTo Proc_Error
   
   Debug.Assert m_hNodeDrag

   ' -- Set AllowedEffects ( == CFSTR_PREFERREDDROPEFFECT) based on user's choice --
   Select Case True
   
      Case GetAsyncKeyState(vbKeyLButton) < 0
         
         Select Case True
            Case GetAsyncKeyState(vbKeyMenu) < 0
               AllowedEffects = vbDropEffectLink
            Case (GetAsyncKeyState(vbKeyShift) < 0) And (GetAsyncKeyState(vbKeyControl) < 0)
               AllowedEffects = vbDropEffectLink
            Case (GetAsyncKeyState(vbKeyShift) < 0) Or (GetAsyncKeyState(vbKeyControl) < 0)
               AllowedEffects = vbDropEffectCopy
            Case Else
               AllowedEffects = vbDropEffectCopy Or vbDropEffectMove Or vbDropEffectLink
         End Select
           
      Case GetAsyncKeyState(vbKeyRButton) < 0
         AllowedEffects = vbDropEffectCopy Or vbDropEffectMove Or vbDropEffectLink
         
      Case Else
         AllowedEffects = vbDropEffectNone
         m_hNodeDrag = 0
         Exit Sub
   End Select
   
   ' -- Which item(s) to drag ? --
   pGetSelectedItemsPidl isfParent, cPidls, aPidls()
   
   ' -- Can item(s) be copied/moved/linked ? --
   ulAttrs = SFGAO_CANCOPY Or SFGAO_CANMOVE Or SFGAO_CANLINK

   If S_OK = isfParent.GetAttributesOf(cPidls, aPidls(1), ulAttrs) Then

      ' SFGAO_CANXXX == vbDropEffectXXX , vbDropEffectLink selfdefined ( == DROPEFFECT_LINK)
      ulAttrs = ulAttrs And (SFGAO_CANCOPY Or SFGAO_CANMOVE Or SFGAO_CANLINK)

      ' restrict user's choice
      AllowedEffects = AllowedEffects And ulAttrs

   Else: Debug.Assert False:  AllowedEffects = vbDropEffectNone:  m_hNodeDrag = 0
   End If
   
   ' -- Set data for formats --
   If AllowedEffects <> vbDropEffectNone Then
      
      ' IDataObject of dragged item(s) (obtained from shell)
      Set m_IDataObject = pIDataObject(m_hOwner, isfParent, cPidls, aPidls(1))
      
      If Not (m_IDataObject Is Nothing) Then
   
#If WIN32_IE >= &H500 Then
         ' Create Drag image
         Dim idsh    As IDragSourceHelper
         Dim tSHDI   As SHDRAGIMAGE
         
         Set idsh = pIDragSourceHelper()
         If Not (idsh Is Nothing) Then
            With tSHDI
               .hbmpDragImage = m_ucTree.DragImageBmp(m_ucTree.SelectedNode, _
                                                      .xOffset, .yOffset, _
                                                      .wDragImage, .hDragImage, .crColorKey)
            End With
            hr = idsh.InitializeFromBitmap(tSHDI, m_IDataObject)
            Debug.Assert hr = S_OK
            Set idsh = Nothing
         End If
      
#End If  ' WIN32_IE >= &H500
         
         ' CFSTR_INDRAGLOOP format: global nonzero DWORD
         ' # not yet implemented #
         
         ' IDropSource implementation for dragged item(s)
         Set m_IDropSource = New clsDropSource
         
         ' Initiate Drag & Drop
         hr = DoDragDrop(ObjPtr(m_IDataObject), ObjPtr(m_IDropSource), AllowedEffects, Effect)
         
         Select Case hr
            Case DRAGDROP_S_DROP
               pOLECompleteDrag Effect
            Case DRAGDROP_S_CANCEL
               Debug.Print "DROP CANCELLED"
            Case Else:  Debug.Assert False: Debug.Print "DODRAGDROP FAILED", Hex$(hr), Effect
         End Select
         
         m_ucTree.Refresh

      Else: Debug.Assert False
      End If   ' Not (m_IDataObject Is Nothing)
      
   End If   ' AllowedEffects <> vbDropEffectNone
   
   m_bInDrag = Unchecked
   m_hNodeDrag = 0

   Set m_IDataObject = Nothing
   Set m_IDropTarget = Nothing
   Set m_IDTHelper = Nothing
   Set m_IDropSource = Nothing

   Exit Sub

Proc_Error:
   Debug.Print "Error: " & Err.Number & ". " & Err.Description, vbOKOnly Or vbCritical, App.Title & ".clsShellTree: Sub pOLEStartDrag"
   If InIDE Then Stop: Resume
End Sub

Private Sub pOLECompleteDrag(Effect As Long)
   ' Effect == CFSTR_PERFORMEDDROPEFFECT value
   ' VB's DataObject is already inaccessible, can use m_IDataObject
   Debug.Assert Not (m_IDataObject Is Nothing)
   
   ' Shell notifs async: received before/after this
   Debug.Print "OLECompleteDrag: " & IIf(Effect, "SUCCESS: " & Effect, "NO DROP or Optimized Move")
   
   Select Case Effect
   
      Case vbDropEffectNone
         ' no drop or optimized move
         
         If (m_eOptions And ftSHNotify) = 0 Then
            ' without Shell change notifies must determine whether(& what) operation happened
            ' use m_IDataObject.GetData()
            Debug.Print "NOT YET IMPLEMENTED"
            
            ' Extract CFSTR_TARGETCLSID, if was deleted by draggging to recycle bin
            
#If WIN32_IE >= &H500 Then
            ' Extract CFSTR_LOGICALPERFORMEDDROPEFFECT
#Else
            ' ???
#End If
         End If
         
      Case vbDropEffectMove
         ' non optimized move: must delete items here
         Debug.Print "NOT YET IMPLEMENTED"
   
   End Select
   
End Sub


' ----------------------------------------------------------------------------------- '
' OLE Drop Target
' ----------------------------------------------------------------------------------- '

Private Sub m_ucTree_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim isfParent     As IShellFolder
   Dim pidlFQ        As Long
   Dim pidlRel       As Long
   Dim hNodeDrop     As Long
   Dim hr            As Long
   Dim tP            As POINTAPI
   Dim RetEffect     As VBRUN.OLEDropEffectConstants
   Dim KeyState      As KEYSTATES
   Dim DBP           As Integer
   DBP = 2
   
   With m_ucTree
      
      RetEffect = vbDropEffectNone
      
      '-- Extract Drop info
      If .OLEGetDropInfo(hNodeDrop) Then
         
         If Not (m_IDataObject Is Nothing) Then
         
            pDecryptNodeKey .NodeKey(hNodeDrop), pidlFQ, pidlRel
            Set isfParent = GetIShellFolderParent(pidlFQ, fRtnDesktop:=True)

            Set m_IDropTarget = pIDropTarget(m_hOwner, isfParent, 1, pidlRel)
            
            If Not (m_IDropTarget Is Nothing) Then
               
               GetCursorPos tP
               KeyState = pConvertKeysVBToAPI(Button, Shift)
               
               ' Need calling DragEnter again, or chosen effect is mangled !?!
               ' (ie right drop menu always highlights link option)
               Call m_IDropTarget.DragEnter(m_IDataObject, KeyState, tP.X, tP.Y, (Effect))
               
               hr = m_IDropTarget.Drop(m_IDataObject, KeyState, tP.X, tP.Y, Effect)
               
               If hr = 1 Then
                  
                  ' SUCESSFUL DROP!
                  .SelectedNode = hNodeDrop
                  
                  ' RetEffect == vbDropEffectNone for optimized move
                  RetEffect = Effect
                  Debug.Assert DbgPrt(DBP >= 1, "SUCESSFUL DROP", Effect)
                  
                  If Not (m_IDTHelper Is Nothing) Then
                     Call m_IDTHelper.DragEnter(.hWnd, m_IDataObject, VarPtr(tP), Effect)
                     Call m_IDTHelper.Drop(m_IDataObject, VarPtr(tP), Effect)
                  End If
               
               End If
               Debug.Assert DbgPrt(hr = 0 And DBP >= 2, "DROP CANCELED BY USER")
               Debug.Assert DbgPrt(hr < 0 And DBP >= 0, "FAILED IDropTarget.Drop()", hr, .NodeText(hNodeDrop))
   
            End If
            Debug.Assert DbgPrt(m_IDropTarget Is Nothing And DBP >= 2, "FAILED IDropTarget", .NodeText(hNodeDrop), dbgShowAttributes(pidlFQ, pidlRel))
         
         End If
         Debug.Assert DbgPrt(m_IDataObject Is Nothing And DBP >= 1, "FAILED IDataObject")
         
      End If
      Debug.Assert DbgPrt(hNodeDrop = 0 And DBP >= 2, "FAILED hNodeDrop")
      
      Set m_IDropTarget = Nothing
      Set m_IDTHelper = Nothing
      
      If (m_hNodeDrag = 0) Then
         ' drag sourced by other app: won't receive OLECompleteDrag
         ' release IDataObject now
         Set m_IDataObject = Nothing
         ' refresh tree now
         .Refresh
    ' Else
    '    ' we sourced drag:
    '    ' keep IDataObject for use in OLECompleteDrag
    '    ' refresh tree in OLECompleteDrag
      End If
      
   End With
      
   Effect = RetEffect
      
End Sub

Private Sub m_ucTree_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
   Dim isfParent     As IShellFolder
   Dim pidlFQ        As Long
   Dim pidlRel       As Long
   Dim pidlFQDrag    As Long
   Dim hNodeDrop     As Long
   Dim RetEffect     As VBRUN.OLEDropEffectConstants
   Dim KeyState      As KEYSTATES
   Dim tP            As POINTAPI
   Dim hr            As Long
   Static hDropLast  As Long
   Static EffectLast As Long
   Dim DBP           As Integer
   DBP = 2
   
   If (Effect = vbDropEffectNone) And (State <> vbLeave) Then Exit Sub
   
   RetEffect = vbDropEffectNone
   
   With m_ucTree
      
      Select Case State
      
         Case vbOver

            '-- Get Drop info:
            If .OLEGetDropInfo(hNodeDrop) Then
               
               If (m_IDataObject Is Nothing) Then
                  ' failed getting interface on vbEnter
                  Effect = vbDropEffectNone
                  Exit Sub
               End If
               
               GetCursorPos tP
               
               ' shortcut,if hovering over same drop node
               If (hNodeDrop <> hDropLast) Then
               
                  ' -- Is item of hNodeDrop a possible drop target ? --
                  pDecryptNodeKey .NodeKey(hNodeDrop), pidlFQ, pidlRel
                  Set isfParent = GetIShellFolderParent(pidlFQ, fRtnDesktop:=True)
                  
                  Set m_IDropTarget = pIDropTarget(m_hOwner, isfParent, 1, pidlRel)
                  
                  ' non exe files have SFGAO_DROPTARGET ?, but no IDropTarget. (OK)
                  ' links have IDropTarget, IDropTarget methods fail for non exe links. (OK)
                  If Not (m_IDropTarget Is Nothing) Then
                  
                     KeyState = pConvertKeysVBToAPI(Button, Shift)
                     
                     hr = m_IDropTarget.DragEnter(m_IDataObject, KeyState, tP.X, tP.Y, Effect)
                     RetEffect = Effect
                     
                     Debug.Assert DbgPrt(hr <> S_OK And DBP >= 0, "FAILED IDropTarget.DragEnter()", hr, .NodeText(hNodeDrop))
                     
                     ' Can't drop on itself or it's children: check
                     If m_hNodeDrag Then
                        ' we sourced drag
                        pDecryptNodeKey .NodeKey(m_hNodeDrag), pidlFQDrag
                        
                        If IsParentPIDL(pidlFQDrag, pidlFQ, bImmediate:=False) Then
                           RetEffect = vbDropEffectNone
                        End If
                   ' Else
                        ' other drag source: need to read IDataObject to obtain pidlFQDrag
                        ' check omitted as explorer does neither
                     End If
                     
                  End If
                  Debug.Assert DbgPrt(m_IDropTarget Is Nothing And DBP >= 1, "FAILED IDropTarget", .NodeText(hNodeDrop), dbgShowAttributes(pidlFQ, pidlRel))
                  
                  ' Calling IDropTargetHelper is OK without valid IDropTarget !
                  If Not (m_IDTHelper Is Nothing) Then
                     hr = m_IDTHelper.DragEnter(.hWnd, m_IDataObject, VarPtr(tP), Effect)
                     Debug.Assert DbgPrt(hr <> S_OK And DBP >= 0, "FAILED IDropTargetHelper.DragEnter()", hr)
                  End If
                  
                  hDropLast = hNodeDrop
                  EffectLast = RetEffect
                  
               Else
                  ' hNodeDrop is hDropLast
                  RetEffect = EffectLast

                  ' Calling DragOver() to move image, if using a IDropTargetHelper
                ' If Not (m_IDropTarget Is Nothing) Then
                '    KeyState = pConvertKeysVBToAPI(Button, Shift)
                '    hr = m_IDropTarget.DragOver(KeyState, tP.X, tP.Y, Effect)
                '    Debug.Assert DbgPrt(hr <> S_OK And DBP >= 0, "FAILED IDropTarget.DragOver()", hr)
                ' End If
                  
                  If Not (m_IDTHelper Is Nothing) Then
                     hr = m_IDTHelper.DragOver(VarPtr(tP), Effect)
                     Debug.Assert DbgPrt(hr <> S_OK And DBP >= 0, "FAILED IDropTargetHelper.DragOver()", hr)
                  End If
                  
               End If   ' (hNodeDrop <> hDropLast)
            
            End If ' .OLEGetDropInfo(hNodeDrop)
            
         Case vbEnter
            
            ' -- reset static vars
            hDropLast = 0&
            EffectLast = 0&
            
            If m_hNodeDrag Then
               ' We started drag and are dragging over ourselves
               Debug.Assert m_hNodeDrag = .SelectedNode
               Debug.Assert Not (m_IDataObject Is Nothing)
               
            Else
               ' Drag over from other source:
               ' Get IDataObject of dragged files once for subsequent vbOver's.
               Debug.Assert (m_IDataObject Is Nothing)
               Set m_IDataObject = pIDataObjectFromVB(Data)
               Debug.Assert DbgPrt(m_IDataObject Is Nothing And DBP >= 0, "FAILED IDataObject")
            End If
            
#If WIN32_IE >= &H500 Then
            ' Create IDropTargetHelper for this D&D operation to show drag image.
            Set m_IDTHelper = pIDropTargetHelper()
#End If
         
         Case vbLeave
            
            ' -- reset static vars
            hDropLast = 0&
            EffectLast = 0&
            
            If Not (m_IDropTarget Is Nothing) Then
               m_IDropTarget.DragLeave
               Set m_IDropTarget = Nothing
            End If
               
            If Not (m_IDTHelper Is Nothing) Then
               m_IDTHelper.DragLeave
               Set m_IDTHelper = Nothing
            End If
            
            ' if we sourced drag, keep IDataObject for possible re-enter
            If m_hNodeDrag = 0 Then Set m_IDataObject = Nothing
            
            .Refresh
            
      End Select

   End With
   
   ' return current allowed effect
   Effect = RetEffect

End Sub

' ----------------------------------------------------------------------------------- '
' OLE Drag & Drop helper functions
' ----------------------------------------------------------------------------------- '

' Gets the COM IDataObject of the dragged shell item(s) from VB's DataObject,
' when drag was sourced by other app.
' Edanmo: http://www.mvps.org/emorcillo/en/code/vb6/getidataobject.shtml
Private Function pIDataObjectFromVB(ByRef Data As DataObject) As IDataObject
   Dim ido  As IDataObject
   
   ' Get an uncounted reference to the IDataObject
   CopyMemory ido, ByVal ObjPtr(Data) + 16, 4&
   
   ' return strong reference
   Set pIDataObjectFromVB = ido
   
   ' Release the weak reference
   CopyMemory ido, 0&, 4
   
End Function

Private Function pConvertKeysVBToAPI(Button As Integer, Shift As Integer) As KEYSTATES

   pConvertKeysVBToAPI = Button Or _
                         IIf(Shift And vbShiftMask, MK_SHIFT, 0) Or _
                         IIf(Shift And vbCtrlMask, MK_CONTROL, 0) Or _
                         IIf(Shift And vbAltMask, MK_ALT, 0)
End Function

' # Option1: going undocumented #
Private Function pAllowAutoExpand(ByVal pidlFQ As Long, ByVal pidlRel As Long) As Boolean
   
   Select Case GetPIDLType(pidlRel)
   
      Case pdtDrive
         pAllowAutoExpand = Not IsFloppyPIDL(pidlRel)
         
      Case pdtB1
         ' Recycled,History,Favorites,My Websites on MSN
         ' # single out My Websites on MSN: want a .NET passport? #
         pAllowAutoExpand = False
         
      Case pdtNetWorkWorkGroup To pdtNetWorkEntire
         ' any network paths
         pAllowAutoExpand = False
      
      Case Else
         pAllowAutoExpand = True
   End Select

End Function

' # Option2: strictly legal #
'Private Function pAllowAutoExpand(ByVal pidlFQ As Long, ByVal pidlRel As Long) As Boolean
'   Dim isfParent  As IShellFolder
'   Dim ulAttrs    As ESFGAO
'
'   Set isfParent = GetIShellFolderParent(pidlFQ, fRtnDesktop:=True)
'   ulAttrs = SFGAO_VALIDATE Or SFGAO_HASSUBFOLDER
'
'   If SUCCEEDED(isfParent.GetAttributesOf(1, pidlRel, ulAttrs)) Then
'
'      pAllowAutoExpand = (ulAttrs And SFGAO_HASSUBFOLDER)
'
'      If pAllowAutoExpand Then
'
'         ' # deny autoexpanding all potential trouble makers #
'         ulAttrs = SFGAO_LINK Or SFGAO_FOLDER Or SFGAO_ISSLOW
'
'         If SUCCEEDED(isfParent.GetAttributesOf(1, pidlRel, ulAttrs)) Then
'
'            Select Case True
'
'               Case (ulAttrs And SFGAO_ISSLOW)
'                  pAllowAutoExpand = False
'
'               Case (CBool(ulAttrs And SFGAO_LINK) And _
'                     CBool(ulAttrs And SFGAO_FOLDER))
'                  ' My Web Sites on MSN: want a .NET passport?
'                  pAllowAutoExpand = False
'
'            End Select
'
'         Else
'            ' sanity
'            pAllowAutoExpand = False
'         End If
'
'         If Not pAllowAutoExpand Then
'            Debug.Print "REJECTED AUTOEXPAND", GetFolderDisplayName(isfParent, pidlRel, SHGDN_NORMAL)
'            Debug.Print dbgShowAttributes(pidlFQ, pidlRel)
'         End If
'
'      End If
'
'   Else
'      ' SFGAO_VALIDATE will fail for empty floppy, without! "Insert.." dialog
'      ' (at least my floppy does with &H80004005 Unknown error)
'      pAllowAutoExpand = False
'      Debug.Print "REJECTED AUTOEXPAND", GetFolderDisplayName(isfParent, pidlRel, SHGDN_NORMAL)
'   End If
'
'End Function


#End If  ' OLEDD

' ----------------------------------------------------------------------------------- '
' Shell File Operation methods
' ----------------------------------------------------------------------------------- '

' Operates on single (virtual) item.
Private Function pShellExecute(pidlFQ As Long, Optional Verb As String = vbNullString) As Boolean
   Dim tSEI    As SHELLEXECUTEINFO

   With tSEI
      .cbSize = Len(tSEI)
      .fMask = SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_DDEWAIT
      .lpIDList = pidlFQ
      .nShow = SW_SHOWNORMAL
      .lpVerb = StrPtr(Verb)  ' Unicode!
   End With

   If ShellExecuteEx(tSEI) Then
      ' hInstApp errorcodes MSDN: SHELLEXECUTEINFO
      pShellExecute = (tSEI.hInstApp > 32)
   End If

'   ' hProcess only with filename, not pidl!
'   tSEI.fMask = SEE_MASK_NOCLOSEPROCESS
'   WaitForSingleObject(tSEI.hProcess, INFINITE)
'   CloseHandle (tSEI.hProcess)
End Function

' Operates on (multiple)(virtual) items by invoking command on item's contextmenu.
Private Function pShellOperation(ByRef isfParent As IShellFolder, _
                                 ByVal cPidls As Long, pidlRel As Long, _
                                 ByVal cmd As SHContextMenuCmdIDs _
                                 ) As Boolean
   Dim CMF        As eCMF
   Dim ulAttrs    As ESFGAO
   Dim ulAttrsRet As ESFGAO
   Dim bOk        As Boolean
   
   CMF = CMF_EXPLORE
   
   Select Case cmd
      Case cmd_Cancel
         Exit Function
      Case cmd_CreateLink
         ulAttrs = SFGAO_CANLINK
      Case cmd_Delete
         ulAttrs = SFGAO_CANDELETE
      Case cmd_Rename
         ulAttrs = SFGAO_CANRENAME
         CMF = CMF_EXPLORE Or CMF_CANRENAME
      Case cmd_Cut
         ulAttrs = SFGAO_CANMOVE
      Case cmd_Copy
         ulAttrs = SFGAO_CANCOPY
      Case cmd_Paste
         ulAttrs = SFGAO_FILESYSANCESTOR
      Case Else
         bOk = True
   End Select
   
   If Not bOk Then
   
      ulAttrsRet = ulAttrs
      
      If S_OK = isfParent.GetAttributesOf(cPidls, pidlRel, ulAttrsRet) Then
         bOk = CBool(ulAttrsRet And ulAttrs)
      End If
   End If
   
   If bOk Then
   
      InvokeShellCommand m_hOwner, isfParent, cPidls, pidlRel, cmd, , CMF
      
      ' # success / failure #
      pShellOperation = True
   Else
   
   End If
   
End Function

' # use CFSTR_PREFERREDDROPEFFECT, CFSTR_PERFORMEDDROPEFFECT, CFSTR_PASTESUCCEEDED #
Private Sub pResetCutItem(ByVal cmd As SHContextMenuCmdIDs, ByVal hNode As Long)
   
   With m_ucTree
   
      If m_hNodeCut Then
      
         Select Case cmd
         
            Case cmd_Cancel
            
            Case cmd_Cut, cmd_Copy, cmd_Paste
            
               .NodeGhosted(m_hNodeCut) = False
               m_hNodeCut = 0
               
            Case cmd_Delete, cmd_Rename
            
               If hNode = m_hNodeCut Then
                  .NodeGhosted(m_hNodeCut) = False
                  m_hNodeCut = 0
               End If
               
         End Select
   
      End If
      
      If cmd = cmd_Cut Then
         .NodeGhosted(hNode) = True
         m_hNodeCut = hNode
      End If
   
   End With
   
End Sub

' ----------------------------------------------------------------------------------- '
' Shell Notify methods
' ----------------------------------------------------------------------------------- '
#If SHCN Then

Private Function pRegisterNotify(ByVal pidl As Long) As Boolean
   Dim eSHCNE  As SHCN_EventIDs
   Dim eStrip  As SHCN_EventIDs
   
   ' set notifications, we want (MSDN: SHChangeNotify function)
   ' need file notifs to update parents regardless of m_eIncludeItems setting.
   
'   eSHCNE = eSHCNE Or SHCNE_MKDIR Or SHCNE_RMDIR Or SHCNE_RENAMEFOLDER Or SHCNE_UPDATEDIR
'   eSHCNE = eSHCNE Or SHCNE_CREATE Or SHCNE_DELETE Or SHCNE_RENAMEITEM Or SHCNE_UPDATEITEM
'   eSHCNE = eSHCNE Or SHCNE_DRIVEADD Or SHCNE_DRIVEADDGUI Or SHCNE_DRIVEREMOVED
'   eSHCNE = eSHCNE Or SHCNE_MEDIAINSERTED Or SHCNE_MEDIAREMOVED
'   eSHCNE = eSHCNE Or SHCNE_UPDATEIMAGE Or SHCNE_ATTRIBUTES
'   eSHCNE = eSHCNE Or SHCNE_NETSHARE Or SHCNE_NETUNSHARE
   
   eSHCNE = SHCNE_ALLEVENTS Or SHCNE_EXTENDED_EVENT   ' # experimenting #
   
   eStrip = eStrip Or SHCNE_FREESPACE  ' Or SHCNE_UPDATEITEM
   
   pUnregisterNotify
   
   If (m_cSN Is Nothing) Then
      Set m_cSN = New clsShellNotify
   End If
   
   With m_cSN
      .EventID = eSHCNE And Not eStrip
      
      If .Initialize(m_hOwner) Then
         If .RegisterNotify(pidl) Then
            pRegisterNotify = True
         End If
      End If
   End With

End Function

Private Sub pUnregisterNotify()
   
   If Not (m_cSN Is Nothing) Then
      m_cSN.UnregisterNotify
   End If
End Sub

Private Sub m_cSN_ShellNotification(ByVal EventID As SHCN_EventIDs, ByVal dwItem1 As Long, ByVal dwItem2 As Long, ByVal bDuplicate As Boolean)
   Dim hParent1      As Long
   Dim hParent2      As Long
   Dim hNode1        As Long
   Dim hNode2        As Long
   Dim isfParent     As IShellFolder
   Dim dwItemParent1 As Long
   Dim dwItemParent2 As Long
'   Dim pidlFQ        As Long
   Dim pidlRel       As Long
   Dim bFindItem     As Boolean
   Dim bRenamed      As Boolean
   Dim bUseTextComp  As Boolean
   Dim bRedraw       As Boolean
   Dim DBP           As Integer
   DBP = 1
   
   On Error GoTo Proc_Error

   ' GetPathFromPIDL doesn't discriminate MyComputer & MyDocuments. WalkPIDL does.
   If EventID <> SHCNE_UPDATEIMAGE Then Debug.Assert DbgPrt(DBP >= 1, dbgWalkPIDL(dwItem1), dbgWalkPIDL(dwItem2), vbCrLf)
'   Debug.Assert DbgPrt(DBP >= 1, GetFolderDisplayName(isfDesktop, dwItem1, SHGDN_FORPARSING Or SHGDN_FORADDRESSBAR), GetFolderDisplayName(isfDesktop, dwItem1, SHGDN_FORPARSING Or SHGDN_FORADDRESSBAR), vbCrLf)

   ' For duplicate notifications stored node handles are reused to save on OpenNode() calls.
   With m_tSHLast
   
      If bDuplicate Then
      
         ' duplicates, we will skip
         ' - # ? any missing functionality  ? #
         ' - # ? act on dupl and skip first ? #
         
         Select Case EventID
            Case SHCNE_RENAMEFOLDER, SHCNE_RENAMEITEM
               Debug.Assert DbgPrt(DBP >= 1, "DUPLICATE SKIPPED")
               Exit Sub
         End Select
      
         ' use stored node handles
         hParent1 = .hParent1
         hParent2 = .hParent2
         hNode1 = .hNode1
         hNode2 = .hNode2
      Else
         ' reset handles
         .hParent1 = 0
         .hParent2 = 0
         .hNode1 = 0
         .hNode2 = 0
      End If
      
   End With
   
   ' #?#
   Debug.Assert (EventID And SHCNE_EXTENDED_EVENT) = 0
   
   ' -- Item dependant   notifs: decide whether searching for item
   ' -- Item independant notifs: handle and exit
   Select Case EventID
         
      Case SHCNE_MKDIR, SHCNE_RMDIR, SHCNE_RENAMEFOLDER, SHCNE_UPDATEDIR
         ' folder events: find item if included
         bFindItem = (m_eIncludeItems And SHCONTF_FOLDERS)
         bUseTextComp = (EventID = SHCNE_RENAMEFOLDER)
   
      Case SHCNE_CREATE, SHCNE_DELETE, SHCNE_RENAMEITEM
         ' file events: find item if included
         bFindItem = (m_eIncludeItems And SHCONTF_NONFOLDERS)
         bUseTextComp = (EventID = SHCNE_RENAMEITEM)
      
      Case SHCNE_UPDATEITEM, SHCNE_ATTRIBUTES
         ' SHCNE_ATTRIBUTES: shell uses SHCNE_UPDATEITEM for folder/file attributes
         Debug.Assert EventID <> SHCNE_ATTRIBUTES
         bFindItem = True
         
      Case SHCNE_DRIVEADD, SHCNE_DRIVEADDGUI, SHCNE_DRIVEREMOVED ' , SHCNE_FREESPACE
         ' disk events: find item
         bFindItem = True
   
      Case SHCNE_MEDIAINSERTED, SHCNE_MEDIAREMOVED
         ' media events: find item if removed
         ' preceded by SHCNE_RENAMEFOLDER & SHCNE_UPDATEIMAGE
         ' no events for floppies
         bFindItem = (EventID = SHCNE_MEDIAREMOVED)
 
      Case SHCNE_NETSHARE, SHCNE_NETUNSHARE, SHCNE_SERVERDISCONNECT
         ' network events
         Debug.Print "# NOT YET IMPLEMENTED #"
         Debug.Assert False
         Exit Sub
         
      Case SHCNE_UPDATEIMAGE
         ' MSDN:    An image in the system image list has changed.
         ' Reality: An application, that was using that particular icon index
         '          in the system image list, is now using something else.
         ' ie  :    Recycle bin (changing between empty and full).
         '          Icon associated with a CD drive. (additional SHCNE_RENAMEFOLDER)
         ' Not for: Document icons that change as a result of a file type association:
         '           -> SHCNE_ASSOCCHANGED event instead.
         '          Folder icon changed by properties.customize.symbol dialog:
         '           -> SHCNE_UPDATEITEM & SHCNE_RENAMEFOLDER events instead.
         ' dwItem1: first byte of abID() is icon index,dummy byte,10 zero(?always?) bytes
         ' dwItem2: 0 or see MSDN for clues: SHChangeDWORDAsIDList, SHChangeUpdateImageIDList
         Debug.Assert dwItem1 <> 0
         Debug.Assert GetPIDLSize(dwItem1) = 12&
         Debug.Print "ICONINDEX: " & GetPIDLType(dwItem1), dbgDumpPIDL(dwItem1)
         If dwItem2 Then
            Debug.Print "ICONINDEX2: " & GetPIDLSize(dwItem2), dbgDumpPIDL(dwItem2)
         End If
         
         ' handled by SHCNE_RENAMEFOLDER,SHCNE_UPDATEITEM,SHCNE_ASSOCCHANGED handlers
         Exit Sub
      
      Case SHCNE_ASSOCCHANGED
         Debug.Print "# NOT YET IMPLEMENTED #"
         Debug.Assert False
         Exit Sub
      
      Case Else
         Debug.Print Hex$(EventID)
         Debug.Assert False
         Exit Sub
         
   End Select
   
   ' -- For all item notifs try finding parent of item in current tree
   dwItemParent1 = GetPIDLParent(dwItem1, fRtnDesktop:=True)
   
   If hParent1 = 0 Then
      hParent1 = OpenNode(dwItemParent1)
      m_tSHLast.hParent1 = hParent1
   End If
   
   ' -- Try finding requested item
   If bFindItem Then
      If hParent1 Then
         If hNode1 = 0 Then
            hNode1 = OpenNode(dwItem1, hParent1, , , , bUseTextComp)
            m_tSHLast.hNode1 = hNode1
         End If
      End If
   End If
   
   ' -- Additional preliminary work for singular events
   Select Case EventID
   
      Case SHCNE_RENAMEFOLDER, SHCNE_RENAMEITEM
         ' find new parent (and item)
         dwItemParent2 = GetPIDLParent(dwItem2, fRtnDesktop:=True)

         If hParent2 = 0 Then
            hParent2 = OpenNode(dwItemParent2)
            m_tSHLast.hParent2 = hParent2
         End If
         
         If bFindItem Then
            If hParent2 Then
               If hNode2 = 0 Then
                  hNode2 = OpenNode(dwItem2, hParent2, , , , bUseTextComp)
                  m_tSHLast.hNode2 = hNode2
               End If
            End If
         End If
         
      Case SHCNE_RMDIR, SHCNE_DELETE
         ' # error if searching dwItemParent2 or dwItem2 in recyclebin #
   
   End Select
         
   ' -- Is notification relevant for current tree contents ?
   
   If (hParent1 Or hNode1 Or hParent2 Or hNode2) = 0 Then
     Debug.Assert DbgPrt(DBP >= 1, "NOT FOR US !?!", bDuplicate)
     GoTo DONE
   End If
   
   ' sanity
   Debug.Assert Not (hParent1 = 0 And hNode1 <> 0)
   Debug.Assert Not (hParent2 = 0 And hNode2 <> 0)
   
   ' -- Act upon notification (no Exit's , no GoTo's anymore)
   
   With m_ucTree
   
      bRedraw = .Redraw
      .Redraw = False
      
      Select Case EventID
      
         Case SHCNE_MKDIR, SHCNE_CREATE
            ' folder/file has been created (dwItem2 == nil)
            
            If (hParent1 <> 0) Then
            
               ' update parent folder, if folder has been created
               If EventID = SHCNE_MKDIR Then
                  .NodePlusMinusButton(hParent1) = True
               End If
               
               If bFindItem Then
                  
                  If (hNode1 = 0) And .NodeExpandedOnce(hParent1) Then
                     ' parent has expanded before: add item
                     
                     Set isfParent = GetIShellFolder(isfDesktop, dwItemParent1)
                     ' create copied relative pidl. Copied again by pInsertItem -> free!
                     pidlRel = GetItemID(dwItem1, GIID_LAST)
                     hNode1 = pInsertItem(isfParent, dwItem1, pidlRel, hParent1)
                     FreePIDL pidlRel
                     Debug.Assert hNode1
                     
                     ' update it(imageindices) and sort parent
                     pUpdateNodePIDL hParent1, dwItemParent1, hNode1, dwItem1, _
                                     bSort:=True
                     Debug.Assert OpenNode(dwItem1, hParent1) = hNode1
                     
                  End If
               End If
               
            End If
            
         Case SHCNE_RMDIR, SHCNE_DELETE
            ' folder/file has been removed: delete from tree
            
            If (hParent1 <> 0) Then
               
               If (hNode1 <> 0) Then
                  ' item is in tree. pidl's are freed in BeforeDelete event.
                  Debug.Assert bFindItem
                  m_ucTree.DeleteNode hNode1
                  hNode1 = 0
               End If
            
               ' update parent folder, if folder has been removed
               If EventID = SHCNE_RMDIR Then
                  If .NodeChild(hParent1) = 0 Then
                     ' ensure it is collapsed !
                     .Collapse hParent1
                     .NodePlusMinusButton(hParent1) = False
                  End If
               End If
            
            End If
            
         Case SHCNE_RENAMEFOLDER, SHCNE_RENAMEITEM
            ' folder/file has been renamed or moved: rename or move in tree
            ' Drag & Drop or Cut/Copy/Paste operations: move in tree
            ' SHCNE_RENAMEFOLDER sent also, if folder icon changed
            
            ' equal parents == item was renamed, otherwise item was moved.
            bRenamed = (isfDesktop.CompareIDs(0&, dwItemParent1, dwItemParent2) = 0)
            
            If (hParent1 <> 0) Then
               
               If bRenamed Then
   
                  If (hNode1 <> 0) Then
                     ' item is in tree: update it with new dwItem2 and sort
                     Debug.Assert bFindItem
                     pUpdateNodePIDL hParent1, dwItemParent2, hNode1, dwItem2, _
                                     bSort:=True, bRename:=True
                  End If
               
               Else
                  ' Moving item: delete old item, if exists
                  If (hNode1 <> 0) Then
                     ' pidl's are freed in BeforeDelete event.
                     Debug.Assert bFindItem
                     m_ucTree.DeleteNode hNode1
                     hNode1 = 0
                  End If
               
                  ' update old parent folder, if folder has been moved
                  If EventID = SHCNE_RENAMEFOLDER Then
                     If .NodeChild(hParent1) = 0 Then
                        ' ensure it is collapsed !
                        .Collapse hParent1
                        .NodePlusMinusButton(hParent1) = False
                     End If
                  End If
               End If
                                   
            End If
            
            If (hParent2 <> 0) Then
            
               If Not bRenamed Then
                  ' Moving item
                  Debug.Assert hNode2 = 0 Or bDuplicate
               
                  If bFindItem Then
                  
                     ' add item as child of hParent2, if has expanded before
                     If (hNode2 = 0) And .NodeExpandedOnce(hParent2) Then
                        
                        Set isfParent = GetIShellFolder(isfDesktop, dwItemParent2)
                        ' create copied relative pidl. Copied again by pInsertItem -> free!
                        pidlRel = GetItemID(dwItem2, GIID_LAST)
                        hNode2 = pInsertItem(isfParent, dwItem2, pidlRel, hParent2)
                        FreePIDL pidlRel
                        Debug.Assert hNode2
                        
                        ' update it(imageindices!) and sort new parent
                        pUpdateNodePIDL hParent2, dwItemParent2, hNode2, dwItem2, _
                                        bSort:=True
                     End If
                  End If
               
                  ' update new parent folder, if folder has been moved
                  If EventID = SHCNE_RENAMEFOLDER Then
                     .NodePlusMinusButton(hParent2) = True
                  End If
               
               End If   ' Not bRenamed
            End If   ' (hParent2 <> 0)
               
            
         Case SHCNE_MEDIAREMOVED
            ' media events: remove drive folder's children
            ' preceded by SHCNE_RENAMEFOLDER & SHCNE_UPDATEIMAGE
            
            If (hNode1 <> 0) Then
               Debug.Assert bFindItem
               'pidl's are freed in BeforeDelete events.
               .Collapse hNode1, , RemoveChildren:=True
               
               .NodePlusMinusButton(hNode1) = False
            End If
            
            
         Case SHCNE_DRIVEADD, SHCNE_DRIVEADDGUI, SHCNE_DRIVEREMOVED
            ' disk events : SHCNE_UPDATEDIR for MyComputer follows
            
            
         Case SHCNE_UPDATEDIR
            ' sent if more than 10 items changed in folder (ie move multiple items)
            If (hNode1 <> 0) Then
               ' item is in tree: update it's contents
               Debug.Assert bFindItem
               pUpdateNodePIDL hParent1, dwItemParent1, hNode1, dwItem1
               
               If .NodeExpandedOnce(hNode1) Then
                  ' compare hNode1's enumeration object with it's current child nodes
                  ' add / delete children then sort
                  pUpdateFolder hNode1
               End If
            End If
               
         Case SHCNE_UPDATEITEM
         
            If (hNode1 <> 0) Then
               ' item is in tree: update it
               Debug.Assert bFindItem
               pUpdateNodePIDL hParent1, dwItemParent1, hNode1, dwItem1
            End If
            
      End Select
      
      .Redraw = bRedraw
   
   End With
   
   ' store new handles
   With m_tSHLast
      .hNode1 = hNode1
      .hNode2 = hNode2
   End With
   
DONE:
   
   FreePIDL dwItemParent1
   FreePIDL dwItemParent2
   
   Exit Sub

Proc_Error:
   Debug.Print "Error: " & Err.Number & ". " & Err.Description
'   MsgBox "Error: " & Err.Number & ". " & Err.Description, vbOKOnly Or vbCritical, App.Title & ".clsShellTree: Sub m_ucTree_ShellNotify"
   If InIDE Then Stop: Resume
End Sub

#End If  ' SHCN

' pidlFQParent & pidlFQ freed by caller.
' Frees existing node pidls and assigns new.
' hParent & pidlFQParent: can pass zero
Private Function pUpdateNodePIDL(ByVal hParent As Long, ByVal pidlFQParent As Long, _
                                 ByVal hNode As Long, ByVal pidlFQ As Long, _
                                 Optional ByVal bSort As Boolean = False, _
                                 Optional ByVal bRename As Boolean = False) As Boolean
   Dim isfParent     As IShellFolder
   Dim ulAttrs       As ESFGAO
   Dim pidlRel       As Long
   Dim idxNode       As Long
   Dim idxImg        As Long
   Dim idxSelImg     As Long
   Dim sNameOld      As String
   Dim sNameNew      As String
   Dim bEnumObj      As Boolean
   Dim DBP           As Integer
   DBP = 0
   On Error GoTo Proc_Error
   
   ' # Problem: Icons of new items/ customized folders                                  #
   ' # Files without registered association have an index m_idxNewFile (=3).            #
   ' # So do created new folders/files/links or new folders with customized icon.       #
   ' # They just exist as streams until Explorer decides to update its contents.        #
   ' # Yet it shows ie a new folder with normal icons.                                  #
   ' # How can we do it?                                                                #
   ' # Using pidls sent in the shell change notifies, always results in m_idxNewFile.   #
   ' # (pidlRel data is zero for size,date,attributes).                                 #
   ' # isf.GetAttributesOf(): even with SFGAO_VALIDATE allows no discrimination of item.#
   ' # File API's return nonsense.                                                      #
   ' # Solution:                                                                        #
   ' # Asking for an enumeration object for the parent folder updates contents.         #
   ' # Use new relative pidl from enumeration and also create a new fully qualified.    #
   
   ' # Problem: Renamed files may be sorted before folders.                             #
   ' # Ask for new enumeration, if sorting is needed and files are included:            #
   
   If pidlFQParent Then
      ' Get the parent's IShellFolder from its fully qualified pidl
      Set isfParent = GetIShellFolder(isfDesktop, pidlFQParent)
   Else
      ' Get the parent's IShellFolder from item's fully qualified pidl
      Set isfParent = GetIShellFolderParent(pidlFQ, fRtnDesktop:=True)
   End If
   
#If WIN32_IE >= &H500 Then
   idxImg = SHMapPIDLToSystemImageListIndex(isfParent, ILFindLastID(pidlFQ), idxSelImg)
'   idxImg = SHMapPIDLToSystemImageListIndex(isfDesktop, pidlFQ, idxSelImg)
#Else
   idxImg = GetFileIconIndexPIDL(pidlFQ, SHGFI_SMALLICON)
   idxSelImg = GetFileIconIndexPIDL(pidlFQ, SHGFI_SMALLICON Or SHGFI_OPENICON)
#End If
   
   ' obtain new enumeration object for new items
   bEnumObj = (idxImg = m_idxNewFile)
   ' obtain new enumeration object, when sorting files
   bEnumObj = bEnumObj Or (bSort And (m_eIncludeItems And SHCONTF_NONFOLDERS))
   
   If bEnumObj Then
      ' get fresh relative pidl.Store as key -> don't free!
      If bRename Then
         ' Passing node's pidlRel, pGetNewPidlRel fails for renamed items, use new pidlFQ.
         pidlRel = pGetNewPidlRel(isfParent, , pidlFQ)
      Else
         pidlRel = pGetNewPidlRel(isfParent, hNode)
      End If
      Debug.Assert pidlRel
      
      ' create new fully qualified pidl.Store as key -> don't free! $SHGetRealIDL
      pidlFQ = CombinePIDLs(pidlFQParent, pidlRel)
      Debug.Assert pidlFQ
      
      ' get icons now with new pidl
      If (idxImg = m_idxNewFile) Then
      
#If WIN32_IE >= &H500 Then
         idxImg = SHMapPIDLToSystemImageListIndex(isfParent, pidlRel, idxSelImg)
#Else
         idxImg = GetFileIconIndexPIDL(pidlFQ, SHGFI_SMALLICON)
         idxSelImg = GetFileIconIndexPIDL(pidlFQ, SHGFI_SMALLICON Or SHGFI_OPENICON)
#End If
      End If
      
   Else
      ' create copied relative pidl.Store as key -> don't free!
      pidlRel = GetItemID(pidlFQ, GIID_LAST)
      ' create new fully qualified pidl.Store as key -> don't free!
      pidlFQ = CopyPIDL(pidlFQ)
   End If
   
   pFreePidls hNode
   
   ' Get the item's attributes, specify what attributes we want.
   ulAttrs = SFGAO_HASSUBFOLDER Or SFGAO_HIDDEN Or _
             SFGAO_SHARE Or SFGAO_LINK ' Or SFGAO_GHOSTED
   
   If m_bShowCompressed Then ulAttrs = ulAttrs Or SFGAO_COMPRESSED
   
   If S_OK <> isfParent.GetAttributesOf(1, pidlRel, ulAttrs) Then
      Debug.Assert False:  Exit Function
   End If
   
   With m_ucTree
      Debug.Assert .Redraw = False
      
      .NodeKey(hNode, idxNode) = CStr(pidlFQ) & ":" & CStr(pidlRel)
      Debug.Assert idxNode
      
      Debug.Assert DbgPrt(DBP >= 2, "OLDCONTENT", .NodeText(, idxNode), .NodeImage(, idxNode), .NodeSelectedImage(, idxNode), .NodePlusMinusButton(hNode), .NodeGhosted(hNode))
      
      If bRename Then
         sNameOld = .NodeText(, idxNode)
         sNameNew = GetFolderDisplayName(isfParent, pidlRel, SHGDN_INFOLDER)
         .NodeText(, idxNode) = sNameNew
      End If
      
      .NodePlusMinusButton(hNode) = CBool(ulAttrs And SFGAO_HASSUBFOLDER)
      .NodeGhosted(hNode) = CBool(ulAttrs And SFGAO_HIDDEN)
      
      .NodeImage(, idxNode) = idxImg
      .NodeSelectedImage(, idxNode) = idxSelImg
      
      If (ulAttrs And SFGAO_SHARE) Then
         .NodeOverlayImage(hNode) = 1
      ElseIf (ulAttrs And SFGAO_LINK) Then
         .NodeOverlayImage(hNode) = 2
      Else
         .NodeOverlayImage(hNode) = 0
      End If
      
#If CUSTDRAW Then
      If (ulAttrs And SFGAO_COMPRESSED) Then
         ' # HowTo get color value for compressed files ? #
         .NodeForecolor(hNode) = vbBlue
      End If
#End If
      
      Debug.Assert DbgPrt(DBP >= 1, "NEWCONTENT", .NodeText(, idxNode), .NodeImage(, idxNode), .NodeSelectedImage(, idxNode), .NodePlusMinusButton(hNode), .NodeGhosted(hNode))
      
      If bSort Then
         If hParent = 0 Then hParent = .NodeParent(hNode)
         bSort = (hParent <> 0)
      End If
      
      If bSort Then
         
         If bRename Then
            If LenB(sNameNew) = LenB(sNameOld) Then
               bSort = (sNameNew <> sNameOld)
            End If
         End If
         
         If bSort Then
            ' Setup the callback and sort the parent folder
            m_ucTree.SortChildrenCB hParent, AddressOf modSortTreeCB.TreeViewCompareProc, _
                                    False, ObjPtr(isfParent)
         End If
      End If
      
      If hNode = .SelectedNode Then
         ' update CurrentXXX properties + any clients by raising selection event
         .SelectedNode = 0
         .SelectedNode = hNode
      End If
      
   End With

   pUpdateNodePIDL = True

   Exit Function

Proc_Error:
   Debug.Print "Error: " & Err.Number & ". " & Err.Description, ".clsShellTree: Function pUpdateNodePIDL"
   If InIDE Then Stop: Resume
End Function

' pass hChild to use node's current pidlRel for comparing.
' Or pass a new pidlFQ (from SHC notifs )for renamed item.
' Returns a new relative pidl, obtained from enumeration object.On failure returns old.
Private Function pGetNewPidlRel(isfParent As IShellFolder, _
                                Optional ByVal hChild As Long, _
                                Optional ByVal pidlFQ As Long) As Long
   Dim iEIDL      As IEnumIDList
   Dim pidlRelEn  As Long
   Dim pidlRel    As Long
   
   Debug.Assert hChild Xor pidlFQ
   
   If hChild Then
      ' use node's current relative pidl
      pDecryptNodeKey m_ucTree.NodeKey(hChild), , pidlRel
   Else
      ' create relative pidl from passed fully qualified.Free afterwards.
      pidlRel = GetItemID(pidlFQ, GIID_LAST)
   End If
      
   ' Create an enumeration object for the parent folder.
   If S_OK = isfParent.EnumObjects(m_hOwner, m_eIncludeItems, iEIDL) Then
                                                                
      ' Enumerate the contents of the parent folder
      Do While (iEIDL.Next(1, pidlRelEn, 0) = NOERROR)
         
         If (S_OK = isfParent.CompareIDs(0&, pidlRelEn, pidlRel)) Then
         
            ' Found node's item: return the new relative pidl from enumeration.
            pGetNewPidlRel = pidlRelEn
            
            Exit Do
         End If
         
         ' Free the relative pidl the enumeration gave us.
         FreePIDL pidlRelEn
      
      Loop   ' ieidl.Next
    
   End If   ' SUCCEEDED(EnumObjects))
   
   If pGetNewPidlRel = 0 Then
      ' failed: return original
      Debug.Assert False
      pGetNewPidlRel = pidlRel
   End If
   
End Function

' ----------------------------------------------------------------------------------- '
' Misc public helper functions
' ----------------------------------------------------------------------------------- '

Public Function NewFolder(Optional bAsSibling As Boolean = False) As Boolean
   Dim isfParent     As IShellFolder
   Dim pidlFQ        As Long
   
   On Error GoTo Proc_Error

   With m_ucTree
      pDecryptNodeKey .NodeKey(.SelectedNode), pidlFQ
      
      If bAsSibling Then
         ' Folder created as sibling of selected
         Set isfParent = GetIShellFolderParent(pidlFQ, fRtnDesktop:=False)
      Else
         ' Folder created as child of selected
         Set isfParent = GetIShellFolder(isfDesktop, pidlFQ, fRtnDesktop:=False)
      End If
      
      If Not (isfParent Is Nothing) Then
         m_hNodeAdded = 0
         
         NewFolder = InvokeShellViewCmd(m_hOwner, isfParent, , scmd_NewFolder)
         
         ' select last inserted node
         If m_hNodeAdded Then
            .SelectedNode = m_hNodeAdded
         End If
      End If
      
   End With

   Exit Function

Proc_Error:
   Debug.Print "Error: " & Err.Number & ". " & Err.Description
   If InIDE Then Stop: Resume
End Function

' Is pidl fully qualified and exists ?
' # use IsBadReadPtr/IsBadWritePtr #
Public Function ValidatePidl(ByVal pidlFQ As Long) As Boolean
   
   If pidlFQ <> 0 Then
   
      Select Case True
         
         Case pidlFQ < 0
            ' Memory pointers are unsigned
         Case HIWORD(pidlFQ) = 0
            ' Memory pointers will always be >64K
         Case Else
            ValidatePidl = pValidatePidl(isfDesktop, pidlFQ)
            
      End Select
   
   Else
      ' pidl = 0 is valid for Desktop!
      ValidatePidl = True
   End If
   
   Debug.Assert ValidatePidl
End Function

Private Function pValidatePidl(isfParent As IShellFolder, pidlRel As Long, _
                               Optional cPidl As Long = 1) As Boolean
   ' # Option1: going undocumented                #
'   ' # Option2: compare GetPathFromPIDL() & "A:\" #
   If Not IsFloppyPIDL(pidlRel) Then
      pValidatePidl = (S_OK = isfParent.GetAttributesOf(cPidl, pidlRel, (SFGAO_VALIDATE)))
   Else
      ' SFGAO_VALIDATE will fail for empty floppy with &H80004005 Unknown error
      pValidatePidl = True
   End If
End Function

' This function tests the validity of the file and path. It works only on the local file system or on a remote drive that has been mounted to a drive letter. It will return FALSE for remote file paths that begin with the UNC names \\server or \\server\share. It will also return FALSE if a mounted remote drive is out of service.
Public Function FileExists(Path As String) As Boolean
   FileExists = PathFileExistsW(StrPtr(Path))
End Function

' Returns True if the folder name is valid.('C:' is valid)
' Returns false for empty drive "A:\"
Public Function FolderExists(Path As String) As Boolean
   FolderExists = PathIsDirectoryW(StrPtr(Path))
End Function

Public Function GetExplorerSettings(ByVal RequestedFlags As SSF_Flags, _
                                    Optional ByRef ReturnedFlags As SSF_Flags) As Boolean
                               
   Dim b(0 To 1) As Byte   ' SHELLFLAGSTATE is actually a 16 bit type
   
   ReturnedFlags = 0
   
   Debug.Assert RequestedFlags > 0 Or RequestedFlags = -1
   
   SHGetSettingsAny b(0), ByVal RequestedFlags
   
   If (b(0) And &H1) = &H1 Then ReturnedFlags = ReturnedFlags Or SSF_SHOWALLOBJECTS
   If (b(0) And &H2) = &H2 Then ReturnedFlags = ReturnedFlags Or SSF_SHOWEXTENSIONS
   If (b(0) And &H4) = &H4 Then ReturnedFlags = ReturnedFlags Or SSF_HIDDENFILEEXTS '???
   If (b(0) And &H10) = &H10 Then ReturnedFlags = ReturnedFlags Or SSF_SHOWCOMPCOLOR
   If (b(0) And &H20) = &H20 Then ReturnedFlags = ReturnedFlags Or SSF_DOUBLECLICKINWEBVIEW
   If (b(0) And &H40) = &H40 Then ReturnedFlags = ReturnedFlags Or SSF_DESKTOPHTML
   If (b(0) And &H80) = &H80 Then ReturnedFlags = ReturnedFlags Or SSF_WIN95CLASSIC
   
   If (b(1) And &H4) = &H4 Then ReturnedFlags = ReturnedFlags Or SSF_NOCONFIRMRECYCLE
   If (b(1) And &H8) = &H8 Then ReturnedFlags = ReturnedFlags Or SSF_SHOWINFOTIP
   If (b(1) And &H80) = &H80 Then ReturnedFlags = ReturnedFlags Or SSF_SHOWSYSFILES
   
   If (ReturnedFlags And RequestedFlags) = RequestedFlags Then
      GetExplorerSettings = True
   End If
   
'   Dim s As String
'   s = _
'      IIf((ReturnedFlags And SSF_SHOWALLOBJECTS), "SHOWALLOBJECTS", "") & _
'      IIf((ReturnedFlags And SSF_SHOWEXTENSIONS), vbTab & "SHOWEXTENSIONS", "") & _
'      IIf((ReturnedFlags And SSF_HIDDENFILEEXTS), vbTab & "HIDDENFILEEXTS", "") & _
'      IIf((ReturnedFlags And SSF_SHOWCOMPCOLOR), vbTab & "SHOWCOMPCOLOR", "") & _
'      IIf((ReturnedFlags And SSF_SORTCOLUMNS), vbTab & "SORTCOLUMNS", "") & _
'      IIf((ReturnedFlags And SSF_SHOWSYSFILES), vbTab & "SHOWSYSFILES", "") & _
'      IIf((ReturnedFlags And SSF_DOUBLECLICKINWEBVIEW), vbTab & "DOUBLECLICKINWEBVIEW", "") & _
'      IIf((ReturnedFlags And SSF_SHOWATTRIBCOL), vbTab & "SHOWATTRIBCOL", "") & _
'      IIf((ReturnedFlags And SSF_DESKTOPHTML), vbTab & "DESKTOPHTML", "") & _
'      IIf((ReturnedFlags And SSF_WIN95CLASSIC), vbTab & "WIN95CLASSIC", "") & _
'      IIf((ReturnedFlags And SSF_DONTPRETTYPATH), vbTab & "DONTPRETTYPATH", "") & _
'      IIf((ReturnedFlags And SSF_MAPNETDRVBUTTON), vbTab & "MAPNETDRVBUTTON", "") & _
'      IIf((ReturnedFlags And SSF_SHOWINFOTIP), vbTab & "SHOWINFOTIP", "") & _
'      IIf((ReturnedFlags And SSF_HIDEICONS), vbTab & "HIDEICONS", "") & _
'      IIf((ReturnedFlags And SSF_NOCONFIRMRECYCLE), vbTab & "NOCONFIRMRECYCLE", "") & _
'      IIf((ReturnedFlags And SSF_FILTER), vbTab & "FILTER", "") & _
'      IIf((ReturnedFlags And SSF_WEBVIEW), vbTab & "WEBVIEW", "") & _
'      IIf((ReturnedFlags And SSF_SHOWSUPERHIDDEN), vbTab & "SHOWSUPERHIDDEN", "") & _
'      IIf((ReturnedFlags And SSF_SEPPROCESS), vbTab & "SEPPROCESS", "") & _
'      IIf((ReturnedFlags And SSF_NONETCRAWLING), vbTab & "NONETCRAWLING", "") & _
'      IIf((ReturnedFlags And SSF_STARTPANELON), vbTab & "STARTPANELON", "") & _
'      IIf((ReturnedFlags And SSF_SHOWSTARTPAGE), vbTab & "SHOWSTARTPAGE", "")
'
'   Debug.Print s
End Function

Public Function IsPidlInMyDocuments(ByVal pidlFQ As Long, _
                                    Optional ByRef IsMyDocuments As Boolean) As Boolean
   Dim pidlCopy      As Long
   Dim pidlMyDoc     As Long
   Dim hr            As Long
   
   On Error GoTo Proc_Error
   
   IsMyDocuments = False

   pidlCopy = GetItemID(pidlFQ, GIID_FIRST)
   
   If pidlCopy Then
      
      pidlMyDoc = GetPIDLFromCSIDL(m_hOwner, CSIDL_PERSONAL)
      
      hr = isfDesktop.CompareIDs(0&, pidlCopy, pidlMyDoc)
      
      FreePIDL pidlCopy
      FreePIDL pidlMyDoc
                     
      If hr = 0 Then
         ' pidl is in MyDocuments folder!
         IsPidlInMyDocuments = True
         IsMyDocuments = (GetItemIDCount(pidlFQ) = 1)
      End If
   End If

   Exit Function

Proc_Error:
   Debug.Print "Error: " & Err.Number & ". " & Err.Description, App.Title & ".clsShellTree: Function IsPidlInMyDocuments"
   If InIDE Then Stop: Resume
End Function

Public Function ConvertPathInMyComputerToMyDocuments(ByRef sPathMyCom As String, _
                                                     ByRef sPathMyDoc As String) As Boolean
   Dim sMyDoc  As String
   Dim pidl    As Long
   
   ' on failure return same path
   sPathMyDoc = sPathMyCom
   
   If LenB(sPathMyCom) Then
      ' # ? sMyDoc = GetPathFromCSIDL(m_hOwner, CSIDL_PERSONAL) returns rubbish ? #
      pidl = GetPIDLFromCSIDL(m_hOwner, CSIDL_PERSONAL)
      sMyDoc = GetPathFromPIDL(pidl)
      FreePIDL pidl
      
      If LenB(sMyDoc) Then
         ' strip physical MyDocuments path & prepend MyDocuments GUID
         If InStr(1, sPathMyCom, sMyDoc) = 1 Then
            sPathMyDoc = Mid$(sPathMyCom, Len(sMyDoc))
            sPathMyDoc = "::{450d8fba-ad25-11d0-98a8-0800361b1103}" & sPathMyDoc
            
            ConvertPathInMyComputerToMyDocuments = True
         End If
      End If
   End If
   
End Function

#If MULSEL Then
   
Private Sub pGetSelectedItemsPidl(isfParent As IShellFolder, cPidls As Long, aPidls() As Long)
   Dim pidlFQ     As Long
   Dim hParent    As Long
   Dim hSel       As Long
   Dim idx        As Long
   
   With m_ucTree
      
      ' ensure SelectedNode is in selection (ucTree might wait for MouseUp to update)
      .NodeSelected(.SelectedNode) = True
      
      hParent = .NodeParent(.SelectedNode)
      
      ' remove all non-siblings of SelectedNode from selection
      For idx = .SelectionCount To 1 Step -1
         
         hSel = .SelectionNode(idx)
         
         If hParent <> .NodeParent(hSel) Then
            .NodeSelected(hSel) = False
         End If
      Next
      
      pDecryptNodeKey .NodeKey(.SelectedNode), pidlFQ
      Set isfParent = GetIShellFolderParent(pidlFQ, fRtnDesktop:=True)
      cPidls = .SelectionCount
      
      If cPidls Then
         
         ReDim aPidls(1 To cPidls)
         
         For idx = 1 To cPidls
            ' aPidls(idx): relative pidl of dragged item
            pDecryptNodeKey .NodeKey(.SelectionNode(idx)), , aPidls(idx)
         Next
         
      Else: Debug.Assert False
      End If
   
   End With
   
End Sub
   
#Else

Private Sub pGetSelectedItemsPidl(isfParent As IShellFolder, cPidls As Long, aPidls() As Long)
   Dim pidlFQ  As Long
   
   cPidls = 1
   ReDim aPidls(1)
   
   With m_ucTree
      pDecryptNodeKey .NodeKey(.SelectedNode), pidlFQ, aPidls(1)
   End With
   
   Set isfParent = GetIShellFolderParent(pidlFQ, fRtnDesktop:=True)
   
End Sub
   
#End If  ' MULSEL



