Attribute VB_Name = "modIShellFolder"
'---------------------------------------------------------------------------------------
' Module    : modIShellFolder
' Author    : OrlandoCurioso 16.05.2005 / Brad Martinez
' Purpose   : mIShellFolder for >=Win2K
' Requires  : IShellFolder Extended Type Library v1.3 OC (ISHF_Ex.tlb)
'             Public Function SUCCEEDED(hr As Long) as Boolean
'---------------------------------------------------------------------------------------
Option Explicit

' Procedure responsibility of pidl memory, # unless specified otherwise #:
' - Calling procedures are solely responsible for freeing pidls they create,
'   or receive as a return value from a called procedure.
' - Called procedures always copy pidls received in their params, and
'   *never* free pidl params.

#If WIN32_IE >= &H500 Then

' GetItemID item ID retrieval constants
Public Const GIID_FIRST = 1
Public Const GIID_LAST = -1

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (pDest As Any, ByVal dwLength As Long, ByVal bFill As Byte)

Private Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

' // shell32

'Private Declare Function SHGetMalloc Lib "shell32" (ppMalloc As IMalloc) As Long
Private Declare Function SHAlloc Lib "shell32" (ByVal cb As Long) As Long
'Private Declare Sub SHFree Lib "shell32" (ByVal pv As Long)

Private Declare Function SHGetDesktopFolder Lib "shell32" (ppshf As IShellFolder) As Long
Private Declare Function SHBindToParent Lib "shell32.dll" (ByVal pidl As Long, riid As IShellFolderEx_TLB.GUID, psf As IShellFolder, ppidlLast As Long) As Long
Private Declare Function SHCloneSpecialIDList Lib "shell32" (ByVal hwndOwner As Long, ByVal csidl As Long, ByVal fCreate As Boolean) As Long
'Private Declare Function SHGetRealIDL Lib "shell32.dll" (psf As IShellFolder, ByVal pidlSimple As Long,ByVal pidlFQ As Long) As Long
'Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal pidl As Long) As Long
Private Declare Function SHILCreateFromPath Lib "shell32" (ByVal pwszPath As Long, pidl As Long, rgflnOut As Long) As Long
Private Declare Function SHGetPathFromIDListW Lib "shell32.dll" (ByVal pidl As Long, ByVal lpszPath As Long) As Long


Public Enum SHGFP_TYPE
   SHGFP_TYPE_CURRENT = &H0      ' current value for user, verify it exists
   SHGFP_TYPE_DEFAULT = &H1
End Enum
Private Declare Function SHGetFolderPathW Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As Long) As Long

' // shell32: pidl functions

'Private Declare Function ILAppendID Lib "shell32" (ByVal pidl As Long, ByVal pmkid As Long, ByVal fAppend As Boolean) As Long
Private Declare Function ILClone Lib "shell32" (ByVal pidl As Long) As Long
Private Declare Function ILCloneFirst Lib "shell32" (ByVal pidl As Long) As Long
Private Declare Function ILCombine Lib "shell32" (ByVal pidl1 As Long, ByVal pidl2 As Long) As Long
Private Declare Function ILCreateFromPathW Lib "shell32" (ByVal pwszPath As Long) As Long
'Private Declare Function ILFindChild Lib "shell32" (ByVal pidlParent As Long, ByVal pidlChild As Long) As Long
Private Declare Function ILFindLastID Lib "shell32" (ByVal pidl As Long) As Long
Private Declare Sub ILFree Lib "shell32" (ByVal pidl As Long)
Private Declare Function ILGetNext Lib "shell32" (ByVal pidl As Long) As Long
Private Declare Function ILGetSize Lib "shell32" (ByVal pidl As Long) As Long
'Private Declare Function ILIsEqual Lib "shell32" (ByVal pidl1 As Long, ByVal pidl2 As Long) As Boolean
Private Declare Function ILIsParent Lib "shell32" (ByVal pidlParent As Long, ByVal pidlBelow As Long, ByVal fImmediate As Boolean) As Boolean
'Private Declare Function ILLoadFromStream Lib "shell32" (pstm As Any,ByVal pidl As Long) As Long
Private Declare Function ILRemoveLastID Lib "shell32" (ByVal pidl As Long) As Boolean
'Private Declare Function ILSaveToStream Lib "shell32" (pstm As Any, ByVal pidl As Long) As Long

' // shlwapi: STRRET conversion
'Private Declare Function StrRetToBufW Lib "shlwapi.dll" (pstr As STRRET, ByVal pidl As Long, szBuf As Long, ByVal cchBuf As Long) As Long
Private Declare Function StrRetToStrW Lib "shlwapi.dll" (pstr As STRRET, ByVal pidl As Long, ppszName As Long) As Long
'

Public Function GetPathFromCSIDL(ByVal hOwner As Long, ByVal csidl As eCSIDL_VALUES, _
                                 Optional ByVal shgfp As SHGFP_TYPE = SHGFP_TYPE_DEFAULT, _
                                 Optional ByVal Flags As eCSIDL_FLAGS = 0) As String
   Dim sPath   As String * MAX_PATH
   Dim lpszW   As Long
   Dim lenS    As Long
   
   lpszW = StrPtr(sPath)
   If SUCCEEDED(SHGetFolderPathW(hOwner, csidl Or Flags, -1, shgfp, lpszW)) Then
      
      lenS = lstrlenW(ByVal lpszW)
      GetPathFromCSIDL = Space$(lenS)
      
      MoveMemory ByVal StrPtr(GetPathFromCSIDL), ByVal lpszW, 2 * lenS
   End If
End Function
   
' Returns a complex pidl (relative to the desktop) from a special folder's ID.
'   hOwner  - handle of window that will own any displayed msg boxes
'   nFolder - special folder index
'   bCreate - folder should be created if it does not already exist.

Public Function GetPIDLFromCSIDL(ByVal hOwner As Long, ByVal nFolder As eCSIDL_VALUES, _
                                 Optional ByVal bCreate As Boolean) As Long
   GetPIDLFromCSIDL = SHCloneSpecialIDList(hOwner, nFolder, bCreate)
End Function

' Returns an absolute pidl (relative to the desktop) from a file system path only (GUID OK).
' Returns zero desktop pidl for empty string.
'   hWnd    - junk (compability with Win98 module)
'   sPath   - fully qualified path (i.e. not from a display name)

Public Function GetPIDLFromPath(hWnd As Long, sPath As String) As Long
   GetPIDLFromPath = ILCreateFromPathW(StrPtr(sPath))
End Function

' Returns an absolute pidl (relative to the desktop) from a file system path only (GUID OK).
' Returns desktop pidl for empty string.
'   sPath   - fully qualified path (i.e. not from a display name)
'   ulAttrs - Attributes of the folder that the caller would like to retrieve along with the PIDL.

Public Function GetPIDLFromPathAttr(sPath As String, Optional ByVal ulAttrs As ESFGAO) As Long
   Dim pidl As Long
   
   If SUCCEEDED(SHILCreateFromPath(StrPtr(sPath), pidl, ulAttrs)) Then
      GetPIDLFromPathAttr = pidl
   End If
End Function

' Returns an absolute pidl's path only (doesn't rtn display names!)

Public Function GetPathFromPIDL(ByVal pidl As Long) As String
   Dim sPath   As String * MAX_PATH
   Dim lpszW   As Long
   Dim lenS    As Long
   
   If pidl Then
      
      lpszW = StrPtr(sPath)
      If SHGetPathFromIDListW(pidl, lpszW) Then
      
         lenS = lstrlenW(ByVal lpszW)
         GetPathFromPIDL = Space$(lenS)
         
         MoveMemory ByVal StrPtr(GetPathFromPIDL), ByVal lpszW, 2 * lenS
      End If
   End If
End Function

' ================================================================
' interface procs

' Returns a reference to the IMalloc interface. (replaced by SHAlloc, SHFree )

'Public Function isMalloc() As IMalloc
'  Static im As IMalloc
'  If (im Is Nothing) Then Call SUCCEEDED(SHGetMalloc(im))
'  Set isMalloc = im
'End Function

' Returns a reference to the desktop folder's IShellFolder interface.

Public Function isfDesktop() As IShellFolder
   Static isf As IShellFolder
   If (isf Is Nothing) Then Call SUCCEEDED(SHGetDesktopFolder(isf))
   Set isfDesktop = isf
End Function

' Returns the IShellFolder interface ID, {000214E6-0000-0000-C000-000000046}

Private Function IID_IShellFolder() As IShellFolderEx_TLB.GUID
   Static iid As IShellFolderEx_TLB.GUID
   If (iid.Data1 = 0) Then iid = GetRIID(rIID_IShellFolder)
   IID_IShellFolder = iid
End Function


' ================================================================
' pidl utility procs

' Determines if the specified pidl is the desktop folder's pidl.
' Returns True if the pidl is the desktop's pidl, returns False otherwise.

' The desktop pidl is only a single item ID whose value is 0 (the 2 byte
' zero-terminator, i.e. SHITEMID.abID is empty). Direct descendents of
' the desktop (My Computer, Network Neighborhood) are absolute pidls
' (relative to the desktop) also with a single item ID, but contain values
' (SHITEMID.abID > 0). Drive folders have 2 item IDs, children of drive
' folders have 3 item IDs, etc. All other single item ID pidls are relative to
' the shell folder in which they reside (just like a relative path).

Public Function IsDesktopPIDL(pidl As Long) As Boolean
  
   ' The GetItemIDSize() call will also return 0 if pidl = 0
   If pidl Then IsDesktopPIDL = (GetItemIDSize(pidl) = 0)

End Function

' Returns the size in bytes of the first item ID in a pidl.
' Returns 0 if the pidl is the desktop's pidl or is the last
' item ID in the pidl (the zero terminator), or is invalid.

Public Function GetItemIDSize(ByVal pidl As Long) As Integer
  
   ' If we try to access memory at address 0 (NULL), then it's bye-bye...
   If pidl Then MoveMemory GetItemIDSize, ByVal pidl, 2

End Function

' Returns the count of item IDs in a pidl.

Public Function GetItemIDCount(ByVal pidl As Long) As Integer
   Dim nItems As Integer
   
   ' If the size of an item ID is 0, then it's the zero
   ' value terminating item ID at the end of the pidl.
   Do While GetItemIDSize(pidl)
      pidl = ILGetNext(pidl)
      nItems = nItems + 1
   Loop
   
   GetItemIDCount = nItems
   
End Function

' Returns a pointer to the next item ID in a pidl.
' Returns 0 if the next item ID is the pidl's zero value terminating 2 bytes.

Public Function GetNextItemID(ByVal pidl As Long) As Long
   GetNextItemID = ILGetNext(pidl)
End Function

' If successful, returns the size in bytes of the memory occcupied by a pidl,
' including it's 2 byte zero terminator. Returns 0 otherwise.

Public Function GetPIDLSize(ByVal pidl As Long) As Integer
   GetPIDLSize = ILGetSize(pidl)
End Function

' Copies and returns the specified item ID from a complex pidl
'   pidl -  pointer to an item ID list from which to copy
'   nItem - 1-based position in the pidl of the item ID to copy   / GIID_FIRST or GIID_LAST

' If successful, returns a new item ID (single pidl)from the specified element positon.
' Returns 0 on failure.
' If nItem exceeds the number of item IDs in the pidl, the last item ID is returned.

Public Function GetItemID(ByVal pidl As Long, ByVal nItem As Integer) As Long
   Dim nCount  As Integer
   Dim i       As Integer
   Dim cb      As Integer
   Dim pidlNew As Long
  
   Select Case nItem
   
      Case GIID_FIRST
         GetItemID = ILCloneFirst(pidl)
   
      Case GIID_LAST
         GetItemID = ILFindLastID(pidl)
         GetItemID = ILClone(GetItemID)
      
      Case Else
      
         nCount = GetItemIDCount(pidl)
         
         If (nItem >= nCount) Then
            GetItemID = GetItemID(pidl, GIID_LAST)
            Exit Function
         End If
         
         ' GetNextItemID returns the 2nd item ID
         For i = 1 To nItem - 1: pidl = GetNextItemID(pidl): Next
           
         ' Get the size of the specified item identifier.
         ' If cb = 0 (the zero terminator), then we'll return a desktop pidl, proceed
         cb = GetItemIDSize(pidl)
         
         ' Allocate a new item identifier list.
         pidlNew = SHAlloc(cb + 2)
         If pidlNew Then
           
           ' Copy the specified item identifier.
           ' and append the zero terminator.
           MoveMemory ByVal pidlNew, ByVal pidl, cb
           MoveMemory ByVal pidlNew + cb, 0, 2
           
           GetItemID = pidlNew
         End If
   End Select

End Function

' Creates a new pidl of the given size

Public Function CreatePIDL(cb As Long) As Long
   Dim pidl As Long
   
   pidl = SHAlloc(cb)
   If pidl Then
      FillMemory ByVal pidl, cb, 0 ' initialize to zero, set by caller
      CreatePIDL = pidl
   End If
End Function

' Returns a copy of a relative or absolute pidl

Public Function CopyPIDL(pidl As Long) As Long
   CopyPIDL = ILClone(pidl)
End Function

' Frees the specified pidl and zeros it.(Passing zero is OK).
' §SHFree ?
Public Sub FreePIDL(pidl As Long)
   ILFree pidl
   pidl = 0
End Sub

' Copies and returns all but the last item ID from the specified absolute pidl.

'   pidl          - pointer to the pidl from which to copy
'   fFreeOldPidl  - optional flag specifying whether to free and zero the passed pidl

' If successful, returns a new absolute pid (relative to the desktop)
' If either a valid single item ID pidl is passed to this proc (either the
' desktop's pidl or a relative pidl), or an invalid pidl is passed, the
' desktop's pidl is returned.

Public Function GetPIDLParent(pidl As Long, _
                              Optional fRtnDesktop As Boolean = False, _
                              Optional fFreeOldPidl As Boolean = False) As Long
   Dim pidlNew  As Long
   Dim pidlLast As Long

   pidlNew = ILClone(pidl)
   pidlLast = ILFindLastID(pidlNew)

   If pidlLast <> pidlNew Then
      If ILRemoveLastID(pidlNew) Then

         If fFreeOldPidl Then Call FreePIDL(pidl)
         GetPIDLParent = pidlNew
      End If
   Else
      FreePIDL pidlNew
   End If

   If (pidlNew = 0) And fRtnDesktop Then
      If fFreeOldPidl Then Call FreePIDL(pidl)
      GetPIDLParent = SHCloneSpecialIDList(0&, 0&, False)
   End If

End Function

'Public Function GetPIDLParent(pidl As Long, _
'                              Optional fRtnDesktop As Boolean = False, _
'                              Optional fFreeOldPidl As Boolean = False) As Long
'   Dim pidlNew  As Long
'   Dim cb       As Long
'
'   ' size of parent without terminator
'   cb = ILFindLastID(pidl) - pidl
'
'   If cb > 0 Then
'
'      pidlNew = CreatePIDL(cb + 2)
'      If pidlNew Then
'
'         MoveMemory ByVal pidlNew, ByVal pidl, cb
'
'         If fFreeOldPidl Then Call FreePIDL(pidl)
'         GetPIDLParent = pidlNew
'      End If
'   End If
'
'   If (pidlNew = 0) And fRtnDesktop Then
'      If fFreeOldPidl Then Call FreePIDL(pidl)
'      GetPIDLParent = SHCloneSpecialIDList(0&, 0&, False)
'   End If
'
'End Function

' Creates a new pidl by prepending pidl2 to pidl1 (i.e pidlNew = pidl1pidl2)
' If pidl1 or pidl2 is zero, returns a clone of the non-zero pidl.
' The two passed pidls are still valid and are not freed unless specified.

Public Function CombinePIDLs(pidl1 As Long, pidl2 As Long, _
                             Optional fFreePidl1 As Boolean = False, _
                             Optional fFreePidl2 As Boolean = False) As Long
   
   CombinePIDLs = ILCombine(pidl1, pidl2)
    
   If fFreePidl1 Then FreePIDL pidl1
   If fFreePidl2 Then FreePIDL pidl2
End Function

' Returns True if pidlFQParent is a parent of pidlFQBelow.
' bImmediate = True : only if it is the immediate parent.
' pidlFQParent == pidlFQBelow : Returns True if bImmediate = False.
Public Function IsParentPIDL(pidlFQParent As Long, pidlFQBelow As Long, _
                             Optional ByVal bImmediate As Boolean) As Boolean
  
   IsParentPIDL = ILIsParent(pidlFQParent, pidlFQBelow, bImmediate)
End Function

' ================================================================
' IShellFolder procs

' Returns a shell item's displayname

'   isfParent  - item's parent folder IShellFolder
'   pidlRel    - item's pidl, relative to isfParent. Simple pidl ! Exceptions see MSDN
'   uFlags     - specifies the type of name to retrieve

Public Function GetFolderDisplayName(isfParent As IShellFolder, pidlRel As Long, _
                                     uFlags As ESHGNO) As String
   Dim lpStr   As STRRET   ' struct filled
   Dim lpszW   As Long     ' string pointer, allocated by StrRetToStrW()
   Dim lenS    As Long
   
   If SUCCEEDED(isfParent.GetDisplayNameOf(pidlRel, uFlags, lpStr)) Then
   
      If SUCCEEDED(StrRetToStrW(lpStr, pidlRel, lpszW)) Then
      
         lenS = lstrlenW(ByVal lpszW)
         GetFolderDisplayName = Space$(lenS)
         
         MoveMemory ByVal StrPtr(GetFolderDisplayName), ByVal lpszW, 2 * lenS
      
         Call CoTaskMemFree(lpszW)
      End If
   End If

End Function

' Returns the IShellFolder for the specified relative pidl

'   isfParent - pidl's parent folder IShellFolder
'   pidlRel   - item's relative pidl we're returning the IShellFolder of.

' If an error occurs, the desktop's IShellFolder is returned.

Public Function GetIShellFolder(isfParent As IShellFolder, pidlRel As Long, _
                                Optional fRtnDesktop As Boolean = True) As IShellFolder
   Dim isf As IShellFolder
   On Error GoTo Out
   
   Call isfParent.BindToObject(pidlRel, 0, IID_IShellFolder, isf)

Out:
   If (Err = 0) And Not (isf Is Nothing) Then
      Set GetIShellFolder = isf
   Else
      If fRtnDesktop Then
         Set GetIShellFolder = isfDesktop
      Else
         Debug.Print "GetIShellFolder FAILED", dbgWalkPIDL(pidlRel)
      End If
   End If
End Function

' Returns a reference to the parent IShellFolder of the last item ID in the specified
' fully qualified pidl.

'  pidlRel - see MSDN: SHBindToParent

' If pidlFQ is zero, or a relative (single item) pidl, then the desktop's IShellFolder
' is returned. If an unexpected error occurs, the object value Nothing is returned.

Public Function GetIShellFolderParent(ByVal pidlFQ As Long, _
                                      Optional fRtnDesktop As Boolean = True, _
                                      Optional pidlRel As Long) As IShellFolder
   Dim isf   As IShellFolder
   
   If SUCCEEDED(SHBindToParent(pidlFQ, IID_IShellFolder, isf, pidlRel)) Then
      Set GetIShellFolderParent = isf
   ElseIf fRtnDesktop Then
      Set GetIShellFolderParent = isfDesktop
   End If
   
End Function

'Public Function GetFreshPIDL(ByVal pidlFQ As Long, _
'                             Optional fFreePidl As Boolean = False) As Long
'   Dim pchEaten      As Long
'   Dim pidlNew       As Long
'   Dim sPath         As String
'
'   If pidlFQ Then
'      sPath = GetFolderDisplayName(isfDesktop, pidlFQ, SHGDN_FORPARSING)
'      If LenB(sPath) Then
'
'         If (S_OK = isfDesktop.ParseDisplayName(0, 0, StrConv(sPath, vbUnicode), _
'                                                pchEaten, pidlNew, 0)) Then
'            GetFreshPIDL = pidlNew
'         Else
'            Debug.Print sPath
'            Debug.Assert False
'         End If
'
'         If fFreePidl Then FreePIDL pidlFQ
'      Else
'         Debug.Assert False
'      End If
'   End If
'End Function

#End If

