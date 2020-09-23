Attribute VB_Name = "modSHGeneral"
'---------------------------------------------------------------------------------------
' Module    : modSHGeneral
' Author    : OrlandoCurioso 14.05.2005
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

Public Const S_OK = 0            ' indicates success
Public Const S_FALSE = 1&        ' special HRESULT value

' Defined as an HRESULT that corresponds to S_OK.
Public Const NOERROR = 0

Public Const MAX_PATH = 260&

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
' dwFlags
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
' dwLanguageId
Private Const LANG_USER_DEFAULT = &H400&

#If WIN32_IE < &H500 Then
Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Private Declare Function lstrcpyW Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
#End If

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

' //

Public Enum eRIID
   rIID_IShellFolder
   rIID_IShellFolder2
   rIID_IShellView
   rIID_IShellView2
   rIID_IShellDetails
   rIID_IContextMenu
   rIID_IContextMenu2
   rIID_IContextMenu3
   rIID_IDataObject
   rIID_IDropTarget
   rIID_IDropSource
   rCLSID_DragDropHelper      ' Class identifier of drag-image manager object
   rIID_IDropTargetHelper
   rIID_IDragSourceHelper
   rIID_IShellIcon
   rIID_IShellIconOverlay
   rIID_IExtractIconA
   rIID_IExtractIconW
End Enum

' // SHGetFileInfo

Private Type SHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Public Enum SHGFI_flags
   SHGFI_LARGEICON = &H0            ' sfi.hIcon is large icon
   SHGFI_SMALLICON = &H1            ' sfi.hIcon is small icon
   SHGFI_OPENICON = &H2             ' sfi.hIcon is open icon
   SHGFI_SHELLICONSIZE = &H4        ' sfi.hIcon is shell size (not system size), rtns BOOL
   SHGFI_PIDL = &H8                 ' pszPath is pidl, rtns BOOL
   SHGFI_USEFILEATTRIBUTES = &H10   ' pretend pszPath exists, rtns BOOL
   SHGFI_ICON = &H100               ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
   SHGFI_DISPLAYNAME = &H200        ' isf.szDisplayName is filled (SHGDN_NORMAL), rtns BOOL
   SHGFI_TYPENAME = &H400           ' isf.szTypeName is filled, rtns BOOL
   SHGFI_ATTRIBUTES = &H800         ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
   SHGFI_ICONLOCATION = &H1000      ' fills sfi.szDisplayName with filename
                                                         ' containing the icon, rtns BOOL
   SHGFI_EXETYPE = &H2000           ' rtns two ASCII chars of exe type
   SHGFI_SYSICONINDEX = &H4000      ' sfi.iIcon is sys il icon index, rtns hImagelist
   SHGFI_LINKOVERLAY = &H8000       ' add shortcut overlay to sfi.hIcon
   SHGFI_SELECTED = &H10000         ' sfi.hIcon is selected icon
   SHGFI_ATTR_SPECIFIED = &H20000   ' get only attributes specified in sfi.dwAttributes
End Enum
'

' ==============================================================
' SHGetFileInfo calls

' If successful returns the specified file's typename, returns an empty string otherwise.
'   pidl  - file's absolute pidl

'Public Function GetFileTypeNamePIDL(pidl As Long) As String
'   Dim sfi As SHFILEINFO
'   If SHGetFileInfo(pidl, 0, sfi, Len(sfi), SHGFI_PIDL Or SHGFI_TYPENAME) Then
'      On Error Resume Next
'      GetFileTypeNamePIDL = Left$(sfi.szTypeName, InStr(sfi.szTypeName, vbNullChar) - 1)
'   End If
'End Function

#If WIN32_IE >= &H500 Then
' use SHMapPIDLToSystemImageListIndex API
#Else

' Returns a file's small or large icon index within the system imagelist.
'   pidl   - file's absolute pidl
'   uType  - either SHGFI_SMALLICON or SHGFI_LARGEICON, and SHGFI_OPENICON

Public Function GetFileIconIndexPIDL(pidl As Long, uType As SHGFI_flags) As Long
   Dim sfi As SHFILEINFO
   If SHGetFileInfo(pidl, 0, sfi, Len(sfi), SHGFI_PIDL Or SHGFI_SYSICONINDEX Or uType) Then
      GetFileIconIndexPIDL = sfi.iIcon
   End If
End Function

#End If

Public Function GetDefaultFileIconIndex(sExtension As String, eAttr As VbFileAttribute, uType As SHGFI_flags) As Long
   Dim sfi As SHFILEINFO
   If SHGetFileInfo(sExtension, eAttr, sfi, Len(sfi), SHGFI_USEFILEATTRIBUTES Or SHGFI_SYSICONINDEX Or uType) Then
      GetDefaultFileIconIndex = sfi.iIcon
   End If
End Function

' Returns the handle of the small or large icon system imagelist.
'   uSize - either SHGFI_SMALLICON or SHGFI_LARGEICON

Public Function GetSystemImagelist(uSize As SHGFI_flags) As Long
   Dim sfi As SHFILEINFO
   
   GetSystemImagelist = SHGetFileInfo(".txt", vbNormal, sfi, Len(sfi), SHGFI_USEFILEATTRIBUTES Or SHGFI_SYSICONINDEX Or uSize)
End Function

' ==============================================================

Public Function GetRIID(ByVal eIID As eRIID) As IShellFolderEx_TLB.GUID
   
   Select Case eIID
   
      Case rIID_IShellFolder     ' {000214E6-0000-0000-C000-000000000046}
         DEFINE_OLEGUID GetRIID, &H214E6, 0, 0
   
      Case rIID_IShellFolder2    ' {93F2F68C-1D1B-11d3-A30E-00C04F79ABD1}
         DEFINE_GUID GetRIID, &H93F2F68C, &H1D1B, &H11D3, &HA3, &HE, &H0, &HC0, &H4F, &H79, &HAB, &HD1
   
      Case rIID_IShellDetails    ' {000214EC-0000-0000-C000-000000000046}
         DEFINE_OLEGUID GetRIID, &H214EC, 0, 0
         
      Case rIID_IShellView       ' {000214E3-0000-0000-C000-000000000046}
         DEFINE_OLEGUID GetRIID, &H214E3, 0, 0
   
      Case rIID_IShellView2      ' {88E39E80-3578-11CF-AE69-08002B2E1262}
         DEFINE_GUID GetRIID, &H88E39E80, &H3578, &H11CF, &HAE, &H69, &H8, &H0, &H2B, &H2E, &H12, &H62
   
      Case rIID_IContextMenu     ' {000214E4-0000-0000-C000-000000000046}
         DEFINE_OLEGUID GetRIID, &H214E4, 0, 0
   
      Case rIID_IContextMenu2    ' {000214F4-0000-0000-C000-000000000046}
         DEFINE_OLEGUID GetRIID, &H214F4, 0, 0
         
      Case rIID_IContextMenu3    ' {BCFCE0A0-EC17-11D0-8D10-00A0C90F2719}
         DEFINE_GUID GetRIID, &HBCFCE0A0, &HEC17, &H11D0, &H8D, &H10, &H0, &HA0, &HC9, &HF, &H27, &H19
   
      Case rIID_IDataObject      ' {0000010e-0000-0000-C000-000000000046}
         DEFINE_OLEGUID GetRIID, &H10E, 0, 0
         
      Case rIID_IDropSource      ' {00000121-0000-0000-C000-000000000046}
         DEFINE_OLEGUID GetRIID, &H121, 0, 0
   
      Case rIID_IDropTarget      ' {00000122-0000-0000-C000-000000000046}
         DEFINE_OLEGUID GetRIID, &H122, 0, 0
   
      Case rCLSID_DragDropHelper  ' {4657278A-411B-11D2-839A-00C04FD918D0}
         DEFINE_GUID GetRIID, &H4657278A, &H411B, &H11D2, &H83, &H9A, &H0, &HC0, &H4F, &HD9, &H18, &HD0
   
      Case rIID_IDropTargetHelper ' {4657278B-411B-11D2-839A-00C04FD918D0}
         DEFINE_GUID GetRIID, &H4657278B, &H411B, &H11D2, &H83, &H9A, &H0, &HC0, &H4F, &HD9, &H18, &HD0
   
      Case rIID_IDragSourceHelper ' {DE5BF786-477A-11D2-839D-00C04FD918D0}
         DEFINE_GUID GetRIID, &HDE5BF786, &H477A, &H11D2, &H83, &H9D, &H0, &HC0, &H4F, &HD9, &H18, &HD0
      
'      Case rIID_IShellIcon       ' {000214E5-0000-0000-C000-000000000046}
'         DEFINE_OLEGUID GetRIID, &H214E5, 0, 0
'
'      Case rIID_IShellIconOverlay ' {7D688A70-C613-11D0-999B-00C04FD655E1}
'         DEFINE_GUID GetRIID, &H7D688A70, &HC613, &H11D0, &H99, &H9B, &H0, &HC0, &H4F, &HD6, &H55, &HE1
'
'      Case rIID_IExtractIconA    ' {000214EB-0000-0000-C000-000000000046}
'         DEFINE_OLEGUID GetRIID, &H214EB, 0, 0
'
'      Case rIID_IExtractIconW    ' {000214FA-0000-0000-C000-000000000046}
'         DEFINE_OLEGUID GetRIID, &H214FA, 0, 0
      
      Case Else:  Debug.Assert False
   End Select
   
End Function

' Fills a GUID

Private Sub DEFINE_GUID(Name As IShellFolderEx_TLB.GUID, _
                        l As Long, w1 As Integer, w2 As Integer, _
                        b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte, _
                        b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
   With Name
      .Data1 = l
      .Data2 = w1
      .Data3 = w2
      .Data4(0) = b0
      .Data4(1) = b1
      .Data4(2) = b2
      .Data4(3) = b3
      .Data4(4) = b4
      .Data4(5) = b5
      .Data4(6) = b6
      .Data4(7) = b7
   End With
End Sub

' Fills an OLE GUID, the Data4 member always is "C000-000000046"

Private Sub DEFINE_OLEGUID(Name As IShellFolderEx_TLB.GUID, l As Long, w1 As Integer, w2 As Integer)
   DEFINE_GUID Name, l, w1, w2, &HC0, 0, 0, 0, 0, 0, 0, &H46
End Sub

' Provides a generic test for success on any status value. (hr = HRESULT)
' Non-negative numbers indicate success.

Public Function SUCCEEDED(hr As Long, _
                          Optional ByRef error As Long, _
                          Optional ByRef ErrStr As String) As Boolean
   Dim DBP  As Integer
   
   error = hr
   
   If (hr >= S_OK) Then
      SUCCEEDED = True
   Else
      ErrStr = "Error: &H" & Hex$(hr) & ", " & GetAPIErrStr(hr)
      
      Select Case hr
         Case &H800704C7   ' user cancelled
            Exit Function
         Case &H80070002   ' can't find file
            Debug.Assert False
         Case &H80004002   ' interface not supported
'            Debug.Assert False
         Case &H80004005   ' unknown error ie empty drive
'            Debug.Assert False
         Case &H80070057   ' false parameter
'            Debug.Assert False
             DBP = 1
         Case Else
             DBP = 2
      End Select
      
      If InIDE Then
         Debug.Assert DbgPrt(DBP >= 1, ErrStr)
      Else
'         If Not frmTest.ucTree.OLEGetDropInfo Then  ' # CAVEAT #
            MsgBox ErrStr, vbExclamation, App.Title & " Succeeded()"
'         End If
      End If
   End If
End Function

' Returns the system-defined description of an API error code

Public Function GetAPIErrStr(dwErrCode As Long) As String
   Dim sErrDesc As String * 256   ' max string resource len
   Dim lenS  As Long
   lenS = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS Or _
                        FORMAT_MESSAGE_MAX_WIDTH_MASK, ByVal 0&, dwErrCode, _
                        LANG_USER_DEFAULT, ByVal sErrDesc, 256, 0)
   GetAPIErrStr = Left$(sErrDesc, lenS)
End Function

' Returns the low 16-bit integer from a 32-bit long integer

Public Function LOWORD(dwValue As Long) As Integer
  MoveMemory LOWORD, dwValue, 2
End Function

' Returns the high 16-bit integer from a 32-bit long integer

Public Function HIWORD(dwValue As Long) As Integer
  MoveMemory HIWORD, ByVal VarPtr(dwValue) + 2, 2
End Function


#If WIN32_IE < &H500 Then

' Returns the string before first null char encountered (if any) from an ANSII string.

Public Function GetStrFromBufferA(sz As String) As String
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would return a zero length string ("").
    GetStrFromBufferA = sz
  End If
End Function

' Returns an ANSII string from a pointer to an ANSII string.

Public Function GetStrFromPtrA(lpszA As Long) As String
  Dim sRtn As String
  sRtn = String$(lstrlenA(ByVal lpszA), 0)
  Call lstrcpyA(ByVal sRtn, ByVal lpszA)
  GetStrFromPtrA = sRtn
End Function

' Returns an ANSII string from a pointer to a Unicode string.

Public Function GetStrFromPtrW(lpszW As Long) As String
  Dim sRtn As String
  sRtn = String$(lstrlenW(ByVal lpszW) * 2, 0)   ' 2 bytes/char
  Call lstrcpyW(ByVal sRtn, ByVal lpszW)
  GetStrFromPtrW = StrConv(sRtn, vbFromUnicode)
End Function

#End If

