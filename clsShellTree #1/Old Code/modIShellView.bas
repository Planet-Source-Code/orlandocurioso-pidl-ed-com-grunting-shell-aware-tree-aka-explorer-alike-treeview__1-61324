Attribute VB_Name = "modIShellView"
'---------------------------------------------------------------------------------------
' Module    : modIShellView
' Author    : OrlandoCurioso 13.05.2005
' Purpose   :
'
'---------------------------------------------------------------------------------------
Option Explicit

Private Type RECT
   Left     As Long
   Top      As Long
   Right    As Long
   Bottom   As Long
End Type

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'

' Returns the IShellView interface ID, {000214E3-0000-0000-C000-000000000046}

'Public Function IID_IShellView() As IShellFolderEx_TLB.GUID
'  Static iid As IShellFolderEx_TLB.GUID
'  If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214E3, 0, 0)
'  IID_IShellView = iid
'End Function


' Q157247 PRB: IShellFolder::CreateViewObject() Causes Access Violation
' catch undocumented WM_GETISHELLBROWSER (== WM_USER + 7) for hOwner ?
' send WM_GETISHELLBROWSER to explorer ?
' TLB: define ISHELLBROWSER
' SHCreateShellFolderView
Public Function CreateView(ByVal hOwner As Long, isfParent As IShellFolder) As Long
   Dim isv     As IShellView
   Dim isvPrev As IShellView
   Dim hr      As Long
   Dim tFS     As FOLDERSETTINGS
   Dim eFF     As FOLDERFLAGS
   Dim eFV     As FOLDERVIEWMODE
   Dim hWnd    As Long
   Dim tR      As RECT
   
   
   If SUCCEEDED(isfParent.CreateViewObject(hOwner, GetRIID(rIID_IShellView), isv)) Then
      
      eFF = FWF_BESTFITWINDOW
      eFV = FVM_DETAILS
      
      With tFS
         .fFlags = eFF
         .ViewMode = eFV
      End With
      
      GetClientRect hOwner, tR
         
      ' GPF with missing ISHELLBROWSER
'      If SUCCEEDED(isv.CreateViewWindow(isvPrev, tFS, 0, VarPtr(tR), hWnd)) Then
'         CreateView = hWnd
'      Else
'         Debug.Assert False
'      End If
      
   Else
      Debug.Assert False
   End If


End Function
