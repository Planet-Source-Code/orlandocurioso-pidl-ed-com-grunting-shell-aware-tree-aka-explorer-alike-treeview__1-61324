Attribute VB_Name = "modSortTreeCB"
'---------------------------------------------------------------------------------------
' Module    : modSortTreeCB
' Author    : OrlandoCurioso 14.04.2005 / Brad Martinez
' Purpose   : ucTreeView/clsShellTree Callback function for TVM_SORTCHILDRENCB.
'
'---------------------------------------------------------------------------------------
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private m_cFolderTree   As Long
'

Property Let FolderTree(ByRef cThis As ucTreeView)
   m_cFolderTree = ObjPtr(cThis)
End Property
Property Get FolderTree() As ucTreeView
   Dim oT As ucTreeView
   If (m_cFolderTree <> 0) Then
      CopyMemory oT, m_cFolderTree, 4
      Set FolderTree = oT
      CopyMemory oT, 0&, 4
   End If
End Property

' Application-defined callback function, which is called during a sort operation each time
' the relative order of two treeview items needs to be compared.
' (see the TVSORTCB struct's lpfnCompare member desciption in the SDK)

' The lParam1 and lParam2 parameters correspond to the lParam member of the TVITEM
' structure for the two items being compared.

'    lParam1     - pointer to the 1st item's TVITEMDATA struct
'    lParam2     - pointer to the 2nd item's TVITEMDATA struct
'    lParamSort  - corresponds to the lParam member of the TVSORTCB structure
'                  that was passed with the TVM_SORTCHILDRENCB message.

' The callback function must return a negative value if the first item should precede the second,
' a positive value if the first item should follow the second, or zero if the two items are equivalent.

' Invoked by TreeView_SortChildrenCB calls in clsShellTree.pInsertFolderItems()
' & pUpdateNodePIDL().

Public Function TreeViewCompareProc(ByVal lParam1 As Long, ByVal lParam2 As Long, _
                                    ByVal lParamSort As Long) As Long
   Dim isfParent  As IShellFolder
   Dim pidlRel1   As Long
   Dim pidlRel2   As Long
   Dim hr         As Long        ' HRESULT
   Dim oClient    As ucTreeView
   
   Set oClient = FolderTree
   
   If Not (oClient Is Nothing) Then

      pDecryptNodeKey oClient.NodeKey(, lParam1), , pidlRel1
      pDecryptNodeKey oClient.NodeKey(, lParam2), , pidlRel2
   
      ' Get the parent folder's un-AddRef'd IShellFolder
      ' from lParamSort that we set in pInsertFolderItems.
      CopyMemory isfParent, lParamSort, 4&
   
      hr = isfParent.CompareIDs(0&, pidlRel1, pidlRel2)
      
      If (hr >= NOERROR) Then
      
         TreeViewCompareProc = LOWORD(hr)
         
'         Debug.Print GetFolderDisplayName(isfParent, pidlRel1, SHGDN_INFOLDER), _
'                     GetFolderDisplayName(isfParent, pidlRel2, SHGDN_INFOLDER), _
'                     TreeViewCompareProc
      End If   ' (hr >= NOERROR)
      
      ' Zero the IShellfolder object variable so it is not Released.
      CopyMemory isfParent, 0&, 4
   End If
End Function

' keep identical to same proc in clsShellTree
Private Sub pDecryptNodeKey(sKey As String, Optional pidlFQ As Long, Optional pidlRel As Long)
   Dim saKey() As String
   saKey = Split(sKey, ":")
   
   pidlFQ = CLng(saKey(0))
   pidlRel = CLng(saKey(1))
End Sub
