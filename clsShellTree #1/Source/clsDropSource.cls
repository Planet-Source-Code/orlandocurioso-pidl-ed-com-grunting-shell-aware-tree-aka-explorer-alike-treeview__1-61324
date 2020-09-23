VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDropSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsDropSource
' Author    : OrlandoCurioso 19.06.2005
' Credits   : J.P.Hamilton (VTable swapping)              http://www.oreilly.com/catalog/vbshell/chapter/ch04.html
' Purpose   : IDropSource implementation needed for calling DoDragDrop API.
' Requires  : IShellFolder Extended Type Library v1.3 OC (ISHF_Ex.tlb)
'---------------------------------------------------------------------------------------
Option Explicit

Private m_pOldGiveFeedback       As Long
Private m_pOldQueryContinueDrag  As Long

Implements IDropSource

Private Sub IDropSource_QueryContinueDrag(ByVal fEscapePressed As Long, ByVal grfKeyState As IShellFolderEx_TLB.KEYSTATES)
   ' Swapped Entry 4
End Sub

Private Sub IDropSource_GiveFeedback(ByVal dwEffect As Long)
   ' Swapped Entry 5
End Sub

Private Sub Class_Initialize()

   Dim pVtable As IDropSource
   Set pVtable = Me
   
   m_pOldQueryContinueDrag = SwapVtableEntry(ObjPtr(pVtable), 4, AddressOf QueryContinueDragVB)
   m_pOldGiveFeedback = SwapVtableEntry(ObjPtr(pVtable), 5, AddressOf GiveFeedbackVB)
   
   Dim pUnk As IShellFolderEx_TLB.IUnknown
   Set pUnk = Me
   pUnk.AddRef
    
End Sub

Private Sub Class_Terminate()

   Dim pVtable As IDropSource
   Set pVtable = Me
   
   m_pOldQueryContinueDrag = SwapVtableEntry(ObjPtr(pVtable), 4, m_pOldQueryContinueDrag)
   m_pOldGiveFeedback = SwapVtableEntry(ObjPtr(pVtable), 5, m_pOldGiveFeedback)
    
End Sub


