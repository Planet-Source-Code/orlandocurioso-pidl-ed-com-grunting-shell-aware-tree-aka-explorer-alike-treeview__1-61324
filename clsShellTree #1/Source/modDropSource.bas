Attribute VB_Name = "modDropSource"
'---------------------------------------------------------------------------------------
' Module    : modDropSource
' Author    : OrlandoCurioso 19.06.2005
' Purpose   : Swapped functions return HRESULT.
'
'---------------------------------------------------------------------------------------
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
   Private Const PAGE_EXECUTE_READWRITE = &H40&
   
Private Const S_OK = 0&
Private Const DRAGDROP_S_DROP = &H40100
Private Const DRAGDROP_S_CANCEL = &H40101
Private Const DRAGDROP_S_USEDEFAULTCURSORS = &H40102
'

Public Function GiveFeedbackVB(ByVal this As IDropSource, ByVal dwEffect As Long) As Long

   GiveFeedbackVB = DRAGDROP_S_USEDEFAULTCURSORS

   If (dwEffect = vbDropEffectMove) Then
      Screen.MousePointer = vbDefault
      GiveFeedbackVB = S_OK
'   ElseIf (dwEffect And (vbDropEffectMove Or vbDropEffectCopy)) Then
'      ' # show only copy pointer #
'      Screen.MousePointer = ###
'      GiveFeedbackVB = S_OK
   End If

End Function

Public Function QueryContinueDragVB(ByVal this As IDropSource, ByVal fEscapePressed As Long, ByVal grfKeyState As IShellFolderEx_TLB.KEYSTATES) As Long

   If fEscapePressed Then
      QueryContinueDragVB = DRAGDROP_S_CANCEL
   ElseIf grfKeyState = S_OK Then
      QueryContinueDragVB = DRAGDROP_S_DROP
 ' Else: QueryContinueDragVB = S_OK
   End If
End Function

Public Function SwapVtableEntry(pObj As Long, EntryNumber As Integer, ByVal lpfn As Long) As Long

   Dim lOldAddr As Long
   Dim lpVtableHead As Long
   Dim lpfnAddr As Long
   Dim lOldProtect As Long

   CopyMemory lpVtableHead, ByVal pObj, 4
   lpfnAddr = lpVtableHead + (EntryNumber - 1) * 4
   CopyMemory lOldAddr, ByVal lpfnAddr, 4

   Call VirtualProtect(lpfnAddr, 4, PAGE_EXECUTE_READWRITE, lOldProtect)
   CopyMemory ByVal lpfnAddr, lpfn, 4
   Call VirtualProtect(lpfnAddr, 4, lOldProtect, lOldProtect)

   SwapVtableEntry = lOldAddr

End Function
