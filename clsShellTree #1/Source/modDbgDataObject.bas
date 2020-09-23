Attribute VB_Name = "modDbgDataObject"
'---------------------------------------------------------------------------------------
' Module    : modDbgDataObject
' Author    : OrlandoCurioso 20.06.2005
' Purpose   :
'
'---------------------------------------------------------------------------------------
Option Explicit

Private Type POINTAPI
   X  As Long
   Y  As Long
End Type

Private Type DROPFILES
   pFiles  As Long
   pt      As POINTAPI
   fNC     As Long
   fWide   As Long
End Type

Private Const CFSTR_SHELLIDLIST = "Shell IDList Array"

'Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare Function RegisterClipboardFormatInt Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Integer
Private Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function ReleaseStgMedium Lib "ole32.dll" (pMedium As STGMEDIUM) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
'

Public Function dbgGetClipboardFormatName(ByVal wFormat As Long) As String
   Dim sBuf As String
   Dim lR   As Long
   
   Select Case LOWORD(wFormat)
   
      Case vbCFFiles
         dbgGetClipboardFormatName = "CF_HDROP"
      Case 1 To 17
         dbgGetClipboardFormatName = "PREDEFINED"
         
      Case Else
         ' custom formats
         sBuf = String$(MAX_PATH, vbNullChar)
         lR = GetClipboardFormatName(wFormat, sBuf, MAX_PATH)
         If (lR <> 0) Then
            dbgGetClipboardFormatName = Left$(sBuf, lR)
         End If
         
   End Select

End Function

' CRASH: for IDataObject of other sources !!!
' Extracts the names of formats from COM IDataObject.
' Extracts the data of formats CFSTR_SHELLIDLIST & CF_HDROP.
Public Function dbgIDataObject(ido As IDataObject) As Boolean
   Dim ief           As IEnumFORMATETC
   Dim tSTGM         As STGMEDIUM
   Dim tFOTC         As FORMATETC
   Dim b()           As Byte
   Dim wFormat       As Long
   Dim hr            As Long
   Dim dvRet         As DV_ERROR
   Dim hGlobal       As Long
   Dim cb            As Long
   Dim DBP           As Integer
   DBP = 2
   
   On Error GoTo Proc_Error

   Debug.Assert DbgPrt(DBP >= 1, vbCrLf & "dbgIDataObject" & String$(60, "-"))

   If Not (ido Is Nothing) Then
   
      ' get IEnumFORMATETC reference
      If SUCCEEDED(ido.EnumFormatEtc(DATADIR_GET, ief)) Then
      
         ' enumerate all formats
         hr = ief.Next(1, tFOTC, 0)
         
         Do While hr = S_OK

            ' get the data tSTGM for the format specified in tFOTC
            dvRet = ido.GetData(tFOTC, tSTGM)

            If dvRet = S_OK Then

               wFormat = LOWORD(tFOTC.cfFormat)

               Debug.Assert DbgPrt(DBP >= 1, dbgGetClipboardFormatName(tFOTC.cfFormat), wFormat)

               If tSTGM.TYMED = TYMED_HGLOBAL Then

                  ' ptr to allocated global memory (maintained by IDataobject)
                  hGlobal = tSTGM.pData

                  If hGlobal <> 0 Then

                     ' copy data to bytearray
                     cb = GlobalSize(hGlobal)
                     ReDim b(0 To cb - 1)

                     CopyMemory b(0), ByVal hGlobal, cb  ' ### CRASH ###

                     If wFormat = RegisterClipboardFormatInt(CFSTR_SHELLIDLIST) Then

                        Debug.Print dbgDumpMEM(VarPtr(b(0)), cb)

                        ' CFSTR_SHELLIDLIST
                        dbgReadIDList b()

                     ElseIf wFormat = vbCFFiles Then

                        ' CF_HDROP :
                        ' - supports ANSI & Unicode, can be detected by DROPFILES struct
                        Dim tDF        As DROPFILES
                        Dim saFiles()  As String
                        Dim idx        As Long

                        CopyMemory tDF, b(0), Len(tDF)

                        With tDF
                           '.pFiles : Offset of file list from the beginning of tDF in bytes.
                           ReDim b(0 To cb - 1 - .pFiles)
                           CopyMemory b(0), ByVal (hGlobal + .pFiles), cb - .pFiles

                           If .fWide Then
                              ' Unicode
                              saFiles = Split(b(), vbNullChar)
                           Else
                              ' ANSI
                              saFiles = Split(StrConv(b(), vbUnicode), vbNullChar)
                           End If
                        End With

                        For idx = 0 To UBound(saFiles) - 1
                           If Len(saFiles(idx)) > 1 Then   ' == saFiles(idx) <> vbNullChar
                              Debug.Assert DbgPrt(DBP >= 2, saFiles(idx))
                           End If
                        Next

                     Else
                        ' other formats

                     End If

                     Erase b()

                  End If   ' hGlobal <> 0

               ElseIf tSTGM.TYMED = TYMED_ISTREAM Then
'                  Dim iis        As IStream
'                  Dim tSTATSTG   As STATSTG
'                  Dim cbRead     As Long
'
'                  ' de-reference ptr to a stream object
'                  CopyMemory iis, tSTGM.pData, 4&
'
'                  If Not (iis Is Nothing) Then
'
'                     If SUCCEEDED(iis.Stat(tSTATSTG, STATFLAG_NONAME)) Then
'
'                        ' copy data to bytearray
'                        cb = tSTATSTG.cbSize * 10000
'
'                        If cb Then
'                           ReDim b(0 To cb - 1)
'
'                           If SUCCEEDED(iis.Read(b(0), cb, cbRead)) Then
'
'                              Debug.Assert cb = cbRead
'
'
'                              Erase b()
'
'                           End If
'                        End If
'                     End If
'
'                  End If
'
'                  CopyMemory iis, 0&, 4

               Else

                  Debug.Assert False
                  Erase b()
                  
               End If   ' tSTGM.TYMED = TYMED_HGLOBAL

               ReleaseStgMedium tSTGM

            End If   ' dvRet = S_OK
            Debug.Assert DbgPrt(dvRet <> S_OK And DBP >= 0, "FAILED ido.GetData()", dvRet)

            If tFOTC.ptd <> 0 Then
               ' If the IEnumFORMATETC::Next method returns a non-NULL DVTARGETDEVICE pointer in the ptd member of the FORMATETC structure, the memory must be freed with the CoTaskMemFree function (or its equivalent). Failure to do so will result in a memory leak.
               CoTaskMemFree tFOTC.ptd
            End If

            ' continue with next format
            hr = ief.Next(1, tFOTC, 0)

         Loop
         
         dbgIDataObject = True
         Debug.Assert DbgPrt(DBP >= 1, "")
         
      End If   ' SUCCEEDED(ido.EnumFormatEtc())
      
   End If
   Debug.Assert DbgPrt(ido Is Nothing And DBP >= 1, "FAILED IDataObject")

   Exit Function

Proc_Error:
   Debug.Print "Error: " & Err.Number & ". " & Err.Description
   If InIDE Then Stop: Resume
End Function

' extracts CFSTR_SHELLIDLIST format. (virtual and filesystem items)
Private Function dbgReadIDList(b() As Byte) As Boolean
   Dim tCIDA         As CIDA
   Dim isfParent     As IShellFolder
   Dim cPidls        As Long
'   Dim pidlRel       As Long
   Dim hGlobal       As Long
   Dim pidlFQ        As Long
   Dim pidlRelIdx    As Long
   Dim i             As Long
   Dim DBP           As Integer
   DBP = 2
   
   On Error GoTo Proc_Error

   ' MSDN: Shell Clipboard Formats: CFSTR_SHELLIDLIST
   '  The data is an STGMEDIUM structure that contains a global memory object.
   '  The structure's hGlobal member points to a CIDA structure.
   
   ' CIDA struct members:
   '  .cidl      as Long == Number of PIDLs that are being transferred, not counting the parent folder
   '  .aoffset() as Long == An array of offsets, relative to the beginning of this structure. The array contains cidl+1 elements. The first element of aoffset contains an offset to the fully-qualified PIDL of a parent folder. If this PIDL is empty, the parent folder is the desktop. Each of the remaining elements of the array contains an offset to one of the PIDLs to be transferred. All of these PIDLs are relative to the PIDL of the parent folder.
   ' To use this structure to retrieve a particular PIDL, add the PIDLs aoffset value to the address of the structure.
   
   ' cast first 4 bytes to long tCIDA.cidl
   CopyMemory cPidls, b(0), 4&
   
   With tCIDA
      .cidl = cPidls
      ReDim .aoffset(0 To cPidls)
      
      For i = 0 To cPidls
         ' cast 4 bytes to long tCIDA.aoffset ( b(4) + 3, b(8) + 3, ...)
         CopyMemory .aoffset(i), b(4 * i + 4), 4&
      Next
      
      ' address of our CIDA
      hGlobal = VarPtr(b(0))  ' !!! not VarPtr(tCIDA) !!!
      
      ' parent's pidl
      pidlFQ = hGlobal + .aoffset(0)
      Set isfParent = GetIShellFolder(isfDesktop, pidlFQ, fRtnDesktop:=True)
      
      Debug.Assert DbgPrt(DBP >= 1, vbCrLf & "dbgReadIDList" & String$(60, "-"))
      Debug.Assert DbgPrt(DBP >= 1, GetFolderDisplayName(isfDesktop, pidlFQ, SHGDN_FORPARSING))
      
      ' pidlRel: create a combined relative pidl of all transfered items
      For i = 1 To cPidls
         pidlRelIdx = hGlobal + .aoffset(i)
         Debug.Assert DbgPrt(DBP >= 2, i, GetFolderDisplayName(isfParent, pidlRelIdx, SHGDN_FORPARSING))
'         pidlRel = CombinePIDLs(pidlRel, pidlRelIdx, fFreePidl1:=True, fFreePidl2:=True)
      Next
      
   End With
   
   dbgReadIDList = True
   
   Debug.Assert DbgPrt(DBP >= 1, "")

   Exit Function

Proc_Error:
   Debug.Print "Error: " & Err.Number & ". " & Err.Description, App.Title & ".clsFolderTree: Function pGetIDataObjectFromVB"
   If InIDE Then Stop: Resume
End Function

