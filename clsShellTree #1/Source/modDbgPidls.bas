Attribute VB_Name = "modDbgPidls"
'---------------------------------------------------------------------------------------
' Module    : modDbgPidls
' Author    : OrlandoCurioso 09.05.2005
' Purpose   : PIDL debug and experimental functions.
'
'---------------------------------------------------------------------------------------
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'

' any type of pidl
Public Function dbgDumpPIDL(ByVal pidl As Long, Optional cStart As Integer) As String
   Dim b(0)    As Byte
   Dim nCount  As Integer
   Dim i       As Integer
   Dim s       As String
   
   If pidl Then
      nCount = GetPIDLSize(pidl)
      
      For i = cStart To nCount - 1
      
         CopyMemory b(0), ByVal (pidl + i), 1&
         dbgDumpPIDL = dbgDumpPIDL & " " & Format$(Hex$(b(0)), "00")
         s = s & " " & Format$(Chr$(b(0)), "#")
      Next
      dbgDumpPIDL = dbgDumpPIDL & vbCrLf & s
   End If
End Function

Public Function dbgDumpMEM(ByVal hMem As Long, cb As Long) As String
   Dim b(0)    As Byte
   Dim i       As Integer
   Dim s       As String
   
   If hMem Then
      
      For i = 0 To cb - 1
      
         CopyMemory b(0), ByVal (hMem + i), 1&
         dbgDumpMEM = dbgDumpMEM & " " & Format$(Hex$(b(0)), "00")
         s = s & " " & Format$(Chr$(b(0)), "#")
      Next
      dbgDumpMEM = dbgDumpMEM & vbCrLf & s
   End If
End Function


Public Function dbgShowAttributes(pidlFQ As Long, pidlRel As Long) As String
   Dim isfParent  As IShellFolder
   Dim ulAttrs    As ESFGAO

   On Error GoTo Proc_Error
   
   ulAttrs = SFGAO_VALIDATE Or _
            SFGAO_CANCOPY Or SFGAO_CANDELETE Or SFGAO_CANLINK Or _
            SFGAO_CANMOVE Or SFGAO_CANRENAME Or SFGAO_CANMONIKER Or _
            SFGAO_COMPRESSED Or SFGAO_DROPTARGET Or SFGAO_ENCRYPTED Or _
            SFGAO_FILESYSANCESTOR Or SFGAO_FILESYSTEM Or _
            SFGAO_FOLDER Or SFGAO_HASSUBFOLDER Or _
            SFGAO_LINK Or SFGAO_READONLY Or SFGAO_HASPROPSHEET Or _
            SFGAO_REMOVABLE Or SFGAO_SHARE Or _
            SFGAO_HIDDEN Or SFGAO_GHOSTED Or _
            SFGAO_BROWSABLE Or SFGAO_ISSLOW Or SFGAO_NEWCONTENT Or _
            SFGAO_STORAGEANCESTOR Or SFGAO_STORAGE Or SFGAO_HASSTORAGE Or SFGAO_STREAM
   
   ' Get the parent's IShellFolder
   Set isfParent = GetIShellFolderParent(pidlFQ, fRtnDesktop:=True)
            
   If (S_OK = isfParent.GetAttributesOf(1, pidlRel, ulAttrs)) Then

      dbgShowAttributes = _
         IIf((ulAttrs And SFGAO_CANCOPY), "CANCOPY", "") & _
         IIf((ulAttrs And SFGAO_CANDELETE), vbTab & "CANDELETE", "") & _
         IIf((ulAttrs And SFGAO_CANLINK), vbTab & "CANLINK", "") & _
         IIf((ulAttrs And SFGAO_CANMOVE), vbTab & "CANMOVE", "") & _
         IIf((ulAttrs And SFGAO_CANRENAME), vbTab & "CANRENAME", "") & _
         IIf((ulAttrs And SFGAO_CANMONIKER), vbTab & "CANMONIKER", "") & _
         IIf((ulAttrs And SFGAO_COMPRESSED), vbTab & "COMPRESSED", "") & _
         IIf((ulAttrs And SFGAO_DROPTARGET), vbTab & "DROPTARGET", "") & _
         IIf((ulAttrs And SFGAO_ENCRYPTED), vbTab & "ENCRYPTED", "") & _
         IIf((ulAttrs And SFGAO_FILESYSANCESTOR), vbTab & "FILESYSANCESTOR", "") & _
         IIf((ulAttrs And SFGAO_FILESYSTEM), vbTab & "FILESYSTEM", "") & _
         IIf((ulAttrs And SFGAO_FOLDER), vbTab & "FOLDER", "") & _
         IIf((ulAttrs And SFGAO_HASSUBFOLDER), vbTab & "HASSUBFOLDER", "") & _
         IIf((ulAttrs And SFGAO_LINK), vbTab & "LINK", "") & _
         IIf((ulAttrs And SFGAO_READONLY), vbTab & "READONLY", "") & _
         IIf((ulAttrs And SFGAO_HASPROPSHEET), vbTab & "HASPROPSHEET", "")
         
      dbgShowAttributes = dbgShowAttributes & _
         IIf((ulAttrs And SFGAO_REMOVABLE), vbTab & "REMOVABLE", "") & _
         IIf((ulAttrs And SFGAO_SHARE), vbTab & "SFGAO_SHARE", "") & _
         IIf((ulAttrs And SFGAO_HIDDEN), vbTab & "HIDDEN", "") & _
         IIf((ulAttrs And SFGAO_GHOSTED), vbTab & "GHOSTED", "") & _
         IIf((ulAttrs And SFGAO_BROWSABLE), vbTab & "BROWSABLE", "") & _
         IIf((ulAttrs And SFGAO_ISSLOW), vbTab & "ISSLOW", "") & _
         IIf((ulAttrs And SFGAO_NEWCONTENT), vbTab & "NEWCONTENT", "") & _
         IIf((ulAttrs And SFGAO_STORAGEANCESTOR), vbTab & "STORAGEANCESTOR", "") & _
         IIf((ulAttrs And SFGAO_HASSTORAGE), vbTab & "HASSTORAGE", "") & _
         IIf((ulAttrs And SFGAO_STORAGE), vbTab & "STORAGE", "") & _
         IIf((ulAttrs And SFGAO_STREAM), vbTab & "STREAM", "")

      If Left$(dbgShowAttributes, Len(vbTab)) = vbTab Then
         dbgShowAttributes = Mid$(dbgShowAttributes, Len(vbTab) + 1)
      End If
'      Debug.Print dbgShowAttributes
      
   Else
      Debug.Print "dbgShowAttributes FAILED"
   End If

   Exit Function

Proc_Error:
   Debug.Print "Error: " & Err.Number & ". " & Err.Description, App.Title & " Function dbgShowAttributes"
   If InIDE Then Stop: Resume
End Function

Public Function dbgWalkPIDL(ByVal pidl As Long, _
                            Optional ByVal uFlags As ESHGNO = SHGDN_NORMAL) As String
   Dim pidlCopy      As Long
   Dim isfParent     As IShellFolder
   Dim sRes          As String
   
   Select Case GetPIDLSize(pidl)
      Case 0
         Exit Function
      Case Is <= 12&
         dbgWalkPIDL = "WalkPIDL SKIPPED"
         Exit Function
   End Select
   
   Set isfParent = isfDesktop
   
   Do While pidl
      
      ' Copy the item identifier to a list by itself (simple pidl)
      pidlCopy = GetItemID(pidl, GIID_FIRST)
      If pidlCopy = 0 Then Exit Do
      
      ' reached terminator (err with GetFolderDisplayName)
      If GetItemIDSize(pidlCopy) = 0 Then
         FreePIDL pidlCopy
         Exit Do
      End If
      
      ' Display the name of the subfolder
      sRes = sRes & "\" & GetFolderDisplayName(isfParent, pidlCopy, uFlags)

      ' Bind to the subfolder
      Set isfParent = GetIShellFolder(isfParent, pidlCopy)
      
      ' Free the copy of the item identifier
      FreePIDL pidlCopy
      
      ' Get the next item ID
      pidl = GetNextItemID(pidl)
   Loop
   
   sRes = Mid$(sRes, 2)
'   Debug.Print sRes
   
   dbgWalkPIDL = sRes
End Function

Public Function dbgContentsPidl(ByVal pidlFQ As Long, Optional ByVal hOwner As Long, _
                                Optional ByVal eIncludeItems As ESHCONTF) As String
   Dim isfParent     As IShellFolder
   Dim iEIDL         As IEnumIDList
   Dim pidlRel       As Long
   Dim sRes          As String
   
   On Error GoTo Proc_Error

   ' Get the parent's IShellFolder from its fully qualified pidl
   Set isfParent = GetIShellFolder(isfDesktop, pidlFQ, fRtnDesktop:=False)
   
   If isfParent Is Nothing Then Err.Raise 91
   
   ' Create an enumeration object for the parent folder.
   If SUCCEEDED(isfParent.EnumObjects(hOwner, eIncludeItems, iEIDL)) Then
                                                                
      ' Enumerate the contents of the parent folder
      Do While (iEIDL.Next(1, pidlRel, 0) = NOERROR)
   
         ' Display the name of the item
         sRes = sRes & "|" & GetFolderDisplayName(isfParent, pidlRel, SHGDN_INFOLDER)
   
         ' Free the relative pidl the enumeration gave us.
         FreePIDL pidlRel
       
      Loop
      
   Else
'      Debug.Print "Error: &H" & Hex(hr) & ", " & GetAPIErrStr(hr)
      Debug.Assert False
   End If
   
   sRes = Mid$(sRes, 2)
'   Debug.Print sRes
   
   dbgContentsPidl = sRes

   Exit Function

Proc_Error:
   Debug.Print "Error: " & Err.Number & ". " & Err.Description, vbOKOnly Or vbCritical, App.Title & " Function dbgContentsPidl"
   If InIDE Then Stop: Resume
End Function

