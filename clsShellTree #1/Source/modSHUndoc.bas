Attribute VB_Name = "modSHUndocumented"
'---------------------------------------------------------------------------------------
' Module    : modSHUndocumented
' Author    : OrlandoCurioso 16.05.2005
' Credits   : Juergen Schmied (internal pidl functions.h)   ' http://cvs.winehq.com/cvsweb/wine/dlls/shell32/pidl.h?rev=1.31
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

Public Enum ePidlDataType        ' undocumented ( see dbgGetPIDLType() )
   pdtDesktopChild = &H1F
   pdtDrive = &H2F
   pdtShellExtension = &H2E      ' Control Panel
   pdtFolder = &H31
   pdtFileOrLink = &H32          ' also new item
   pdtLink = &H3A
   pdtB1 = &HB1                  ' Recycled,History,Favorites,My Websites on MSN
   pdtNetWorkWorkGroup = &H41
   pdtNetWorkComputer = &H42
   pdtNetWorkProvider = &H46
   pdtNetWorkEntire = &H47
End Enum

'  ' item identifier (relative pidl), allocated by the shell
'  Public Type SHITEMID
'    cb As Integer         ' size of struct, including cb itself
'    abID() As Byte        ' variable length item identifier
'  End Type

Private Type PIDLDATA_DRIVE
   pidlType    As Byte
   DrivePath   As String * 20
   Unknown     As Long
End Type

Private Type PIDLDATA_GENERIC
   pidlType       As Byte
   Dummy          As Byte
   dwFileSizeL    As Integer     ' or dwFileSize As Long & switch HIWORD/LOWORD
   dwFileSizeH    As Integer
   uFileDate      As Integer
   uFileTime      As Integer
   uFileAttribs   As Integer
   DosName        As String      ' + 12 bytes
   '                             # ~20 bytes of unknown contents #
   Name           As String
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
'

Public Function GetPIDLType(ByVal pidlRel As Long) As ePidlDataType
   ' undocumented type identifier
   CopyMemory GetPIDLType, ByVal (pidlRel + 2), 1&
End Function

Public Function IsFloppyPIDL(ByVal pidlRel As Long) As Boolean
   Dim b()     As Byte
   
   If GetPIDLType(pidlRel) = pdtDrive Then
   
      ReDim b(0 To 1)
      CopyMemory b(0), ByVal (pidlRel + 3), 2&
      
      IsFloppyPIDL = (b(0) = 65&) Or (b(0) = 66&)  ' Asc("A"),Asc("B")
      
      Debug.Assert b(1) = Asc(":")
'      Dim sPath   As String
'      sPath = StrConv(b(), vbUnicode)
'      Debug.Print sPath
   End If

End Function

' # coded on WinXP SP2 #
Public Function dbgGetPIDLType(ByVal pidlRel As Long) As String
   Dim tDRV    As PIDLDATA_DRIVE
   Dim tGEN    As PIDLDATA_GENERIC
   Dim b()     As Byte
   Dim s       As String
   Dim s2      As String
   Dim i       As Integer
   Dim cb      As Integer
   Dim eType   As Integer
   Dim pAddr   As Long
   
   On Error GoTo Proc_Error
   
   ' first  byte of abID() is some undocumented type identifier
   ' second byte is always zero
   
   If pidlRel Then
         
      If Not IsDesktopPIDL(pidlRel) Then
      
         CopyMemory eType, ByVal (pidlRel + 2), 1&
         dbgGetPIDLType = Hex$(eType)
         
         Select Case eType
         
            Case &H1F:        s = "Child of Desktop"
            
            Case &H2F:        s = "Drive"
            
               CopyMemory tDRV, ByVal (pidlRel + 2), Len(tDRV)
               
               With tDRV
                  Debug.Assert .pidlType = &H2F
                  
                  ' DrivePath is ANSI string
                  s = s & vbTab & _
                      Left$(.DrivePath, InStr(.DrivePath & vbNullChar, vbNullChar) - 1) _
                      & vbTab & .Unknown
                  
                  If GetPIDLSize(pidlRel) <> 2 + Len(tDRV) Then
                     ' only my cd burner has PIDLSize = 43 bytes and tDRV.Unknown = 138
                     s = s & vbTab & "DIFFERENT DRIVE SIZE"
                ' else
                     ' my floppy, 2 harddisks , several partions, a USB removable, cdrom:
                     ' all have PIDLSize = 27 bytes and tDRV.Unknown = 0
                  End If
               End With
               
            Case &H2E:        s = "Control Panel"
            
            Case &H31, &H32:  s = IIf(eType = &H31, "Folder", "File or Link or New")
               
               ' addr(abID(0)) == pidlRel + SHITEMID.cb
               pAddr = pidlRel + 2
               
               ' copy fixed part into udt (not the strings)
               CopyMemory tGEN, ByVal (pAddr), 12&
               With tGEN
               
                  Debug.Assert (.pidlType = &H31) Or (.pidlType = &H32)
                  Debug.Assert .Dummy = 0
                  
                  s2 = (.dwFileSizeH * 2 ^ 16 + .dwFileSizeL) & vbTab & _
                        .uFileDate & vbTab & .uFileTime & vbTab & .uFileAttribs
               
                  ' addr(DOS name 8+3) = addr + pidltype,dummy,size,date,time,attr
                  pAddr = pAddr + 12
                  
                  ' buffer 12 bytes (8 + "." + 3)
                  ReDim b(0 To 11)
                  CopyMemory b(0), ByVal (pAddr), 12&
                  ' DOS name is ANSI string
                  .DosName = StrConv(b, vbUnicode) & vbNullChar
                  .DosName = Left$(.DosName, InStr(1, .DosName, vbNullChar) - 1)
                  
                  ' after 21 or 22 bytes follows the long name as Unicode string
                  ' len of DOS name: even 22 bytes / uneven 21 bytes
                  ' # ~20 bytes of unknown contents #
                  i = 22 - Len(.DosName) Mod 2
                  
                  ' address of long name
                  pAddr = pAddr + Len(.DosName) + i
                  
                  ' size : subtract SHITEMID.cb + 2 bytes for the zero terminating item ID
                  cb = GetPIDLSize(pidlRel) - 4
                  
                  ' remaining maximum bytes for long name
                  cb = pidlRel + cb - pAddr
                  Debug.Assert cb > 0
                  
                  .Name = String$(cb \ 2 + 1, vbNullChar)
                  CopyMemory ByVal StrPtr(.Name), ByVal pAddr, cb
                  .Name = Left$(.Name, lstrlenW(StrPtr(.Name)))
               
                  ' File or Link ? can now be determined by DOS name
                  If eType = &H32 Then
                     
                     s = "File  "
                     
                     If Len(.DosName) > 4 Then
                        If Right$(.DosName, 4) = ".LNK" Then
                           s = "Link  "
                        End If
                     End If
                     
                     With tGEN
                        If (.dwFileSizeH = 0) And (.dwFileSizeL = 0) Then
                           If (.uFileDate) And (.uFileTime = 0) Then
                              If .uFileAttribs = 0 Then s = "New Item"
                           End If
                        End If
                     End With
                     
                  End If
                  
                  s = s & vbTab & Space$(12 - Len(.DosName)) & .DosName & vbTab & .Name & vbTab & s2
               
               End With
               
            Case &H3A:  s = "Link  "
            
            Case &H40 To &H4F: s = "NetWork"
            
            Case &HB1:  s = "Recycled,History,Favorites,My Websites on MSN"
            
            Case Else:  s = "UNKNOWN"
            
         End Select
      
      Else
         dbgGetPIDLType = "0 ": s = "DeskTop"
      End If
      
      dbgGetPIDLType = dbgGetPIDLType & vbTab & Format$(GetPIDLSize(pidlRel), "0##") & vbTab & s
      
   End If
   
   Exit Function

Proc_Error:
   Debug.Print "Error: " & Err.Number & ". " & Err.Description
   If InIDE Then Stop: Resume
End Function


