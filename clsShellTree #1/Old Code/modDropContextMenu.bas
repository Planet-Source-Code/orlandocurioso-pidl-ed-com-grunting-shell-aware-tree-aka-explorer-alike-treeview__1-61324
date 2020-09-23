Attribute VB_Name = "modDropContextMenu"
Option Explicit

Public Const vbDropEffectLink = 4&  ' DROPEFFECT_LINK

'Private Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As String) As Long
Private Declare Function LoadMenuLong Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As Long) As Long
Private Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long

Private Const MF_BYCOMMAND = &H0&
'Private Const MF_BYPOSITION = &H400&
'Private Const RT_MENU = 4&

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
   Private Const LOAD_LIBRARY_AS_DATAFILE = &H2
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
'

' x,y in hOwner client coords
Public Function ShowDropContextMenu(ByVal hOwner As Long, _
                                    ByRef Effect As OLEDropEffectConstants, _
                                    ByVal x As Long, ByVal y As Long, _
                                    Optional DefaultEffect As OLEDropEffectConstants _
                                    ) As Boolean
   Dim hMenuRes As Long
   Dim hSubMenu As Long
   Dim idCmd    As Long
   Dim tP       As POINTAPI
    
'   200 Menu
'   LANGUAGE LANG_GERMAN, SUBLANG_GERMAN
'   {
'   POPUP ""
'   {
'      MENUITEM "Hierher &verschieben", 2             ' Move
'      MENUITEM "Hierher &kopieren", 1                ' Copy
'      MENUITEM "Verknüpfungen hier &erstellen", 3    ' Link
'      MENUITEM Separator
'      MENUITEM "Abbrechen", 0                        ' Cancel
'   }
'   }
   
   If Effect = vbDropEffectNone Then Exit Function
   
   ' Extract menu 200 from shell32.dll
   hMenuRes = GetMenuResource("Shell32.dll", 200)
   
   If IsMenu(hMenuRes) Then
      
      hSubMenu = GetSubMenu(hMenuRes, 0&)
      
      ' remove menuitems of unsupported effects
      If (Effect And vbDropEffectMove) = 0 Then
         RemoveMenu hSubMenu, vbDropEffectMove, MF_BYCOMMAND
      End If
      If (Effect And vbDropEffectCopy) = 0 Then
         RemoveMenu hSubMenu, vbDropEffectCopy, MF_BYCOMMAND
      End If
      If (Effect And vbDropEffectLink) = 0 Then
         RemoveMenu hSubMenu, vbDropEffectLink - 1, MF_BYCOMMAND  ' CMD = 3
      End If

      Select Case DefaultEffect
         Case vbDropEffectMove, vbDropEffectCopy
            SetMenuDefaultItem hSubMenu, DefaultEffect, MF_BYCOMMAND
         Case vbDropEffectLink
            SetMenuDefaultItem hSubMenu, vbDropEffectLink - 1, MF_BYCOMMAND   ' CMD = 3
      End Select
      
      tP.x = x:   tP.y = y
      Call ClientToScreen(hOwner, tP)
      
      ' Show the drop context menu.
      idCmd = TrackPopupMenu(hSubMenu, _
                              TPM_LEFTBUTTON Or TPM_RIGHTBUTTON Or _
                              TPM_LEFTALIGN Or TPM_TOPALIGN Or _
                              TPM_HORIZONTAL Or TPM_RETURNCMD, _
                              tP.x, tP.y, 0, hOwner, 0)
      DestroyMenu hSubMenu
      DestroyMenu hMenuRes
      
      ' return false, if user cancelled
      ShowDropContextMenu = (idCmd <> 0)
      
   Else
      ' Win9X
      Debug.Assert False
   End If
   
   ' Return selected menu command == chosen drop effect
   If idCmd = 3 Then idCmd = vbDropEffectLink
   Effect = idCmd

End Function

' http://vb.mvps.org/articles/ap200005.htm
Private Function GetMenuResource(ByVal ModuleName As String, ResID As Integer) As Long
   Dim hModule As Long
   Dim FreeLib As Boolean
   Static Busy As Boolean

   Debug.Assert LenB(ModuleName)

   ' This routine is *not* re-entrant!
   If Not Busy Then Busy = True
   
   ' Check first to see if the module is already mapped into this process.
   hModule = GetModuleHandle(ModuleName)
   If hModule = 0 Then
      ' Load library
      hModule = LoadLibraryEx(ModuleName, 0&, LOAD_LIBRARY_AS_DATAFILE)
      If hModule = 0 Then
         Debug.Print "LoadLibraryEx error: " & Err.LastDllError
         Debug.Assert False
         Busy = False
      Else
         ' Set a flag that reminds us to free this handle.
         FreeLib = True
      End If
   End If
      
   ' Only load menu if no problems loading module.
   If Busy Then
      
      ' # >=Win2K    : LoadLibraryEx , LoadMenu OK                                     #
      ' # Win9X/NT   : FindResourceEx/RT_MENU,LoadResource,LockResource                #
      ' # NOT YET IMPLEMENTED #
      GetMenuResource = LoadMenuLong(hModule, ResID)    ' MAKEINTRESOURCE(ResID)
      
      ' Close handle for any module we loaded.
      If FreeLib Then
         Call FreeLibrary(hModule)
         ' Clear re-entry flag, and return success.
         Busy = False
      End If
      
   End If
   
End Function

