Attribute VB_Name = "modSHContextMenu"
'---------------------------------------------------------------------------------------
' Module    : modSHContextMenu
' Author    : OrlandoCurioso 13.05.2005 / Brad Martinez
' Purpose   :
' Bugs      : Some CanonicalVerb's work (properties), some not (open) ???
' CRASH     : ISF.CreateViewObject on Win98
'---------------------------------------------------------------------------------------
Option Explicit

Private Const idCF = 1&

' MenuOffset: all enum values below are relative to idCmdFirst
Public Const idCmdFirst = idCF
Public Const idCmdLast = &H7FFF

Public Enum SHContextMenuCmdIDs
   cmd_Cancel = 0
   cmd_CreateLink = idCF + 16
   cmd_Delete = idCF + 17
   cmd_Rename = idCF + 18
   cmd_Properties = idCF + 19
   cmd_Cut = idCF + 24
   cmd_Copy = idCF + 25
   cmd_Paste = idCF + 26
End Enum

' Shell item verbs
'Public Const scmd_CreateLink = "link"
'Public Const scmd_Delete = "delete"
'Public Const scmd_Rename = "rename"
'Public Const scmd_Properties = "properties"
'Public Const scmd_Cut = "cut"
'Public Const scmd_Copy = "copy"
'Public Const scmd_Paste = "paste"

' Shell view background verbs
Public Const scmd_NewFolder = "NewFolder"
'Public Const scmd_CreateLink = "link"


' Global IContextMenu2/3 ownerdrawn interfaces:
' Menu messages must be forwarded to interfaces HandleMenuMsgX methods in owner's wndproc.
' WM_INITMENUPOPUP, WM_DRAWITEM, WM_MEASUREITEM + WM_MENUCHAR (IContextMenu3)
Public ICtxMenu2 As IContextMenu2
Public ICtxMenu3 As IContextMenu3

' //

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long

' // Menu

Private Enum TPM_wFlags
   TPM_LEFTBUTTON = &H0
   TPM_RIGHTBUTTON = &H2
   TPM_LEFTALIGN = &H0
   TPM_CENTERALIGN = &H4
   TPM_RIGHTALIGN = &H8
   TPM_TOPALIGN = &H0
   TPM_VCENTERALIGN = &H10
   TPM_BOTTOMALIGN = &H20
   
   TPM_HORIZONTAL = &H0
   TPM_VERTICAL = &H40
   TPM_NONOTIFY = &H80
   TPM_RETURNCMD = &H100
End Enum

Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As TPM_wFlags, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As Any) As Long
'

' ----------------------------------------------------------------------------------- '
'  IShellFolder context menu
' ----------------------------------------------------------------------------------- '

' Displays the specified items' shell context menu.
'
'    hOwner       - window handle that owns context menu and any err msgboxes
'    isfParent    - pointer to the items' parent shell folder
'    cPidls       - count of pidls at, and after, pidlRel
'    pidlRel      - the first item's pidl, relative to isfParent
'    x,y          - location of the context menu, in hOwner client coords
'    CMF          - options which commands are added -> MSDN: IContextMenu::QueryContextMenu
' Returns selected menu command, 0 if cancelled.

Public Function ShowShellContextMenu(ByVal hOwner As Long, isfParent As IShellFolder, _
                                     ByVal cPidls As Long, pidlRel As Long, _
                                     ByVal X As Long, ByVal Y As Long, _
                                     Optional ByVal CMF As eCMF = CMF_EXPLORE _
                                     ) As SHContextMenuCmdIDs

   Dim icm     As IContextMenu
   Dim idxICM  As Long

   ' Get a reference to the item's IContextMenuX interface
   idxICM = pGetIContextMenu(hOwner, isfParent, cPidls, pidlRel, icm)
   
   ShowShellContextMenu = pShowSHContextMenu(hOwner, icm, idxICM, X, Y, CMF)
End Function

' Invokes a commmand present on item's IContextMenu, but without showing menu.
' # HowTo indicate sucess/failure #
Public Function InvokeShellCommand(ByVal hOwner As Long, isfParent As IShellFolder, _
                                   ByVal cPidls As Long, pidlRel As Long, _
                                   Optional ByVal idCmd As SHContextMenuCmdIDs, _
                                   Optional ByVal CanonicalVerb As String = vbNullString, _
                                   Optional ByVal CMF As eCMF = CMF_EXPLORE) As Boolean
   Dim icm     As IContextMenu
   
   If idCmd = 0 And LenB(CanonicalVerb) = 0 Then Exit Function
   
   ' Get a reference to the item's IContextMenu interface (any version)
   If pGetIContextMenu(hOwner, isfParent, cPidls, pidlRel, icm, bStandard:=True) Then
      
      InvokeShellCommand = pInvokeSHCmd(hOwner, icm, idCmd, CanonicalVerb, CMF)
   End If
End Function

' Returns a reference to the item's highest version IContextMenu interface.
' Returns version as number (3,2, 1 for icm, 0 if failed).
' bStandard = True : don't retrieve higher version
' cPidls >1 : pidlRel as immediate children of isfParent.
Private Function pGetIContextMenu(ByVal hOwner As Long, isfParent As IShellFolder, _
                                  ByVal cPidls As Long, pidlRel As Long, _
                                  icm As IContextMenu, Optional ByVal bStandard As Boolean _
                                  ) As Long
   Dim icm2    As IContextMenu2
   Dim icm3    As IContextMenu3
   Dim hr      As Long
    
   Set icm = Nothing
   
   ' Get a reference to the item's IContextMenu interface
   hr = isfParent.GetUIObjectOf(hOwner, cPidls, pidlRel, GetRIID(rIID_IContextMenu), 0, icm)
   If SUCCEEDED(hr) Then
      
      pGetIContextMenu = 1
      If bStandard Then Exit Function
      
      ' Try obtaining the higher version interfaces (supersets of IContextMenu):
      ' Needed so submenus get filled from the HandleMenuMsg calls in owner's wndproc.
      
      hr = icm.QueryInterface(GetRIID(rIID_IContextMenu3), icm3)
      
      If Not (icm3 Is Nothing) Then
         Set icm = icm3
         pGetIContextMenu = 3
         
      Else
         hr = icm.QueryInterface(GetRIID(rIID_IContextMenu2), icm2)
         
         If Not (icm2 Is Nothing) Then
            Set icm = icm2
            pGetIContextMenu = 2
         End If
      
      End If
      
   End If

End Function

Public Function FindShellCommand(ByVal hOwner As Long, isfParent As IShellFolder, _
                                 ByVal cPidls As Long, pidlRel As Long, _
                                 ByVal CanonicalVerb As String, _
                                 Optional ByVal CMF As eCMF = CMF_EXPLORE) As SHContextMenuCmdIDs
   Dim icm     As IContextMenu
   
   ' Get a reference to the item's IContextMenu interface (any version)
   If pGetIContextMenu(hOwner, isfParent, cPidls, pidlRel, icm, bStandard:=True) Then
      
      FindShellCommand = pFindSHCmd(icm, CanonicalVerb, CMF)
   End If
End Function

' ----------------------------------------------------------------------------------- '
'  IShellView context menu
' ----------------------------------------------------------------------------------- '
   
Public Function ShowSHCtxMenuViewBK(ByVal hOwner As Long, isfParent As IShellFolder, _
                                    ByVal X As Long, ByVal Y As Long, _
                                    Optional ByVal CMF As eCMF) As Long
   Dim isv     As IShellView
   Dim icm     As IContextMenu
   Dim idxICM  As Long
   
   Debug.Assert Not (isfParent Is Nothing)
   
   ' Get a reference to the folder's view object
#If WIN32_IE >= &H500 Then
   If SUCCEEDED(isfParent.CreateViewObject(hOwner, GetRIID(rIID_IShellView), isv)) Then
#Else
   ' # CRASH ??? #
   If False Then
#End If
   
      ' Get a reference to the view background's IContextMenuX interface
      idxICM = pGetCtxMenuView(isv, icm)
   
      ShowSHCtxMenuViewBK = pShowSHContextMenu(hOwner, icm, idxICM, X, Y, CMF)
   End If
End Function

' Invokes a commmand present on view background's IContextMenu, but without showing menu.
' # HowTo indicate sucess/failure #
Public Function InvokeShellViewCmd(ByVal hOwner As Long, isfParent As IShellFolder, _
                                   Optional ByVal idCmd As Long, _
                                   Optional ByVal CanonicalVerb As String = vbNullString, _
                                   Optional ByVal CMF As eCMF) As Boolean
   Dim isv     As IShellView
   Dim icm     As IContextMenu
  
   If idCmd = 0 And LenB(CanonicalVerb) = 0 Then Exit Function
    
   ' Get a reference to the folder's view object
#If WIN32_IE >= &H500 Then
   If SUCCEEDED(isfParent.CreateViewObject(hOwner, GetRIID(rIID_IShellView), isv)) Then
#Else
   ' # CRASH ??? #
   If False Then
#End If
      ' Get a reference to the view background's IContextMenu interface (any version)
      If pGetCtxMenuView(isv, icm, bStandard:=True) Then
         
         InvokeShellViewCmd = pInvokeSHCmd(hOwner, icm, idCmd, CanonicalVerb, CMF)
      End If
   End If
   
End Function

' Returns a reference to the view background's highest version IContextMenu interface.
' Returns version as number (3,2, 1 for icm, 0 if failed).
' bStandard = True : don't retrieve higher version
Private Function pGetCtxMenuView(isv As IShellView, icm As IContextMenu, _
                                 Optional ByVal bStandard As Boolean) As Long
   Dim icm2    As IContextMenu2
   Dim icm3    As IContextMenu3
   Dim hr      As Long
    
   Set icm = Nothing
    
   ' Get a reference to the view background's IContextMenu interface
   hr = isv.GetItemObject(SVGIO_BACKGROUND, GetRIID(rIID_IContextMenu), icm)
   If SUCCEEDED(hr) Then
      
      pGetCtxMenuView = 1
      If bStandard Then Exit Function
      
      ' Try obtaining the higher version interfaces (supersets of IContextMenu):
      ' Needed so submenus get filled from the HandleMenuMsg calls in owner's wndproc.
      
      hr = icm.QueryInterface(GetRIID(rIID_IContextMenu3), icm3)
      
      If Not (icm3 Is Nothing) Then
         Set icm = icm3
         pGetCtxMenuView = 3
         
      Else
         hr = icm.QueryInterface(GetRIID(rIID_IContextMenu2), icm2)
         
         If Not (icm2 Is Nothing) Then
            Set icm = icm2
            pGetCtxMenuView = 2
         End If
      
      End If
      
   End If

End Function

' ----------------------------------------------------------------------------------- '
'  Shell context menus common functions
' ----------------------------------------------------------------------------------- '

' # ? GCS_VALIDATE always returns error ,although correct verb and helptext are retrieved ? #
Public Function GetShellCommandString(ByRef icm As IContextMenu, _
                                      ByVal idCmd As SHContextMenuCmdIDs, _
                                      Optional ByVal GCS As eGCS = GCS_VERB, _
                                      Optional ByRef sReturn As String) As Boolean
   Dim lRes As Long
   
   idCmd = idCmd - idCmdFirst
   
   Select Case GCS

      Case GCS_VERB, GCS_HELPTEXT, GCS_VERB Or GCS_UNICODE, GCS_HELPTEXT Or GCS_UNICODE
         sReturn = String$(MAX_PATH, vbNullChar)
         If (0 <= icm.GetCommandString(idCmd, GCS, 0, sReturn, MAX_PATH)) Then
            Select Case GCS
               Case GCS_VERB Or GCS_UNICODE, GCS_HELPTEXT Or GCS_UNICODE
                  sReturn = StrConv(sReturn, vbFromUnicode)
            End Select
            If InStr(sReturn, vbNullChar) Then
               sReturn = Left$(sReturn, InStr(sReturn, vbNullChar) - 1)
            End If
            GetShellCommandString = True
         End If

      Case GCS_VALIDATE, GCS_VALIDATE Or GCS_UNICODE
         ' Returns S_OK if the menu item exists, or S_FALSE otherwise
         lRes = icm.GetCommandString(idCmd, GCS, 0, vbNullString, 0)

         If lRes = S_OK Then
            Debug.Assert False
            ' # ? never ever ? #
            GetShellCommandString = True
         Else
            ' # despite MSDN doc: returns HResult (not implemented), not S_FALSE #
            Debug.Assert lRes = &H80004001
         End If

      Case Else:  Debug.Assert False
   End Select
   
   If GetShellCommandString = False Then
      sReturn = vbNullString
   End If
End Function

Private Function pShowSHContextMenu(ByVal hOwner As Long, icm As IContextMenu, ByVal idxICM As Long, _
                                    ByVal X As Long, ByVal Y As Long, _
                                    Optional ByVal CMF As eCMF) As Long
   Dim hr      As Long
   Dim hMenu   As Long
   Dim idCmd   As Long
   Dim tP      As POINTAPI
   Dim cmi     As CMINVOKECOMMANDINFO
   
   If idxICM Then
   
      Debug.Assert Not (icm Is Nothing)
   
      Select Case idxICM
         Case 3
            Set ICtxMenu3 = icm
         Case 2
            Set ICtxMenu2 = icm
      End Select
   
      ' Create a new popup menu...
      hMenu = CreatePopupMenu()
      If hMenu Then

         ' Add the item's shell commands to the popup menu.
         hr = icm.QueryContextMenu(hMenu, 0, idCmdFirst, idCmdLast, CMF)
         
         If SUCCEEDED(hr) Then
        
            tP.X = X:   tP.Y = Y
            Call ClientToScreen(hOwner, tP)
            
            ' Show the item's context menu.
            idCmd = TrackPopupMenu(hMenu, _
                                   TPM_LEFTBUTTON Or TPM_RIGHTBUTTON Or _
                                   TPM_LEFTALIGN Or TPM_TOPALIGN Or _
                                   TPM_HORIZONTAL Or TPM_RETURNCMD, _
                                   tP.X, tP.Y, 0, hOwner, 0)
        
            ' If a menu command is selected...
            If idCmd Then
        
               ' Fill the struct with the selected command's information.
               With cmi
                  .cbSize = Len(cmi)
                  .hWnd = hOwner
                  .lpVerb = idCmd - idCmdFirst   ' MAKEINTRESOURCE(idCmd-idCmdFirst)
                  .nShow = SW_SHOWNORMAL
               End With
               
               ' Invoke the shell's context menu command. The call itself does
               ' not err if the pidlRel item is invalid, but depending on the selected
               ' command, Explorer *may* raise an err. We don't need the return
               ' val, which should always be NOERROR anyway...
               Call icm.InvokeCommand(cmi)
          
            End If   ' idCmd
            
         End If   ' hr >= NOERROR (QueryContextMenu)

         Call DestroyMenu(hMenu)
    
      End If   ' hMenu
      
   End If   ' idxICM
   
   ' Release the item's IContextMenuX from the global variables.
   Set ICtxMenu2 = Nothing
   Set ICtxMenu3 = Nothing
   
   ' Return selected menu command
   pShowSHContextMenu = idCmd

End Function


' Invokes a commmand present on item's IContextMenu, but without showing menu.
' # HowTo indicate sucess/failure #
Private Function pInvokeSHCmd(ByVal hOwner As Long, icm As IContextMenu, _
                              ByVal idCmd As SHContextMenuCmdIDs, _
                              Optional ByVal CanonicalVerb As String = vbNullString, _
                              Optional ByVal CMF As eCMF = CMF_EXPLORE) As Boolean
   
   Dim hr      As Long   ' HRESULT
   Dim hMenu   As Long
   Dim cmi     As CMINVOKECOMMANDINFO
   
   ' Create a new popup menu...
   hMenu = CreatePopupMenu()
   If hMenu Then
   
      ' Add the item's shell commands to the popup menu.
      hr = icm.QueryContextMenu(hMenu, 0, idCmdFirst, idCmdLast, CMF)
      
      If SUCCEEDED(hr) Then
     
         ' hr == highest offset of the largest command identifier that was assigned, plus one
         ' hr == idCmdMax - idCmdFirst + 1
         Debug.Assert (idCmd - idCmdFirst) < hr
         
         ' Fill the struct with the selected command's information.
         With cmi
            .cbSize = Len(cmi)
            .hWnd = hOwner
            If LenB(CanonicalVerb) = 0 Then
               .lpVerb = idCmd - idCmdFirst   ' MAKEINTRESOURCE(idCmd-idCmdFirst)
               Debug.Assert HIWORD(.lpVerb) = 0
            Else
               .lpVerb = StrPtr(StrConv(CanonicalVerb & vbNullChar, vbFromUnicode))
               Debug.Assert HIWORD(.lpVerb) <> 0
            End If
            .nShow = SW_SHOWNORMAL
'           .fMask = CMIC_MASK_FLAG_NO_UI  ' # similiar to SEE_XXX flags for ShellExecuteEx
         End With
         
         ' Invoke the shell's context menu command. The call itself does
         ' not err if the pidlRel item is invalid, but depending on the selected
         ' command, Explorer *may* raise an err. We don't need the return
         ' val, which should always be NOERROR anyway...
         Call icm.InvokeCommand(cmi)
         Debug.Assert hr >= 0
         pInvokeSHCmd = (hr >= 0)
       
      End If   ' hr >= NOERROR (QueryContextMenu)
   
      Call DestroyMenu(hMenu)
   
   End If   ' hMenu

End Function

Private Function pFindSHCmd(icm As IContextMenu, sFindVerb As String, _
                            Optional ByVal CMF As eCMF = CMF_EXPLORE) As SHContextMenuCmdIDs
   Dim hr         As Long   ' HRESULT
   Dim hMenu      As Long
   Dim idCmd      As SHContextMenuCmdIDs
   Dim idCmdEnd   As SHContextMenuCmdIDs
   Dim sVerb      As String
   
   If Not (icm Is Nothing) Then
    
      ' Create a new popup menu...
      hMenu = CreatePopupMenu()
      If hMenu Then

         ' Add the item's shell commands to the popup menu.
         hr = icm.QueryContextMenu(hMenu, 0, idCmdFirst, idCmdLast, CMF)
         
         If SUCCEEDED(hr) Then
            
            ' hr == highest offset of the largest command identifier that was assigned, plus one
            idCmdEnd = hr
            
            For idCmd = idCmdFirst To idCmdEnd
            
               Call GetShellCommandString(icm, idCmd, GCS_VERB, sVerb)
            
               If LCase$(sVerb) = LCase$(sFindVerb) Then
               
                  pFindSHCmd = idCmd - idCmdFirst + 1
                  Exit For
               End If
            
            Next idCmd
            
         End If   ' hr >= NOERROR (QueryContextMenu)

         Call DestroyMenu(hMenu)
    
      End If   ' hMenu
   End If   ' idxICM

End Function

' ----------------------------------------------------------------------------------- '
'  Shell context menus debug functions
' ----------------------------------------------------------------------------------- '

Public Function dbgEnumShellCommands(ByVal hOwner As Long, isfParent As IShellFolder, _
                        ByVal cPidls As Long, pidlRel As Long, _
                        Optional ByVal idCmdStart As SHContextMenuCmdIDs = idCmdFirst, _
                        Optional ByVal CMF As eCMF = CMF_EXPLORE) As String
   Dim icm        As IContextMenu
   Dim idxICM     As Long
   
   ' Get a reference to the item's IContextMenu interfaces
   idxICM = pGetIContextMenu(hOwner, isfParent, cPidls, pidlRel, icm)
                        
   If idxICM Then
      dbgEnumShellCommands = dbgpEnumCommands(icm, idCmdStart, CMF)
   End If
End Function

Public Function dbgEnumShellViewCommands(ByVal hOwner As Long, isfParent As IShellFolder, _
                        Optional ByVal idCmdStart As SHContextMenuCmdIDs = idCmdFirst, _
                        Optional ByVal CMF As eCMF = CMF_NORMAL) As String
   Dim isv        As IShellView
   Dim icm        As IContextMenu
   Dim idxICM     As Long
                        
   ' Get a reference to the folder's view object
#If WIN32_IE >= &H500 Then
   If SUCCEEDED(isfParent.CreateViewObject(hOwner, GetRIID(rIID_IShellView), isv)) Then
#Else
   ' # CRASH ??? #
   If False Then
#End If
      ' Get a reference to the view background's IContextMenuX interface
      idxICM = pGetCtxMenuView(isv, icm)
      If idxICM Then
         dbgEnumShellViewCommands = dbgpEnumCommands(icm, idCmdStart, CMF)
      End If
   End If
                        
End Function

Private Function dbgpEnumCommands(icm As IContextMenu, _
                        Optional ByVal idCmdStart As SHContextMenuCmdIDs = idCmdFirst, _
                        Optional ByVal CMF As eCMF = CMF_EXPLORE) As String
   Dim hr         As Long   ' HRESULT
   Dim hMenu      As Long
   Dim idCmd      As SHContextMenuCmdIDs
   Dim idCmdEnd   As SHContextMenuCmdIDs
   Dim sVerb      As String
   Dim sHelpText  As String
   
   If Not (icm Is Nothing) Then
    
      ' Create a new popup menu...
      hMenu = CreatePopupMenu()
      If hMenu Then

         ' Add the item's shell commands to the popup menu.
         hr = icm.QueryContextMenu(hMenu, 0, idCmdFirst, idCmdLast, CMF)
         
         If SUCCEEDED(hr) Then
            
            ' hr == highest offset of the largest command identifier that was assigned, plus one
            idCmdEnd = hr
            
            dbgpEnumCommands = vbCrLf
            
            For idCmd = idCmdStart To idCmdEnd
            
               Call GetShellCommandString(icm, idCmd, GCS_VERB, sVerb)
               Call GetShellCommandString(icm, idCmd, GCS_HELPTEXT, sHelpText)
            
               If LenB(sVerb) Or LenB(sHelpText) Then
                  dbgpEnumCommands = dbgpEnumCommands & _
                     (idCmd - idCmdFirst + 1) & _
                     Space$(5 - Len(CStr(idCmd - idCmdFirst + 1))) & vbTab & _
                     "COMMAND: " & sVerb & Space$(Abs(10 - Len(sVerb))) & vbTab & _
                     "HELPTEXT: " & sHelpText & vbCrLf
               End If
            
            Next idCmd
            
            dbgpEnumCommands = dbgpEnumCommands & vbCrLf
            
         End If   ' hr >= NOERROR (QueryContextMenu)

         Call DestroyMenu(hMenu)
    
      End If   ' hMenu
   End If   ' idxICM

End Function

