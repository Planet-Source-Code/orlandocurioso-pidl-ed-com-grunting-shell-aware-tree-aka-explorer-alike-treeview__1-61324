VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Test clsShellTree"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin TestFolderTree.ucTreeView ucTree 
      Height          =   4095
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   2535
      _extentx        =   4471
      _extenty        =   7223
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "New Folder"
      Height          =   300
      Index           =   8
      Left            =   60
      TabIndex        =   14
      Top             =   4080
      Width           =   1150
   End
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CheckBox chk 
      Caption         =   "Open in My Computer"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
   End
   Begin VB.ListBox lstRoot 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "ContentsPIDL"
      Height          =   300
      Index           =   6
      Left            =   1200
      TabIndex        =   13
      Top             =   3780
      Width           =   1150
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Delete selected (just node)"
      Height          =   300
      Index           =   5
      Left            =   60
      TabIndex        =   8
      Top             =   2580
      Width           =   2295
   End
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   900
      Width           =   2295
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Selected as new root"
      Height          =   300
      Index           =   4
      Left            =   60
      TabIndex        =   11
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Reopen root && selected"
      Height          =   300
      Index           =   3
      Left            =   60
      TabIndex        =   10
      Top             =   3180
      Width           =   2295
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "WalkPIDL"
      Height          =   300
      Index           =   2
      Left            =   60
      TabIndex        =   12
      Top             =   3780
      Width           =   1150
   End
   Begin VB.CheckBox chk 
      Caption         =   "OnlySubDirs"
      Height          =   195
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Open App.Path  in  My  ..."
      Height          =   300
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtPath 
      Height          =   300
      Left            =   60
      TabIndex        =   1
      Top             =   4560
      Width           =   5175
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Reopen selected"
      Height          =   300
      Index           =   0
      Left            =   60
      TabIndex        =   9
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Options"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tree Root"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IncludeItems"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   660
      Width           =   900
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmTest
' Author    : OrlandoCurioso 15.04.2005
' Purpose   : Testcase for clsShellTree. Settings stored in registry.
'
' Caveat    : To ensure memory deallocation (free pidls) don't use ucTreeView's Clear method.
'             Use clsShellTree.Clear instead or terminate clsShellTree.
' Info      :
'             Shell and Common Controls Versions, CSIDL Constants       http://vbnet.mvps.org/index.html?code/browse/csidlversions.htm
'             Using SHGetFolderPath to Find Popular Shell Folders       http://vbnet.mvps.org/index.html?code/browse/csidl.htm
'---------------------------------------------------------------------------------------
Option Explicit

Private Enum eChk
   chkSubDirs = 0
   chkMyDocs
End Enum

'Private Enum eCmd
''
'End Enum

Private Enum eLst
   lstIncludeItems = 0
   lstOptions
End Enum

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
   Private Const LB_SETTABSTOPS As Long = &H192

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long

Private m_cST           As clsShellTree
Private m_eOptions(1)   As Long
Private m_bInProc       As Boolean
Private m_sLastPath     As String
'

Private Sub pAddRoot()
   Dim sPath   As String
   
   m_cST.IncludeItems = m_eOptions(lstIncludeItems)
   m_cST.Options = m_eOptions(lstOptions)
   
   sPath = m_cST.CurrentPath
   With lstRoot(0)
      If .ListIndex = -1 Then .ListIndex = 0
      m_cST.AddRoot .ItemData(.ListIndex), , , OnlySubDirs:=chk(chkSubDirs)
      ucTree.SelectedNode = ucTree.NodeRoot
   End With
   m_cST.CurrentPath = sPath
End Sub

Private Sub chk_Click(Index As Integer)
   Select Case Index
      Case chkSubDirs
         pAddRoot
         lstRoot_LostFocus 0
      Case chkMyDocs
         chk(Index).Caption = "Open in My " & IIf(chk(Index), "Documents", "Computer")
   End Select
End Sub

Private Sub cmdTest_Click(Index As Integer)
   Dim pidl    As Long
   Dim sPath   As String
   On Error GoTo Proc_Error

   With m_cST
   
      Select Case Index

         Case 0   ' reopen selected
            pidl = .CurrentPidl
            If pidl Then
               ucTree.SelectedNode = 0
               .CurrentPidl = pidl
            End If
              
         Case 1   ' open App.Path node
            .ConvertPathInMyComputerToMyDocuments App.Path, sPath
            sPath = IIf(chk(chkMyDocs), sPath, App.Path)
            
'            .CurrentPath = sPath
            pidl = GetPIDLFromPath(Me.hWnd, sPath)
            .CurrentPidl = pidl
            FreePIDL pidl
            
         Case 2   ' WalkPIDL
            If InIDE Then
               Debug.Print dbgWalkPIDL(.CurrentPidl)
            Else
               MsgBox (dbgWalkPIDL(.CurrentPidl))
            End If
         
         Case 3   ' reopen root & selected
            pidl = CopyPIDL(.CurrentPidl)
            .AddRoot lstRoot(0).ItemData(lstRoot(0).ListIndex), , , chk(chkSubDirs)
            .CurrentPidl = pidl
            FreePIDL pidl
      
         Case 4   ' selected as new root
            pidl = CopyPIDL(.CurrentPidl)
            .AddRoot , pidl, , chk(chkSubDirs)
            FreePIDL pidl
            
         Case 5   ' delete selected node
            ucTree.DeleteNode ucTree.SelectedNode
         
         Case 6   ' dbgContentsPidl
            If InIDE Then
               Debug.Print dbgContentsPidl(.CurrentPidl)
            Else
               MsgBox (dbgContentsPidl(.CurrentPidl))
            End If
            
         Case 8  ' New Folder
            .NewFolder

      End Select
   
   End With
   
   On Error Resume Next
   ucTree.SetFocus
   
   Exit Sub

Proc_Error:
   MsgBox "Error: " & Err.Number & ". " & Err.Description, vbOKOnly Or vbCritical, App.Title & ".frmTest: Sub cmdTest_Click"
   If InIDE Then Stop: Resume
End Sub

Private Sub Form_Load()
   Dim sCmd As String
   
   ' survive ... (at least to look at the call stack)
   SetUnhandledExceptionFilter AddressOf MyExceptionHandler
   
   With ucTree
      .Initialize
      
      .BorderStyle = bsFixedSingle
      .ItemHeight = 18
      
      .HasButtons = True
      
#If WIN32_IE >= &H500 Then
      .HasRootLines = True
      .TrackSelect(UseStandardCursor:=True) = True
#Else
      .HasLines = True
      .HasRootLines = True
#End If

'      .BackColor = &HADDBF7
      ' cdColor option needed, if
      ' - specifying non system BackColor
      ' - or show compressed files in other ForeColor (handled by clsShellTree)
      ' - or using Multiselect
'      .DoCustomDraw = cdColor
         
   End With
   
#If WIN32_IE >= &H500 Then
   Dim f As StdFont
   Set f = New StdFont
   f.Name = "Tahoma"
   Set ucTree.Font = f
   Set f = New StdFont
   f.Name = "Tahoma"
   f.Size = 7
   Set txtPath.Font = f
   txtPath.Height = 255
#End If
   
   Set m_cST = New clsShellTree
   With m_cST
      .Initialize ucTree, Me.hWnd
      
      ' done by pAddRoot()
'      .Options = ftContextMenu Or ftSHNotify Or ftCollapseReset
'      .IncludeItems = SHCONTF_FOLDERS Or SHCONTF_INCLUDEHIDDEN
'      .AddRoot CSIDL_DRIVES, , , OnlySubDirs:=True
   End With
   
   pFillLstRoot lstRoot(0)
   lstRoot(0).Height = 255
   
   With lst(lstIncludeItems)
      .AddItem "FOLDERS":                 .ItemData(.NewIndex) = SHCONTF_FOLDERS
      .AddItem "NONFOLDERS":              .ItemData(.NewIndex) = SHCONTF_NONFOLDERS
      .AddItem "INCLUDEHIDDEN":           .ItemData(.NewIndex) = SHCONTF_INCLUDEHIDDEN
      .AddItem "SHAREABLE":               .ItemData(.NewIndex) = SHCONTF_SHAREABLE
      .AddItem "STORAGE":                 .ItemData(.NewIndex) = SHCONTF_STORAGE
      .AddItem "NETPRINTERSRCH":          .ItemData(.NewIndex) = SHCONTF_NETPRINTERSRCH
      .AddItem "INITONFIRSTNEXT":         .ItemData(.NewIndex) = SHCONTF_INIT_ON_FIRST_NEXT
      .Height = 255
   End With
   
   With lst(lstOptions)
      .AddItem "FileOperations":          .ItemData(.NewIndex) = ftFileOperations
      .AddItem "ContextMenu":             .ItemData(.NewIndex) = ftContextMenu
      .AddItem "SHNotify":                .ItemData(.NewIndex) = ftSHNotify
      .AddItem "CollapseReset":           .ItemData(.NewIndex) = ftCollapseReset
      .Height = 255
   End With
   
   m_bInProc = True
   pStateRead
   If m_eOptions(lstIncludeItems) = 0 Then lst(lstIncludeItems).Selected(0) = True
   m_bInProc = False
   
   pAddRoot   ' add root with stored lstRoot/lst/OnlySubDirs options
   
   ' remove quotes from command line
   sCmd = Replace$(Command$(), """", vbNullString)
   
   If LenB(sCmd) = 0 Then
      ' open last stored path
      m_cST.CurrentPath = m_sLastPath
      
   ElseIf m_cST.FolderExists(sCmd) Or m_cST.FileExists(sCmd) Then
      ' open from path in command line (ie left drop on exe):
      ' NoGo for virtual items (windows standard, not VB shortcome)
      m_cST.CurrentPath = sCmd
      
   Else
      m_cST.CurrentPath = m_sLastPath
   End If
   
   On Error Resume Next
   ucTree.SetFocus
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   With lst(0)
      ucTree.Move .Left + .Width + 60, 0, ScaleWidth - .Width - .Left - 60, ScaleHeight - txtPath.Height
   End With
   txtPath.Move 0, ScaleHeight - txtPath.Height, ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   pStateWrite
   Set m_cST = Nothing
   ' Remove the hook
   SetUnhandledExceptionFilter ByVal 0&
End Sub

' =========================================================================================
Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
   Dim i  As Long
   
   m_eOptions(Index) = 0
   
   With lst(Index)
      If .SelCount Then
         For i = 0 To .ListCount - 1
            If .Selected(i) Then
               m_eOptions(Index) = m_eOptions(Index) Or .ItemData(i)
            End If
         Next
      Else
         Select Case Index
           Case lstIncludeItems
               .Selected(0) = True
               m_eOptions(Index) = SHCONTF_FOLDERS
'           Case lstOptions
'               m_eOptions(Index) = 0
         End Select
      End If
   End With
   
   If Not m_bInProc Then
      pAddRoot
   End If
End Sub
Private Sub lst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim i As Long
   With lst(Index)
      If (X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height) Then
         ' MOUSELEAVE
         ReleaseCapture
         ' reset
         .Height = 255
         ' show highest selection
         .ListIndex = -1
         For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
               .TopIndex = i
               Exit For
            End If
         Next
      ElseIf GetCapture() <> .hWnd Then
         ' MOUSEENTER
         SetCapture .hWnd
         ' dropdown
         .ZOrder
         .Height = .ListCount * 255
      End If
   End With
End Sub

Private Sub lstRoot_Click(Index As Integer)
   Dim w As Long, h As Long
   On Error GoTo Proc_Error
   
   With lstRoot(Index)
   
      If Not m_bInProc Then
         If .Height < 500 Then
            ' dropdown
            h = (txtPath.Top - .Top)
            If h > .ListCount * 255 Then h = .ListCount * 255
            w = ScaleWidth - .Left
            If w > 10000 Then w = 10000
            .Height = h:      .Width = w
            .ZOrder
            .ToolTipText = ""
            
         Else
            .Height = 255:    .Width = lst(0).Width
            If .ListIndex <> -1 Then .TopIndex = .ListIndex
            .ToolTipText = .Text
            
            pAddRoot
         End If
      End If
   
   End With
Proc_Error:
End Sub
Private Sub lstRoot_LostFocus(Index As Integer)
   With lstRoot(Index)
      .Height = 255:    .Width = lst(0).Width
      .ToolTipText = .Text
   End With
End Sub

' =========================================================================================
Private Sub txtPath_Change()

   If Not m_bInProc Then
      ' if user input evaluates to existing folder, select in tree
      If (LenB(txtPath) <> 0) Then
         If m_cST.FolderExists(txtPath) Then
            m_cST.CurrentPath = txtPath
         End If
      End If
   End If
End Sub
Private Sub txtPath_GotFocus()
   lstRoot_LostFocus 0
   
   txtPath.SelStart = 0
   txtPath.SelLength = Len(txtPath)
End Sub

Private Sub ucTree_AfterSelectionChange()
   ' set txtPath to selected node's path
   m_bInProc = True
   txtPath = m_cST.DisplayName
   txtPath.SelStart = Len(txtPath)
   m_bInProc = False
End Sub

' =========================================================================================
' =========================================================================================

Private Sub pStateRead()
   Dim sa() As String
   Dim s    As String
   Dim i    As Long
   Dim k    As Long
'   Dim c    As Control
'   On Error Resume Next
   
   ppStateRead chk
   
   s = VBA.GetSetting(App.Title, "State", "Form")
   If LenB(s) Then
      sa = Split(s, ":")
      If UBound(sa) = 4 Then
         Move sa(1), sa(2), sa(3), sa(4)
         WindowState = CInt(sa(0))
      End If
      Erase sa
   End If
   
   m_sLastPath = VBA.GetSetting(App.Title, "State", "LastPath")
   lstRoot(0).ListIndex = CLng(VBA.GetSetting(App.Title, "State", "lstRoot0", "0"))
   
   With lst
      For k = .LBound To .UBound
         s = VBA.GetSetting(App.Title, "State", .Item(k).Name & k)
         If LenB(s) Then
            sa = Split(s, ":")
            For i = 0 To .Item(k).ListCount - 1
               .Item(k).Selected(i) = CInt(sa(i))
            Next
         End If
      Next
   End With
End Sub

Private Sub pStateWrite()
   Dim s    As String
   Dim i    As Long
   Dim k    As Long
'   On Error Resume Next
   
   ppStateWrite chk
   
   i = WindowState:   If i <> vbNormal Then WindowState = vbNormal
   VBA.SaveSetting App.Title, "State", "Form", i & ":" & Left & ":" & Top & ":" & Width & ":" & Height
   VBA.SaveSetting App.Title, "State", "LastPath", m_cST.CurrentPath
   VBA.SaveSetting App.Title, "State", "lstRoot0", CStr(lstRoot(0).ListIndex)
   
   With lst
      For k = .LBound To .UBound
         s = vbNullString
         For i = 0 To .Item(k).ListCount - 1
            s = s & ":" & Abs(.Item(k).Selected(i))
         Next
         s = Mid$(s, 2)
         VBA.SaveSetting App.Title, "State", .Item(k).Name & k, s
      Next
   End With
End Sub

' for Check-/Optionbutton !arrays!
Private Sub ppStateRead(chk As Object, Optional ByVal Section As String = "State")
   Dim sa() As String
   Dim s    As String
   Dim c    As Control
   
   s = VBA.GetSetting(App.Title, Section, chk(chk.LBound).Name)
   If LenB(s) Then
      sa = Split(s, ":")
      For Each c In chk
         c.Value = CInt(sa(c.Index))
      Next
   End If
End Sub
Private Sub ppStateWrite(chk As Object, Optional ByVal Section As String = "State")
   Dim s    As String
   Dim c    As Control
   
   For Each c In chk
      s = s & ":" & c.Value
   Next
   s = Mid$(s, 2)
   VBA.SaveSetting App.Title, Section, chk(chk.LBound).Name, s
End Sub

' =========================================================================================
' http://vbnet.mvps.org/index.html?code/browse/csidl.htm
Private Sub pFillLstRoot(cbo As Control)

   ReDim TabArray(0 To 2) As Long

   TabArray(0) = 146
   TabArray(1) = 176
   TabArray(2) = 213
      
   'Clear existing tabs and set the list tabstop
   Call SendMessage(cbo.hWnd, LB_SETTABSTOPS, 0&, ByVal 0&)
   Call SendMessage(cbo.hWnd, LB_SETTABSTOPS, 4&, TabArray(0))
   
   With cbo
   
      .AddItem "CSIDL_DESKTOP" & vbTab & "virtual" & vbTab & vbTab & ""
      .ItemData(.NewIndex) = CSIDL_DESKTOP
      
      .AddItem "CSIDL_INTERNET" & vbTab & "virtual" & vbTab & vbTab & "Internet Explorer (icon on desktop)"
      .ItemData(.NewIndex) = CSIDL_INTERNET
      
      .AddItem "CSIDL_PROGRAMS" & vbTab & "file" & vbTab & vbTab & "Start Menu\Programs"
      .ItemData(.NewIndex) = CSIDL_PROGRAMS
      
      .AddItem "CSIDL_CONTROLS" & vbTab & "virtual" & vbTab & vbTab & "My Computer\Control Panel"
      .ItemData(.NewIndex) = CSIDL_CONTROLS
      
      .AddItem "CSIDL_PRINTERS" & vbTab & "virtual" & vbTab & vbTab & "My Computer\Printers"
      .ItemData(.NewIndex) = CSIDL_PRINTERS
      
      .AddItem "CSIDL_PERSONAL" & vbTab & "file" & vbTab & vbTab & "My Documents"
      .ItemData(.NewIndex) = CSIDL_PERSONAL
      
      .AddItem "CSIDL_FAVORITES" & vbTab & "file" & vbTab & vbTab & "\Favorites"
      .ItemData(.NewIndex) = CSIDL_FAVORITES
      
      .AddItem "CSIDL_STARTUP" & vbTab & "file" & vbTab & vbTab & "Start Menu\Programs\Startup"
      .ItemData(.NewIndex) = CSIDL_STARTUP
      
      .AddItem "CSIDL_RECENT" & vbTab & "file" & vbTab & vbTab & "\Recent"
      .ItemData(.NewIndex) = CSIDL_RECENT
      
      .AddItem "CSIDL_SENDTO" & vbTab & "file" & vbTab & vbTab & "\SendTo"
      .ItemData(.NewIndex) = CSIDL_SENDTO
      
      .AddItem "CSIDL_BITBUCKET" & vbTab & "virtual" & vbTab & vbTab & "\Recycle Bin"
      .ItemData(.NewIndex) = CSIDL_BITBUCKET
      
      .AddItem "CSIDL_STARTMENU" & vbTab & "file" & vbTab & vbTab & "\Start Menu"
      .ItemData(.NewIndex) = CSIDL_STARTMENU
      
      .AddItem "CSIDL_MYDOCUMENTS" & vbTab & "virtual" & vbTab & vbTab & "\My Documents\"
      .ItemData(.NewIndex) = CSIDL_MYDOCUMENTS
      
      .AddItem "CSIDL_MYMUSIC" & vbTab & "file" & vbTab & vbTab & "\My Documents\My Music"
      .ItemData(.NewIndex) = CSIDL_MYMUSIC
      
      .AddItem "CSIDL_MYVIDEO" & vbTab & "file" & vbTab & vbTab & "\My Documents\My Video"
      .ItemData(.NewIndex) = CSIDL_MYVIDEO
      
      .AddItem "CSIDL_DESKTOPDIRECTORY" & vbTab & "file" & vbTab & vbTab & "\Desktop"
      .ItemData(.NewIndex) = CSIDL_DESKTOPDIRECTORY
      
      .AddItem "CSIDL_DRIVES" & vbTab & "virtual" & vbTab & vbTab & "My Computer"
      .ItemData(.NewIndex) = CSIDL_DRIVES
      
      .AddItem "CSIDL_NETWORK" & vbTab & "virtual" & vbTab & vbTab & "Network Neighborhood"
      .ItemData(.NewIndex) = CSIDL_NETWORK
      
      .AddItem "CSIDL_NETHOOD" & vbTab & "file" & vbTab & vbTab & "\nethood (may dupe My Network Places)"
      .ItemData(.NewIndex) = CSIDL_NETHOOD
      
      .AddItem "CSIDL_FONTS" & vbTab & "virtual" & vbTab & vbTab & "windows\fonts"
      .ItemData(.NewIndex) = CSIDL_FONTS
      
      .AddItem "CSIDL_TEMPLATES" & vbTab & "file" & vbTab & vbTab & "\templates"
      .ItemData(.NewIndex) = CSIDL_TEMPLATES
      
      .AddItem "CSIDL_COMMON_STARTMENU" & vbTab & "file" & vbTab & vbTab & "\Start Menu"
      .ItemData(.NewIndex) = CSIDL_COMMON_STARTMENU
      
      .AddItem "CSIDL_COMMON_PROGRAMS" & vbTab & "file" & vbTab & vbTab & "\Programs"
      .ItemData(.NewIndex) = CSIDL_COMMON_PROGRAMS
      
      .AddItem "CSIDL_COMMON_STARTUP" & vbTab & "file" & vbTab & vbTab & "\Startup"
      .ItemData(.NewIndex) = CSIDL_COMMON_STARTUP
      
      .AddItem "CSIDL_COMMON_DESKTOPDIRECTORY" & vbTab & "file" & vbTab & vbTab & "\Desktop"
      .ItemData(.NewIndex) = CSIDL_COMMON_DESKTOPDIRECTORY
      
      .AddItem "CSIDL_APPDATA" & vbTab & "file" & vbTab & "v4.71" & vbTab & "\Application Data"
      .ItemData(.NewIndex) = CSIDL_APPDATA
      
      .AddItem "CSIDL_PRINTHOOD" & vbTab & "file" & vbTab & vbTab & "\PrintHood"
      .ItemData(.NewIndex) = CSIDL_PRINTHOOD
      
      .AddItem "CSIDL_LOCAL_APPDATA" & vbTab & "file" & vbTab & "v5.0" & vbTab & "\Local Settings\Application Data (non roaming)"
      .ItemData(.NewIndex) = CSIDL_LOCAL_APPDATA
      
      .AddItem "CSIDL_ALTSTARTUP" & vbTab & "file" & vbTab & vbTab & "nonlocalized startup program group"
      .ItemData(.NewIndex) = CSIDL_ALTSTARTUP
      
      .AddItem "CSIDL_COMMON_ALTSTARTUP" & vbTab & "file" & vbTab & "NT only" & vbTab & "nonlocalized Startup group for all users"
      .ItemData(.NewIndex) = CSIDL_COMMON_ALTSTARTUP
      
      .AddItem "CSIDL_COMMON_FAVORITES" & vbTab & "file" & vbTab & "NT only" & vbTab & "all user's favorite items"
      .ItemData(.NewIndex) = CSIDL_COMMON_FAVORITES
      
      .AddItem "CSIDL_INTERNET_CACHE" & vbTab & "file" & vbTab & "v4.72" & vbTab & "temporary Internet files"
      .ItemData(.NewIndex) = CSIDL_INTERNET_CACHE
      
      .AddItem "CSIDL_COOKIES" & vbTab & "file" & vbTab & "NT only" & vbTab & "Internet cookies"
      .ItemData(.NewIndex) = CSIDL_COOKIES
      
      .AddItem "CSIDL_HISTORY" & vbTab & "file" & vbTab & "NT only" & vbTab & "Internet history items"
      .ItemData(.NewIndex) = CSIDL_HISTORY
      
      .AddItem "CSIDL_COMMON_APPDATA" & vbTab & "file" & vbTab & "v5.0" & vbTab & "\Application Data"
      .ItemData(.NewIndex) = CSIDL_COMMON_APPDATA
      
      .AddItem "CSIDL_WINDOWS" & vbTab & "file" & vbTab & "v5.0" & vbTab & "Windows directory or SYSROOT"
      .ItemData(.NewIndex) = CSIDL_WINDOWS
      
      .AddItem "CSIDL_SYSTEM" & vbTab & "file" & vbTab & "v5.0" & vbTab & "GetSystemDirectory()"
      .ItemData(.NewIndex) = CSIDL_SYSTEM
      
      .AddItem "CSIDL_PROGRAM_FILES" & vbTab & "file" & vbTab & "v5.0" & vbTab & "C:\Program Files"
      .ItemData(.NewIndex) = CSIDL_PROGRAM_FILES
      
      .AddItem "CSIDL_MYPICTURES " & vbTab & "file" & vbTab & "v5.0" & vbTab & "\My Documents\My Pictures"
      .ItemData(.NewIndex) = CSIDL_MYPICTURES
      
      .AddItem "CSIDL_PROFILE" & vbTab & "file" & vbTab & "v5.0" & vbTab & "\"
      .ItemData(.NewIndex) = CSIDL_PROFILE
      
      .AddItem "CSIDL_SYSTEMX86" & vbTab & "file" & vbTab & vbTab & "x86 system directory on RISC"
      .ItemData(.NewIndex) = CSIDL_SYSTEMX86
      
      .AddItem "CSIDL_PROGRAM_FILESX86" & vbTab & "file" & vbTab & vbTab & "x86 Program Files folder on RISC"
      .ItemData(.NewIndex) = CSIDL_PROGRAM_FILESX86
      
      .AddItem "CSIDL_PROGRAM_FILES_COMMON" & vbTab & "file" & vbTab & "v5.0" & vbTab & "C:\Program Files\Common"
      .ItemData(.NewIndex) = CSIDL_PROGRAM_FILES_COMMON
      
      .AddItem "CSIDL_PROGRAM_FILES_COMMONX86" & vbTab & "file" & vbTab & vbTab & "x86 Program Files Common folder on RISC"
      .ItemData(.NewIndex) = CSIDL_PROGRAM_FILES_COMMONX86
      
      .AddItem "CSIDL_COMMON_TEMPLATES" & vbTab & "file" & vbTab & vbTab & "\Templates"
      .ItemData(.NewIndex) = CSIDL_COMMON_TEMPLATES
      
      .AddItem "CSIDL_COMMON_DOCUMENTS" & vbTab & "file" & vbTab & vbTab & "\Documents"
      .ItemData(.NewIndex) = CSIDL_COMMON_DOCUMENTS
      
      .AddItem "CSIDL_COMMON_ADMINTOOLS" & vbTab & "file" & vbTab & "v5.0" & vbTab & "\Start Menu\Programs\Administrative Tools"
      .ItemData(.NewIndex) = CSIDL_COMMON_ADMINTOOLS
      
      .AddItem "CSIDL_ADMINTOOLS " & vbTab & "file" & vbTab & "v5.0" & vbTab & "\Start Menu\Programs\Administrative Tools"
      .ItemData(.NewIndex) = CSIDL_ADMINTOOLS
      
      .AddItem "CSIDL_CONNECTIONS" & vbTab & "virtual" & vbTab & vbTab & "Network and dial-up connections folder"
      .ItemData(.NewIndex) = CSIDL_CONNECTIONS
      
      .AddItem "CSIDL_COMMON_MUSIC" & vbTab & "file" & vbTab & vbTab & "My Music folder for all users"
      .ItemData(.NewIndex) = CSIDL_COMMON_MUSIC
      
      .AddItem "CSIDL_COMMON_PICTURES" & vbTab & "file" & vbTab & vbTab & "My Pictures folder for all users"
      .ItemData(.NewIndex) = CSIDL_COMMON_PICTURES
      
      .AddItem "CSIDL_COMMON_VIDEO" & vbTab & "file" & vbTab & vbTab & "My Video folder for all users"
      .ItemData(.NewIndex) = CSIDL_COMMON_VIDEO
      
      .AddItem "CSIDL_RESOURCES" & vbTab & "file" & vbTab & vbTab & "System resource directory"
      .ItemData(.NewIndex) = CSIDL_RESOURCES
      
      .AddItem "CSIDL_RESOURCES_LOCALIZED" & vbTab & "file" & vbTab & vbTab & "Localized resource directory"
      .ItemData(.NewIndex) = CSIDL_RESOURCES_LOCALIZED
      
      .AddItem "CSIDL_COMMON_OEM_LINKS" & vbTab & "file" & vbTab & vbTab & "Links to OEM specific apps for all users"
      .ItemData(.NewIndex) = CSIDL_COMMON_OEM_LINKS
      
      .AddItem "CSIDL_CDBURN_AREA" & vbTab & "file" & vbTab & vbTab & "\Local Settings\Application Data\Microsoft\CD Burning"
      .ItemData(.NewIndex) = CSIDL_CDBURN_AREA
      
      .AddItem "CSIDL_COMPUTERSNEARME" & vbTab & "virtual" & vbTab & vbTab & "Computers Near Me folder"
      .ItemData(.NewIndex) = CSIDL_COMPUTERSNEARME
      
   End With
End Sub
