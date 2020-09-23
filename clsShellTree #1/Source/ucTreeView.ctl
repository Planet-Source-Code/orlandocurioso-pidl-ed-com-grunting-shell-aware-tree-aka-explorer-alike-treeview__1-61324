VERSION 5.00
Begin VB.UserControl ucTreeView 
   BackColor       =   &H80000005&
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1980
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   132
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "ucTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : ucTreeView
' Author    : OrlandoCurioso 22.02.2005
' Credits   : '.. standing on the shoulders of giants ..'
'             Carles P.V. (ucTreeView.ctl 1.3.3)                           http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57047&lngWId=1
'             Paul Caton  (ASM subclassing & MouseLeave/MouseEnter code)   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'             Steve Mahon (vbAccelerator TreeView Control)                 http://www.vbaccelerator.com/home/VB/Code/Controls/TreeView/TreeView_Control/VB6_TreeView_Full_Source_zip_vbalTreeView_ctl.asp
'             Vlad Vissoultchev (OleGuids3.tlb)                            http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=41506&lngWId=1
'             Mike Gainer,Matt Curland,Bill Storage (IOleInPlaceActivate)  VBPJ ???
'             Jean-Edouard Lachand-Robert (Drawing dotted lines)           http://www.codeguru.com/Cpp/G-M/gdi/article.php/c147/
'             Eric Kohl/Wine Project (concept of node geometry)            http://cvs.winehq.com/cvsweb/wine/dlls/comctl32/treeview.c
'             Shog ? (drag image in popup window)                          http://blogs.wdevs.com/shog9/archive/2004/09/22/597.aspx
'
' Purpose   : Virtual OwnerDraw API Treeview (comctl32 >= 4.71)
'
'             Virtual:
'             - NodeText & images callback implemented, children stay in control of comctl.
'             - Since comctl bases ItemRect,HitTest & Scrollbar behaviour on the tree font,
'               a dummy text is supplied for individual NodeFont's or non-default drawing.
'             - For optimal appearance keep tree font size smaller than any NodeFont size.
'             - -> zSubclass_Proc()/TVN_GETDISPINFO
'
'             OwnerDraw:
'             - Draws the complete tree(pTreeOwnerDraw()), depending on DoCustomDraw property.
'             - If drawing departs from node geometry comctl expects, feed comctl with
'               mouse msg's altered by a customized pvTVHitTest function and change
'               dummy text calculation.(ie NodeTextIndent)
'
'             OwnerDraw currently implemented (additional) drawing:
'             - DoCustomDraw property,CustomizeNode method.
'             - NodeForeColor,NodeBackcolor,NodeFont,NodeExpandedImage,NodeItemdata props.
'             - NodeTextIndent prop: adds extra space between NodeImage and NodeText.
'             - TreeColor(tvTreeColors) prop: SelectedColor,HotColor etc.
'             - Mix nodes with/without image in tree (cdMixNoImage)
'             - SelectedImage & ExpandedImage: SelectionStyle prop.
'             - pTreeOwnerDraw: add project-specific code (cdProject).
'
'             Deviations from comctl behaviour:
'             - NodeText drawn vertically centered in ItemRect. (comctl: top aligned)
'             - State images vertically centered in ItemRect.   (comctl: bottom aligned)
'             - ToolTip adjusted for Node Font,Fore & BackColor.
'             - LabelEditBox uses NodeFont.
'
'             Multiple Selection:
'             - MultiSelect,SelectionCount,SelectionNode(Index),NodeSelected(hNode) props.
'             - MultipleSelection & SelectedImage: SelectionStyle prop.
'
'             OLE Drag&Drop:
'             - tvDataFormats enum for ownerdrawn & multiselected nodes.
'             - tvDataOptions enum determines extent of data written.
'             - Use of PropertyBag is flexible for changes in Node_Data UDT.
'             - DragAutomatic: left drag only, disturbes MultiSelect operation.
'             - DragManual: left & right drag (use OLEStartDrag method), MultiSelect OK.
'             - Additional data with OLESetDataInfo, OLEGetDataInfoEx methods.
'             - OLEDragInsertStyle: disAutomatic switches styles, based on cursor position.
'
'             Misc:
'             - Unicode support.
'             - AutoFont prop, AdjustFont event: if NodeFonts are unknown at design-time.
'             - NoDataText event: display text on empty treeview.
'             - NodeDblClick event permits preventing node expansion on doubleclick.
'             - External (State) imagelists: hImagelist,ImageListCount,NodeStateImage props.
'             - Expand method: EventOnly option for Load on Demand clients.
'             - Collapse method: RemoveChildren option for Load on Demand clients
'                                and to reset NodeExpandedOnce prop.
'
' Requires  : mIOIPAOTreeView.bas & OleGuids3.tlb (in IDE only)
'
' Caveat    : LockWindowUpdate in zSubclass_Proc()/TVN_ITEMEXPANDING & TVN_ITEMEXPANDED
'             + Collapse method. Outcomment,if container flickers.
'
' Bugs      :
'             OLESetData event doesn't fire ???
'             Artifacts of DragImage (-> UserControl_OLEDragOver())
'             Crossprocess dragging: DragImage clips target NodeRect,when redrawing.
'             IDE End button crash (IOLEInPlaceActiveObject) -> http://groups.google.de/groups?q=g:thl1549974307d&dq=&hl=de&lr=&selm=%23kM1a7sPAHA.231%40cppssbbsa02.microsoft.com
'
' ToDo      :
'             HowTo update a single ItemRect ? (-> zSubclass_Proc()/TVN_SETDISPINFO, pRefreshNodeRects())
'             Buffer DC of complete client area for background picture/scrolling.
' Marked    : # ... # for options & special attent
'---------------------------------------------------------------------------------------
' Update#1  :
'
' Caveat    : While dragging any popup of a (system) msgbox causes a system hangup!
'             (ie with OLEDragAutoExpand trying to expand empty floppy disk)
'
' Additions :
'             #Const FLDBR: if used as Folder Browser with clsShellTree.
'             NodeOverlayImage prop(use ImageList_SetOverlayImage API to set overlays in Iml).
'             NodeKey prop read/write for re-keying nodes.Take care!
'             NodeIndex prop: access internal idxNode.
'             NodeXXX props which are stored internal, optional access faster with idxNode.
'             BeforeCollapse,AfterCollapse events: needed for TVE_COLLAPSERESET (-> Collapse())
'             BeforeDelete event (not raised with Clear method or on termination).
'             SortChildrenCB method: sorting with callback.
'             TrackSelect prop: UseStandardCursor option.
'             EnsureVisible method: NoScrollRight option.
'             Expand method: Escape key breaks ExpandChildren loop.
'             DragImageBmp prop: returns bitmap of Drag image.
'
' Removed   : Collape event.
'
' Modified  :
' MOD1      : Calmer dragging,if no-drop nodes are hilit as well.
' MOD2      : BeforeLabelEdit event: EditString if only a part of NodeText may be edited.
' MOD3      : OLEGetDropInfo as function.Without valid DropNode, returns false.
'             Use OLEGetDropInfo() to detect dragging mode.
' MOD4      : OLEStartDrag as function indicating success/failure.
' MOD5      : TVHT_ONITEM includes TVHT_ONITEMTEXTINDENT (CUSTDRAW = 1).
' MOD6      : Imagelists created with actual colour depth of screen.
' MOD7      : Drag image without hOwner.
' MOD8      : Drag image solely moved by timer.
' MOD9      : Autoscroll horizontally instead of EnsureVisible in UserControl_OLEDragOver.
'
' BUGFIX1   : m_colSelected in pSelectedNodeChanged (enmity)
' BUGFIX2   : After unsucessful drag operation dragging same node fails then suceeds.
' BUGFIX3   : If after BeforeExpand event node has no children, then selection can't
'             be changed until another button is hit.
' BUGFIX4   : DragAutoExpand enhanced and adapted for Load on Demand clients.
' BUGFIX5   : TrackSelect: missing underlined font for nodes using tree font.
' BUGFIX6   : Implemented drawing of ghosted icons.
' BUGFIX7   : Ensure a OLEDragOver event with State = vbEnter.
' BUGFIX8   : Resource leak: release m_HDC.
'
'---------------------------------------------------------------------------------------
' Update#2  :
'
' Caveat    : Drag image supported only on >=Win2K (SetLayeredWindowAttributes)
'             Degrades gracefully for <Win2K.
'
' Modified  : Drag image uses extra window. Complete DDIMG section reworked.
'             See DDIMG section comments (CLIPBRDWNDCLASS)!
'---------------------------------------------------------------------------------------

Option Explicit

' Project conditional compilation constants
'#Const UNICODE = 1

' Private conditional compilation constants
'                         basic control (compiled No Optimizations)               88KB
#Const CUSTDRAW = 1     ' OwnerDraw                                              +24KB
#Const MULSEL = 1       ' Multiple Selection       (requires CUSTDRAW)           + 8KB
#Const AUTOFNT = 1      ' AutoFont                 (requires CUSTDRAW)           + 4KB
#Const OLEDD = 1        ' OLE Drag & Drop                                        +16KB
#Const DDIMG = 0        ' Drag & Drop Image        (requires OLEDD + CUSTDRAW)
#Const FLDBR = 1        ' Folder Browser support

#Const DRAW_DBG = 0     ' debug OwnerDraw code
#Const MULSEL_DBG = 0   ' debug MultiSelect code
#Const FNT_DBG = 0      ' debug Font handling code
#Const HEXORG = 0       ' original Subclass_Start

'//

Private Type RECT2
   X1 As Long
   Y1 As Long
   X2 As Long
   Y2 As Long
End Type

Private Type POINTAPI
   X  As Long
   Y  As Long
End Type

Private Type SIZEAPI
   cX As Long
   cY As Long
End Type

'// Window styles & msg's

Private Const GWL_STYLE              As Long = (-16)
Private Const WS_TABSTOP             As Long = &H10000
'Private Const WS_BORDER              As Long = &H800000
Private Const WS_CHILD               As Long = &H40000000
'Private Const WS_HSCROLL             As Long = &H100000
'Private Const WS_VSCROLL             As Long = &H200000

Private Const GWL_EXSTYLE            As Long = (-20)
Private Const WS_EX_CLIENTEDGE       As Long = &H200

Private Const WM_SIZE                As Long = &H5
Private Const WM_SETFOCUS            As Long = &H7
Private Const WM_KILLFOCUS           As Long = &H8
Private Const WM_SETREDRAW           As Long = &HB
Private Const WM_SETTEXT             As Long = &HC
Private Const WM_PAINT               As Long = &HF
Private Const WM_ERASEBKGND          As Long = &H14
Private Const WM_SETCURSOR           As Long = &H20
Private Const WM_SETFONT             As Long = &H30
Private Const WM_GETFONT             As Long = &H31
Private Const WM_MOUSEACTIVATE       As Long = &H21
Private Const WM_NOTIFY              As Long = &H4E
Private Const WM_KEYDOWN             As Long = &H100
Private Const WM_KEYUP               As Long = &H101
Private Const WM_CHAR                As Long = &H102
Private Const WM_TIMER               As Long = &H113
Private Const WM_HSCROLL             As Long = &H114
Private Const WM_VSCROLL             As Long = &H115
Private Const WM_MOUSEMOVE           As Long = &H200
Private Const WM_LBUTTONDOWN         As Long = &H201
Private Const WM_LBUTTONUP           As Long = &H202
Private Const WM_LBUTTONDBLCLK       As Long = &H203
Private Const WM_RBUTTONDOWN         As Long = &H204
Private Const WM_RBUTTONUP           As Long = &H205
Private Const WM_RBUTTONDBLCLK       As Long = &H206
Private Const WM_MBUTTONDOWN         As Long = &H207
Private Const WM_MBUTTONUP           As Long = &H208
Private Const WM_MBUTTONDBLCLK       As Long = &H209
Private Const WM_CAPTURECHANGED      As Long = &H215

#If FLDBR Then
Private Const WM_INITMENUPOPUP       As Long = &H117
Private Const WM_DRAWITEM            As Long = &H2B
Private Const WM_MEASUREITEM         As Long = &H2C
Private Const WM_MENUCHAR            As Long = &H120
#End If


'// Common Control

Private Type NMHDR
   hwndFrom As Long
   idfrom   As Long
   code     As Long
End Type

Private Type NMCUSTOMDRAW
   hdr         As NMHDR
   dwDrawStage As Long
   hdc         As Long
   rc          As RECT2
   dwItemSpec  As Long
   uItemState  As Long
   lItemlParam As Long
End Type

Private Type tagINITCOMMONCONTROLSEX
   dwSize   As Long
   dwICC    As Long
End Type
Private Declare Function InitCommonControlsEx Lib "COMCTL32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
Private Declare Sub InitCommonControls Lib "COMCTL32" ()
   Private Const ICC_TREEVIEW_CLASSES = &H2&                      ' treeview, tooltips

Private Const COMCTL32_VERSION       As Long = 5

' Common control shared messages
Private Const CCM_FIRST              As Long = &H2000
'Private Const CCM_SETBKCOLOR         As Long = (CCM_FIRST + 1)    ' lParam = bkColor
'Private Const CCM_SETCOLORSCHEME     As Long = (CCM_FIRST + 2)    ' lParam = COLORSCHEME struct ptr
'Private Const CCM_GETCOLORSCHEME     As Long = (CCM_FIRST + 3)    ' lParam = COLORSCHEME struct ptr
'Private Const CCM_GETDROPTARGET      As Long = (CCM_FIRST + 4)
Private Const CCM_SETUNICODEFORMAT   As Long = (CCM_FIRST + 5)
Private Const CCM_GETUNICODEFORMAT   As Long = (CCM_FIRST + 6)
Private Const CCM_SETVERSION         As Long = (CCM_FIRST + 7)
Private Const CCM_GETVERSION         As Long = (CCM_FIRST + 8)
'Private Const CCM_SETNOTIFYWINDOW    As Long = (CCM_FIRST + 9)    ' wParam = hwndParent
'Private Const CCM_SETWINDOWTHEME     As Long = (CCM_FIRST + 11)
'Private Const CCM_DPISCALE           As Long = (CCM_FIRST + 12)

' Common notification codes (WM_NOTIFY)
Private Const H_MAX                  As Long = &HFFFF + 1
Private Const NM_FIRST               As Long = H_MAX
'Private Const NM_OUTOFMEMORY         As Long = (NM_FIRST - 1)
Private Const NM_CLICK               As Long = (NM_FIRST - 2)
'Private Const NM_DBLCLK              As Long = (NM_FIRST - 3)
'Private Const NM_RETURN              As Long = (NM_FIRST - 4)
'Private Const NM_RCLICK              As Long = (NM_FIRST - 5)
'Private Const NM_RDBLCLK             As Long = (NM_FIRST - 6)
Private Const NM_SETFOCUS            As Long = (NM_FIRST - 7)
Private Const NM_KILLFOCUS           As Long = (NM_FIRST - 8)
Private Const NM_CUSTOMDRAW          As Long = (NM_FIRST - 12)
'Private Const NM_HOVER               As Long = (NM_FIRST - 13)
'Private Const NM_NCHITTEST           As Long = (NM_FIRST - 14)
'Private Const NM_KEYDOWN             As Long = (NM_FIRST - 15)
'Private Const NM_RELEASEDCAPTURE     As Long = (NM_FIRST - 16)
'Private Const NM_SETCURSOR           As Long = (NM_FIRST - 17)
'Private Const NM_CHAR                As Long = (NM_FIRST - 18)
'Private Const NM_TOOLTIPSCREATED     As Long = (NM_FIRST - 19)

' CustomDraw paint stages
Private Const CDDS_PREPAINT = &H1&
'Private Const CDDS_POSTPAINT = &H2&
'Private Const CDDS_PREERASE = &H3&
'Private Const CDDS_POSTERASE = &H4&
Private Const CDDS_ITEM = &H10000
'Private Const CDDS_SUBITEM = &H20000
Private Const CDDS_ITEMPREPAINT = CDDS_ITEM Or CDDS_PREPAINT
'Private Const CDDS_ITEMPOSTPAINT = CDDS_ITEM Or CDDS_POSTPAINT
'Private Const CDDS_ITEMPREERASE = CDDS_ITEM Or CDDS_PREERASE
'Private Const CDDS_ITEMPOSTERASE = CDDS_ITEM Or CDDS_POSTERASE
' CustomDraw Item states.
Private Const CDIS_SELECTED = &H1&
'Private Const CDIS_GRAYED = &H2&
'Private Const CDIS_DISABLED = &H4&
'Private Const CDIS_CHECKED = &H8&
Private Const CDIS_FOCUS = &H10&
'Private Const CDIS_DEFAULT = &H20&
Private Const CDIS_HOT = &H40&
'Private Const CDIS_MARKED = &H80&
'Private Const CDIS_INDETERMINATE = &H100&
' CustomDraw return values.
Private Const CDRF_DODEFAULT = &H0&
'Private Const CDRF_NEWFONT = &H2&
Private Const CDRF_SKIPDEFAULT = &H4&
'Private Const TVCDRF_NOIMAGES = &H10000         '  valid on CDRF_NOTIFYITEMPREPAINT: don't draw images

'Private Const CDRF_NOTIFYPOSTPAINT = &H10&
Private Const CDRF_NOTIFYITEMDRAW = &H20&
'Private Const CDRF_NOTIFYPOSTERASE = &H40&      ' not implemented by treeview
'Private Const CDRF_NOTIFYITEMERASE = &H80&
'Private Const CDRF_NOTIFYSUBITEMDRAW = &H20&



'// TreeView

Private Const WC_TREEVIEW  As String = "SysTreeView32"

Private Type TVITEM
   mask           As Long
   hItem          As Long
   State          As Long
   stateMask      As Long
   pszText        As Long
   cchTextMax     As Long
   iImage         As Long
   iSelectedImage As Long
   cChildren      As Long
   lParam         As Long
End Type

Private Type TVITEMEX
   mask           As Long
   hItem          As Long
   State          As Long
   stateMask      As Long
   pszText        As Long
   cchTextMax     As Long
   iImage         As Long
   iSelectedImage As Long
   cChildren      As Long
   lParam         As Long
   iIntegral      As Long
End Type

Private Type TVINSERTSTRUCT
   hParent      As Long
   hInsertAfter As Long
   Item         As TVITEMEX   ' >= 4.71
End Type

Private Type TVSORTCB
   hParent     As Long
   lpfnCompare As Long
   lParam      As Long
End Type

Private Type NMTREEVIEW
   hdr     As NMHDR
   action  As Long
   itemOld As TVITEM
   itemNew As TVITEM
   ptDrag  As POINTAPI
End Type

Private Type NMTVDISPINFO
   hdr  As NMHDR
   Item As TVITEM
End Type

Private Type TVHITTESTINFO
   pt    As POINTAPI
   flags As Long
   hItem As Long
End Type

Private Type NMTVCUSTOMDRAW
   NMCD        As NMCUSTOMDRAW
   clrText     As Long
   clrTextBk   As Long
   iLevel      As Long         ' IE >= 4.0
End Type

Private Const TVI_ROOT               As Long = &HFFFF0000
Private Const TVI_FIRST              As Long = &HFFFF0001
Private Const TVI_LAST               As Long = &HFFFF0002
Private Const TVI_SORT               As Long = &HFFFF0003

Private Const TVE_COLLAPSE           As Long = &H1
Private Const TVE_EXPAND             As Long = &H2
'Private Const TVE_TOGGLE             As Long = &H3
Private Const TVE_EXPANDPARTIAL      As Long = &H4000    ' >= 4.70
Private Const TVE_COLLAPSERESET      As Long = &H8000

Private Const TVGN_ROOT              As Long = &H0
Private Const TVGN_NEXT              As Long = &H1
Private Const TVGN_PREVIOUS          As Long = &H2
Private Const TVGN_PARENT            As Long = &H3
Private Const TVGN_CHILD             As Long = &H4
Private Const TVGN_DROPHILITE        As Long = &H8
Private Const TVGN_CARET             As Long = &H9

Private Const TVGN_FIRSTVISIBLE      As Long = &H5
Private Const TVGN_NEXTVISIBLE       As Long = &H6
Private Const TVGN_PREVIOUSVISIBLE   As Long = &H7
Private Const TVGN_LASTVISIBLE       As Long = &HA        ' >= 4.71

Private Const TVIF_TEXT              As Long = &H1
Private Const TVIF_IMAGE             As Long = &H2
Private Const TVIF_PARAM             As Long = &H4
Private Const TVIF_STATE             As Long = &H8
Private Const TVIF_HANDLE            As Long = &H10
Private Const TVIF_SELECTEDIMAGE     As Long = &H20
Private Const TVIF_CHILDREN          As Long = &H40
'Private Const TVIF_INTEGRAL          As Long = &H80      ' >= 4.71
Private Const TVIF_ALL               As Long = TVIF_TEXT Or TVIF_IMAGE Or TVIF_PARAM Or TVIF_SELECTEDIMAGE

'Private Const TVIS_SELECTED          As Long = &H2
Private Const TVIS_CUT               As Long = &H4
Private Const TVIS_DROPHILITED       As Long = &H8
Private Const TVIS_BOLD              As Long = &H10
Private Const TVIS_EXPANDED          As Long = &H20
Private Const TVIS_EXPANDEDONCE      As Long = &H40
'Private Const TVIS_EXPANDPARTIAL     As Long = &H80      ' >= 4.70

Private Const TVIS_OVERLAYMASK       As Long = &HF00
Private Const TVIS_STATEIMAGEMASK    As Long = &HF000

Private Const TV_FIRST               As Long = &H1100

#If UNICODE Then
Private Const TVM_INSERTITEM         As Long = (TV_FIRST + 50)
Private Const TVM_GETITEM            As Long = (TV_FIRST + 62)
Private Const TVM_SETITEM            As Long = (TV_FIRST + 63)
'Private Const TVM_GETISEARCHSTRING   As Long = (TV_FIRST + 64)
Private Const TVM_EDITLABEL          As Long = (TV_FIRST + 65)
#Else
Private Const TVM_INSERTITEM         As Long = (TV_FIRST + 0)
Private Const TVM_GETITEM            As Long = (TV_FIRST + 12)
Private Const TVM_SETITEM            As Long = (TV_FIRST + 13)
'Private Const TVM_GETISEARCHSTRING   As Long = (TV_FIRST + 23)
Private Const TVM_EDITLABEL          As Long = (TV_FIRST + 14)
#End If

Private Const TVM_DELETEITEM         As Long = (TV_FIRST + 1)
Private Const TVM_EXPAND             As Long = (TV_FIRST + 2)
Private Const TVM_GETITEMRECT        As Long = (TV_FIRST + 4)
Private Const TVM_GETCOUNT           As Long = (TV_FIRST + 5)
Private Const TVM_GETINDENT          As Long = (TV_FIRST + 6)
Private Const TVM_SETINDENT          As Long = (TV_FIRST + 7)
Private Const TVM_GETIMAGELIST       As Long = (TV_FIRST + 8)
Private Const TVM_SETIMAGELIST       As Long = (TV_FIRST + 9)
Private Const TVM_GETNEXTITEM        As Long = (TV_FIRST + 10)
Private Const TVM_SELECTITEM         As Long = (TV_FIRST + 11)
Private Const TVM_GETEDITCONTROL     As Long = (TV_FIRST + 15)
Private Const TVM_GETVISIBLECOUNT    As Long = (TV_FIRST + 16)
Private Const TVM_HITTEST            As Long = (TV_FIRST + 17)
'Private Const TVM_CREATEDRAGIMAGE    As Long = (TV_FIRST + 18)
Private Const TVM_SORTCHILDREN       As Long = (TV_FIRST + 19)
Private Const TVM_ENSUREVISIBLE      As Long = (TV_FIRST + 20)
Private Const TVM_SORTCHILDRENCB     As Long = (TV_FIRST + 21)
Private Const TVM_ENDEDITLABELNOW    As Long = (TV_FIRST + 22)
'Private Const TVM_SETTOOLTIPS        As Long = (TV_FIRST + 24)
'Private Const TVM_GETTOOLTIPS        As Long = (TV_FIRST + 25)
Private Const TVM_SETINSERTMARK      As Long = (TV_FIRST + 26)
Private Const TVM_SETITEMHEIGHT      As Long = (TV_FIRST + 27)
Private Const TVM_GETITEMHEIGHT      As Long = (TV_FIRST + 28)
Private Const TVM_SETBKCOLOR         As Long = (TV_FIRST + 29)
Private Const TVM_SETTEXTCOLOR       As Long = (TV_FIRST + 30)
Private Const TVM_GETBKCOLOR         As Long = (TV_FIRST + 31)
Private Const TVM_GETTEXTCOLOR       As Long = (TV_FIRST + 32)
'Private Const TVM_SETSCROLLTIME      As Long = (TV_FIRST + 33)
'Private Const TVM_GETSCROLLTIME      As Long = (TV_FIRST + 34)
'Private Const TVM_SETBORDER          As Long = (TV_FIRST + 35)   ' ???
'Private Const TVM_GETBORDER          As Long = (TV_FIRST + 36)   ' ???
Private Const TVM_SETINSERTMARKCOLOR As Long = (TV_FIRST + 37)
Private Const TVM_GETINSERTMARKCOLOR As Long = (TV_FIRST + 38)
'Private Const TVM_GETITEMSTATE       As Long = (TV_FIRST + 39)
Private Const TVM_SETLINECOLOR       As Long = (TV_FIRST + 40)
Private Const TVM_GETLINECOLOR       As Long = (TV_FIRST + 41)
'Private Const TVM_SETUNICODEFORMAT   As Long = CCM_SETUNICODEFORMAT
'Private Const TVM_GETUNICODEFORMAT   As Long = CCM_GETUNICODEFORMAT

Private Const TVN_FIRST              As Long = -400

#If UNICODE Then
Private Const TVN_SELCHANGING        As Long = (TVN_FIRST - 50)
Private Const TVN_SELCHANGED         As Long = (TVN_FIRST - 51)
Private Const TVN_GETDISPINFO        As Long = (TVN_FIRST - 52)
Private Const TVN_SETDISPINFO        As Long = (TVN_FIRST - 53)
Private Const TVN_ITEMEXPANDING      As Long = (TVN_FIRST - 54)
Private Const TVN_ITEMEXPANDED       As Long = (TVN_FIRST - 55)
Private Const TVN_BEGINDRAG          As Long = (TVN_FIRST - 56)
'Private Const TVN_BEGINRDRAG         As Long = (TVN_FIRST - 57)
Private Const TVN_DELETEITEM         As Long = (TVN_FIRST - 58)
Private Const TVN_BEGINLABELEDIT     As Long = (TVN_FIRST - 59)
Private Const TVN_ENDLABELEDIT       As Long = (TVN_FIRST - 60)
'Private Const TVN_GETINFOTIP         As Long = (TVN_FIRST - 14)
#Else
Private Const TVN_SELCHANGING        As Long = (TVN_FIRST - 1)
Private Const TVN_SELCHANGED         As Long = (TVN_FIRST - 2)
Private Const TVN_GETDISPINFO        As Long = (TVN_FIRST - 3)
Private Const TVN_SETDISPINFO        As Long = (TVN_FIRST - 4)
Private Const TVN_ITEMEXPANDING      As Long = (TVN_FIRST - 5)
Private Const TVN_ITEMEXPANDED       As Long = (TVN_FIRST - 6)
Private Const TVN_BEGINDRAG          As Long = (TVN_FIRST - 7)
'Private Const TVN_BEGINRDRAG         As Long = (TVN_FIRST - 8)
Private Const TVN_DELETEITEM         As Long = (TVN_FIRST - 9)
Private Const TVN_BEGINLABELEDIT     As Long = (TVN_FIRST - 10)
Private Const TVN_ENDLABELEDIT       As Long = (TVN_FIRST - 11)
'Private Const TVN_GETINFOTIP         As Long = (TVN_FIRST - 13)
#End If

'Private Const TVN_KEYDOWN            As Long = (TVN_FIRST - 12)
'Private Const TVN_SINGLEEXPAND       As Long = (TVN_FIRST - 15)

Private Const TVS_HASBUTTONS         As Long = &H1
Private Const TVS_HASLINES           As Long = &H2
Private Const TVS_LINESATROOT        As Long = &H4
Private Const TVS_EDITLABELS         As Long = &H8
Private Const TVS_DISABLEDRAGDROP    As Long = &H10
Private Const TVS_SHOWSELALWAYS      As Long = &H20
'Private Const TVS_RTLREADING         As Long = &H40
Private Const TVS_NOTOOLTIPS         As Long = &H80
Private Const TVS_CHECKBOXES         As Long = &H100
Private Const TVS_TRACKSELECT        As Long = &H200
Private Const TVS_SINGLEEXPAND       As Long = &H400
Private Const TVS_INFOTIP            As Long = &H800
Private Const TVS_FULLROWSELECT      As Long = &H1000
'Private Const TVS_NOSCROLL           As Long = &H2000
'Private Const TVS_NONEVENHEIGHT      As Long = &H4000
'Private Const TVS_NOHSCROLL          As Long = &H8000

'Private Const TVS_SHAREDIMAGELISTS   As Long = &H0       ' ???
'Private Const TVS_PRIVATEIMAGELISTS  As Long = &H400     ' ???

Private Const TVSIL_NORMAL           As Long = &H0
Private Const TVSIL_STATE            As Long = &H2

' treeview sends TVN_SETDISPINFO,TVN_GETDISPINFO notifications to parent window
Private Const LPSTR_TEXTCALLBACK = -1&     ' TVITEM.pszText pointer
Private Const I_IMAGECALLBACK = -1&        ' TVITEM.iImage & .iSelectedImage
'Private Const I_CHILDRENCALLBACK = -1&     ' TVITEM.cChildren


' // ToolTip

'Private Type NMTTDISPINFO
'    hdr        As NMHDR
'    lpszText   As Long
'    szText(0 To 79) As Byte
'    hInst      As Long
'    uFlags     As Long
'    lParam     As Long
'End Type

'Private Type NMTVGETINFOTIP
'   hdr         As NMHDR
'   pszText     As Long
'   cchTextMax  As Long
'   hItem       As Long
'   lParam      As Long
'End Type

'Private Type TOOLINFO
'   cbSize      As Long
'   uFlags      As TT_Flags
'   hwnd        As Long
'   uId         As Long
'   rc          As RECT2
'   hInst       As Long
'   lpszText    As String
'   lParam      As Long
'End Type

' ToolTip Flags
'Private Enum TT_Flags
'   TTF_IDISHWND = &H1
'   TTF_CENTERTIP = &H2
'   TTF_RTLREADING = &H4
'   TTF_SUBCLASS = &H10
'   TTF_TRACK = &H20
'   TTF_ABSOLUTE = &H80
'   TTF_TRANSPARENT = &H100
'   TTF_DI_SETITEM = &H8000&
'End Enum

' ToolTip Notifications
Private Const TTN_FIRST             As Long = (-520)
Private Const TTN_GETDISPINFOA      As Long = (TTN_FIRST - 0)     ' aka TTN_NEEDTEXT
Private Const TTN_GETDISPINFOW      As Long = (TTN_FIRST - 10)
Private Const TTN_SHOW              As Long = (TTN_FIRST - 1)
'Private Const TTN_POP               As Long = (TTN_FIRST - 2)
'Private Const TTN_LINKCLICK         As Long = (TTN_FIRST - 3)

' ToolTip Messages
Private Const WM_USER               As Long = &H400
'Private Const TTM_ACTIVATE          As Long = (WM_USER + 1)    ' crashed
'Private Const TTM_SETDELAYTIME      As Long = (WM_USER + 3)
'Private Const TTM_SETTOOLINFOA      As Long = (WM_USER + 9)
'Private Const TTM_UPDATETIPTEXTA    As Long = (WM_USER + 12)
'Private Const TTM_GETCURRENTTOOLA   As Long = (WM_USER + 15)
Private Const TTM_SETTIPBKCOLOR     As Long = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR   As Long = (WM_USER + 20)
'Private Const TTM_POP               As Long = (WM_USER + 28)
'Private Const TTM_POPUP             As Long = (WM_USER + 34)   ' WinXP only
'Private Const TTM_SETTOOLINFOW      As Long = (WM_USER + 54)
'Private Const TTM_UPDATETIPTEXTW     As Long = (WM_USER + 57)
'Private Const TTM_GETCURRENTTOOLW   As Long = (WM_USER + 59)

'Private Const TT_BufferSize         As Long = 80&              ' self defined

'// ImageList

Private Const CLR_DEFAULT  As Long = &HFF000000
Private Const CLR_NONE     As Long = &HFFFFFFFF
Private Const ILC_MASK     As Long = &H1
'Private Const ILC_COLORDDB As Long = &HFE

Private Declare Function ImageList_Create Lib "COMCTL32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Add Lib "COMCTL32" (ByVal hImageList As Long, ByVal hBitmap As Long, ByVal hBitmapMask As Long) As Long
Private Declare Function ImageList_AddMasked Lib "COMCTL32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_AddIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_Destroy Lib "COMCTL32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_GetImageCount Lib "COMCTL32" (ByVal hImageList As Long) As Long
'Private Declare Function ImageList_GetIcon Lib "COMCTL32" (ByVal hIml As Long, ByVal i As Long, ByVal diFlags As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cX As Long, cY As Long) As Long
'Private Declare Function ImageList_Replace Lib "COMCTL32" (ByVal hIml As Long, ByVal i As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Boolean
Private Declare Function ImageList_ReplaceIcon Lib "COMCTL32" (ByVal hIml As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_Remove Lib "COMCTL32" (ByVal hIml As Long, ByVal ImgIndex As Long) As Long
Private Declare Function ImageList_Copy Lib "COMCTL32" (ByVal himlDst As Long, ByVal iDst As Long, ByVal himlSrc As Long, ByVal iSrc As Long, ByVal uFlags As Long) As Long
'   Private Const ILCF_MOVE = &H0&
   Private Const ILCF_SWAP = &H1&
Private Declare Function ImageList_Draw Lib "COMCTL32" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_DrawEx Lib "COMCTL32" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
   Private Const ILD_NORMAL = 0&
   Private Const ILD_TRANSPARENT = 1&
'   Private Const ILD_BLEND25 = 2&
   Private Const ILD_SELECTED = 4&
'   Private Const ILD_FOCUS = 4&
'   Private Const ILD_MASK = &H10&
'   Private Const ILD_IMAGE = &H20&
'   Private Const ILD_ROP = &H40&
'   Private Const ILD_OVERLAYMASK = 3840&

'Private Declare Function ImageList_BeginDrag Lib "COMCTL32" (ByVal himlTrack As Long, ByVal iTrack As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long
'Private Declare Sub ImageList_EndDrag Lib "COMCTL32" ()
'Private Declare Function ImageList_DragEnter Lib "COMCTL32" (ByVal hwndLock As Long, ByVal X As Long, ByVal Y As Long) As Long
'Private Declare Function ImageList_DragLeave Lib "COMCTL32" (ByVal hwndLock As Long) As Long
'Private Declare Function ImageList_DragMove Lib "COMCTL32" (ByVal X As Long, ByVal Y As Long) As Long
'Private Declare Function ImageList_DragShowNolock Lib "COMCTL32" (ByVal fShow As Long) As Long
'Private Declare Function ImageList_SetDragCursorImage Lib "COMCTL32" (ByVal himlDrag As Long, ByVal iDrag As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long
'Private Declare Function ImageList_GetDragImage Lib "COMCTL32" (ppt As POINTAPI, pptHotspot As POINTAPI) As Long

'// Mouse leave/enter support

Private Enum TRACKMOUSEEVENT_FLAGS
    [TME_HOVER] = &H1
    [TME_LEAVE] = &H2
    [TME_QUERY] = &H40000000
    [TME_CANCEL] = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize      As Long
    dwFlags     As TRACKMOUSEEVENT_FLAGS
    hwndTrack   As Long
    dwHoverTime As Long
End Type

Private Const WM_MOUSELEAVE As Long = &H2A3

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "COMCTL32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

'// ScrollBar

'Private Type SCROLLINFO
'    cbSize As Long
'    fMask As Long
'    nMin As Long
'    nMax As Long
'    nPage As Long
'    npos As Long
'    nTrackPos As Long
'End Type

'Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
'Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
'Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
'Private Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
'Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal npos As Long, ByVal bRedraw As Long) As Long
'Private Declare Function SetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
'Private Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT2, lprcClip As RECT2, ByVal hrgnUpdate As Long, lprcUpdate As RECT2) As Long
'Private Declare Function ScrollDCNull Lib "user32" Alias "ScrollDC" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As Long, lprcClip As Long, ByVal hrgnUpdate As Long, lprcUpdate As Long) As Long

'Private Const SB_HORZ = 0
'Private Const SB_VERT = 1
'Private Const SB_CTL = 2
'Private Const SB_BOTH = 3

Private Const SB_LINEUP    As Long = 0
Private Const SB_LINELEFT  As Long = 0
Private Const SB_LINEDOWN  As Long = 1
Private Const SB_LINERIGHT As Long = 1
Private Const SB_PAGEUP    As Long = 2
Private Const SB_PAGELEFT  As Long = 2
Private Const SB_PAGEDOWN  As Long = 3
Private Const SB_PAGERIGHT As Long = 3
Private Const SB_TOP       As Long = 6
Private Const SB_LEFT      As Long = 6
Private Const SB_BOTTOM    As Long = 7
Private Const SB_RIGHT     As Long = 7
'Private Const SB_ENDSCROLL As Long = 8

'Private Const SB_THUMBPOSITION = 4
'Private Const SB_THUMBTRACK = 5

'Private Const SIF_RANGE = &H1
'Private Const SIF_PAGE = &H2
'Private Const SIF_POS = &H4
'Private Const SIF_DISABLENOSCROLL = &H8
'Private Const SIF_TRACKPOS = &H10
'Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
    
' // XP Visual Styles

'Private Declare Function IsThemeActive Lib "UxTheme" () As Long
'Private Declare Function IsAppThemed Lib "UxTheme.dll" () As Boolean
'Private Declare Function ActivateWindowTheme Lib "uxtheme.dll" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByVal pszSubAppName As Long = 0, Optional ByVal pszSubIdList As Long = 0) As Long
'Private Declare Function DeactivateWindowTheme Lib "uxtheme.dll" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByRef pszSubAppName As String = "", Optional ByRef pszSubIdList As String = "") As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT2, pClipRect As RECT2) As Long
'Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pBoundingRect As RECT2, pContentRect As RECT2) As Long
'Private Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlag As Long, ByVal dwTextFlags2 As Long, pRect As RECT2) As Long
'Private Declare Function DrawThemeIcon Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT2, ByVal hIml As Long, ByVal iImageIndex As Long) As Long
'Private Declare Function DrawThemeEdge Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pDestRect As RECT2, ByVal uEdge As Long, ByVal uFlags As Long, pContentRect As RECT2) As Long
'Private Declare Function GetThemePartSize Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT2, ByVal eSize As THEME_SIZE, pSize As SIZEAPI) As Long
'Private Enum THEME_SIZE
'    TS_MIN = 0             '// minimum size
'    TS_TRUE = 1            '// size without stretching
'    TS_DRAW = 2            '// size that theme mgr will use to draw part
'End Enum
'Private Enum TMT_SIZINGTYPE
'   TM_ST_TRUESIZE = 0
'   TM_ST_STRETCH = 1
'   TM_ST_TILE = 2
'End Enum
'Private Declare Function GetThemeInt Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, piVal As Long) As Long

Private Enum TM_CLASS_PARTS_TREEVIEW
   TVP_TREEITEM = 1
   TVP_GLYPH
   TVP_BRANCH
End Enum
'Private Enum TM_PART_STATES_TREEITEM
'   TREIS_NORMAL = 1
'   TREIS_HOT
'   TREIS_SELECTED
'   TREIS_DISABLED
'   TREIS_SELECTEDNOTFOCUS
'End Enum
Private Enum TM_PART_STATES_GLYPH
   GLPS_CLOSED = 1
   GLPS_OPENED
End Enum

    
'// Misc
#If UNICODE Then
Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function lstrcpyW Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
#Else
'Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
#End If

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
'Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
'Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT2) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT2) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT2, lpSrc1Rect As RECT2, lpSrc2Rect As RECT2) As Long
Private Declare Function UnionRect Lib "user32" (lpDestRect As RECT2, lpSrc1Rect As RECT2, lpSrc2Rect As RECT2) As Long
Private Declare Function EqualRect Lib "user32" (lpRect1 As RECT2, lpRect2 As RECT2) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal X As Long, ByVal Y As Long) As Long
'Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT2, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT2, ByVal bErase As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
   Private Const SW_SHOW           As Long = 5
'   Private Const SW_HIDE           As Long = 0

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
   Private Const SWP_NOMOVE         As Long = &H2
   Private Const SWP_NOSIZE         As Long = &H1
   Private Const SWP_NOOWNERZORDER  As Long = &H200
   Private Const SWP_NOZORDER       As Long = &H4
   Private Const SWP_FRAMECHANGED   As Long = &H20
   Private Const SWP_NOACTIVATE     As Long = &H10
'   Private Const SWP_NOSENDCHANGING As Long = &H400
'   Private Const SWP_HIDEWINDOW     As Long = &H80
'   Private Const HWND_NOTOPMOST = -2&
'   Private Const HWND_TOP = 0&
'   Private Const HWND_TOPMOST = -1&
'   Private Const HWND_BOTTOM = 1&

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Declare Function GetCapture Lib "user32" () As Long

Private Declare Function timeGetTime Lib "winmm" () As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long


' // GDI + Drawing stuff

Private Enum eDrawTextFormat
'   DT_TOP = &H0
'   DT_LEFT = &H0
'   DT_CENTER = &H1
'   DT_RIGHT = &H2
   DT_VCENTER = &H4
'   DT_BOTTOM = &H8
   DT_WORDBREAK = &H10&
   DT_SINGLELINE = &H20
   DT_EXPANDTABS = &H40
'   DT_TABSTOP = &H80
   DT_NOCLIP = &H100
'   DT_EXTERNALLEADING = &H200
   DT_CALCRECT = &H400
   DT_NOPREFIX = &H800&
'   DT_INTERNAL = &H1000
'   DT_EDITCONTROL = &H2000
'   DT_PATH_ELLIPSIS = &H4000
'   DT_END_ELLIPSIS = &H8000
'   DT_MODIFYSTRING = &H10000
'   DT_RTLREADING = &H20000
'   DT_WORD_ELLIPSIS = &H40000
End Enum
#If UNICODE Then
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT2, ByVal wFormat As eDrawTextFormat) As Long
Private Declare Function GetTextExtentPoint32W Lib "gdi32" (ByVal hdc As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As SIZEAPI) As Long
#Else
Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT2, ByVal wFormat As eDrawTextFormat) As Long
Private Declare Function GetTextExtentPoint32A Lib "gdi32" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZEAPI) As Long
#End If

Private Enum eSysColor
'   COLOR_SCROLLBAR = 0
'   COLOR_DESKTOP = 1
'   COLOR_ACTIVECAPTION = 2
'   COLOR_INACTIVECAPTION = 3
'   COLOR_MENU = 4
   COLOR_WINDOW = 5
'   COLOR_WINDOWFRAME = 6
'   COLOR_MENUTEXT = 7
   COLOR_WINDOWTEXT = 8
'   COLOR_CAPTIONTEXT = 9
'   COLOR_ACTIVEBORDER = 10
'   COLOR_INACTIVEBORDER = 11
'   COLOR_APPWORKSPACE = 12
   COLOR_HIGHLIGHT = 13
   COLOR_HIGHLIGHTTEXT = 14
   COLOR_3DFACE = 15               ' vbButtonFace
'   COLOR_3DSHADOW = 16
   COLOR_GRAYTEXT = 17             ' default LineColor
'   COLOR_BTNTEXT = 18
'   COLOR_INACTIVECAPTIONTEXT = 19
   COLOR_3DHIGHLIGHT = 20
   COLOR_3DDKSHADOW = 21
   COLOR_3DLIGHT = 22
'   COLOR_INFOTEXT = 23
'   COLOR_INFOBK = 24
'  #if(WINVER >= 0x0500)
   COLOR_HOTLIGHT = 26
'   COLOR_GRADIENTACTIVECAPTION = 27
'   COLOR_GRADIENTINACTIVECAPTION = 28
'  #if(WINVER >= 0x0501)
'   COLOR_ALTBTNFACE = 25
'   COLOR_MENUHILIGHT = 29
'   COLOR_MENUBAR = 30
End Enum

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As eSysColor) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As eSysColor) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

Private Type LOGFONT
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
#If UNICODE Then
    lfFaceName As String * 31 'LF_FACESIZE
#Else
    lfFaceName(0 To 31) As Byte
#End If
End Type

'Private Const FW_NORMAL         As Long = 400
Private Const FW_BOLD           As Long = 700

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
'   Private Const BITSPIXEL      As Long = 12    'Number of bits per pixel
   Private Const COLORRES       As Long = 108   'Actual color resolution
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

'Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
   Private Const TRANSPARENT = 1&
   Private Const OPAQUE = 2&
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, pDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
'Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long


'Private Type LOGPEN
'   lopnStyle As Long
'   lopnWidth As POINTAPI
'   lopnColor As Long
'End Type

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'Private Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
   Private Const PS_SOLID = 0
   
'Private Declare Function MoveToEx Lib "GDI32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function MoveToExNull Lib "gdi32.dll" Alias "MoveToEx" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT2) As Long
'Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long


Private Type BITMAP '24 bytes
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type
Private Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long

'Private Declare Function GetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
'Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
   Private Enum EPBRasterOperations
      PATCOPY = &HF00021         ' (DWORD) dest = pattern
      PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
      DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
   End Enum

Private Declare Function RectVisible Lib "gdi32" (ByVal hdc As Long, lpRect As RECT2) As Long
'Private Declare Function PtVisible Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'Private Declare Function GetClipBox Lib "gdi32" (ByVal hdc As Long, lpRect As RECT2) As Long
'Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function IntersectClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type RGNDATAHEADER
   dwSize As Long
   iType As Long
   nCount As Long
   nRgnSize As Long
   rcBound As RECT2
End Type

Private Type RGNDATA
   rdh As RGNDATAHEADER
   Buffer() As RECT2
End Type

Private Enum eRegionComplexity
   NULLREGION = 1
   SIMPLEREGION = 2
   COMPLEXREGION = 3
'   ERROR
End Enum

Private Declare Function GetUpdateRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal fErase As Long) As Long
'Private Declare Function GetUpdateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT2, ByVal bErase As Long) As Long
'Private Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
'Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
'   Private Const RGN_AND = 1
'   Private Const RGN_COPY = 5
'   Private Const RGN_OR = 2
'   Private Const RGN_XOR = 3
'   Private Const RGN_DIFF = 4

' // Drag image Version 2

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
   Private Const HWND_MESSAGE As Long = -3
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private Declare Function RegisterClass Lib "user32.dll" Alias "RegisterClassA" (lpWndClass As WNDCLASS) As Long
Private Declare Function UnregisterClass Lib "user32.dll" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
'Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type WNDCLASS
  style As Long
  lpfnWndProc As Long
  cbClsExtra As Long
  cbWndExtra As Long
  hInstance As Long
  hIcon As Long
  hCursor As Long
  hbrBackground As Long
  lpszMenuName As String
  lpszClassName As String
End Type

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Private Const WS_POPUP           As Long = &H80000000
Private Const WS_DISABLED        As Long = &H8000000

Private Const WS_EX_TOPMOST      As Long = &H8
Private Const WS_EX_TRANSPARENT  As Long = &H20
Private Const WS_EX_TOOLWINDOW   As Long = &H80
Private Const WS_EX_LAYERED      As Long = &H80000

Private Const CS_DBLCLKS         As Long = &H8
Private Const CS_SAVEBITS        As Long = &H800

'========================================================================================
' ucTreeView Declarations
'========================================================================================

'-- Public enums:

Public Enum tvBorderStyleConstants
   [bsNone] = 0
   [bsFixedSingle]
End Enum

Public Enum tvRelationConstants
   [rLast] = 0
   [rFirst]
   [rSort]
   [rNext]
   [rPrevious]
End Enum

Public Enum tvScrollConstants
   [sHome] = 0
   [sPageUp]
   [sUp]
   [sDown]
   [sPageDown]
   [sEnd]
   [sLeft]
   [sPageLeft]
   [sLineLeft]
   [sLineRight]
   [sPageRight]
   [sRight]
End Enum

' for public HitTest function
Public Enum tvHitTestConstants
   TVHT_NOWHERE = &H1&
   TVHT_ONITEMICON = &H2&
   TVHT_ONITEMLABEL = &H4&
   TVHT_ONITEMINDENT = &H8&
   TVHT_ONITEMBUTTON = &H10&
   TVHT_ONITEMRIGHT = &H20&
   TVHT_ONITEMSTATEICON = &H40&
   TVHT_ABOVE = &H100&
   TVHT_BELOW = &H200&
   TVHT_TORIGHT = &H400&
   TVHT_TOLEFT = &H800&
   TVHT_ONITEMTEXTINDENT = &H8000&  ' selfdefined for cdDrawLabel: TextIndent space in Label
#If CUSTDRAW Then
   ' MOD5 : TVHT_ONITEM includes TVHT_ONITEMTEXTINDENT
   TVHT_ONITEM = TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON Or TVHT_ONITEMTEXTINDENT
#Else
   TVHT_ONITEM = TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON
#End If
End Enum

Public Enum tvSelectionFocusConstants
   sfNormal = 0
   sfHideSelection
   sfShowSelectionAlways
End Enum

Public Enum tvImagelistConstants
   ilNormal = 0
   ilState = 1
End Enum

'-- Property variables:
Private m_bCheckBoxes         As Boolean
Private m_bFullRowSelect      As Boolean
Private m_bHasButtons         As Boolean
Private m_bHasLines           As Boolean
Private m_bHasRootLines       As Boolean
Private m_bLabelEdit          As Boolean
Private m_bSingleExpand       As Boolean
Private m_bTrackSelect        As Boolean
Private m_bUseStandardCursor  As Boolean
Private m_bRedraw             As Boolean
Private m_eHideSelection      As tvSelectionFocusConstants

'-- Private constants:

'Private Const MAX_PATH        As Long = 260
Private Const PATH_SEPARATOR  As String = "\"
Private Const SII_SWAPMASK    As Long = 3
Private Const SII_CHECKED     As Long = 2
Private Const SII_UNCHECKED   As Long = 1
Private Const ALLOCATE_SIZE   As Long = 100
Private Const DUMCHAR         As String = "."
Private Const IMG_NONE        As Long = -1  ' no image: used in public properties & methods + NodeData UDT
Private Const SPACE_IL        As Long = 3   ' space between Image and Label

'-- Private types:

Private Type NODE_DATA
    hNode      As Long
    sText      As String
    sKey       As String
    sTag       As String
    idxImg     As Long
    idxSelImg  As Long

#If CUSTDRAW Then
    lForeColor As Long
    lBackColor As Long
    idxFont    As Long
    lItemData  As Long
    idxExpImg  As Long
    lIndent    As Long
    ptFont     As SIZEAPI     ' .cx = TextWidth, .cy = TextHeight with NodeFont
    xFont      As Long        ' TextWidth with tree font
#End If

End Type

'-- Private variables:

Private m_bInitialized           As Boolean
Private m_lComctlVersion         As Long
Private m_hModShell32            As Long
Private m_uIPAO                  As IPAOHookStructTreeView
Private m_hUC                    As Long      ' UserControl.hWnd
Private m_hTreeView              As Long
Private m_hEdit                  As Long
Private m_HDC                    As Long
Private m_hImageList(1)          As Long      ' Normal = 0/ State = 1
Private m_lImageListCount(1)     As Long      '    ""
Private m_bExtImagelist(1)       As Boolean   '    ""

Private WithEvents m_oFont       As StdFont   ' external font coupled to Font property
Attribute m_oFont.VB_VarHelpID = -1
Private m_iFont                  As IFont     ' internal font
Private m_hFont                  As Long
Private m_iFontLF                As LOGFONT

Private m_bHoldDeletePostProcess As Boolean
Private m_cKey                   As Collection
Private m_uNodeData()            As NODE_DATA
Private m_lNodeCount             As Long

Private m_bTrack                 As Boolean
Private m_bTrackUser32           As Boolean
Private m_bInControl             As Boolean
Private m_lx                     As Long        ' cursor position in client coords:
Private m_ly                     As Long        ' set by WM_MOUSEMOVE or WM_TIMER

Private m_bCustomDraw            As Boolean     ' == False for CUSTDRAW = 0
Private m_bFocus                 As Boolean
Private m_bInClear               As Boolean
Private m_bInTTip                As Boolean
Private m_bInPaint               As Boolean
Private m_bInSort                As Boolean
Private m_hInExpand              As Long

Private m_lDumCharW              As Long


#If CUSTDRAW Then

Public Enum tvCustomDraw     ' OR'ed
   cdOff = 0
   cdColor = 1
   cdFont = 2
   cdExpandedImage = 4
   cdMixNoImage = 8
   cdLabel = 16
   cdLabelIndent = 32
   cdAll = cdColor Or cdFont Or cdExpandedImage Or cdMixNoImage Or cdLabel Or cdLabelIndent
   cdProject = 256
End Enum

Public Enum tvSelectionStyle
   ' exclusive
   ssPrioritySelected = 0    ' Selected image has priority over expanded (default)
   ssPriorityExpanded = 1    ' Expanded image has priority over selected
   ' OR'ed
#If MULSEL Then
   ' exclusive
   ssImageSelected = 2       ' only real selected node displays it's selected image (default)
   ssImageMultiSelected = 4  ' all multiselected nodes display their selected images
#End If

End Enum

Private Const clrTree = 0&
Private Const clrTreeBK = 1&

Public Enum tvTreeColors
   clrSelected = clrTreeBK + 1
   clrSelectedBK
   clrSelectedNoFocusBk
   clrHot
   clrHilit
   clrHilitBK
End Enum

Private m_lTreeColors(clrTree To clrHilitBK) As Long

Private m_bDrawColor          As Boolean
Private m_bDrawFont           As Boolean
Private m_bDrawExpanded       As Boolean
Private m_bDrawLabel          As Boolean
Private m_bDrawLabelTI        As Boolean
Private m_bMixNoImage         As Boolean
Private m_bDrawProject        As Boolean
Private m_bPriorityExpanded   As Boolean

Private m_FntLF()             As LOGFONT
Private m_hFnt()              As Long
Private m_iFontCount          As Long

Private m_lImageH             As Long
Private m_lImageW             As Long
Private m_hBrDot              As Long

Private Type udtMemDC
   hdc     As Long
   hBmp    As Long
   hBmpOld As Long
   lWidth  As Long
   lHeight As Long
End Type

Private m_tMemDC()            As udtMemDC    ' Buffer DC as array for future additions
Private m_tpOffset            As POINTAPI

#If MULSEL Then
Private m_colSelected         As Collection
Private m_bMultiSelect        As Boolean
Private m_bMultiSelectImage   As Boolean
Private m_bInProc             As Boolean
Private m_hSelectionRoot      As Long
Private m_bKeyDown            As Boolean
#End If

#If AUTOFNT Then
Private Const TESTEXTENT      As String = "Xy" ' "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Private m_bAutoFont           As Boolean
#End If  ' AUTOFNT

#End If 'CUSTDRAW


#If OLEDD Then

Public Enum tvOLEDragConstants
   drgNone = -1
   drgManual = vbOLEDragManual
   drgAutomatic = vbOLEDragAutomatic
End Enum

Public Enum tvOLEDropConstants
   drpNone = vbOLEDropNone
   drpManual = vbOLEDropManual
End Enum

Public Enum tvOLEDragInsertStyleConstants
   disInsertMark = 0
   disDropHilite
   disAutomatic
End Enum

Public Enum tvDataOptions     ' OR'ed, apply to ucTreeView as OLE Drag source
   daMinimal = 1              ' used by drgAutomatic: hNode,Key,Tag of single drag node
#If CUSTDRAW Then
   daCustomData = 2           ' CustomDraw data (as in CustomizeNode method)
#End If
#If MULSEL Then
   daMultipleSelection = 4    ' use with drgManual only, eventually trim selection first
#End If
   daCurrentChildren = 8      ' includes present children
   daChildren = 16            ' raises Expand events to include all children (Load on Demand)
   daInterProcess = 32        ' drag across processes
End Enum

Public Enum tvDataFormats     ' apply to ucTreeView as OLE Drag source & Drop target
   OLE_FORMAT_ID = &HFFFFABCD ' daMinimal,daMultipleSelection
   OLE_FORMAT_ID1             ' + daCustomData
   OLE_FORMAT_ID2             ' ++ daCurrentChildren,daChildren,
   OLE_FORMAT_ID3             ' +++ daInterProcess
End Enum

Private Const T_DRAG_DELAY       As Long = 100
Private Const T_SCROLL_DELAY_X   As Long = 100
Private Const T_SCROLL_DELAY_Y   As Long = 250
Private Const T_EXPAND_DELAY     As Long = 1000
Private Const DELIM_PB           As String = ":"

Private m_eOLEDragMode           As tvOLEDragConstants
Private m_eOLEDropMode           As tvOLEDropConstants
Private m_eOLEDragInsertStyle    As tvOLEDragInsertStyleConstants
Private m_bNodeDropInsertAfter   As Boolean
Private m_bOLEDragAutoInsert     As Boolean
Private m_bOLEDragAutoExpand     As Boolean

Private m_pbDrag                 As PropertyBag
Private m_eButton                As MouseButtonConstants
Private m_hNodeDrag              As Long
Private m_hNodeDrop              As Long
Private m_lDragCounter           As Long        ' ! reused for T_DRAG_DELAY & T_SCROLL_DELAY
Private m_lExpandCounter         As Long
Private m_bStateEnter            As Boolean

#If DDIMG Then
Private m_hClipBrd               As Long        ' CLIPBRDWNDCLASS
Private m_hDDImg                 As Long
Private m_bInDrag                As Boolean
Private m_idxTimer               As Long
Private m_tpDDOffset             As POINTAPI
Private Const TIM_IVALL          As Long = 10&
Private Const DD_WNDCLS          As String = "DragImgWnd"
#End If

#End If  ' OLEDD

'-- Event declarations:

Public Event Click(ByVal Button As Integer)
Public Event DblClick(ByVal Button As Integer)
Public Event NodeClick(ByVal hNode As Long)
Public Event BeforeNodeCheck(ByVal hNode As Long, NewStateImage As Long, OldStateImage As Long)
Public Event AfterNodeCheck(ByVal hNode As Long)
Public Event NodeDblClick(ByVal hNode As Long, ByVal Button As Integer, bCancelExpansion As Boolean)
Public Event BeforeSelectionChange(ByVal hNodeNew As Long, ByVal hNodeOld As Long, Cancel As Integer)
Public Event AfterSelectionChange()
Public Event BeforeExpand(ByVal hNode As Long, ByVal ExpandedOnce As Boolean, Cancel As Integer)
Public Event AfterExpand(ByVal hNode As Long, ByVal ExpandedOnce As Boolean)
Public Event BeforeCollapse(ByVal hNode As Long, Cancel As Integer)
Public Event AfterCollapse(ByVal hNode As Long)
Public Event BeforeDelete(ByVal hNode As Long, ByVal idxNode As Long)
Public Event BeforeLabelEdit(ByVal hNode As Long, Cancel As Integer, EditString As String, ByVal hEdit As Long)   ' MOD2
Public Event AfterLabelEdit(ByVal hNode As Long, Cancel As Integer, NewString As String)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event MouseEnter()
Public Event MouseLeave()
Public Event Resize()

#If CUSTDRAW Then
Public Event NoDataText(ByRef Text As String, ByRef Font As StdFont)
#End If

#If OLEDD Then
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLECompleteDrag(Effect As Long)
#End If

#If AUTOFNT Then
Public Event AdjustFont(ByRef Height As Long, ByRef bHeight As Boolean, ByVal bMax As Boolean)
#End If

'========================================================================================
' Subclasser declarations
'========================================================================================

Private Enum eMsgWhen
    [MSG_AFTER] = 1                                  'Message calls back after the original (previous) WndProc
    [MSG_BEFORE] = 2                                 'Message calls back before the original (previous) WndProc
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE 'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES     As Long = -1          'All messages added or deleted
Private Const CODE_LEN         As Long = 200         'Length of the machine code in bytes
Private Const GWL_WNDPROC      As Long = -4          'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04         As Long = 88          'Table B (before) address patch offset
Private Const PATCH_05         As Long = 93          'Table B (before) entry count patch offset
Private Const PATCH_08         As Long = 132         'Table A (after) address patch offset
Private Const PATCH_09         As Long = 137         'Table A (after) entry count patch offset

Private Type tSubData                                'Subclass data type
    hwnd                       As Long               'Handle of the window being subclassed
    nAddrSub                   As Long               'The address of our new WndProc (allocated memory).
    nAddrOrig                  As Long               'The address of the pre-existing WndProc
    nMsgCntA                   As Long               'Msg after table entry count
    nMsgCntB                   As Long               'Msg before table entry count
    aMsgTblA()                 As Long               'Msg after table array
    aMsgTblB()                 As Long               'Msg Before table array
End Type

Private sc_aSubData()          As tSubData           'Subclass data array
#If HEXORG = 1 Then
Private sc_aBuf(1 To CODE_LEN) As Byte               'Code buffer byte array
#Else
Private sc_aBuf(1 To 50)       As Long               'Code buffer Long array
#End If
Private sc_pCWP                As Long               'Address of the CallWindowsProc
Private sc_pEbMode             As Long               'Address of the EbMode IDE break/stop/running function
Private sc_pSWL                As Long               'Address of the SetWindowsLong function

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

'========================================================================================
' Usercontrol
'========================================================================================

Private Sub UserControl_Initialize()

    '(*) KBID 309366 (http://support.microsoft.com/default.aspx?scid=kb;en-us;309366)
    m_hModShell32 = LoadLibraryA("shell32.dll")
    '-- Initialize common controls
    Call pIsNewComctl32
    
    '-- Initialize font object
    m_oFont_FontChanged vbNullString

    '-- Initialize Node data array ('1' based) and Key collection
    ReDim m_uNodeData(0 To 1)
    Set m_cKey = New Collection

    ' Redraw default
    m_bRedraw = True
    
#If CUSTDRAW Then
    Dim idx As tvTreeColors
    
    ' initialize with system default colors
    For idx = clrSelected To clrHilitBK
        TreeColor(idx) = CLR_NONE
    Next
    
    m_lTreeColors(clrTree) = GetSysColor(COLOR_WINDOWTEXT)
    m_lTreeColors(clrTreeBK) = GetSysColor(COLOR_WINDOW)
    
    ' dim used memory DC's
    ReDim m_tMemDC(0)
#End If

#If OLEDD Then
    ' disabled OLEDrag as default
    m_eOLEDragMode = [drgNone]
#End If

'(*) From vbAccelerator
'    http://www.vbaccelerator.com/home/VB/Code/Libraries/XP_Visual_Styles/Preventing_Crashes_at_Shutdown/article.asp
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Set Font = Ambient.Font
End Sub

Private Sub UserControl_Terminate()

   On Error Resume Next

   Set m_oFont = Nothing
   Set m_iFont = Nothing
   
   If (m_hTreeView) Then
   
      '-- Stop subclassing and destroy all
      Call Subclass_StopAll
      Call mIOIPAOTreeView.TerminateIPAO(m_uIPAO)
      Call pvDestroyImageList
      Call pvDestroyTreeView
   
      '-- Free node data array and key collection
      Erase m_uNodeData()
      Set m_cKey = Nothing
   
      '-- Free library
      Call FreeLibrary(m_hModShell32)
   End If
   
#If CUSTDRAW Then
   Dim idxDC   As Long
   
   For idxDC = LBound(m_tMemDC) To UBound(m_tMemDC)
      pDestroyDC idxDC
   Next
   
   pDestroyDotBrush
   ClearFonts
#End If
#If MULSEL Then
   Set m_colSelected = Nothing
#End If

   On Error GoTo 0
End Sub

' Display usercontrol name in IDE
' http://www.aboutvb.de/khw/artikel/khwshowdisplayname.htm
Private Sub UserControl_AmbientChanged(PropertyName As String)
   Select Case LCase$(PropertyName)
      Case "displayname"
         UserControl_Show
      Case "font"
         Set Font = Ambient.Font
   End Select
End Sub

Private Sub UserControl_Show()
  If Not Ambient.UserMode Then
    lblName.Caption = Ambient.DisplayName
  End If
End Sub

'========================================================================================
' Subclass handler: MUST be the first Public routine in this file.
'                   That includes public properties also.
'========================================================================================

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

   Static bNodeClick As Boolean
   Static bDblClick  As Boolean

   Dim uNMH          As NMHDR
   Dim uNMTV         As NMTREEVIEW
   Dim uNMTVDI       As NMTVDISPINFO
   Dim hNode         As Long
   Dim lfHit         As tvHitTestConstants
   Dim hEdit         As Long
   Dim nCancel       As Integer
   Dim sText         As String
   Dim X             As Long
   Dim Y             As Long
   Dim lS            As Integer  'ShiftConstants
   Dim KeyCode       As Integer
   Dim lFalseHit     As Long
   Dim idxNewState   As Long
   Dim idxOldState   As Long
   Dim tP            As POINTAPI
   Dim idxNode       As Long
   
   Select Case lng_hWnd
      
      Case m_hUC  ' UserControl.hWnd

         ' *** Usercontrol msg's ***
         Select Case uMsg
            
            Case WM_NOTIFY

               Call CopyMemory(uNMH, ByVal lParam, Len(uNMH))

               If (uNMH.hwndFrom <> m_hTreeView) Then Exit Sub

               ' don't process TVN_DELETEITEM, if TreeView is cleared
               If m_bInClear Then Exit Sub

               Select Case uNMH.code
               
#If CUSTDRAW Then
                  Case NM_CUSTOMDRAW
                  
                     If m_bCustomDraw Then
                        lReturn = pTreeOwnerDraw(m_hTreeView, lParam)
                     Else
                        lReturn = CDRF_DODEFAULT
                     End If
#End If

                  Case NM_SETFOCUS
                     
                     Call pvSetIPAO
                     m_bFocus = True
#If MULSEL Then
                     If m_bMultiSelect Then pRedrawSelection
#End If
                  Case NM_KILLFOCUS
                     
                     Debug.Assert Not m_bFocus            ' reset in m_hTreeView/WM_KILLFOCUS
#If MULSEL Then
                     If m_bMultiSelect Then pRedrawSelection
#End If
                  Case NM_CLICK ', NM_RCLICK

                     If pvTVHitTest(hNode, lfHit, X, Y) Then
                        If ((lfHit And (TVHT_ONITEMICON Or TVHT_ONITEMLABEL)) Or m_bFullRowSelect) Then
                           RaiseEvent NodeClick(hNode)
                           ' Debug.Print "NM_CLICK"
                           bNodeClick = True
                           ' ####
'                           RaiseEvent MouseUp((uNMH.code = NM_CLICK) + 2, pvShiftState(), x, y)
                        End If
                        ' Else : handled in WM_XBUTTONUP
                     End If
                     ' NM_RCLICK never fires, if WM_RBUTTONDOWN eaten
                     ' Debug.Assert uNMH.code <> NM_RCLICK
                     
                  Case TVN_DELETEITEM
                     
                     Call CopyMemory(uNMTV, ByVal lParam, Len(uNMTV))
                     RaiseEvent BeforeDelete(uNMTV.itemOld.hItem, uNMTV.itemOld.lParam)
                     m_uNodeData(uNMTV.itemOld.lParam).hNode = 0&
                     ' no cancel: return value is ignored.

                  Case TVN_SELCHANGING

                     Call CopyMemory(uNMTV, ByVal lParam, Len(uNMTV))
                     RaiseEvent BeforeSelectionChange(uNMTV.itemNew.hItem, _
                                                      uNMTV.itemOld.hItem, nCancel)
                     ' return true to prevent selection change
                     lReturn = Abs(nCancel <> 0&)
#If MULSEL Then
                     If m_bKeyDown Then
                        If (nCancel = 0&) And Not m_bInProc Then
                           If uNMTV.itemNew.hItem <> 0& Then
                              ' update selection for arrow keys
                              pSelectedNodeChanged uNMTV.itemNew.hItem, pvShiftState(), _
                                                   bMouseUp:=False
                           End If
                        End If
                        m_bKeyDown = False
                     End If
#End If  ' MULSEL

                  Case TVN_SELCHANGED
#If MULSEL Then
                     ' don't process notification caused by pDeselectNode proc
                     If (Not m_bInProc) Then
                        If (Not bNodeClick) Then
                           Call CopyMemory(uNMTV, ByVal lParam, Len(uNMTV))
                           If uNMTV.itemNew.hItem Then
                              ' node selected by code or key
                              RaiseEvent NodeClick(uNMTV.itemNew.hItem)
                              ' Debug.Print "Node_CLICK by TVN_SELCHANGED", uNMTV.action
                           End If
                        End If
                     End If
#Else
                     If (Not bNodeClick) Then
                        Call CopyMemory(uNMTV, ByVal lParam, Len(uNMTV))
                        RaiseEvent NodeClick(uNMTV.itemNew.hItem)
                     End If
#End If  ' MULSEL
                     bNodeClick = False
                     RaiseEvent AfterSelectionChange

                  Case TVN_ITEMEXPANDING
                     
                     Call CopyMemory(uNMTV, ByVal lParam, Len(uNMTV))
                   ' Debug.Print "TVN_ITEMEXPANDING", IIf(uNMTV.action = TVE_EXPAND, "EXPAND", "COLLAPSE")
                     
                     ' # outcomment if flicker occurs #
                     ' m_hInExpand set in Expand() / Collapse()
                     If (m_hInExpand = 0&) Then
                        LockWindowUpdate m_hTreeView
                     End If
                     
                     With uNMTV
                     
                        Select Case .action
                        
                           Case TVE_EXPAND, TVE_EXPAND Or TVE_EXPANDPARTIAL
                           
                              RaiseEvent BeforeExpand(.itemNew.hItem, _
                                                      CBool(.itemNew.State And TVIS_EXPANDEDONCE), _
                                                      nCancel)
                              ' BUGFIX3: If node still has no children, must collapse it explicitly.
                              If NodeChild(.itemNew.hItem) = 0 Then
                                 Collapse .itemNew.hItem
                              End If
                           
                          Case TVE_COLLAPSE
                          
                              RaiseEvent BeforeCollapse(.itemNew.hItem, nCancel)
                              
                         ' Case TVE_TOGGLE
                         
                         ' Case Else: Debug.Assert False
                              ' TVE_COLLAPSE or TVE_COLLAPSERESET sends no notifs
                        End Select
                        
                        ' # outcomment if flicker occurs #
                        If (m_hInExpand = 0&) Then
                           If (nCancel <> 0&) Then
                              LockWindowUpdate 0
                           ElseIf (.action = TVE_COLLAPSE) Then
                              If Not pTVHasChildren(.itemNew.hItem) Then
                                 ' pushed button,but node has no children -> no TVN_ITEMEXPANDED notif
                                 LockWindowUpdate 0
                              End If
                           End If
                        End If
                        
                     End With
                     ' return True to prevent node expansion/collapsing
                     lReturn = nCancel

                  Case TVN_ITEMEXPANDED
                     
                     Call CopyMemory(uNMTV, ByVal lParam, Len(uNMTV))
                   '  Debug.Print "TVN_ITEMEXPANDED", IIf(uNMTV.action = TVE_EXPAND, "EXPAND", "COLLAPSE")

                     With uNMTV.itemNew
                        Select Case uNMTV.action
                           Case TVE_EXPAND, TVE_EXPAND Or TVE_EXPANDPARTIAL
                              RaiseEvent AfterExpand(.hItem, CBool(.State And TVIS_EXPANDEDONCE))
                           Case TVE_COLLAPSE
                              RaiseEvent AfterCollapse(.hItem)
                         ' Case Else: Debug.Assert False
                           ' TVE_COLLAPSE or TVE_COLLAPSERESET sends no notifs
                        End Select
                     End With
                     
                     ' # outcomment if flicker occurs #
                     If (m_hInExpand = 0&) Then
                        LockWindowUpdate 0&
                     End If
                     
                  Case TVN_GETDISPINFO
                     ' LPSTR_TEXTCALLBACK, I_IMAGECALLBACK, I_CHILDRENCALLBACK used in TVITEM
                     
                     ' LPSTR_TEXTCALLBACK: Based on the first TVN_GETDISPINFO notif
                     ' comctl calculates ItemRect for all following operations.
                     ' Force an update with pRefreshNodeRects().
                     ' While expanding comctl sends for each child a TVN_GETDISPINFO
                     ' notif(TEXT) and a WM_ERASEBKGND msg, unless there was a first call.
                     ' After expansion it asks again 2x TVN_GETDISPINFO (TEXT,TEXT & IMAGE)
                     ' for all visible children in respond to a WM_PAINT msg.
                     ' -> Skip processing TVN_GETDISPINFO in Paint cycle, we supply text.
                     If m_bInPaint And m_bCustomDraw Then Exit Sub
                     
                     Call CopyMemory(uNMTVDI, ByVal lParam, Len(uNMTVDI))
                        
                     If uNMTVDI.hdr.hwndFrom <> m_hTreeView Then Exit Sub
                        
                     Dim xDum As Long
                     Dim sDum As String
                        
                     With uNMTVDI.Item
#If DRAW_DBG Then
                        Debug.Print "TVN_GETDISPINFO", m_uNodeData(.lParam).sText, (.mask And TVIF_TEXT) = TVIF_TEXT, (.mask And TVIF_IMAGE) = TVIF_IMAGE, (.mask And TVIF_SELECTEDIMAGE) = TVIF_SELECTEDIMAGE
#End If
                        If m_bInSort Or m_bInTTip Or Not m_bCustomDraw Then
                           ' Supply real text, if either:
                           ' - ToolTip sends TTN_GETDISPINFO
                           '   -- I found no way to alter tooltip text,once set by comctl.
                           ' - SortChildren() or AddNode(Relation = rSort) called
                           '   -- To use comctl sorting.
                           ' - Drawing done by comctl.
                           pStringToPointer m_uNodeData(.lParam).sText, .cchTextMax, .pszText
                           
                           ' Supply image indices, if drawing done by comctl
                           ' I_IMAGECALLBACK prevented in pvTVAddNode()
                           If (.mask And TVIF_IMAGE) = TVIF_IMAGE Then
                              Debug.Assert (.mask And TVIF_SELECTEDIMAGE) = TVIF_SELECTEDIMAGE
                              Debug.Assert (.mask And TVIF_TEXT) = TVIF_TEXT
                              .iImage = m_uNodeData(.lParam).idxImg
                              .iSelectedImage = m_uNodeData(.lParam).idxSelImg
                              ' copy values back
                              Call CopyMemory(ByVal lParam, uNMTVDI, Len(uNMTVDI))
                           End If

                        Else
                           ' To achieve standard HitTest,Scrollbar behaviour with different NodeFont's
                           ' add dummy text which will produce equal ItemRect's.

                           If m_uNodeData(.lParam).idxFont <> 0& Then
                              ' dummy textwidth needed = textwidth(NodeFont) - textwidth(TreeFont)
                              xDum = m_uNodeData(.lParam).ptFont.cX - m_uNodeData(.lParam).xFont
#If FNT_DBG Then
                              If xDum < 0& Then
                                 Debug.Print "!!! Decrease Tree Font size !!!", m_uNodeData(.lParam).sText
                              End If
#End If
                         ' Else: node uses tree font.Allow for TextIndent only.
                           End If
                           
                           If m_bDrawLabelTI Then
                              xDum = xDum + m_uNodeData(.lParam).lIndent
                           End If
                           
                           ' round up xDum: better appearance with horizontal scrollbar
                           If xDum >= m_lDumCharW Then
                              sDum = String$(1 + xDum \ m_lDumCharW, DUMCHAR)
                           ElseIf xDum > 0& Then
                              sDum = DUMCHAR
                           End If
                           
                           pStringToPointer m_uNodeData(.lParam).sText & sDum, _
                                            .cchTextMax, .pszText
                        End If
                        
                     End With ' uNMTVDI.Item
                     
                     ' unneeded: write to passed pszText pointer
                   ' Call CopyMemory(ByVal lParam, uNMTVDI, Len(uNMTVDI))

                  Case TVN_SETDISPINFO
                     ' called when NodeText changed (ie LabelEdit)
                     Call CopyMemory(uNMTVDI, ByVal lParam, Len(uNMTVDI))
                     If uNMTVDI.hdr.hwndFrom <> m_hTreeView Then Exit Sub
                        
                     With uNMTVDI.Item
                        ' update text and textwidth
                        m_uNodeData(.lParam).sText = pStringFromPointer(.pszText)
#If CUSTDRAW Then
                        If m_uNodeData(.lParam).idxFont <> 0& Then
                           pCalculateRcFont .lParam, m_uNodeData(.lParam).sText
                           m_uNodeData(.lParam).xFont = pGetRcTreeFont(m_uNodeData(.lParam).sText)
                        End If
#End If
                        ' # Problem: comctl must update its ItemRect for this single node   #
                        ' # This updates ItemRect, but horizontal scrollbar is not affected #
                        .pszText = LPSTR_TEXTCALLBACK
                        Call SendMessage(m_hTreeView, TVM_SETITEM, 0&, uNMTVDI.Item)
                        ' # This causes correct scrollbar behaviour,but all nodes refresh   #
'                        pRefreshNodeRects
                        
#If DRAW_DBG Then
                        Debug.Assert (.mask And TVIF_TEXT) = TVIF_TEXT
                        Debug.Assert (.mask And TVIF_IMAGE) = 0&
                        Debug.Print "TVN_SETDISPINFO", m_uNodeData(.lParam).sText
#End If
                     End With
                     
                  Case TVN_BEGINLABELEDIT
                     
                     If (pvShiftState() <> 0&) Then
                        ' prevent LabelEdit if Shift,CTRL keys pressed
                        lReturn = 1&
                        Exit Sub
                     End If
                     
                     Call CopyMemory(uNMTVDI, ByVal lParam, Len(uNMTVDI))
                     hEdit = SendMessageLong(m_hTreeView, TVM_GETEDITCONTROL, 0&, 0&)
                     
                     ' MOD2: sText can be changed ByRef as editable part of NodeText
                     sText = m_uNodeData(uNMTVDI.Item.lParam).sText
                     
                     RaiseEvent BeforeLabelEdit(uNMTVDI.Item.hItem, nCancel, sText, hEdit)
                     ' Return TRUE (<> 0) to cancel label editing.
                     lReturn = Abs(nCancel)

                     If nCancel = 0& Then
                     
                        m_hEdit = hEdit
                        
#If CUSTDRAW Then
                        ' set real text for editbox
#If UNICODE Then
                        SendMessageW hEdit, WM_SETTEXT, 0&, ByVal StrPtr(sText)
#Else
                        SendMessage hEdit, WM_SETTEXT, 0&, ByVal sText
#End If
                        If m_bDrawFont Then
                           Dim idxFont As Long
                           idxFont = m_uNodeData(uNMTVDI.Item.lParam).idxFont
                           If idxFont <> 0& Then
                              ' editbox uses tree font, set node font
                              SendMessageLong hEdit, WM_SETFONT, m_hFnt(idxFont), 1&
                           End If
                        End If
#End If  ' CUSTDRAW
#If MULSEL Then
                        ' redraw & clear selection
                        pSelectedNodeChanged uNMTVDI.Item.hItem, 0&, bMouseUp:=False
#End If
                     End If

                  Case TVN_ENDLABELEDIT

                     Call CopyMemory(uNMTVDI, ByVal lParam, Len(uNMTVDI))
                     With uNMTVDI.Item
                        
                        sText = pStringFromPointer(.pszText)
                        
                        RaiseEvent AfterLabelEdit(.hItem, nCancel, sText)
                        
                        If (.pszText <> 0&) Then   ' or GPF!
                           If Not (nCancel Or GetAsyncKeyState(vbKeyEscape)) Then
                           
                              pStringToPointer sText, .cchTextMax, .pszText
                              
                              ' return TRUE to set the item's label to the edited text.
                              lReturn = 1&
                           End If
                        End If
                     End With
                     m_hEdit = 0&

#If OLEDD Then
'                  Case TVN_BEGINRDRAG
'                     ' TVN_BEGINRDRAG sent even with TVS_DISABLEDRAGDROP !!!
'                     ' TVN_BEGINRDRAG never fires, if WM_RBUTTONDOWN eaten
'                     Debug.Assert False
                     
                  Case TVN_BEGINDRAG
                  
                     ' Preventing inadvertant drag & drop
                     If (timeGetTime() - m_lDragCounter > T_DRAG_DELAY) Then
                        Call CopyMemory(uNMTV, ByVal lParam, Len(uNMTV))
   
                        m_hNodeDrag = uNMTV.itemNew.hItem
   
                        If (m_hNodeDrag) Then
                           m_hNodeDrop = 0&
                           m_lDragCounter = 0&
                           Call SendMessageLong(m_hTreeView, TVM_SELECTITEM, TVGN_CARET, m_hNodeDrag)
                           Call SetCapture(m_hTreeView)
                           Call UserControl.OLEDrag
                        End If
                     End If
#End If  ' OLEDD

'                  Case NM_DBLCLK, NM_RDBLCLK
'                     ' see m_hTreeView/WM_LBUTTONDBLCLK
'

               End Select  ' WM_NOTIFY uNMH.code

            Case WM_SETFOCUS

               Call SetFocus(m_hTreeView)

            Case WM_MOUSEACTIVATE

               Call pvSetIPAO

            Case WM_SIZE
               
               MoveWindow m_hTreeView, 0, 0, ScaleWidth, ScaleHeight, bRepaint:=True
               RaiseEvent Resize

            End Select  ' uMsg Usercontrol


      Case m_hTreeView

         ' *** TreeView msg's ***
         Select Case uMsg
         
#If CUSTDRAW Then
            Case WM_ERASEBKGND
            
               If m_bCustomDraw Then
                  ' Erase background of (complex) update region
                  If m_bRedraw Then
                     pTreeEraseBK
                  End If
                  ' eat msg & return True to indicate no further erasing is required
                  bHandled = True
                  lReturn = 1&
               End If
#If DRAW_DBG Then
               Debug.Print "WM_ERASEBKGND"
#End If

            Case WM_PAINT
            
               ' flag for sucessive NM_CUSTOMDRAW notifs
               m_bInPaint = bBefore
            
               If Not bBefore Then
                  ' Display NoData text for Nodecount = 0
                  If (m_lNodeCount = 0&) Then
                     pDrawNoDataText
                  End If
                  
'                  ' # fake a CDDS_POSTPAINT #
'                  Dim NMTVCD        As NMTVCUSTOMDRAW
'                  NMTVCD.NMCD.dwDrawStage = CDDS_POSTPAINT
'                  pTreeOwnerDraw m_hTreeView, VarPtr(NMTVCD)
                  
               End If
               
#If DRAW_DBG Then
               If bBefore Then Debug.Print "WM_PAINT"
#End If
               
#End If  ' CUSTDRAW
               
#If DDIMG Then
            Case WM_TIMER
'               Debug.Print "WM_TIMER", wParam
               
               Select Case wParam
                  Case m_idxTimer
                     ' no WM_MouseXXX while dragging
                     GetCursorPos tP
                     ScreenToClient m_hTreeView, tP
                     
                     If (tP.X <> m_lx Or tP.Y <> m_ly) Then
                        m_lx = tP.X
                        m_ly = tP.Y
                        pDragImageMove
                     End If
               End Select
#End If

            Case WM_MOUSEMOVE

               If (m_bInControl = False) Then
                  m_bInControl = True
                  Call pvTrackMouseLeave(lng_hWnd)
                  RaiseEvent MouseEnter
               End If
               X = lParam And 65535:   Y = lParam \ 65536
               
               If (X <> m_lx Or Y <> m_ly) Then
                  m_lx = X
                  m_ly = Y
                  RaiseEvent MouseMove(pvButton(uMsg), pvShiftState(), X, Y)
               End If


            Case WM_MOUSELEAVE
               ' sent too, if dragging starts
               m_bInControl = False
               m_lx = -1
               m_ly = -1
               RaiseEvent MouseLeave

            Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN
               
               lS = pvShiftState()
               X = lParam And 65535:   Y = lParam \ 65536
               
               RaiseEvent MouseDown(pvButton(uMsg), lS, X, Y)

               If (uMsg = WM_LBUTTONDOWN) Then
                  
                  If pvTVHitTest(hNode, lfHit, X, Y, lFalseHit) Then
                  
                     If ((lfHit And TVHT_ONITEMSTATEICON) = TVHT_ONITEMSTATEICON) Then
                        idxOldState = NodeStateImage(hNode)
                        idxNewState = SII_SWAPMASK - idxOldState ' only valid for SII_(UN)CHECKED
                        RaiseEvent BeforeNodeCheck(hNode, idxNewState, idxOldState)
                        If (idxNewState <> idxOldState) Then
                           pTVStateImage(hNode) = idxNewState
                           RaiseEvent AfterNodeCheck(hNode)
                        End If
                        bHandled = True
#If CUSTDRAW Then
                     Else
                        If (lFalseHit <> 0&) Then
                           Select Case lfHit
                              Case TVHT_ONITEMRIGHT
                                 ' eat false hit
                                 bHandled = True
                              Case TVHT_ONITEMLABEL
                                 ' create a hit: shift mouse position to left
                                 ' (in msg only, not moving the cursor)
                                 lParam = lParam + lFalseHit
                                 ' Debug.Print "SHIFTED LEFT", lFalseHit
                           End Select
                        End If
#If MULSEL Then
                        If m_bMultiSelect And Not bHandled Then
                           pSelectedNodeChanged hNode, lS, bMouseUp:=False
                        End If
#End If  ' MULSEL
#End If  ' CUSTDRAW
                     End If
                  End If   ' pvTVHitTest(hNode, lfHit, x, y)
               End If   ' (uMsg = WM_LBUTTONDOWN)
               
               ' http://groups.google.de/groups?hl=de&lr=&selm=HLUkOAnJyfFVw81v04n9dOVHPWPq%404ax.com
               ' TreeView handles WM_RBUTTONDOWN directly instead of dispatching it.
               ' -> no MouseUp msg
               ' eat WM_RBUTTONDOWN: no NM_RCLICK,TVN_BEGINRDRAG notifications sent
               bHandled = bHandled Or (uMsg <> WM_LBUTTONDOWN)  ' (pvButton(uMsg) <> vbLeftButton)

#If OLEDD Then
               ' Preventing inadvertant drag & drop
               m_lDragCounter = timeGetTime()
#End If

            Case WM_LBUTTONUP, WM_RBUTTONUP, WM_MBUTTONUP
               
               lS = pvShiftState()
               X = lParam And 65535:   Y = lParam \ 65536
               
               RaiseEvent MouseUp(pvButton(uMsg), lS, X, Y)
                  
               ' Click event fires only if no node hit
               Call pvTVHitTest(hNode, lfHit, X, Y)
               If (lfHit And (TVHT_ONITEM Or TVHT_ONITEMBUTTON)) = 0& Then
                  If (lfHit And TVHT_ONITEMRIGHT) = 0& Or Not m_bFullRowSelect Then
                     ' Click on treeview and no node hit
                     If Not bDblClick Then
                        ' eat trailing Click after DblClick
                        RaiseEvent Click(pvButton(uMsg))
                     End If
                  End If
               End If
               bDblClick = False
               m_hEdit = 0&
#If OLEDD Then
               m_hNodeDrag = 0&
#End If
#If MULSEL Then
               If m_bMultiSelect And (uMsg = WM_LBUTTONUP) Then
                  If (lfHit And TVHT_ONITEM) Or _
                     (((lfHit And TVHT_ONITEMRIGHT) = TVHT_ONITEMRIGHT) And m_bFullRowSelect) Then
                     pSelectedNodeChanged hNode, lS, bMouseUp:=True
                  End If
               End If
#End If

#If CUSTDRAW Then

            Case WM_NOTIFY

               Dim rc As RECT2, rcItem As RECT2, rcItemC As RECT2, rcText As RECT2
               Dim lColor     As Long
               
               Call CopyMemory(uNMH, ByVal lParam, Len(uNMH))

               Select Case uNMH.code

                  Case TTN_SHOW
                     ' ToolTip is about to be displayed: customize TT.
                     
                     If Not bBefore Then
                        
                        If Not pvTVHitTest(hNode) Then Exit Sub
                        idxNode = pTVlParam(hNode)
                        
                        pTVItemRect hNode, rcItem, OnlyText:=True
                        rcItemC = pGetItemRectReal(hNode, rcItem, rcText, idxNode)
                     
                        ' works only after tree processed msg
                        If m_bDrawColor Then
                           ' # outcomment unwanted behaviour #
                           ' set NodeForeColor as ToolTip textcolor
                           lColor = m_uNodeData(idxNode).lForeColor
                           If lColor = CLR_NONE Then
                              ' reset to tree textcolor
                              lColor = GetSysColor(COLOR_WINDOWTEXT)
                           End If
                           SendMessageLong uNMH.hwndFrom, TTM_SETTIPTEXTCOLOR, lColor, 0&
                           
                           ' set NodeBackColor as ToolTip backcolor
                           lColor = m_uNodeData(idxNode).lBackColor
                           If lColor = CLR_NONE Then
                              ' reset to tree backcolor
                              lColor = m_lTreeColors(clrTreeBK)
                           End If
                           SendMessageLong uNMH.hwndFrom, TTM_SETTIPBKCOLOR, lColor, 0&
                        End If

                        If m_bDrawFont Then
                           If m_uNodeData(idxNode).idxFont <> 0& Then
                              ' set NodeFont as ToolTip font: sets also correct size
                              Dim hFont As Long
                              hFont = m_hFnt(m_uNodeData(idxNode).idxFont)
                              If hFont <> SendMessageLong(uNMH.hwndFrom, WM_GETFONT, 0, 0&) Then
                                 SendMessageLong uNMH.hwndFrom, WM_SETFONT, m_hFnt(m_uNodeData(idxNode).idxFont), 1&
                              End If
                           End If
                        End If

                        ' size is OK, match ToolTip position to node text:
                        ' allow for TextIndent and center vertical in NodeItemRect.
                        tP.X = rcItemC.X1 - 1&
                        If m_bDrawLabel Then
                           ' center vertical in NodeItemRect
                           tP.Y = rcText.Y1 - 2&
                        Else
                           ' NodeItemRect top
                           tP.Y = rcItemC.Y1 - 2&
                        End If
                        ClientToScreen m_hTreeView, tP
                        ' ensure TT will be fully visible even at right of screen
                        If (tP.X + (rcItemC.X2 - rcItemC.X1)) > Screen.Width \ Screen.TwipsPerPixelX Then
                           ' shift right into screen
                           tP.X = Screen.Width \ Screen.TwipsPerPixelX - (rcItemC.X2 - rcItemC.X1)
                        End If
                        SetWindowPos uNMH.hwndFrom, 0&, tP.X, tP.Y, 0&, 0&, _
                                     SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_NOZORDER
                     End If
                     
                  Case TTN_GETDISPINFOA, TTN_GETDISPINFOW
                     ' TT requests text
                     ' COMCTL calculates with ItemRect, whether ToolTip should be displayed:
                     ' Incorrect behaviour for NodeFont & TextIndent.
                     If bBefore Then
                        
                        pvTVHitTest hNode
                        If hNode Then
                           idxNode = pTVlParam(hNode)
                           
                           pTVItemRect hNode, rc, OnlyText:=False
                           pTVItemRect hNode, rcItem, OnlyText:=True
                           rcItemC = pGetItemRectReal(hNode, rcItem, rcText, idxNode)
      
                           ' decide whether ToolTip should be displayed
                           bHandled = (rcItem.X2 > rc.X2) And Not (rcItemC.X2 > rc.X2)
                        End If
                        m_bInTTip = True
                        
                     Else
                        m_bInTTip = False
                     End If
                     
               End Select  ' uNMH.code
               
#End If  ' CUSTDRAW

            Case WM_KILLFOCUS
            
               ' reset before painting occurs
               m_bFocus = False
            
            Case WM_KEYDOWN

               ' Toggle node expansion for Enter & Space(not CheckBoxes) keys.
               '       Inhibit toggling with KeyDown event(KeyCode = 0).
               KeyCode = wParam And &H7FFF&

               RaiseEvent KeyDown(KeyCode, pvShiftState())
               wParam = (wParam And Not &H7FFF&) Or (KeyCode And &H7FFF&)

               If ((KeyCode = vbKeySpace) And Not m_bCheckBoxes) _
                  Or (KeyCode = vbKeyReturn) Then
                  hNode = SelectedNode
                  If hNode Then
                     If NodeExpanded(hNode) Then
                        Collapse hNode
                     Else
                        Expand hNode
                     End If
                  End If
               ElseIf ((KeyCode = vbKeySpace) And m_bCheckBoxes) Then
                  hNode = SelectedNode
                  idxOldState = NodeStateImage(hNode)
                  idxNewState = SII_SWAPMASK - idxOldState ' only valid for SII_(UN)CHECKED
                  RaiseEvent BeforeNodeCheck(hNode, idxNewState, idxOldState)
                  If idxNewState <> idxOldState Then
                     pTVStateImage(hNode) = idxNewState
                     RaiseEvent AfterNodeCheck(hNode)
                  End If
                  bHandled = True
               End If
#If MULSEL Then
               If m_bMultiSelect Then
                  Select Case KeyCode
                     Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyPageDown, vbKeyPageUp
                        ' update selection for arrow keys on TVN_SELCHANGING
                        m_bKeyDown = True
                     Case vbKeyA
                        If pvShiftState = vbCtrlMask Then
                           ' select all nodes (CTRL + A)
                           For idxNode = 1 To m_lNodeCount
                              NodeSelected(m_uNodeData(idxNode).hNode) = True
                           Next
                        End If
                  End Select
               End If
#End If

            Case WM_CHAR
               
               KeyCode = wParam And &H7FFF&
               RaiseEvent KeyPress(wParam And &H7FFF&)
               wParam = (wParam And Not &H7FFF&) Or (KeyCode And &H7FFF&)

            Case WM_KEYUP
               
               KeyCode = wParam And &H7FFF&
               lS = pvShiftState()
               RaiseEvent KeyUp(KeyCode, lS)
               wParam = (wParam And Not &H7FFF&) Or (KeyCode And &H7FFF&)
#If MULSEL Then
               If m_bMultiSelect Then
                  Select Case KeyCode
                     Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyPageDown, vbKeyPageUp
                        hNode = SelectedNode
                        If hNode Then
                           If (lS And vbCtrlMask) = 0& Then
                              pSelectedNodeChanged hNode, lS, bMouseUp:=True
                         ' Else: COMCTL scrolls, if Ctrl key is pressed: no selection change
                           End If
                        End If
                  End Select
               End If
#End If

            Case WM_LBUTTONDBLCLK, WM_RBUTTONDBLCLK, WM_MBUTTONDBLCLK
               ' DblClick and NodeDblClick events:
               ' WM_XBUTTONDBLCLK used instead of NM_XDBLCLK:
               ' Treeview NM_RDBLCLK sent only, if left doubleclicked while right button pressed.
               ' Fix this undocumented 'behaviour by design' with WM_RBUTTONDBLCLK msg.
               ' WM_MBUTTONDBLCLK provides missing NM_MDBLCLK functionality.
               ' Provides TreeView DblClick event: no node hit.

               Call pvTVHitTest(hNode, lfHit)

               If hNode Then
                  If (lfHit And TVHT_ONITEMSTATEICON) Then
                     ' DblClick on 'CheckBox'
                     bHandled = True
                  ElseIf (lfHit And TVHT_ONITEM) Then
                     ' permits preventing node expansion on doubleclick
                     RaiseEvent NodeDblClick(hNode, pvButton(uMsg), bHandled)
                  End If

               Else
                  ' DblClick in Treeview, but no node hit
                  RaiseEvent DblClick(pvButton(uMsg))
                  ' eat trailing Click after DblClick
                  bDblClick = True
               End If
                  
            Case WM_SETCURSOR
               ' The low-order word of lParam specifies the hit-test code.
               ' The high-order word of lParam specifies the identifier of the mouse message.
               ' see TrackSelect prop
               If m_bUseStandardCursor Then
                  If wParam = m_hTreeView Then
                     lParam = lParam And 65535        ' == LOWORD(lParam)
                     Select Case lParam
                        Case 1, 2                     ' == HTCLIENT,HTCAPTION ' if contextmenu invoked -> HTCAPTION ???
                           If HitTest(, , False, lfHit) Then
                              If (lfHit And (TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMTEXTINDENT)) Then
                                 Const IDC_ARROW = 32512&
                                 SetCursor LoadCursor(0, IDC_ARROW)
                                 lReturn = 1
                                 bHandled = True
                              End If
                           End If
                     End Select
                  End If
               End If

#If FLDBR Then

            Case WM_INITMENUPOPUP, WM_DRAWITEM, WM_MEASUREITEM
               ' Handle ownerdrawn shell context menu messages
            
               If Not (ICtxMenu3 Is Nothing) Then
                  Call ICtxMenu3.HandleMenuMsg2(uMsg, wParam, lParam, lReturn)
                  bHandled = True
               ElseIf Not (ICtxMenu2 Is Nothing) Then
                  Call ICtxMenu2.HandleMenuMsg(uMsg, wParam, lParam)
                  lReturn = 1
                  bHandled = True
               End If
            
            Case WM_MENUCHAR
            
               If Not (ICtxMenu3 Is Nothing) Then
                  Call ICtxMenu3.HandleMenuMsg2(uMsg, wParam, lParam, lReturn)
                  bHandled = True
               End If
            
          ' Case WM_MENUSELECT
               ' Show descriptive help text in StatusBar: ICtxMenuX.GetCommandString()
               
#End If  ' FLDBR
            
         End Select  ' uMsg TreeView

#If DDIMG Then
      
      Case m_hDDImg  ' Drag image window
         
         Select Case uMsg
         
            Case WM_PAINT
               Dim hdc  As Long
               
               hdc = GetDC(m_hDDImg)
               
               ' m_tMemDC(1) holds drag image (-> pDragImageStart())
               With m_tMemDC(1)
                  BitBlt hdc, 0&, 0&, .lWidth, .lHeight, .hdc, 0&, 0&, vbSrcCopy
               End With
               
               ReleaseDC m_hDDImg, hdc
               
          ' Case WM_ERASEBKGND
            
          ' Case Else
          '    Debug.Assert uMsg = WM_KILLFOCUS
          '  ' DefWindowProc m_hDDImg, uMsg, wParam, lParam
         End Select
            
      Case m_hClipBrd      ' CLIPBRDWNDCLASS window
      
         Select Case uMsg
         
            Case WM_CAPTURECHANGED
               ' drag complete or !! drag contextmenu popped !!
               pDragImageComplete
            
'            Case WM_USER  ' wParam = 0 /lParam == hwnd of new ole-active control in app ?
'               Debug.Print "CLIPBRDWNDCLASS   WM_USER", Hex$(lParam)
'            Case &H307  ' WM_DESTROYCLIPBOARD   wParam = lParam = 0
'               Debug.Print "CLIPBRDWNDCLASS   WM_DESTROYCLIPBOARD"
'            Case 0   ' ??? wParam = lParam = 0
'            Case Else
'               Debug.Print "CLIPBRDWNDCLASS", Hex$(uMsg), wParam, lParam:  Debug.Assert False
         End Select
      
#End If  ' DDIMG

   End Select ' lng_hWnd

End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function Initialize() As Boolean

    If (m_bInitialized = False) Then

        Initialize = pvCreateTreeView()

        If (m_hTreeView) Then
            
            m_hUC = UserControl.hwnd   ' MSDN: "never store the hWnd value in a variable",but does it change after creation ???
            
            '-- Subclass UserControl (parent)
            Call Subclass_Start(m_hUC)
            Call Subclass_AddMsg(m_hUC, WM_MOUSEACTIVATE, MSG_AFTER)
            Call Subclass_AddMsg(m_hUC, WM_SETFOCUS, MSG_AFTER)
            Call Subclass_AddMsg(m_hUC, WM_SIZE, MSG_AFTER)
            Call Subclass_AddMsg(m_hUC, WM_NOTIFY, MSG_AFTER)

            '-- Subclass TreeView (child)
            Call Subclass_Start(m_hTreeView)
            Call Subclass_AddMsg(m_hTreeView, WM_KILLFOCUS, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_KEYDOWN, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_CHAR, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_KEYUP, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_LBUTTONDOWN, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_RBUTTONDOWN, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_MBUTTONDOWN, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_MOUSEMOVE, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_LBUTTONUP, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_RBUTTONUP, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_MBUTTONUP, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_LBUTTONDBLCLK, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_RBUTTONDBLCLK, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_MBUTTONDBLCLK, MSG_BEFORE)
#If CUSTDRAW Then
            Call Subclass_AddMsg(m_hTreeView, WM_PAINT, MSG_BEFORE_AND_AFTER)
            Call Subclass_AddMsg(m_hTreeView, WM_ERASEBKGND, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_NOTIFY, MSG_BEFORE_AND_AFTER)
#End If
#If DDIMG Then
            Call Subclass_AddMsg(m_hTreeView, WM_TIMER, MSG_BEFORE)
            
            Dim hwnd As Long
            Dim pID  As Long
      
            Do
               hwnd = FindWindowEx(HWND_MESSAGE, hwnd, "CLIPBRDWNDCLASS", vbNullString)
               
               If App.ThreadID = GetWindowThreadProcessId(hwnd, pID) Then
                  m_hClipBrd = hwnd
                  
                  '-- Subclass CLIPBRDWNDCLASS window
                  Call Subclass_Start(m_hClipBrd)
                  Call Subclass_AddMsg(m_hClipBrd, WM_CAPTURECHANGED, MSG_BEFORE)
'                  Call Subclass_AddMsg(m_hClipBrd, ALL_MESSAGES, MSG_BEFORE)
                  Exit Do
               End If
               
            Loop While hwnd <> 0
            
            Debug.Assert m_hClipBrd
#End If
#If FLDBR Then
            Call Subclass_AddMsg(m_hTreeView, WM_INITMENUPOPUP, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_DRAWITEM, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_MEASUREITEM, MSG_BEFORE)
            Call Subclass_AddMsg(m_hTreeView, WM_MENUCHAR, MSG_BEFORE)
#End If

            '-- TreeView mouse enter/leave support and mouse pos. initialization
            m_bTrack = True
            m_bTrackUser32 = pvIsFunctionExported("TrackMouseEvent", "User32")
            If (Not m_bTrackUser32) Then
                If (Not pvIsFunctionExported("_TrackMouseEvent", "Comctl32")) Then
                    m_bTrack = False
                End If
            End If
            If (m_bTrack) Then
                Call Subclass_AddMsg(m_hTreeView, WM_MOUSELEAVE, MSG_BEFORE)
            End If
            m_lx = -1
            m_ly = -1

            '-- Initialize OLEInPlaceActiveObject interface
            Call mIOIPAOTreeView.InitIPAO(m_uIPAO, Me)
            
            m_bInitialized = True
        End If
    End If
End Function

Public Function InitializeImageList(Optional ByVal ImageWidth As Long = 16, _
                                    Optional ByVal ImageHeight As Long = 16) As Boolean

   If (m_hTreeView) Then
      Call pvDestroyImageList
       
      ' MOD6
      m_hImageList(ilNormal) = ImageList_Create(ImageWidth, ImageHeight, _
                                  GetDeviceCaps(GetDC(0), COLORRES) Or ILC_MASK, 0&, 0&)

      If (m_hImageList(ilNormal)) Then
         Call pvSetImageList(m_hImageList(ilNormal))
         InitializeImageList = True
      End If

#If CUSTDRAW Then
      ' used in pTreeCustomDraw
      m_lImageW = ImageWidth
      m_lImageH = ImageHeight
#End If

   End If
End Function

Public Function AddBitmap(ByVal hBitmap As Long, _
                          Optional ByVal MaskColor As OLE_COLOR = CLR_NONE, _
                          Optional ByVal bDestroy As Boolean = True) As Long
   If m_hImageList(ilNormal) Then
      If (MaskColor <> CLR_NONE) Then
         AddBitmap = ImageList_AddMasked(m_hImageList(ilNormal), hBitmap, _
                                         pTranslateColor(MaskColor))
      Else
         AddBitmap = ImageList_Add(m_hImageList(ilNormal), hBitmap, 0&)
      End If
      m_lImageListCount(ilNormal) = ImageList_GetImageCount(m_hImageList(ilNormal))
      If bDestroy Then
         DeleteObject hBitmap
      End If
   End If
End Function

Public Function AddIcon(ByVal hIcon As Long, Optional ByVal bDestroy As Boolean = True _
                        ) As Long
   If m_hImageList(ilNormal) Then
      AddIcon = ImageList_AddIcon(m_hImageList(ilNormal), hIcon)
      m_lImageListCount(ilNormal) = ImageList_GetImageCount(m_hImageList(ilNormal))
      If bDestroy Then
         DestroyIcon hIcon
      End If
   End If
End Function

'alternate use ImageList_Replace API (involves creating mask)
Public Function ReplaceBitmap(ByVal hBitmap As Long, ByVal idxReplace As Long, _
                              Optional ByVal MaskColor As OLE_COLOR = CLR_NONE, _
                              Optional ByVal bDestroy As Boolean = True) As Long
   If m_hImageList(ilNormal) Then
      If idxReplace > IMG_NONE And idxReplace < ImageListCount Then
         ' add new as last
         If AddBitmap(hBitmap, MaskColor, bDestroy) <> IMG_NONE Then
            ' swap new with replacable
            ImageList_Copy m_hImageList(ilNormal), m_lImageListCount(ilNormal) - 1, _
                           m_hImageList(ilNormal), idxReplace, ILCF_SWAP
            ' remove replacable as last
            ImageList_Remove m_hImageList(ilNormal), m_lImageListCount(ilNormal) - 1
            ReplaceBitmap = idxReplace
         End If
      End If
   End If
   Debug.Assert (ReplaceBitmap = idxReplace)
End Function

' appends at end for idxReplace:=IMG_NONE
Public Function ReplaceIcon(ByVal hIcon As Long, ByVal idxReplace As Long, _
                            Optional ByVal bDestroy As Boolean = True) As Long
   If m_hImageList(ilNormal) Then
      If idxReplace >= IMG_NONE And idxReplace < ImageListCount Then
         ReplaceIcon = ImageList_ReplaceIcon(m_hImageList(ilNormal), idxReplace, hIcon)
         m_lImageListCount(ilNormal) = ImageList_GetImageCount(m_hImageList(ilNormal))
      End If
      If bDestroy Then
         DestroyIcon hIcon
      End If
   End If
   Debug.Assert (ReplaceIcon = idxReplace) Or (ReplaceIcon = m_lImageListCount(ilNormal) - 1) And (idxReplace = IMG_NONE)
End Function

Public Property Get Redraw() As Boolean
   Redraw = m_bRedraw
End Property
Public Property Let Redraw(ByVal bRedraw As Boolean)
   If m_bRedraw <> bRedraw Then
      m_bRedraw = bRedraw
      pSetRedrawMode bRedraw
   End If
End Property

Private Sub pSetRedrawMode(ByVal Enable As Boolean)
   If (m_hTreeView) Then
      Call SendMessageLong(m_hTreeView, WM_SETREDRAW, -Enable, 0&)
   End If
End Sub

Public Sub Refresh(Optional hNode As Long)
   Dim tR   As RECT2
   If hNode Then
      tR.X1 = hNode
      SendMessage m_hTreeView, TVM_GETITEMRECT, 0&, tR
   Else
      GetClientRect m_hTreeView, tR
   End If
   InvalidateRect m_hTreeView, tR, 1&
End Sub

'== Adding / removing nodes
Public Function AddNode(Optional ByVal hRelative As Long, _
                        Optional ByVal Relation As tvRelationConstants, _
                        Optional ByVal Key As String, _
                        Optional ByVal Text As String, _
                        Optional ByVal Image As Long = IMG_NONE, _
                        Optional ByVal SelectedImage As Long = IMG_NONE, _
                        Optional ByVal PlusButton As Boolean = False, _
                        Optional ByVal Tag As String = vbNullString _
                        ) As Long
   
   If (m_hTreeView) Then
        
      ' for TVN_GETDISPINFO calls need to populate m_uNodeData before adding node
        
      m_lNodeCount = m_lNodeCount + 1
      If (m_lNodeCount Mod ALLOCATE_SIZE = 1) Then
          ' Err 10
          ReDim Preserve m_uNodeData(0 To m_lNodeCount + ALLOCATE_SIZE)
      End If
      
      With m_uNodeData(m_lNodeCount)
         .sKey = Key
         .sTag = Tag
         .sText = Text
         
         If (Image > IMG_NONE And Image < ImageListCount) Then
            .idxImg = Image
         Else
            .idxImg = IMG_NONE
         End If
         If (SelectedImage > IMG_NONE And SelectedImage < m_lImageListCount(ilNormal)) Then
            .idxSelImg = SelectedImage
         Else
            ' SelectedImage = Image, if unspecified
            .idxSelImg = .idxImg
         End If

#If CUSTDRAW Then
         .lForeColor = CLR_NONE
         .lBackColor = CLR_NONE
         .idxExpImg = IMG_NONE
         .xFont = pGetRcTreeFont(Text)
         ' .idxFont = 0
         ' ptFont.cX = 0: ptFont.cY = 0
         ' .lItemdata = 0
#End If
      
         AddNode = pvTVAddNode(m_lNodeCount, hRelative, Relation, PlusButton)
         
         If AddNode <> 0& Then
            ' permit adding unkeyed nodes
            If LenB(Key) Then
               On Error GoTo errH
               ' Err 457
               Call m_cKey.Add(AddNode, Key)
            End If
            
            .hNode = AddNode
            
         Else
            m_lNodeCount = m_lNodeCount - 1
         End If
         
      End With

   End If
   Exit Function

errH:
   Debug.Print Err.Number, Err.Description
   If Err = 10& Then
      ' Failed to redim m_uNodeData(): This array is fixed or temporarily locked.
      ' Propably added node inside a With m_uNodeData() clause, see pOLEWrite proc.
      Debug.Assert AddNode = 0
      Debug.Assert False
      m_lNodeCount = m_lNodeCount - 1
   ElseIf Err = 457& Then
      ' This key is already associated with an element of this collection
      Debug.Assert AddNode <> 0
      m_lNodeCount = m_lNodeCount - 1
      Call SendMessageLong(m_hTreeView, TVM_DELETEITEM, 0, AddNode)
      AddNode = 0
   Else
      Debug.Assert False
   End If
End Function

' don't process TVN_DELETEITEM, if TreeView is cleared
Public Sub Clear()

   If (m_hTreeView) Then

      If m_bRedraw Then pSetRedrawMode False
      m_bInClear = True

      '-- Delete all TreeView nodes
      Call SendMessageLong(m_hTreeView, TVM_DELETEITEM, 0, TVI_ROOT)

      m_bInClear = False
      If m_bRedraw Then pSetRedrawMode True

      '-- Erase node data array and key collection
      ReDim m_uNodeData(0 To 1)
      Set m_cKey = Nothing
      Set m_cKey = New Collection
      '-- Reset count
      m_lNodeCount = 0
   End If
End Sub

' set Redraw too
Public Property Get HoldDeletePostProcess() As Boolean
   HoldDeletePostProcess = m_bHoldDeletePostProcess
End Property
Public Property Let HoldDeletePostProcess(ByVal Hold As Boolean)
' Use this carefully!
' Use when need to do multiple calls to DeleteNode() function in same routine/loop.
' This will prevent an update for each call done to DeleteNode() function.
' Ex.:
'
'   Call ucTreeView.HoldDeletePostProcess(True)  '-> Stops internal update process of collection and array
'   Call ucTreeView.DeleteNode(hNode1)
'   [...]
'   Call ucTreeView.DeleteNode(hNodeN)
'   Call ucTreeView.HoldDeletePostProcess(False) '-> Proceeds to update

    '-- Hold or proceed ?
    m_bHoldDeletePostProcess = Hold
    If (Hold = False) Then
        Call pvDoDeletePostProcess
'        Redraw = True
'    Else
'        Redraw = False
    End If
End Property

' COMCTL default behavior, when single child is deleted:
' - parent's expanded/expandedonce states don't change.
' - if parent button was set with NodePlusMinusButton,it is not removed.
Public Function DeleteNode(ByVal hNode As Long) As Boolean
  
    If (m_hTreeView) Then
        
        If (SendMessageLong(m_hTreeView, TVM_DELETEITEM, 0, hNode)) Then
            
            '-- Waiting for multiple DeleteNode() calls ?
            If (m_bHoldDeletePostProcess = False) Then
                '-- Let's proceed to update our collection and array
                Call pvDoDeletePostProcess
            End If
            '-- Success
            DeleteNode = True
        End If
    End If
End Function

'== Validating key / Retrieving Key and hNode

Public Function IsValidKey(ByVal Key As String) As Boolean

    On Error GoTo errH

    Call m_cKey.Add(0, Key)
    Call m_cKey.Remove(m_cKey.Count)
    IsValidKey = True
    Exit Function

errH:
End Function

'== Editing labels (Start/End edition from code)

Public Function StartLabelEdit(ByVal hNode As Long) As Boolean
    If (m_hTreeView) Then
        StartLabelEdit = CBool(SendMessageLong(m_hTreeView, TVM_EDITLABEL, 0&, hNode))
    End If
End Function
Public Sub EndLabelEdit(ByVal Cancel As Boolean)
    If (m_hTreeView) Then
        Call SendMessageLong(m_hTreeView, TVM_ENDEDITLABELNOW, -Cancel, 0&)
    End If
End Sub

'== Visibility / Expanding and Collapsing / Sorting / Hereditary checking

Public Sub EnsureVisible(ByVal hNode As Long, Optional NoScrollRight As Boolean)
   If (m_hTreeView) Then
      Call SendMessageLong(m_hTreeView, TVM_ENSUREVISIBLE, 0&, hNode)
      If NoScrollRight Then
         ' default behaviour: if necessary, scrolls to right
         Scroll sLeft
      End If
   End If
End Sub

' EventOnly: raises Expand event(s) without actually expanding node(s) ' # Risky #
Public Sub Expand(ByVal hNode As Long, Optional ByVal ExpandChildren As Boolean = False, _
                  Optional ByVal EventOnly As Boolean = False)
   Dim hNext As Long
   
   If hNode = 0 Then
      Debug.Assert False
      Exit Sub
   End If
   
   If (m_hTreeView) Then
      If Not EventOnly Then
         
         If (m_hInExpand = 0&) Then
            ' LockWindowUpdate once
            ' m_hInExpand prevents LockWindowUpdate's for expanding children in zSubclassProc/TVN_ITEMEXPANDING
            m_hInExpand = hNode
            LockWindowUpdate m_hTreeView
         End If
         Call SendMessageLong(m_hTreeView, TVM_EXPAND, TVE_EXPAND, hNode)
         
      Else
         If pTVcChildren(hNode) Then  ' == NodePlusMinusButton
            Dim Cancel As Integer
            
            RaiseEvent BeforeExpand(hNode, pTVState(hNode, TVIS_EXPANDEDONCE), Cancel)
            
            If (Cancel = 0) Then
               pTVState(hNode, TVIS_EXPANDEDONCE) = True     ' # Risky #
               
               RaiseEvent AfterExpand(hNode, True)
               
            Else
               ExpandChildren = False
            End If
            
         Else
            ExpandChildren = False
         End If
      End If
      
      If (ExpandChildren) Then
         hNext = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_CHILD, hNode)
         Do While hNext
            Call Expand(hNext, ExpandChildren:=True, EventOnly:=EventOnly)
            hNext = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_NEXT, hNext)
            If GetAsyncKeyState(vbKeyEscape) < 0 Then Exit Do
         Loop
      End If
      
      If m_hInExpand = hNode Then
         ' release LockWindowUpdate, after all children have expanded
         m_hInExpand = 0&
         LockWindowUpdate 0&
      End If
   End If
End Sub

' RemoveChildren: proper way to reset NodeExpandedOnce by TVE_COLLAPSERESET
' # TVE_COLLAPSERESET must be called !before! node is collapsed (BeforeCollapse event). #
' # Calling afterwards resets NodeExpandedOnce, but ExpandedOnce var in successive      #
' # BeforeExpand events is always set.                                                  #
Public Sub Collapse(ByVal hNode As Long, Optional ByVal CollapseChildren As Boolean = False, _
                    Optional ByVal RemoveChildren As Boolean = False)

   Dim hNext As Long

   If hNode = 0 Then Debug.Assert False:  Exit Sub

   If (m_hTreeView) Then
      If (m_hInExpand = 0&) Then
         ' LockWindowUpdate once
         ' m_hInExpand prevents LockWindowUpdate's for collapsing children in zSubclassProc/TVN_ITEMEXPANDING
         m_hInExpand = hNode
         LockWindowUpdate m_hTreeView
      End If
      If (Not RemoveChildren) Then
         Call SendMessageLong(m_hTreeView, TVM_EXPAND, TVE_COLLAPSE, hNode)
      Else
         ' sends TVN_DELETEITEM but no TVN_ITEMEXPANDING/TVN_ITEMEXPANDED notifs
         Call SendMessageLong(m_hTreeView, TVM_EXPAND, TVE_COLLAPSE Or TVE_COLLAPSERESET, hNode)
         Debug.Assert NodeChild(hNode) = 0
         Debug.Assert NodeExpandedOnce(hNode) = False
         pvDoDeletePostProcess
         m_hInExpand = 0&
         LockWindowUpdate 0&
         Exit Sub
      End If
      If (CollapseChildren) Then
         hNext = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_CHILD, hNode)
         Do While hNext
             Call Collapse(hNext, CollapseChildren:=True)
             hNext = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_NEXT, hNext)
         Loop
      End If
      If m_hInExpand = hNode Then
         ' release LockWindowUpdate, after all children have collapsed
         m_hInExpand = 0&
         LockWindowUpdate 0&
      End If
   End If
End Sub

Public Sub SortChildren(ByVal hNode As Long, Optional ByVal SortAllLevels As Boolean = False)
'Don't know why, but fRecurse param. is not working.
'So, sort recursively using recursive call.
   Dim hNext   As Long
   
   If (m_hTreeView) Then
   
      m_bInSort = True
      Call SendMessageLong(m_hTreeView, TVM_SORTCHILDREN, 0, hNode)
      
      If (SortAllLevels) Then
         hNext = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_CHILD, hNode)
         Do While hNext
            Call SortChildren(hNext, SortAllLevels:=True)
            hNext = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_NEXT, hNext)
         Loop
      End If
      m_bInSort = False
   End If
End Sub

Public Sub SortChildrenCB(ByVal hNode As Long, ByVal CallBackAddress As Long, _
                          Optional ByVal SortAllLevels As Boolean = False, _
                          Optional ByVal lParamSort As Long = 0&)
'Don't know why, but fRecurse param. is not working.
'So, sort recursively using recursive call.
   Dim hNext   As Long
   Dim tTVCB   As TVSORTCB
   
   If (m_hTreeView) Then
      
      m_bInSort = True
      
      With tTVCB
         .hParent = hNode
         .lParam = lParamSort
         .lpfnCompare = CallBackAddress
      End With
      
      Call SendMessage(m_hTreeView, TVM_SORTCHILDRENCB, 0&, tTVCB)
      
      If (SortAllLevels) Then
         hNext = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_CHILD, hNode)
         Do While hNext
            Call SortChildrenCB(hNext, CallBackAddress, True, lParamSort)
            hNext = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_NEXT, hNext)
         Loop
      End If
      m_bInSort = False
   
   End If
End Sub

Public Sub CheckChildren(ByVal hNode As Long, ByVal New_NodeChecked As Boolean)

  Dim hNext As Long

    If (m_hTreeView) Then

        hNext = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_CHILD, hNode)
        Do While hNext
            If (New_NodeChecked) Then
                pTVStateImage(hNext) = SII_CHECKED
              Else
                pTVStateImage(hNext) = SII_UNCHECKED
            End If
            Call CheckChildren(hNext, New_NodeChecked)
            hNext = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_NEXT, hNext)
        Loop
    End If
End Sub

'== HitTest / Insertion mark / Hilited node

' x,y in pixel client coordinates (always returned), returns lfHit
Public Function HitTest(Optional ByVal X As Long, Optional ByVal Y As Long, _
                        Optional ByVal FullRowHit As Boolean = True, _
                        Optional ByRef lfHit As tvHitTestConstants) As Long
  Dim hNode As Long

    If (m_hTreeView) Then

        Call pvTVHitTest(hNode, lfHit, X, Y)

        If (FullRowHit) Then
            HitTest = hNode
        Else
            If (lfHit And TVHT_ONITEM) Then
                HitTest = hNode
            End If
        End If
    End If
End Function

' x,y in client coordinates
Private Function pvTVHitTest(hNode As Long, Optional lfHit As tvHitTestConstants, _
                             Optional X As Long, Optional Y As Long, _
                             Optional lFalseHit As Long) As Boolean

   Dim uTVHI As TVHITTESTINFO

   If (X > 0&) And (Y > 0&) Then
      uTVHI.pt.X = X:   uTVHI.pt.Y = Y
   Else
      Call GetCursorPos(uTVHI.pt)
      Call ScreenToClient(m_hTreeView, uTVHI.pt)
   End If

   pvTVHitTest = SendMessage(m_hTreeView, TVM_HITTEST, 0&, uTVHI)
   With uTVHI
      lfHit = .flags
      hNode = .hItem
   End With
   
   X = uTVHI.pt.X:   Y = uTVHI.pt.Y
   
#If CUSTDRAW Then
   lFalseHit = 0&
   
   If m_bDrawLabel Then
      If ((lfHit And TVHT_ONITEMLABEL) = TVHT_ONITEMLABEL) Or _
         ((lfHit And TVHT_ONITEMRIGHT) = TVHT_ONITEMRIGHT) Then
         ' comctl may report false Hittest for NodeFont <> TreeFont
         Dim rcItem As RECT2, rcItemC As RECT2, rcJunk As RECT2

         pTVItemRect hNode, rcItem, OnlyText:=True
         rcItemC = pGetItemRectReal(hNode, rcItem, rcJunk)
         
         If ((lfHit And TVHT_ONITEMLABEL) = TVHT_ONITEMLABEL) Then
            If X > rcItemC.X2 Then
               lfHit = TVHT_ONITEMRIGHT
               lFalseHit = rcItem.X2 - rcItemC.X2  ' positive
               ' Debug.Print "FALSE HIT: ONITEMLABEL, correct: ONITEMRIGHT", lFalseHit
            ElseIf m_bDrawLabelTI Then
               If (X > rcItem.X1 And X < rcItemC.X1) Then
                  ' in Label, but TextIndent space
                  lfHit = lfHit Or TVHT_ONITEMTEXTINDENT
                  ' lFalseHit stays zero
                  ' Debug.Print "HIT: TVHT_ONITEMTEXTINDENT"
               End If
            End If
         Else
            ' TVHT_ONITEMRIGHT
            If X <= rcItemC.X2 Then
               lfHit = TVHT_ONITEMLABEL
               lFalseHit = rcItem.X2 - rcItemC.X2  ' negative: shift left by lFalseHit pixels
               ' Debug.Print "FALSE HIT: ONITEMRIGHT, correct: ONITEMLABEL", lFalseHit
            End If
         End If
         
      End If
   End If
#End If

End Function

Public Sub SetInsertionMark(ByVal hNode As Long, Optional ByVal InsertAfter As Boolean = True)

    If (m_hTreeView) Then

        Call SendMessageLong(m_hTreeView, TVM_SETINSERTMARK, -InsertAfter, hNode)
    End If
End Sub

Public Sub SetHilitedNode(ByVal hNode As Long)

    If (m_hTreeView) Then

        Call SendMessageLong(m_hTreeView, TVM_SELECTITEM, TVGN_DROPHILITE, hNode)
    End If
End Sub

'== Scrolling

Public Sub Scroll(ByVal Direction As tvScrollConstants)

    Select Case Direction

        Case [sHome]:      Call SendMessageLong(m_hTreeView, WM_VSCROLL, SB_TOP, 0)
        Case [sPageUp]:    Call SendMessageLong(m_hTreeView, WM_VSCROLL, SB_PAGEUP, 0)
        Case [sUp]:        Call SendMessageLong(m_hTreeView, WM_VSCROLL, SB_LINEUP, 0)
        Case [sDown]:      Call SendMessageLong(m_hTreeView, WM_VSCROLL, SB_LINEDOWN, 0)
        Case [sPageDown]:  Call SendMessageLong(m_hTreeView, WM_VSCROLL, SB_PAGEDOWN, 0)
        Case [sEnd]:       Call SendMessageLong(m_hTreeView, WM_VSCROLL, SB_BOTTOM, 0)
        Case [sLeft]:      Call SendMessageLong(m_hTreeView, WM_HSCROLL, SB_LEFT, 0)
        Case [sPageLeft]:  Call SendMessageLong(m_hTreeView, WM_HSCROLL, SB_PAGELEFT, 0)
        Case [sLineLeft]:  Call SendMessageLong(m_hTreeView, WM_HSCROLL, SB_LINELEFT, 0)
        Case [sLineRight]: Call SendMessageLong(m_hTreeView, WM_HSCROLL, SB_LINERIGHT, 0)
        Case [sPageRight]: Call SendMessageLong(m_hTreeView, WM_HSCROLL, SB_PAGERIGHT, 0)
        Case [sRight]:     Call SendMessageLong(m_hTreeView, WM_HSCROLL, SB_RIGHT, 0)
    End Select
End Sub

'========================================================================================
' Imagelists
'========================================================================================

' assign external imagelist (or use internal imagelist elsewhere)
Public Property Get hImageList(Optional ByVal ILS As tvImagelistConstants = ilNormal) As Long
   hImageList = m_hImageList(ILS)
End Property
Public Property Let hImageList(Optional ByVal ILS As tvImagelistConstants = ilNormal, _
                               ByVal newImageList As Long)
   Dim lImageW As Long, lImageH As Long
   
   If newImageList <> m_hImageList(ILS) Then
      pvDestroyImageList ILS
      
      If newImageList <> 0& Then
         m_lImageListCount(ILS) = ImageList_GetImageCount(newImageList)
         ImageList_GetIconSize newImageList, lImageW, lImageH
         If (lImageW > 0&) And (lImageH > 0&) Then
            pvSetImageList newImageList, ILS
            m_hImageList(ILS) = newImageList
            m_bExtImagelist(ILS) = True
         Else
            Debug.Assert False
            Debug.Assert m_bExtImagelist(ILS) = False
            m_hImageList(ILS) = 0&
            Err.Raise 380
         End If
      End If

#If CUSTDRAW Then
      ' used in pTreeCustomDraw
      If ILS = ilNormal Then
         m_lImageW = lImageW
         m_lImageH = lImageH
      End If
#End If

   End If
End Property

' internally needed to set (state-)imageindices with external imagelists
Public Property Get ImageListCount(Optional ByVal ILS As tvImagelistConstants = ilNormal) As Long

   If m_bExtImagelist(ILS) And (m_hImageList(ILS) <> 0&) Then
      m_lImageListCount(ILS) = ImageList_GetImageCount(m_hImageList(ILS))
   End If
   ImageListCount = m_lImageListCount(ILS)
End Property

'========================================================================================
' Properties: Node (stored in m_uNodeData / m_cKey)
'========================================================================================

Public Property Get NodeHandle(Optional ByVal Key As String, _
                               Optional ByVal idxNode As Long) As Long
   If (m_hTreeView) Then
      If LenB(Key) Then
         On Error Resume Next
         NodeHandle = m_cKey(Key)
         On Error GoTo 0
      ElseIf (idxNode > 0) And (idxNode <= m_lNodeCount) Then
         NodeHandle = m_uNodeData(idxNode).hNode
      Else: Debug.Assert False
      End If
   End If
End Property

Public Property Get NodeKey(Optional ByVal hNode As Long, _
                            Optional ByRef idxNode As Long) As String
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      NodeKey = m_uNodeData(idxNode).sKey
   End If
End Property
Public Property Let NodeKey(Optional ByVal hNode As Long, _
                            Optional ByRef idxNode As Long, New_Key As String)
   Dim oldKey As String
   On Error GoTo Proc_Error

   If (m_hTreeView) Then
      pIndex hNode, idxNode
      hNode = m_uNodeData(idxNode).hNode
      oldKey = m_uNodeData(idxNode).sKey
      If LenB(oldKey) Then
         On Error Resume Next
         m_cKey.Remove oldKey
         On Error GoTo Proc_Error
      End If
      If LenB(New_Key) Then
         m_cKey.Add hNode, New_Key
      End If
      m_uNodeData(idxNode).sKey = New_Key
   End If

   Exit Property

Proc_Error:
   Debug.Print "Error: " & Err.Number & ". " & Err.Description
   On Error GoTo 0
   Err.Raise 380
End Property

' (== Not IsValidKey(Key))
Public Property Get NodeExists(ByVal Key As String) As Boolean
   On Error Resume Next
   NodeExists = (m_cKey(Key) > 0)
End Property

Public Property Get NodeText(Optional ByVal hNode As Long, _
                             Optional ByRef idxNode As Long) As String
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      NodeText = m_uNodeData(idxNode).sText
   End If
End Property
Public Property Let NodeText(Optional ByVal hNode As Long, Optional ByRef idxNode As Long, _
                             ByVal New_NodeText As String)
   Dim uTVI       As TVITEM
   
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      m_uNodeData(idxNode).sText = New_NodeText
   
#If CUSTDRAW Then
      pCalculateRcFont idxNode, New_NodeText
      m_uNodeData(idxNode).xFont = pGetRcTreeFont(New_NodeText)
#End If
   
      With uTVI
         .hItem = hNode
         .mask = TVIF_TEXT
         .pszText = LPSTR_TEXTCALLBACK
      End With
   
      Call SendMessage(m_hTreeView, TVM_SETITEM, 0, uTVI)
   End If
End Property

Public Property Get NodeImage(Optional ByVal hNode As Long, _
                              Optional ByRef idxNode As Long) As Long
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      NodeImage = m_uNodeData(idxNode).idxImg
   End If
End Property
Public Property Let NodeImage(Optional ByVal hNode As Long, Optional ByRef idxNode As Long, _
                              ByVal lIndex As Long)
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      With m_uNodeData(idxNode)
         If (lIndex > IMG_NONE And lIndex < ImageListCount) Then
               .idxImg = lIndex
               If .idxSelImg = IMG_NONE Then
                  .idxSelImg = lIndex
               End If
            Else
               .idxImg = IMG_NONE
               .idxSelImg = IMG_NONE
         End If
      End With
      Refresh hNode
   End If
End Property

Public Property Get NodeSelectedImage(Optional ByVal hNode As Long, _
                                      Optional ByRef idxNode As Long) As Long
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      NodeSelectedImage = m_uNodeData(idxNode).idxSelImg
   End If
End Property
Public Property Let NodeSelectedImage(Optional ByVal hNode As Long, Optional ByRef idxNode As Long, _
                                      ByVal lIndex As Long)
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      With m_uNodeData(idxNode)
         If (lIndex > IMG_NONE And lIndex < ImageListCount) Then
            .idxSelImg = lIndex
         Else
            .idxSelImg = IMG_NONE
         End If
      End With
      Refresh hNode
   End If
End Property

Public Property Get NodeTag(Optional ByVal hNode As Long, _
                            Optional ByRef idxNode As Long) As String
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      NodeTag = m_uNodeData(idxNode).sTag
   End If
End Property
Public Property Let NodeTag(Optional ByVal hNode As Long, Optional ByRef idxNode As Long, _
                            ByVal New_NodeTag As String)
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      m_uNodeData(idxNode).sTag = New_NodeTag
   End If
End Property

Public Property Get NodeIndex(ByVal hNode As Long) As Long
   NodeIndex = pTVlParam(hNode)
   Debug.Assert NodeIndex <> 0
End Property

Private Sub pIndex(ByVal hNode As Long, ByRef idxNode As Long)
   If (m_hTreeView) Then
      If idxNode = 0 Then
         ' get idxNode from comctl
         idxNode = pTVlParam(hNode)
      ElseIf (idxNode > 0) And (idxNode <= m_lNodeCount) Then
         ' idxNode valid
    ' Else
         ' invalid idxNode
      End If
   End If
   If idxNode = 0 Then
'      Debug.Assert False
      Debug.Print "ucTree.pIndex FAILED: ", hNode
      Err.Raise 380
   End If
End Sub

'========================================================================================
' Properties: Node (stored by comctl)
'========================================================================================

Public Property Get SelectedNode() As Long
   If (m_hTreeView) Then
      SelectedNode = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_CARET, 0)
   End If
End Property
Public Property Let SelectedNode(ByVal hNode As Long)
   If (m_hTreeView) Then
      Call SendMessageLong(m_hTreeView, TVM_SELECTITEM, TVGN_CARET, hNode)
   End If
End Property

' 0 <= NodeStateImage <= 15
Public Property Get NodeStateImage(ByVal hNode As Long) As Long
  Dim lIndex As Long

    If (m_hTreeView) Then
        lIndex = pTVStateImage(hNode)
        If (lIndex >= 0& And lIndex < ImageListCount(ilState)) Then
            NodeStateImage = lIndex
          Else
            NodeStateImage = 0&
        End If
    End If
End Property
' setting NodeStateImage = 0 removes any StateImage and shifts NodeRect to left
Public Property Let NodeStateImage(ByVal hNode As Long, ByVal lIndex As Long)
    If (m_hTreeView) Then
        If (lIndex >= 0& And lIndex < ImageListCount(ilState) And lIndex <= 15) Then
            pTVStateImage(hNode) = lIndex
          Else
            pTVStateImage(hNode) = 0&
        End If
    End If
End Property

' NodeBold is irrelevant,if NodeFont is used!
Public Property Get NodeBold(ByVal hNode As Long) As Boolean
   If (m_hTreeView) Then
      NodeBold = pTVState(hNode, TVIS_BOLD)
   End If
End Property
Public Property Let NodeBold(ByVal hNode As Long, ByVal New_NodeBold As Boolean)
   If (m_hTreeView) Then
      pTVState(hNode, TVIS_BOLD) = New_NodeBold
   End If
End Property

Public Property Get NodeGhosted(ByVal hNode As Long) As Boolean
   If (m_hTreeView) Then
      NodeGhosted = pTVState(hNode, TVIS_CUT)
   End If
End Property
Public Property Let NodeGhosted(ByVal hNode As Long, ByVal New_NodeGhosted As Boolean)
   If (m_hTreeView) Then
      pTVState(hNode, TVIS_CUT) = New_NodeGhosted
   End If
End Property

Public Property Get NodeHilited(ByVal hNode As Long) As Boolean
   If (m_hTreeView) Then
      NodeHilited = pTVState(hNode, TVIS_DROPHILITED)
   End If
End Property
Public Property Let NodeHilited(ByVal hNode As Long, ByVal New_NodeHilited As Boolean)
   If (m_hTreeView) Then
      pTVState(hNode, TVIS_DROPHILITED) = New_NodeHilited
   End If
End Property

Public Property Get NodePlusMinusButton(ByVal hNode As Long) As Boolean
   If (m_hTreeView) Then
      NodePlusMinusButton = CBool(pTVcChildren(hNode))
   End If
End Property
Public Property Let NodePlusMinusButton(ByVal hNode As Long, ByVal New_NodePlusMinusButton As Boolean)
   If (m_hTreeView) Then
      pTVcChildren(hNode) = IIf(New_NodePlusMinusButton, 1, 0)
   End If
End Property

' with external State imagelist only valid for NodeStateImage = 1 or 2
Public Property Get NodeChecked(ByVal hNode As Long) As Boolean
    If (m_hTreeView And m_bCheckBoxes) Then
        NodeChecked = (pTVStateImage(hNode) = SII_CHECKED)
    End If
End Property
Public Property Let NodeChecked(ByVal hNode As Long, ByVal New_NodeChecked As Boolean)
    If (m_hTreeView And m_bCheckBoxes) Then
        If (New_NodeChecked) Then
            pTVStateImage(hNode) = SII_CHECKED
          Else
            pTVStateImage(hNode) = SII_UNCHECKED
        End If
    End If
End Property

' 0 <= NodeOverlayImage <= 15 (4.70: 0 to 4)
Public Property Get NodeOverlayImage(ByVal hNode As Long) As Long
  Dim lIndex As Long

    If (m_hTreeView) Then
        lIndex = pTVOverlayImage(hNode)
        If (lIndex >= 0& And lIndex <= 15) Then
            NodeOverlayImage = lIndex
          Else
            NodeOverlayImage = 0&
        End If
    End If
End Property
Public Property Let NodeOverlayImage(ByVal hNode As Long, ByVal lIndex As Long)
    If (m_hTreeView) Then
        If (lIndex >= 0& And lIndex <= 15) Then
            pTVOverlayImage(hNode) = lIndex
          Else
            pTVOverlayImage(hNode) = 0&
        End If
    End If
End Property

'== Node navigation (R.O.)

Public Property Get NodeRoot() As Long
    If (m_hTreeView) Then
        NodeRoot = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_ROOT, 0)
    End If
End Property

Public Property Get NodeParent(ByVal hNode As Long) As Long
    If (m_hTreeView) Then
        NodeParent = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_PARENT, hNode)
    End If
End Property

Public Property Get NodeChild(ByVal hNode As Long) As Long
    If (m_hTreeView) Then
        NodeChild = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_CHILD, hNode)
    End If
End Property

Public Property Get NodeFirstVisible() As Long
    If (m_hTreeView) Then
        NodeFirstVisible = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_FIRSTVISIBLE, 0)
    End If
End Property

Public Property Get NodeLastVisible(Optional bLastNode As Boolean = False) As Long
   Dim uTVHI         As TVHITTESTINFO
   Dim tR            As RECT2
   Dim rcItem        As RECT2
   Dim lItemHeight   As Long
   
   If (m_hTreeView) Then
      
      ' >= 4.71: Retrieves the last expanded item in the tree, not the last visible !!!
      NodeLastVisible = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_LASTVISIBLE, 0&)

      If Not bLastNode Then
      
         Call GetClientRect(m_hTreeView, tR)
         pTVItemRect NodeLastVisible, rcItem, OnlyText:=False
         
         If rcItem.Y2 >= tR.Y2 Then
            ' find real last visible
            lItemHeight = SendMessageLong(m_hTreeView, TVM_GETITEMHEIGHT, 0&, 0&)
            uTVHI.pt.Y = tR.Y2 - lItemHeight 'First fully visible
            
            Do While uTVHI.pt.Y > -1
            
               Call SendMessage(m_hTreeView, TVM_HITTEST, 0, uTVHI)
               If (uTVHI.hItem) Then
                  NodeLastVisible = uTVHI.hItem
                  Exit Do
               End If
               uTVHI.pt.Y = uTVHI.pt.Y - lItemHeight
            Loop
         
      '  Else: ' TVGN_LASTVISIBLE returned real last visible
         End If
      End If
   End If
'   Debug.Print IIf(bLastNode, "LASTNODE", "LASTVISIBLE"), NodeText(NodeLastVisible)
End Property

Public Property Get NodePrevious(ByVal hNode As Long) As Long
    If (m_hTreeView) Then
        NodePrevious = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_PREVIOUSVISIBLE, hNode)
    End If
End Property

Public Property Get NodeNext(ByVal hNode As Long) As Long
    If (m_hTreeView) Then
        NodeNext = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, hNode)
    End If
End Property

Public Property Get NodeFirstSibling(ByVal hNode As Long) As Long

  Dim hPrev As Long

    hPrev = hNode
    Do
        NodeFirstSibling = hPrev
        hPrev = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_PREVIOUS, hPrev)
    Loop Until hPrev = 0
End Property

Public Property Get NodeLastSibling(ByVal hNode As Long) As Long

  Dim hNext As Long

    hNext = hNode
    Do
        NodeLastSibling = hNext
        hNext = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_NEXT, hNext)
    Loop Until hNext = 0
End Property

Public Property Get NodePreviousSibling(ByVal hNode As Long) As Long
    If (m_hTreeView) Then
        NodePreviousSibling = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_PREVIOUS, hNode)
    End If
End Property

Public Property Get NodeNextSibling(ByVal hNode As Long) As Long
    If (m_hTreeView) Then
        NodeNextSibling = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_NEXT, hNode)
    End If
End Property

Public Property Get NodeVisible(ByVal hNode As Long, Optional ByVal OnlyText As Boolean = False) As Boolean

  Dim uRctWnd As RECT2
  Dim uRctItm As RECT2
  Dim uRctInt As RECT2

    If (m_hTreeView) Then

        Call GetClientRect(m_hTreeView, uRctWnd)
        Let uRctItm.X1 = hNode
        Call SendMessage(m_hTreeView, TVM_GETITEMRECT, -OnlyText, uRctItm)
        Call IntersectRect(uRctInt, uRctWnd, uRctItm)

        NodeVisible = CBool(EqualRect(uRctInt, uRctItm))
    End If
End Property

Public Property Get NodeExpanded(ByVal hNode As Long) As Boolean
   If (m_hTreeView) Then
      NodeExpanded = pTVState(hNode, TVIS_EXPANDED)
   End If
End Property
'' # Risky #
'Friend Property Let NodeExpanded(ByVal hNode As Long, ByVal New_NodeExpanded As Boolean)
'   If (m_hTreeView) Then
'      pTVState(hNode, TVIS_EXPANDED) = New_NodeExpanded
'   End If
'End Property

Public Property Get NodeExpandedOnce(ByVal hNode As Long) As Boolean
   If (m_hTreeView) Then
      NodeExpandedOnce = pTVState(hNode, TVIS_EXPANDEDONCE)
   End If
End Property
'' # Risky #: see Collapse method
'Friend Property Let NodeExpandedOnce(ByVal hNode As Long, ByVal New_NodeExpandedOnce As Boolean)
'   If (m_hTreeView) Then
'      pTVState(hNode, TVIS_EXPANDEDONCE) = New_NodeExpandedOnce
'   End If
'End Property

'== Node count (Total/Children)

Public Property Get NodeCount() As Long
   NodeCount = m_lNodeCount
   Debug.Assert (m_lNodeCount = SendMessageLong(m_hTreeView, TVM_GETCOUNT, 0, 0))
End Property

Public Property Get NodeChildren(ByVal hNode As Long) As Long

  Dim hNext As Long

    If (m_hTreeView) Then
        hNext = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_CHILD, hNode)
        If (hNext) Then
            Do
                NodeChildren = NodeChildren + 1
                hNext = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_NEXT, hNext)
            Loop Until hNext = 0
        End If
    End If
End Property

'== Node full path / Node level

Public Property Get NodeFullPath(ByVal hNode As Long, _
                                 Optional ByVal PathSeparator As String = PATH_SEPARATOR _
                                 ) As String

  Dim hParent As Long

    If (m_hTreeView) Then

        If (hNode) Then

            NodeFullPath = m_uNodeData(pTVlParam(hNode)).sText

            hParent = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_PARENT, hNode)
            Do While hParent
                NodeFullPath = m_uNodeData(pTVlParam(hParent)).sText & _
                               PathSeparator & NodeFullPath
                hParent = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_PARENT, hParent)
            Loop
        End If
    End If
End Property

Public Property Get NodeLevel(ByVal hNode As Long) As Long

  Dim hParent As Long

    If (m_hTreeView) Then

        If (hNode) Then

            NodeLevel = 1

            hParent = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_PARENT, hNode)
            Do While hParent
                NodeLevel = NodeLevel + 1
                hParent = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_PARENT, hParent)
            Loop
        End If
    End If
End Property

' RealFont: returns rect based on NodeFont & TextIndent.
'           Left & Right may differ from normal rcItem.
Public Function NodeRect(ByVal hNode As Long, _
                         Optional Left As Long, Optional Top As Long, _
                         Optional Right As Long, Optional Bottom As Long, _
                         Optional ByVal OnlyText As Boolean = False, _
                         Optional ByVal RealFont As Boolean = False, _
                         Optional ByVal ScaleMode As ScaleModeConstants = vbPixels _
                         ) As Boolean
   Dim rcItem As RECT2
   Dim rcJunk As RECT2

   NodeRect = pTVItemRect(hNode, rcItem, OnlyText)
   
#If CUSTDRAW Then
   If RealFont And OnlyText Then
      rcItem = pGetItemRectReal(hNode, rcItem, rcJunk)
   End If
#End If

   Left = rcItem.X1:  Top = rcItem.Y1:  Right = rcItem.X2:  Bottom = rcItem.Y2
   If ScaleMode <> vbPixels Then
      Left = UserControl.ScaleX(Left, vbPixels, ScaleMode)
      Top = UserControl.ScaleY(Top, vbPixels, ScaleMode)
      Right = UserControl.ScaleX(Right, vbPixels, ScaleMode)
      Bottom = UserControl.ScaleY(Bottom, vbPixels, ScaleMode)
   End If
End Function


'========================================================================================
' Properties: TreeView (appearance/styles)
'========================================================================================

Public Property Get hwnd() As Long
    hwnd = m_hTreeView
End Property

Public Property Get hdc() As Long
   hdc = m_HDC
End Property

Public Property Get BorderStyle() As tvBorderStyleConstants
    If (m_hTreeView) Then
        BorderStyle = -((GetWindowLong(m_hTreeView, GWL_EXSTYLE) And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE)
    End If
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As tvBorderStyleConstants)
    If (m_hTreeView) Then
        Select Case New_BorderStyle
            Case [bsNone]
                Call SetWindowLong(m_hTreeView, GWL_EXSTYLE, 0)
            Case [bsFixedSingle]
                Call SetWindowLong(m_hTreeView, GWL_EXSTYLE, WS_EX_CLIENTEDGE)
        End Select
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    If (m_hTreeView) Then
        BackColor = SendMessageLong(m_hTreeView, TVM_GETBKCOLOR, 0, 0)
    End If
End Property
Public Property Let BackColor(ByVal New_Backcolor As OLE_COLOR)
    If (m_hTreeView) Then
        Call SendMessageLong(m_hTreeView, TVM_SETBKCOLOR, 0, pTranslateColor(New_Backcolor))
    End If
#If CUSTDRAW Then
    m_lTreeColors(clrTreeBK) = pTranslateColor(New_Backcolor)
#End If
End Property

Public Property Get ForeColor() As OLE_COLOR
    If (m_hTreeView) Then
        ForeColor = SendMessageLong(m_hTreeView, TVM_GETTEXTCOLOR, 0, 0)
    End If
End Property
Public Property Let ForeColor(ByVal New_Forecolor As OLE_COLOR)
    If (m_hTreeView) Then
        Call SendMessageLong(m_hTreeView, TVM_SETTEXTCOLOR, 0, pTranslateColor(New_Forecolor))
    End If
#If CUSTDRAW Then
    m_lTreeColors(clrTree) = pTranslateColor(New_Forecolor)
#End If
End Property

Public Property Get LineColor() As OLE_COLOR
    If (m_hTreeView) Then
        LineColor = SendMessageLong(m_hTreeView, TVM_GETLINECOLOR, 0, 0)
        If LineColor = CLR_DEFAULT Then
            ' not specified, return used default color
            LineColor = GetSysColor(COLOR_GRAYTEXT)
        End If
    End If
End Property
Public Property Let LineColor(ByVal New_LineColor As OLE_COLOR)
    If (m_hTreeView) Then
        Call SendMessageLong(m_hTreeView, TVM_SETLINECOLOR, 0, pTranslateColor(New_LineColor))
    End If
End Property

Public Property Get InsertMarkColor() As OLE_COLOR
    If (m_hTreeView) Then
        InsertMarkColor = SendMessageLong(m_hTreeView, TVM_GETINSERTMARKCOLOR, 0, 0)
    End If
End Property
Public Property Let InsertMarkColor(ByVal New_InsertMarkColor As OLE_COLOR)
    If (m_hTreeView) Then
        Call SendMessageLong(m_hTreeView, TVM_SETINSERTMARKCOLOR, 0, pTranslateColor(New_InsertMarkColor))
    End If
End Property

' UC uses initially StdFont for TreeView.
' If Font is set, TreeView responds to any subsequent changes in passed font.
' Pass Nothing to break the coupling -> initial StdFont is reused.
' Set the passed font to Nothing     -> tree keeps properties of font.
' # COMCTL6: Setting Font may change ItemHeight & ItemIndent #
Public Property Get Font() As StdFont
   Set Font = m_oFont
End Property
Public Property Set Font(ByVal New_Font As StdFont)
   ' set external m_oFont
   Set m_oFont = New_Font
   m_oFont_FontChanged vbNullString
End Property

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)

   If (m_oFont Is Nothing) Then
      Set m_oFont = New StdFont
      m_oFont.Name = "MS Shell Dlg"
   End If
   
   ' internal m_iFont de-coupled from external m_oFont
   Set m_iFont = Nothing
   Set m_iFont = pFontClone(m_oFont)
   m_hFont = m_iFont.hFont

#If CUSTDRAW Then
   pvStdFontToLogFont m_oFont, m_iFontLF
   
   If m_HDC Then
      pUpdateRcTreeFonts
      m_lDumCharW = pGetRcTreeFont(DUMCHAR)
   End If
#End If
   
   If m_hTreeView Then
      Call SendMessage(m_hTreeView, WM_SETFONT, m_hFont, ByVal 1&)
   End If
    
#If FNT_DBG Then
    Debug.Print "NEWFONT", m_oFont.Name, m_oFont.Size, m_oFont.Bold
#End If
End Sub

Private Function pFontClone(fnt As IFont) As IFont
   fnt.Clone pFontClone
End Function

' ItemHeight rounded down to next even or set TVS_NONEVENHEIGHT
Public Property Get ItemHeight() As Long
   If (m_hTreeView) Then
      ItemHeight = SendMessageLong(m_hTreeView, TVM_GETITEMHEIGHT, 0, 0)
   End If
End Property
Public Property Let ItemHeight(ByVal New_ItemHeight As Long)
   If (m_hTreeView) Then
      Call SendMessageLong(m_hTreeView, TVM_SETITEMHEIGHT, New_ItemHeight, 0)
   End If
End Property

Public Property Get ItemIndent() As Long
   If (m_hTreeView) Then
      ItemIndent = SendMessageLong(m_hTreeView, TVM_GETINDENT, 0, 0)
   End If
End Property
Public Property Let ItemIndent(ByVal New_ItemIndent As Long)
   If (m_hTreeView) Then
      Call SendMessageLong(m_hTreeView, TVM_SETINDENT, New_ItemIndent, 0)
   End If
End Property

'//

Public Property Get CheckBoxes() As Boolean
   CheckBoxes = m_bCheckBoxes
End Property
Public Property Let CheckBoxes(ByVal New_CheckBoxes As Boolean)

   Dim lNode As Long

   If (m_hTreeView) Then
      m_bCheckBoxes = New_CheckBoxes
      If (m_bCheckBoxes) Then
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_CHECKBOXES, 0&)
         m_hImageList(ilState) = SendMessageLong(m_hTreeView, TVM_GETIMAGELIST, TVSIL_STATE, 0&)
         m_lImageListCount(ilState) = ImageList_GetImageCount(m_hImageList(ilState))
      Else
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0&, TVS_CHECKBOXES)
         For lNode = 1 To m_lNodeCount
            pTVStateImage(m_uNodeData(lNode).hNode) = 0&
         Next lNode
         pvDestroyImageList ilState
         '* In fact, TVS_CHECKBOXES is not removed.
         '  TreeView window should be destroyed and created again without this style.
         '  This is an intermediate solution to avoid re-add all nodes and their data.
         '  You'll observe that window max. right coordinate is not updated.
      End If
   End If
End Property

' MSDN: TVS_FULLROWSELECT cannot be used in conjunction with the TVS_HASLINES style.
Public Property Get FullRowSelect() As Boolean
   FullRowSelect = m_bFullRowSelect
End Property
Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
   If (m_hTreeView) Then
      m_bFullRowSelect = New_FullRowSelect
      If (m_bFullRowSelect) Then
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_FULLROWSELECT, 0)
      Else
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_FULLROWSELECT)
      End If
   End If
End Property

Public Property Get HasButtons() As Boolean
   HasButtons = m_bHasButtons
End Property
Public Property Let HasButtons(ByVal New_HasButtons As Boolean)
   If (m_hTreeView) Then
      m_bHasButtons = New_HasButtons
      If (m_bHasButtons) Then
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_HASBUTTONS, 0)
      Else
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_HASBUTTONS)
      End If
   End If
End Property

Public Property Get HasLines() As Boolean
   HasLines = m_bHasLines
End Property
Public Property Let HasLines(ByVal New_HasLines As Boolean)
   If (m_hTreeView) Then
      m_bHasLines = New_HasLines
      If m_bHasLines Then
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_HASLINES, 0)
      Else
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_HASLINES)
      End If
   End If
#If CUSTDRAW Then
   If m_bHasLines Then
      ' Create pattern brush for dotted lines once
      pCreateDotBrush
   Else
      pDestroyDotBrush
   End If
#End If
End Property

Public Property Get HasRootLines() As Boolean
   HasRootLines = m_bHasRootLines
End Property
Public Property Let HasRootLines(ByVal New_HasRootLines As Boolean)
   If (m_hTreeView) Then
      m_bHasRootLines = New_HasRootLines
      If (m_bHasRootLines) Then
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_LINESATROOT, 0)
      Else
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_LINESATROOT)
      End If
   End If
End Property

Public Property Get HideSelection() As tvSelectionFocusConstants
   HideSelection = m_eHideSelection
End Property
Public Property Let HideSelection(ByVal New_HideSelection As tvSelectionFocusConstants)
   If (m_hTreeView) Then
      m_eHideSelection = New_HideSelection
      If (m_eHideSelection = sfHideSelection) Then
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_SHOWSELALWAYS)
      Else
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_SHOWSELALWAYS, 0)
      End If
   End If
End Property

Public Property Get LabelEdit() As Boolean
   LabelEdit = m_bLabelEdit
End Property
Public Property Let LabelEdit(ByVal New_LabelEdit As Boolean)
   If (m_hTreeView) Then
      m_bLabelEdit = New_LabelEdit
      If (m_bLabelEdit) Then
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_EDITLABELS, 0)
      Else
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_EDITLABELS)
      End If
   End If
End Property

Public Property Get SingleExpand() As Boolean
   SingleExpand = m_bSingleExpand
End Property
Public Property Let SingleExpand(ByVal New_SingleExpand As Boolean)
   If (m_hTreeView) Then
      m_bSingleExpand = New_SingleExpand
      If (m_bSingleExpand) Then
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_SINGLEEXPAND, 0)
      Else
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_SINGLEEXPAND)
      End If
   End If
End Property

' comctl default: hand cursor
Public Property Get TrackSelect(Optional ByVal UseStandardCursor As Boolean) As Boolean
   TrackSelect = m_bTrackSelect
End Property
Public Property Let TrackSelect(Optional ByVal UseStandardCursor As Boolean, _
                                ByVal New_TrackSelect As Boolean)
   If (m_hTreeView) Then
      m_bTrackSelect = New_TrackSelect
      m_bUseStandardCursor = UseStandardCursor And m_bTrackSelect
      If (m_bTrackSelect) Then
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_TRACKSELECT, 0)
         If m_bUseStandardCursor Then
            Call Subclass_AddMsg(m_hTreeView, WM_SETCURSOR, MSG_BEFORE)
         End If
      Else
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_TRACKSELECT)
      End If
   End If
End Property

Public Property Get ToolTips() As Boolean
   If (m_hTreeView) Then
      ToolTips = ((GetWindowLong(m_hTreeView, GWL_STYLE) And TVS_NOTOOLTIPS) = 0)
   Else
      ' treeview created with tooltips as default
      ToolTips = True
   End If
End Property
Public Property Let ToolTips(ByVal bToolTips As Boolean)
   If (m_hTreeView) Then
      If (bToolTips) Then
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_NOTOOLTIPS)
      Else
         Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_NOTOOLTIPS, 0)
      End If
   End If
End Property



'========================================================================================
' Private
'========================================================================================

'== Creating/destroying TreeView / Changing window styles

Private Function pIsNewComctl32() As Boolean
   ' ensures that the Comctl32.dll library is loaded (Brad Martinez)
   Dim icc As tagINITCOMMONCONTROLSEX

   On Error GoTo Err_InitOldVersion

   icc.dwSize = Len(icc)
   icc.dwICC = ICC_TREEVIEW_CLASSES

   ' err 453 "Specified DLL function not found", if the new version isn't installed
   pIsNewComctl32 = InitCommonControlsEx(icc)

   Exit Function

Err_InitOldVersion:
   InitCommonControls
End Function

Private Function pvCreateTreeView() As Boolean

   Dim lExStyle As Long
   Dim lTVStyle As Long

   '-- Define window style
   lTVStyle = WS_CHILD Or WS_TABSTOP Or TVS_SHOWSELALWAYS Or TVS_DISABLEDRAGDROP _
              Or TVS_INFOTIP ' Or TVS_RTLREADING
   lExStyle = WS_EX_CLIENTEDGE

   '-- Create TreeView window
   m_hTreeView = CreateWindowEx(lExStyle, WC_TREEVIEW, vbNullString, lTVStyle, _
                                0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight, _
                                UserControl.hwnd, 0&, App.hInstance, ByVal 0&)
   
   '-- Success [?]
   If (m_hTreeView) Then
      
#If UNICODE Then
      Call SendMessageLong(m_hTreeView, CCM_SETUNICODEFORMAT, 1&, 0&)
      Debug.Assert SendMessageLong(m_hTreeView, CCM_GETUNICODEFORMAT, 0&, 0&) <> 0&
#End If
      
      m_lComctlVersion = COMCTL32_VERSION
      ' m_lComctlVersion will be zero for COMCTL < 5.0
      m_lComctlVersion = SendMessageLong(m_hTreeView, CCM_SETVERSION, m_lComctlVersion, 0&)
      ' Win2K : 0 / WinXP : 0 / 6 (manifest)
      Debug.Print "CCM_SETVERSION: " & m_lComctlVersion
      m_lComctlVersion = SendMessageLong(m_hTreeView, CCM_GETVERSION, 0&, 0&)
      ' Win2K : 5 / WinXP : 5 / 6 (manifest)
      Debug.Print "CCM_GETVERSION: " & m_lComctlVersion
      
      '-- Get DC handle
      m_HDC = GetDC(m_hTreeView)

      '-- Get internal State Imagelist handle
      m_hImageList(ilState) = SendMessageLong(m_hTreeView, TVM_GETIMAGELIST, TVSIL_STATE, 0&)
      Debug.Assert m_bExtImagelist(ilState) = False
      
      '-- System background and foreground colors
      Call SendMessageLong(m_hTreeView, TVM_SETBKCOLOR, 0, GetSysColor(COLOR_WINDOW))
      Call SendMessageLong(m_hTreeView, TVM_SETTEXTCOLOR, 0, GetSysColor(COLOR_WINDOWTEXT))

      '-- Set (ambient) font
      Call SendMessageLong(m_hTreeView, WM_SETFONT, m_hFont, 1&)
#If CUSTDRAW Then
      m_lDumCharW = pGetRcTreeFont(DUMCHAR)
#End If
      
      '-- Show TreeView
      Call ShowWindow(m_hTreeView, SW_SHOW)
      pvCreateTreeView = True
   End If
End Function

Private Sub pvDestroyTreeView()

   If (m_hTreeView) Then
      If m_HDC Then
         If ReleaseDC(m_hTreeView, m_HDC) Then  ' BUGFIX8
            m_HDC = 0
         End If
      End If
      If (DestroyWindow(m_hTreeView)) Then
         m_hTreeView = 0
      End If
   End If
   Debug.Assert (m_hTreeView + m_HDC) = 0
End Sub

Private Sub pvSetWndStyle(ByVal hwnd As Long, ByVal lType As Long, ByVal lStyle As Long, ByVal lStyleNot As Long)

  Dim lS As Long

    lS = GetWindowLong(hwnd, lType)
    lS = (lS And Not lStyleNot) Or lStyle
    Call SetWindowLong(hwnd, lType, lS)
    Call SetWindowPos(hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub

'== Image list and logical font

Private Function pvSetImageList(ByVal hImageList As Long, _
                                Optional ByVal ILS As tvImagelistConstants = ilNormal) As Long
   If ILS = ilNormal Then
      '-- 'Normal' image list
      pvSetImageList = SendMessageLong(m_hTreeView, TVM_SETIMAGELIST, TVSIL_NORMAL, hImageList)
   Else
      '-- State image list
      pvSetImageList = SendMessageLong(m_hTreeView, TVM_SETIMAGELIST, TVSIL_STATE, hImageList)
   End If
End Function

Private Function pvDestroyImageList(Optional ByVal ILS As tvImagelistConstants = ilNormal) As Boolean
   
   If m_hImageList(ILS) Then
      If m_bExtImagelist(ILS) Then
         ' don't destroy previous external imagelist
         pvDestroyImageList = True
      Else
         pvDestroyImageList = CBool(ImageList_Destroy(m_hImageList(ILS)))
      End If
      Debug.Assert pvDestroyImageList
   End If
   
   m_bExtImagelist(ILS) = False
   m_hImageList(ILS) = 0&
   m_lImageListCount(ILS) = 0&
End Function

Private Sub pvStdFontToLogFont(fnt As StdFont, tLF As LOGFONT)
   Const PointsPerTwip = 20& ' (1440 / 72)
   
   ' Convert an OLE StdFont to a LOGFONT structure:
   With tLF
#If UNICODE Then
      .lfFaceName = fnt.Name & vbNullChar
#Else
      Dim b()   As Byte
      b = StrConv(fnt.Name, vbFromUnicode)
      ReDim Preserve b(0 To 31) As Byte
      CopyMemory .lfFaceName(0), b(0), 32&
#End If
      .lfHeight = -fnt.Size * PointsPerTwip \ Screen.TwipsPerPixelY
      .lfItalic = fnt.Italic
      .lfWeight = fnt.Weight ' IIf(fnt.Bold, FW_BOLD, FW_NORMAL)
      .lfUnderline = fnt.Underline
      .lfStrikeOut = fnt.Strikethrough
      .lfCharSet = fnt.Charset
'      .lfQuality = 3 'NONANTIALIASED_QUALITY
   End With
End Sub

' Extracts metrics from LOGFONT structure to create and return a new StdFont object.
Private Function pvCreateFont(ByRef lf As LOGFONT) As StdFont
   Const FW_BOLD = 700&
   Const PointsPerTwip = 20& ' (1440 / 72)
   Dim fnt  As StdFont
   
   Set fnt = New StdFont
   With fnt
      .Bold = (lf.lfWeight >= FW_BOLD)
      .Charset = lf.lfCharSet
      .Italic = CBool(lf.lfItalic)
#If UNICODE Then
      .Name = Left$(lf.lfFaceName, InStr(lf.lfFaceName, vbNullChar) - 1)
#Else
      .Name = StrConv(lf.lfFaceName, vbUnicode)
#End If
      .Size = -(lf.lfHeight * Screen.TwipsPerPixelY \ PointsPerTwip)
      .Strikethrough = lf.lfStrikeOut
      .Underline = lf.lfUnderline
      .Weight = lf.lfWeight
   End With
   
   Set pvCreateFont = fnt
End Function

Private Function pTranslateColor(ByVal clr As OLE_COLOR) As Long
    If OleTranslateColor(clr, 0&, pTranslateColor) Then
        pTranslateColor = CLR_NONE
    End If
End Function


'//

Private Function pvTVAddNode(idxNode As Long, hRelative As Long, _
                             eRelation As Long, bForcePlusButton As Boolean) As Long

   Dim uTVIS     As TVINSERTSTRUCT
   Dim hPrevious As Long

   With uTVIS

      With .Item
         .mask = TVIF_ALL
         .lParam = idxNode
         
         .pszText = LPSTR_TEXTCALLBACK
         
         If m_bCustomDraw Then
            ' prevent TVN_GETDISPINFO calls for images
            .iImage = -2&
            .iSelectedImage = -2&
         Else
            .iImage = I_IMAGECALLBACK
            .iSelectedImage = I_IMAGECALLBACK
         End If
         
         If (bForcePlusButton) Then
             .cChildren = 1
             .mask = .mask Or TVIF_CHILDREN
         End If
      End With
        
      If (hRelative) Then
         .hParent = hRelative
      Else
         .hParent = TVI_ROOT
      End If
      
      Select Case eRelation
         Case [rFirst]
            .hInsertAfter = TVI_FIRST
         Case [rLast]
            .hInsertAfter = TVI_LAST
         Case [rSort]
            .hInsertAfter = TVI_SORT
            m_bInSort = True
         Case [rNext]
            .hParent = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_PARENT, hRelative)
            .hInsertAfter = hRelative
         Case [rPrevious]
            .hParent = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_PARENT, hRelative)
            hPrevious = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_PREVIOUS, hRelative)
            If (hPrevious) Then
               .hInsertAfter = hPrevious
            Else
               .hInsertAfter = TVI_FIRST
            End If
         Case Else
            .hInsertAfter = TVI_LAST
      End Select
   End With

   pvTVAddNode = SendMessage(m_hTreeView, TVM_INSERTITEM, 0, uTVIS)
   m_bInSort = False
End Function

Private Property Get pTVlParam(hNode As Long) As Long
   Dim uTVI As TVITEM

   With uTVI
      .hItem = hNode
      .mask = TVIF_PARAM
   End With
   
   If (SendMessage(m_hTreeView, TVM_GETITEM, 0, uTVI)) Then
      pTVlParam = uTVI.lParam
   End If
End Property

Private Property Get pTVStateImage(hNode As Long) As Long
   Dim uTVI As TVITEM

   With uTVI
      .hItem = hNode
      .mask = TVIF_HANDLE Or TVIF_STATE
      .stateMask = TVIS_STATEIMAGEMASK
   End With
   
   If (SendMessage(m_hTreeView, TVM_GETITEM, 0, uTVI)) Then
      pTVStateImage = pvSTATEIMAGEMASKTOINDEX(uTVI.State And TVIS_STATEIMAGEMASK)
   End If
End Property
Private Property Let pTVStateImage(hNode As Long, lIndex As Long)
   Dim uTVI As TVITEM
   
   With uTVI
      .hItem = hNode
      .mask = TVIF_HANDLE Or TVIF_STATE
      .stateMask = TVIS_STATEIMAGEMASK
      .State = pvINDEXTOSTATEIMAGEMASK(lIndex)
   End With
   
   Call SendMessage(m_hTreeView, TVM_SETITEM, 0, uTVI)
End Property
Private Function pvINDEXTOSTATEIMAGEMASK(lIndex As Long) As Long
   pvINDEXTOSTATEIMAGEMASK = lIndex * (2 ^ 12)
End Function
Private Function pvSTATEIMAGEMASKTOINDEX(lState As Long) As Long
   pvSTATEIMAGEMASKTOINDEX = lState / (2 ^ 12)
End Function

Private Property Get pTVState(hNode As Long, lState As Long) As Boolean
   Dim uTVI As TVITEM
   
   Debug.Assert hNode > 0
   
   With uTVI
      .hItem = hNode
      .mask = TVIF_HANDLE Or TVIF_STATE
   End With
   
   If (SendMessage(m_hTreeView, TVM_GETITEM, 0, uTVI)) Then
      pTVState = (uTVI.State And lState)
   End If
End Property
Private Property Let pTVState(hNode As Long, lState As Long, fAdd As Boolean)

  Dim uTVI As TVITEM

    With uTVI
        .hItem = hNode
        .mask = TVIF_HANDLE Or TVIF_STATE
        .stateMask = lState
        .State = fAdd And lState
    End With

    Call SendMessage(m_hTreeView, TVM_SETITEM, 0, uTVI)
End Property

Private Property Get pTVcChildren(hNode As Long) As Long
   Dim uTVI As TVITEM

   With uTVI
      .hItem = hNode
      .mask = TVIF_HANDLE Or TVIF_CHILDREN
   End With
   
   If (SendMessage(m_hTreeView, TVM_GETITEM, 0, uTVI)) Then
      pTVcChildren = uTVI.cChildren
   End If
End Property
Private Property Let pTVcChildren(hNode As Long, cChildren As Long)
   Dim uTVI As TVITEM

   With uTVI
      .hItem = hNode
      .mask = TVIF_HANDLE Or TVIF_CHILDREN
      .cChildren = cChildren
   End With
   
   Call SendMessage(m_hTreeView, TVM_SETITEM, 0, uTVI)
End Property

Private Property Get pTVHasChildren(hNode As Long) As Boolean

    pTVHasChildren = CBool(SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_CHILD, hNode))
End Property

Private Function pTVItemRect(hNode As Long, uRect As RECT2, _
                                 Optional ByVal OnlyText As Boolean = False) As Boolean
    Let uRect.X1 = hNode
    pTVItemRect = SendMessage(m_hTreeView, TVM_GETITEMRECT, -OnlyText, uRect)
End Function

Private Property Get pTVOverlayImage(hNode As Long) As Long
   Dim uTVI As TVITEM

   With uTVI
      .hItem = hNode
      .mask = TVIF_HANDLE Or TVIF_STATE
      .stateMask = TVIS_OVERLAYMASK
   End With
   
   If (SendMessage(m_hTreeView, TVM_GETITEM, 0, uTVI)) Then
      pTVOverlayImage = pvOVERLAYMASKTOINDEX(uTVI.State And TVIS_OVERLAYMASK)
   End If
End Property
Private Property Let pTVOverlayImage(hNode As Long, lIndex As Long)
   Dim uTVI As TVITEM
   
   With uTVI
      .hItem = hNode
      .mask = TVIF_HANDLE Or TVIF_STATE
      .stateMask = TVIS_OVERLAYMASK
      .State = pvINDEXTOOVERLAYMASK(lIndex)
   End With
   
   Call SendMessage(m_hTreeView, TVM_SETITEM, 0, uTVI)
End Property
Private Function pvINDEXTOOVERLAYMASK(lIndex As Long) As Long
   pvINDEXTOOVERLAYMASK = lIndex * (2 ^ 8)
End Function
Private Function pvOVERLAYMASKTOINDEX(lState As Long) As Long
   pvOVERLAYMASKTOINDEX = lState / (2 ^ 8)
End Function

'//

'== Post-processing of deleted node(s)

Private Sub pvDoDeletePostProcess()

  Dim lNode As Long
  Dim lLast As Long 'Last empty
  Dim uTVI  As TVITEM

    '-- Remove collection items
    For lNode = m_lNodeCount To 1 Step -1
        If (m_uNodeData(lNode).hNode = 0) Then
            ' permit adding unkeyed nodes
            If LenB(m_uNodeData(lNode).sKey) Then
               Call m_cKey.Remove(m_uNodeData(lNode).sKey)
            End If
        End If
    Next lNode

    '-- Remove (move) array items
    uTVI.mask = TVIF_PARAM
    lLast = 0
    For lNode = 1 To m_lNodeCount
        If (m_uNodeData(lNode).hNode = 0) Then
            If lLast = 0 Then lLast = lNode
          Else
            If (lLast) Then
                m_uNodeData(lLast) = m_uNodeData(lNode)
                With uTVI
                    .hItem = m_uNodeData(lLast).hNode
                    .lParam = lLast
                End With
                Call SendMessage(m_hTreeView, TVM_SETITEM, 0, uTVI)
                lLast = lLast + 1
            End If
        End If
    Next lNode
    If (lLast) Then
        m_lNodeCount = lLast - 1
    End If

    '-- Resize array
    ReDim Preserve m_uNodeData(0 To m_lNodeCount + ((ALLOCATE_SIZE + 1) - m_lNodeCount Mod ALLOCATE_SIZE))
End Sub

'== String

' lstrlen assumes that lpString is a NULL-terminated string !!!
Private Function pStringFromPointer(ByVal lpString As Long) As String

#If UNICODE Then
   Dim nLen As Long
   
   If lpString Then
      nLen = lstrlenW(ByVal lpString)
      If nLen Then
         ' allocate string with nLen chars
         pStringFromPointer = String$(nLen, 0)
'         ' copy 2x nLen bytes for Unicode
'         CopyMemory ByVal StrPtr(pStringFromPointer), ByVal lpString, 2 * nLen
         lstrcpyW ByVal StrPtr(pStringFromPointer), ByVal lpString
      End If
   End If
   
#Else
   Dim nLen As Long
   Dim b()  As Byte
   
   If lpString Then
      nLen = lstrlenA(ByVal lpString)
      If nLen Then
         ' allocate buffer with nLen bytes
         ReDim b(0 To nLen - 1) As Byte
         ' copy nLen bytes for ANSI
         CopyMemory b(0), ByVal lpString, nLen
         pStringFromPointer = StrConv(b(), vbUnicode)
      End If
   End If
#End If
End Function

' copy string to !existing! pointer with buffer size nLen
Private Sub pStringToPointer(sText As String, ByVal nLen As Long, ByRef lpString As Long)

#If UNICODE Then
   ' pad to buffer size nLen chars
   If nLen > Len(sText) Then
      lstrcpyW ByVal lpString, ByVal StrPtr(sText & String$(nLen - Len(sText), vbNullChar))
   Else: Debug.Assert False
   End If
#Else
   Dim b()  As Byte
   
   b = StrConv(sText, vbFromUnicode)
   ' pad to buffer size nLen bytes
   ReDim Preserve b(0 To nLen - 1) As Byte
   CopyMemory ByVal lpString, b(0), nLen
#End If

End Sub

'== Misc

Private Function pvButton(ByVal uMsg As Long) As MouseButtonConstants

    Select Case uMsg
        Case WM_LBUTTONDOWN, WM_LBUTTONUP, WM_LBUTTONDBLCLK ' NM_DBLCLK
            pvButton = vbLeftButton
        Case WM_RBUTTONDOWN, WM_RBUTTONUP, WM_RBUTTONDBLCLK ' NM_RDBLCLK
            pvButton = vbRightButton
        Case WM_MBUTTONDOWN, WM_MBUTTONUP, WM_MBUTTONDBLCLK
            pvButton = vbMiddleButton
        Case WM_MOUSEMOVE
            Select Case True
                Case GetAsyncKeyState(vbKeyLButton) < 0
                    pvButton = vbLeftButton
                Case GetAsyncKeyState(vbKeyRButton) < 0
                    pvButton = vbRightButton
                Case GetAsyncKeyState(vbKeyMButton) < 0
                    pvButton = vbMiddleButton
            End Select
    End Select
End Function

Private Function pvShiftState() As ShiftConstants

  Dim lS As ShiftConstants

    If (GetAsyncKeyState(vbKeyShift) < 0) Then
        lS = lS Or vbShiftMask
    End If
    If (GetAsyncKeyState(vbKeyMenu) < 0) Then
        lS = lS Or vbAltMask
    End If
    If (GetAsyncKeyState(vbKeyControl) < 0) Then
        lS = lS Or vbCtrlMask
    End If
    pvShiftState = lS
End Function

'----------------------------------------------------------------------------------------
' MouseEnter/Leave support
'----------------------------------------------------------------------------------------

Private Function pvIsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
'-- Determine if the passed function is supported

  Dim hMod       As Long
  Dim bLibLoaded As Boolean

    hMod = GetModuleHandleA(sModule)

    If (hMod = 0) Then
        hMod = LoadLibraryA(sModule)
        If (hMod) Then
            bLibLoaded = True
        End If
    End If

    If (hMod) Then
        If (GetProcAddress(hMod, sFunction)) Then
            pvIsFunctionExported = True
        End If
    End If

    If (bLibLoaded) Then
        Call FreeLibrary(hMod)
    End If
End Function

Private Sub pvTrackMouseLeave(ByVal lng_hWnd As Long)
'-- Track the mouse leaving the indicated window

  Dim uTME As TRACKMOUSEEVENT_STRUCT

    If (m_bTrack) Then

        With uTME
            .cbSize = Len(uTME)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With

        If (m_bTrackUser32) Then
            Call TrackMouseEvent(uTME)
          Else
            Call TrackMouseEventComCtl(uTME)
        End If
    End If
End Sub

'----------------------------------------------------------------------------------------
' OLEInPlaceActivateObject interface
'----------------------------------------------------------------------------------------

Private Sub pvSetIPAO()

  Dim pOleObject          As IOleObject
  Dim pOleInPlaceSite     As IOleInPlaceSite
  Dim pOleInPlaceFrame    As IOleInPlaceFrame
  Dim pOleInPlaceUIWindow As IOleInPlaceUIWindow
  Dim rcPos               As RECT2
  Dim rcClip              As RECT2
  Dim uFrameInfo          As OLEINPLACEFRAMEINFO

    On Error Resume Next

    Set pOleObject = Me
    Set pOleInPlaceSite = pOleObject.GetClientSite

    If (Not pOleInPlaceSite Is Nothing) Then
        Call pOleInPlaceSite.GetWindowContext(pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(rcPos), VarPtr(rcClip), VarPtr(uFrameInfo))
        If (Not pOleInPlaceFrame Is Nothing) Then
            Call pOleInPlaceFrame.SetActiveObject(m_uIPAO.ThisPointer, vbNullString)
        End If
        If (Not pOleInPlaceUIWindow Is Nothing) Then 'And Not m_bMouseActivate
            Call pOleInPlaceUIWindow.SetActiveObject(m_uIPAO.ThisPointer, vbNullString)
          Else
            Call pOleObject.DoVerb(OLEIVERB_UIACTIVATE, 0, pOleInPlaceSite, 0, m_hUC, VarPtr(rcPos))
        End If
    End If

    On Error GoTo 0
End Sub

Friend Function frTranslateAccel(pMsg As Msg) As Boolean

   Dim pOleObject      As IOleObject
   Dim pOleControlSite As IOleControlSite

   On Error Resume Next
   
   With pMsg

      Select Case .message

         Case WM_KEYDOWN, WM_KEYUP

            Select Case .wParam
               Case vbKeyReturn, vbKeyEscape
                  ' a button on a form with Cancel(Escape) or Default(Return) property
                  ' set true fires, if msg's not handled here. # Outcomment otherwise #
                  If m_hEdit Then
                     ' end or cancel label editing
                     Call SendMessageLong(m_hEdit, .message, .wParam, .lParam)
                     frTranslateAccel = True
                  Else
                     Call SendMessageLong(m_hTreeView, .message, .wParam, .lParam)
                     frTranslateAccel = True
                  End If

               Case vbKeyTab

                  If (pvShiftState() And vbCtrlMask) Then
                     Set pOleObject = Me
                     Set pOleControlSite = pOleObject.GetClientSite
                     If (Not pOleControlSite Is Nothing) Then
                        Call pOleControlSite.TranslateAccelerator(VarPtr(pMsg), pvShiftState() And vbShiftMask)
                     End If
                  End If
                  frTranslateAccel = False

               Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyPageDown, vbKeyPageUp
                  If m_hEdit Then
                     Call SendMessageLong(m_hEdit, .message, .wParam, .lParam)
                  Else
                     Call SendMessageLong(m_hTreeView, .message, .wParam, .lParam)
                  End If
                  frTranslateAccel = True
            
            End Select

      End Select

   End With
End Function

'----------------------------------------------------------------------------------------
' Subclass code - The programmer may call any of the following Subclass_??? routines
'----------------------------------------------------------------------------------------

Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
'Parameters:
'   lng_hWnd - The handle of the window for which the uMsg is to be added to the callback table
'   uMsg     - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
'   When     - Whether the msg is to callback before, after or both with respect to the the default (previous) handler

    With sc_aSubData(zIdx(lng_hWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
''Delete a message from the table of those that will invoke a callback.
''Parameters:
''   lng_hWnd - The handle of the window for which the uMsg is to be removed from the callback table
''   uMsg     - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
''   When     - Whether the msg is to be removed from the before, after or both callback tables
'
'    With sc_aSubData(zIdx(lng_hWnd))
'        If (When And eMsgWhen.MSG_BEFORE) Then
'            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
'        End If
'        If (When And eMsgWhen.MSG_AFTER) Then
'            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
'        End If
'    End With
'End Sub

Private Function Subclass_InIDE() As Boolean
'Return whether we're running in the IDE.

    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Start subclassing the passed window handle
'Parameters:
'   lng_hWnd - The handle of the window to be subclassed
'Returns;
'   The sc_aSubData() index

  Dim i                        As Long                       'Loop index
  Dim nSubIdx                  As Long                       'Subclass data index

  Const PUB_CLASSES            As Long = 0                   'The number of UserControl public classes
  Const GMEM_FIXED             As Long = 0                   'Fixed memory GlobalAlloc flag
  Const PAGE_EXECUTE_READWRITE As Long = &H40&               'Allow memory to execute without violating XP SP2 Data Execution Prevention
  Const PATCH_01               As Long = 18                  'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02               As Long = 68                  'Address of the previous WndProc
  Const PATCH_03               As Long = 78                  'Relative address of SetWindowsLong
  Const PATCH_06               As Long = 116                 'Address of the previous WndProc
  Const PATCH_07               As Long = 121                 'Relative address of CallWindowProc
  Const PATCH_0A               As Long = 186                 'Address of the owner object
  Const FUNC_CWP               As String = "CallWindowProcA" 'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM               As String = "EbMode"          'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL               As String = "SetWindowLongA"  'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER               As String = "user32"          'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5               As String = "vba5"            'Location of the EbMode function if running VB5
  Const MOD_VBA6               As String = "vba6"            'Location of the EbMode function if running VB6

    'If it's the first time through here..
    If (sc_aBuf(1) = 0) Then

#If HEXORG = 1 Then
  Dim sSubCode                 As String                     'Subclass code string
  Dim j                        As Long                       'Loop index

        'Build the hex pair subclass string
        sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
                   "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
                   "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
                   "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90" & _
                   Hex$(&HA4 + (PUB_CLASSES * 12)) & "070000C3"

        'Convert the string from hex pairs to bytes and store in the machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            sc_aBuf(j) = CByte("&H" & Mid$(sSubCode, i, 2))                       'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                      'Next pair of hex characters

        'Get API function addresses
        If (Subclass_InIDE) Then                                                  'If we're running in the VB IDE
            sc_aBuf(16) = &H90                                                    'Patch the code buffer to enable the IDE state code
            sc_aBuf(17) = &H90                                                    'Patch the code buffer to enable the IDE state code
            sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                            'Get the address of EbMode in vba6.dll
            If (sc_pEbMode = 0) Then                                              'Found?
                sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                        'VB5 perhaps
            End If
        End If

#Else
       ' Load OpCode from Long array (25 x faster)
       ' (PUB_CLASSES != 0 !!!)
        
        sc_aBuf(1) = &H83E58955:            sc_aBuf(26) = &HFFC30000
        sc_aBuf(2) = &H3157F8C4:            sc_aBuf(27) = &H75FF1475
        sc_aBuf(3) = &HFC4589C0:            sc_aBuf(28) = &HC75FF10
        sc_aBuf(4) = &HEBF84589:            sc_aBuf(29) = &H680875FF
        sc_aBuf(5) = &HE80E:                sc_aBuf(30) = &H0
        sc_aBuf(6) = &HF8830000:            sc_aBuf(31) = &HE8
        sc_aBuf(7) = &H85217402:            sc_aBuf(32) = &HFC458900
        sc_aBuf(8) = &HE82474C0:            sc_aBuf(33) = &HBFD231C3
        sc_aBuf(9) = &H30:                  sc_aBuf(34) = &H0
        sc_aBuf(10) = &HF87D83:             sc_aBuf(35) = &HB9
        sc_aBuf(11) = &H38E80A75:           sc_aBuf(36) = &H1E800
        sc_aBuf(12) = &HE8000000:           sc_aBuf(37) = &HE3C30000
        sc_aBuf(13) = &H4D:                 sc_aBuf(38) = &H78C90932
        sc_aBuf(14) = &HFC458B5F:           sc_aBuf(39) = &HC458B07
        sc_aBuf(15) = &H10C2C9:             sc_aBuf(40) = &H2775AFF2
        sc_aBuf(16) = &H26E8:               sc_aBuf(41) = &H5014458D
        sc_aBuf(17) = &H68F1EB00:           sc_aBuf(42) = &H5010458D
        sc_aBuf(18) = &H0:                  sc_aBuf(43) = &H500C458D
        sc_aBuf(19) = &H75FFFC6A:           sc_aBuf(44) = &H5008458D
        sc_aBuf(20) = &HE808:               sc_aBuf(45) = &H50FC458D
        sc_aBuf(21) = &HE0EB0000:           sc_aBuf(46) = &H50F8458D
        sc_aBuf(22) = &HBF4AD231:           sc_aBuf(47) = &HB852
        sc_aBuf(23) = &H0:                  sc_aBuf(48) = &H8B500000
        sc_aBuf(24) = &HB9:                 sc_aBuf(49) = &HA490FF00
        sc_aBuf(25) = &H2DE800:             sc_aBuf(50) = &HC3000007
  
        'Get API function addresses
        If (Subclass_InIDE) Then                                                  'If we're running in the VB IDE
            sc_aBuf(4) = &H90F84589                                               'Patch the code buffer to enable the IDE state code
            sc_aBuf(5) = &HE890                                                   'Patch the code buffer to enable the IDE state code
            sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                            'Get the address of EbMode in vba6.dll
            If (sc_pEbMode = 0) Then                                              'Found?
                sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                        'VB5 perhaps
            End If
        End If

        i = 2 * CODE_LEN + 1   ' for calling VirtualProtect
#End If

        Call zPatchVal(VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me))                  'Patch the address of this object instance into the static machine code buffer

        sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                   'Get the address of the CallWindowsProc function
        sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                   'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                     'Create the first sc_aSubData element

      Else
        nSubIdx = zIdx(lng_hWnd, True)
        If (nSubIdx = -1) Then                                                    'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                                   'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                  'Create a new sc_aSubData element
        End If

        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)

        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                             'Allocate memory for the machine code WndProc
        Call VirtualProtect(ByVal .nAddrSub, CODE_LEN, PAGE_EXECUTE_READWRITE, i)  'Mark memory as executable
        Call RtlMoveMemory(ByVal .nAddrSub, sc_aBuf(1), CODE_LEN)                 'Copy the machine code from the static byte array to the code array in sc_aSubData
        .hwnd = lng_hWnd                                                          'Store the hWnd
        .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                'Set our WndProc in place

        Call zPatchRel(.nAddrSub, PATCH_01, sc_pEbMode)                           'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                           'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, sc_pSWL)                              'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                           'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, sc_pCWP)                              'Patch the relative address of the CallWindowProc api function
    End With
End Function

Private Sub Subclass_StopAll()
'Stop all subclassing

  Dim i As Long

    i = UBound(sc_aSubData())                                                     'Get the upper bound of the subclass data array
    Do While i >= 0                                                               'Iterate through each element
        With sc_aSubData(i)
            If (.hwnd <> 0) Then                                                  'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hwnd)                                         'Subclass_Stop
            End If
        End With

        i = i - 1                                                                 'Next element
    Loop
End Sub

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Stop subclassing the passed window handle
'Parameters:
'   lng_hWnd - The handle of the window to stop being subclassed

    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)                       'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                                    'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                                    'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                                                'Release the machine code memory
        .hwnd = 0                                                                 'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                             'Clear the before table
        .nMsgCntA = 0                                                             'Clear the after table
        Erase .aMsgTblB                                                           'Erase the before table
        Erase .aMsgTblA                                                           'Erase the after table
    End With
End Sub

'----------------------------------------------------------------------------------------
'These z??? routines are exclusively called by the Subclass_??? routines.
'----------------------------------------------------------------------------------------

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
'Worker sub for Subclass_AddMsg

  Dim nEntry  As Long                                                             'Message table entry index
  Dim nOff1   As Long                                                             'Machine code buffer offset 1
  Dim nOff2   As Long                                                             'Machine code buffer offset 2

    If (uMsg = ALL_MESSAGES) Then                                                 'If all messages
        nMsgCnt = ALL_MESSAGES                                                    'Indicates that all messages will callback
      Else                                                                        'Else a specific message number
        Do While nEntry < nMsgCnt                                                 'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1

            If (aMsgTbl(nEntry) = 0) Then                                         'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                            'Re-use this entry
                Exit Sub                                                          'Bail
            ElseIf (aMsgTbl(nEntry) = uMsg) Then                                  'The msg is already in the table!
                Exit Sub                                                          'Bail
            End If
        Loop                                                                      'Next entry

        nMsgCnt = nMsgCnt + 1                                                     'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                              'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                   'Store the message number in the table
    End If

    If (When = eMsgWhen.MSG_BEFORE) Then                                          'If before
        nOff1 = PATCH_04                                                          'Offset to the Before table
        nOff2 = PATCH_05                                                          'Offset to the Before table entry count
      Else                                                                        'Else after
        nOff1 = PATCH_08                                                          'Offset to the After table
        nOff2 = PATCH_09                                                          'Offset to the After table entry count
    End If

    If (uMsg <> ALL_MESSAGES) Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                          'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                                         'Patch the appropriate table entry count
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
'Return the memory address of the passed function in the passed dll

    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                                        'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
''Worker sub for Subclass_DelMsg
'
'  Dim nEntry As Long
'
'    If (uMsg = ALL_MESSAGES) Then                                                 'If deleting all messages
'        nMsgCnt = 0                                                               'Message count is now zero
'        If When = eMsgWhen.MSG_BEFORE Then                                        'If before
'            nEntry = PATCH_05                                                     'Patch the before table message count location
'          Else                                                                    'Else after
'            nEntry = PATCH_09                                                     'Patch the after table message count location
'        End If
'        Call zPatchVal(nAddr, nEntry, 0)                                          'Patch the table message count to zero
'      Else                                                                        'Else deleteting a specific message
'        Do While nEntry < nMsgCnt                                                 'For each table entry
'            nEntry = nEntry + 1
'            If (aMsgTbl(nEntry) = uMsg) Then                                      'If this entry is the message we wish to delete
'                aMsgTbl(nEntry) = 0                                               'Mark the table slot as available
'                Exit Do                                                           'Bail
'            End If
'        Loop                                                                      'Next entry
'    End If
'End Sub

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the sc_aSubData() array index of the passed hWnd
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start

    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                            'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If (.hwnd = lng_hWnd) Then                                            'If the hWnd of this element is the one we're looking for
                If (Not bAdd) Then                                                'If we're searching not adding
                    Exit Function                                                 'Found
                End If
            ElseIf (.hwnd = 0) Then                                               'If this an element marked for reuse.
                If (bAdd) Then                                                    'If we're adding
                    Exit Function                                                 'Re-use it
                End If
            End If
        End With
        zIdx = zIdx - 1                                                           'Decrement the index
    Loop

    If (Not bAdd) Then
        Debug.Assert False                                                        'hWnd not found, programmer error
    End If

'If we exit here, we're returning -1, no freed elements were found
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
'Patch the machine code buffer at the indicated offset with the relative address to the target address.

    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
'Patch the machine code buffer at the indicated offset with the passed value

    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
'Worker function for Subclass_InIDE

    zSetTrue = True
    bValue = True
End Function


#If OLEDD Then

'========================================================================================
' OLE Drag & Drop support
'========================================================================================

Public Property Get OLEDragAutoExpand() As Boolean
    OLEDragAutoExpand = m_bOLEDragAutoExpand
End Property
Public Property Let OLEDragAutoExpand(ByVal New_OLEDragAutoExpand As Boolean)
    m_bOLEDragAutoExpand = New_OLEDragAutoExpand
End Property

' disAutomatic: Switches between InsertMark (cursor near borders of drop node)
'               and DropHilite (node interior).
' OLEDragInsertStyle returns actual style, not disAutomatic !
Public Property Get OLEDragInsertStyle() As tvOLEDragInsertStyleConstants
    OLEDragInsertStyle = m_eOLEDragInsertStyle
End Property
Public Property Let OLEDragInsertStyle(ByVal New_OLEDragInsertStyle As tvOLEDragInsertStyleConstants)
    If New_OLEDragInsertStyle <> disAutomatic Then
      m_eOLEDragInsertStyle = New_OLEDragInsertStyle
      m_bOLEDragAutoInsert = False
    Else
      m_bOLEDragAutoInsert = True
    End If
End Property

Public Property Get OLEDragMode() As tvOLEDragConstants
    OLEDragMode = m_eOLEDragMode
End Property
Public Property Let OLEDragMode(ByVal New_OLEDragMode As tvOLEDragConstants)
    If (m_hTreeView) Then
        m_eOLEDragMode = New_OLEDragMode
        Select Case m_eOLEDragMode
            Case [drgNone], [drgManual]
                Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_DISABLEDRAGDROP, 0)
            Case [drgAutomatic]
                Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_DISABLEDRAGDROP)
        End Select
    End If
End Property

Public Property Get OLEDropMode() As tvOLEDropConstants
    OLEDropMode = m_eOLEDropMode
End Property
Public Property Let OLEDropMode(ByVal New_OLEDropMode As tvOLEDropConstants)
    If (m_hTreeView) Then
        m_eOLEDropMode = New_OLEDropMode
        UserControl.OLEDropMode = m_eOLEDropMode
    End If
End Property

'//

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
   
   Set m_pbDrag = Nothing
   
   If m_hNodeDrag Then
      If m_eOLEDragMode = drgAutomatic Then
         '-- set minimal Data (clears previous Data)
         OLESetDataInfo Data, daMinimal
    ' Else
         ' drgManual: client sets data in OLEStartDrag or OLESetData event
    '    Data.Clear
      End If
      '-- This gives the user the opportunity to set AllowedEffects (Source)
      RaiseEvent OLEStartDrag(Data, AllowedEffects)
   End If

   '-- Check
   If (AllowedEffects = vbDropEffectNone) Then
      Call Data.Clear
      m_hNodeDrag = 0&
      Set m_pbDrag = Nothing
#If DDIMG Then
   Else
      pDragImageStart m_hNodeDrag
#End If
   End If
End Sub

' fires only with OLEDragMode = drgManual, when data is requested but not yet set.
' # ??? but it never fired for me ??? #
Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
   RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

   Dim eNewStyle     As tvOLEDragInsertStyleConstants
   Dim bInsertAfter  As Boolean
   Dim tR            As RECT2
   Dim lItemHeight   As Long
   Dim hNode         As Long
   Dim lfHit         As Long
   Dim lScroll       As Long
   Dim dx            As Long
   Dim dy            As Long
   Dim lH            As Long
   Dim lW            As Long
   Static hLastExpanded As Long

   If State = vbLeave Then
      m_bStateEnter = False
      hLastExpanded = 0&
      m_hNodeDrop = 0&
      RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
      Call pvHideDropPointer
      Exit Sub
   End If
    
   If (m_eOLEDropMode <> drpNone) Then
      
      '-- Get hit-info
      Call GetClientRect(m_hUC, tR) ' not m_hTreeView
      
      If PtInRect(tR, X, Y) Then
         Call pvTVHitTest(hNode, lfHit, (X), (Y))
      Else
         ' outside of control,yet pvTVHitTest can return valid hNode (y pos captured)
         m_hNodeDrop = 0&
         Exit Sub
      End If
      
      '-- Need to scroll ?
      lItemHeight = SendMessageLong(m_hTreeView, TVM_GETITEMHEIGHT, 0&, 0&)
      
      ' dy: distance from top/bottom, where auto scrolling occurs.
      ' Scrolling speed increases linear towards borders.
      ' T_SCROLL_DELAY_Y sets slowest speed at dy.
      dy = 2 * lItemHeight
      
      ' use theoretical lastvisible noderect for bottom border, not clientrect
      lH = lItemHeight * SendMessageLong(m_hTreeView, TVM_GETVISIBLECOUNT, 0&, 0&)
      
      lScroll = -1&
      
      Select Case True
         Case (timeGetTime() - m_lDragCounter) < T_SCROLL_DELAY_Y
            ' skip scrolling
         
         Case Y < dy
            lScroll = SB_LINEUP
            m_lDragCounter = timeGetTime() + T_SCROLL_DELAY_Y * Y \ dy - T_SCROLL_DELAY_Y
            
         Case Y > (lH - dy)
            ' Avoid a second scroll done by TreeView itself when a partially
            ' visible last item is selected.
            If (NodeVisible(hNode)) Then
               lScroll = SB_LINEDOWN
            End If
            m_lDragCounter = timeGetTime() + T_SCROLL_DELAY_Y * (lH - Y) \ dy - T_SCROLL_DELAY_Y
      
      End Select
      
      ' check max/min scrollbar position to avoid flicker
      If (lScroll = SB_LINEUP) Then
         If NodePrevious(hNode) = 0& Then lScroll = -1&
      ElseIf (lScroll = SB_LINEDOWN) Then
         If NodeNext(hNode) = 0& Then lScroll = -1&
      End If
      
      If (lScroll <> -1&) Then
         Call SendMessageLong(m_hTreeView, WM_VSCROLL, lScroll, 0)
      End If
      
      ' -- Horizontal scrolling MOD9
      lW = ScaleWidth
      dx = lW \ 4
      
      lScroll = -1&
      
      Select Case True
         Case (timeGetTime() - m_lDragCounter) < T_SCROLL_DELAY_X
            ' skip scrolling
         
         Case X < dx
            lScroll = SB_LINELEFT
            m_lDragCounter = timeGetTime() + T_SCROLL_DELAY_X * X \ dx - T_SCROLL_DELAY_X
            
         Case X > (lW - dx)
            lScroll = SB_LINERIGHT
            m_lDragCounter = timeGetTime() + T_SCROLL_DELAY_X * (lW - X) \ dx - T_SCROLL_DELAY_X
      
      End Select
      
      If (lScroll <> -1&) Then
         Call SendMessageLong(m_hTreeView, WM_HSCROLL, lScroll, 0)
      End If
      
      '-- Get hit-info, again: Is there a node ?
      If pvTVHitTest(hNode, lfHit, (X), (Y)) Then

         ' disAutomatic: automatic insertstyle switching
         If m_bOLEDragAutoInsert Then
            
            pTVItemRect hNode, tR, OnlyText:=False
               
            ' dy: distance from node borders  # change as appropriate #
            dy = (tR.Y2 - tR.Y1) \ 4
            
            If Y > (tR.Y1 + dy) And Y < (tR.Y2 - dy) Then
               ' well inside of hNodeDrop: use DropHilite
               eNewStyle = disDropHilite
               bInsertAfter = True
            Else
               ' near borders of hNodeDrop: use InsertMark
               eNewStyle = disInsertMark
               bInsertAfter = (Y > (tR.Y1 + tR.Y2) \ 2)
            End If
            If (eNewStyle <> m_eOLEDragInsertStyle) Then
               Call pvHideDropPointer
               m_eOLEDragInsertStyle = eNewStyle
               If (bInsertAfter <> m_bNodeDropInsertAfter) Then
                  m_bNodeDropInsertAfter = bInsertAfter
               End If
               Call pvShowDropPointer(hNode)
            ElseIf (bInsertAfter <> m_bNodeDropInsertAfter) Then
               m_bNodeDropInsertAfter = bInsertAfter
               Call pvShowDropPointer(hNode)
            End If
         
         ' Insert-mark ?
         ElseIf (m_eOLEDragInsertStyle = [disInsertMark]) Then
         
            pTVItemRect hNode, tR, OnlyText:=False
            bInsertAfter = (Y > (tR.Y1 + tR.Y2) \ 2)
            If (bInsertAfter <> m_bNodeDropInsertAfter) Then
               m_bNodeDropInsertAfter = bInsertAfter
               Call pvShowDropPointer(hNode)
            End If
         End If

         '-- New drop-Node ?
         If (m_hNodeDrop <> hNode) Then
            
            m_hNodeDrop = hNode
            Call pvShowDropPointer(hNode)

            '-- Timing Expand ...
            m_lExpandCounter = timeGetTime()

         Else
            '-- Expand ?
            If m_hNodeDrop <> hLastExpanded Then
            
               If (lfHit And (TVHT_ONITEMICON Or TVHT_ONITEMLABEL)) Then
               
                  If (m_bOLEDragAutoExpand) Then
                     ' BUGFIX4: DragAutoExpand adapted for Load on Demand clients
                     If pTVcChildren(hNode) Then ' == NodePlusMinusButton, <> pTVHasChildren
                        If Not (pTVState(hNode, TVIS_EXPANDED)) Then
                           If Not (pTVState(hNode, TVIS_EXPANDEDONCE)) Then
                              If (timeGetTime() - m_lExpandCounter > T_EXPAND_DELAY) Then
                                 Call pvHideDropPointer
                                 ' Try to expand only once, TVIS_EXPANDEDONCE is not reliable.
                                 hLastExpanded = m_hNodeDrop
                                 Call Expand(hNode, ExpandChildren:=False)
                                 Call pvShowDropPointer(hNode)
                              End If
                           End If
                        End If
                        
                     End If
                  End If
   
               Else
                  '-- Timing Expand ...
                  m_lExpandCounter = timeGetTime()
               End If
            
            End If   ' m_hNodeDrop <> hLastExpanded
         End If   ' (m_hNodeDrop <> hNode)

      Else
         '-- No drop-Node
         m_hNodeDrop = 0&
         Call pvHideDropPointer
         
      End If   ' pvTVHitTest()

      ' # Fix: Button info lost in UserControl_OLEDragDrop #
      m_eButton = Button
      
      ' # BUGFIX7: Ensure a OLEDragOver event with State = vbEnter #
      ' #          (may have exited on real vbEnter)               #
      If m_bStateEnter = False Then
         State = vbEnter
         m_bStateEnter = True
      End If
      
      '-- Raise event now (Target)
      RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)

'      '-- Check cancel [Esc]
'      If (Effect = vbDropEffectNone) Then
'         ' leave m_hNodeDrop or flicker occurs if hovering over same node
'         ' m_hNodeDrop = 0
'         ' MOD1: calmer dragging,if no-drop nodes are hilit as well
'         ' Call pvHideDropPointer
'      End If
   End If
   
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)

   '-- Source event: after every OLEDragOver() event on Target
   RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)

End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

#If DDIMG Then
   pDragImageComplete
#End If
    
   Call pvHideDropPointer
   Call ReleaseCapture
   m_bStateEnter = False
    
   ' # Shift OK, Button info is lost ???, Mouse is up -> pvButton(WM_MOUSEMOVE) useless #
   Debug.Assert Button = 0
   Button = m_eButton
   
   '-- Target last event...
   RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)

   '-- ...clear drop-info
   m_hNodeDrop = 0&
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)

#If DDIMG Then
   pDragImageComplete
#End If

   Call pvHideDropPointer
   Call ReleaseCapture
   m_bStateEnter = False
   
   '-- Source last event
   RaiseEvent OLECompleteDrag(Effect)
   Set m_pbDrag = Nothing
    
   If Effect = vbDropEffectNone Then
      ' BUGFIX2: After unsucessful drag operation dragging same node fails then suceeds.
      Dim lParam As Long, rcItem As RECT2
         
      If pTVItemRect(SelectedNode, rcItem, OnlyText:=True) Then
         With rcItem
            .X1 = (.X1 + .X2) \ 2
            .Y1 = (.Y1 + .Y2) \ 2
            lParam = .X1 + .Y1 * 65536
         End With
'         SendMessageLong m_hTreeView, WM_LBUTTONDOWN, 0, lParam
         SendMessageLong m_hTreeView, WM_LBUTTONUP, 0, lParam
      End If
   End If

End Sub

'========================================================================================
' OLE Drag & Drop Source
'========================================================================================

Private Sub pvShowDropPointer(hNode As Long)

    Select Case m_eOLEDragInsertStyle
        Case [disDropHilite]
            Call SendMessageLong(m_hTreeView, TVM_SELECTITEM, TVGN_DROPHILITE, hNode)
        Case [disInsertMark]
            Call SendMessageLong(m_hTreeView, TVM_SETINSERTMARK, -m_bNodeDropInsertAfter, hNode)
    End Select
End Sub

Private Sub pvHideDropPointer()

    Select Case m_eOLEDragInsertStyle
        Case [disDropHilite]
            Call SendMessageLong(m_hTreeView, TVM_SELECTITEM, TVGN_DROPHILITE, 0)
        Case [disInsertMark]
            Call SendMessageLong(m_hTreeView, TVM_SETINSERTMARK, 0, 0)
    End Select
End Sub

'//

' MOD4
Public Function OLEStartDrag(ByVal hNodeDrag As Long) As Boolean
   On Error GoTo Proc_Error

   Debug.Assert Me.OLEDragMode = drgManual
   Debug.Assert hNodeDrag
   
   If (m_hNodeDrag = 0&) Then
      m_hNodeDrag = hNodeDrag
      m_hNodeDrop = 0&
      Call SetCapture(m_hTreeView)
      Call UserControl.OLEDrag
      OLEStartDrag = True
   End If
   
   Exit Function

Proc_Error:
   Call ReleaseCapture
   Debug.Print "Error: " & Err.Number & ". " & Err.Description, vbOKOnly Or vbCritical, App.Title & ".ucTreeView: Function OLEStartDrag"
   If InIDE Then Stop: Resume
End Function

' Writes data to DataObject, extent of data written depends on DataOptions:
' - In drgAutomatic mode daMinimal data is already set on UserControl_OLEStartDrag.
' - Call from client OLEStartDrag/OLESetData event to write extended data.
Public Sub OLESetDataInfo(Data As DataObject, _
                          Optional ByVal DataOptions As tvDataOptions = daMinimal)
   Dim PB            As PropertyBag
   Dim eDF           As tvDataFormats
   Dim idxDragNode   As Long
   
   Debug.Assert m_hNodeDrag
   
   idxDragNode = pTVlParam(m_hNodeDrag)
   
   eDF = OLE_FORMAT_ID
   Set PB = New PropertyBag
   
   PB.WriteProperty "hTree", m_hTreeView
   PB.WriteProperty "hIls", m_hImageList(ilNormal), 0&
'   PB.WriteProperty "hIls2", m_hImageList(ilState), 0&
      
#If MULSEL Then
      
   If (DataOptions And daMultipleSelection) Then
      Dim idx     As Long
      Dim IdxPB   As Long
      Dim hNode   As Long
      
      If SelectionCount > 1& Then
         ' ### sort selection ###
      End If
      
      For idx = 1& To SelectionCount
         hNode = SelectionNode(idx)
         If hNode <> m_hNodeDrag Then
            ' multiselected nodes written with an index as suffix (1 To SelectionCount -1)
            IdxPB = IdxPB + 1&
            pOLEWrite PB, pTVlParam(hNode), DELIM_PB & IdxPB, DataOptions
         Else
            ' drag node (included in selection) written with index 0
            pOLEWrite PB, idxDragNode, DELIM_PB & "0", DataOptions
         End If
      Next
      
      If Not NodeSelected(m_hNodeDrag) Then
         ' drag node (not included in selection) written with index 0
         pOLEWrite PB, idxDragNode, DELIM_PB & "0", DataOptions
      End If
      
      ' Count without suffix denotes count of 'root' nodes in DataObject
      PB.WriteProperty "Count", IdxPB + 1&
      
   Else
      ' drag node written with index 0
      pOLEWrite PB, idxDragNode, DELIM_PB & "0", DataOptions
      ' Count = 1 for single drag node
      PB.WriteProperty "Count", 1&
   End If
   
#Else

   ' drag node
   pOLEWrite PB, idxDragNode, DELIM_PB & "0", DataOptions
   ' Count = 1 for single drag node
   PB.WriteProperty "Count", 1&
#End If  ' MULSEL
          
#If CUSTDRAW Then
   If (DataOptions And daCustomData) Then eDF = OLE_FORMAT_ID1
#End If
   If (DataOptions And (daCurrentChildren Or daChildren)) Then eDF = OLE_FORMAT_ID2
   If (DataOptions And daInterProcess) Then eDF = OLE_FORMAT_ID3
          
   ' Clear previous Data
   Call Data.Clear
   ' Drag node text as extra format
   Call Data.SetData(m_uNodeData(idxDragNode).sText, vbCFText)
   ' write propertybag contents to DataObject as specified format
   Data.SetData PB.Contents, eDF
   
   Debug.Print "Data bytes: " & LenB(PB.Contents)
   Set PB = Nothing
End Sub

' recursive for DataOptions = daCurrentChildren Or daChildren
' # add/remove properties as fit to reduce transfered data #
Private Sub pOLEWrite(PB As PropertyBag, ByVal idxNode As Long, _
                      ByVal sIdxPB As String, ByVal DataOptions As tvDataOptions)
   
   ' Debug.Print sIdxPB, NodeFullPath(m_uNodeData(idxNode).hNode)
   
   If (DataOptions And daChildren) Then
      ' populate with all children, even if not yet expanded
      ' # With m_uNodeData() locks the array -> err 10 in AddNode, when redimmed #
      ' # Add or expand(Load on Demand) nodes outside of With clause             #
      Dim hNode   As Long
      Dim bRedraw    As Boolean

      hNode = m_uNodeData(idxNode).hNode
      If Not NodeExpandedOnce(hNode) Then
         bRedraw = m_bRedraw
         Redraw = False
         ' raises Expand events without actually expanding
         Expand hNode, ExpandChildren:=True, EventOnly:=True
         Redraw = bRedraw
      End If
   End If
   
   With m_uNodeData(idxNode)
      ' data for DataOptions >= daMinimal
      PB.WriteProperty "#1" & sIdxPB, .hNode, 0&
      PB.WriteProperty "#2" & sIdxPB, .sKey, vbNullString
      PB.WriteProperty "#3" & sIdxPB, .sTag, vbNullString
      
      If (DataOptions And daInterProcess) Then
         PB.WriteProperty "-1" & sIdxPB, .sText, vbNullString
         PB.WriteProperty "-2" & sIdxPB, .idxImg, IMG_NONE
         PB.WriteProperty "-3" & sIdxPB, .idxSelImg, IMG_NONE
      End If
      
#If CUSTDRAW Then
      If (DataOptions And daCustomData) Then
         PB.WriteProperty "+1" & sIdxPB, .lForeColor, CLR_NONE
         PB.WriteProperty "+2" & sIdxPB, .lBackColor, CLR_NONE
         PB.WriteProperty "+3" & sIdxPB, .idxFont, 0&
         PB.WriteProperty "+4" & sIdxPB, NodeFont(, idxNode), Nothing
         PB.WriteProperty "+5" & sIdxPB, .lItemData, 0&
         PB.WriteProperty "+6" & sIdxPB, .idxExpImg, IMG_NONE
         PB.WriteProperty "+7" & sIdxPB, .lIndent, 0&
      End If
#End If
   
   End With
   
   If (DataOptions And (daCurrentChildren Or daChildren)) Then
      Dim hChild        As Long
      Dim lChildCount   As Long
      ' # recurse outside of With clause #
      hChild = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_CHILD, _
                               m_uNodeData(idxNode).hNode)
      Do While hChild
         lChildCount = lChildCount + 1
         ' recurse: write child data with added suffix to parent suffix
         pOLEWrite PB, pTVlParam(hChild), _
                   sIdxPB & DELIM_PB & lChildCount, DataOptions
         hChild = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_NEXT, hChild)
      Loop
      ' Count with suffix denotes ChildCount of node
      PB.WriteProperty "Count" & sIdxPB, lChildCount  ' !no default value!
   Else
      ' used by pOleValidate for all formats
      PB.WriteProperty "Count" & sIdxPB, 0&           ' !no default value!
   End If
   
   ' Debug.Print sIdxPB, lChildCount, m_uNodeData(idxNode).sKey, NodeFullPath(m_uNodeData(idxNode).hNode)
End Sub


'========================================================================================
' OLE Drag & Drop Target
'========================================================================================

Public Function OLEGetDropInfo(Optional hNodeDrop As Long, Optional InsertAfter As Boolean) _
                               As Boolean
   hNodeDrop = m_hNodeDrop
   InsertAfter = IIf(m_eOLEDragInsertStyle = [disInsertMark], m_bNodeDropInsertAfter, True)
   OLEGetDropInfo = (hNodeDrop <> 0)   ' MOD3
End Function

' if DataFormat is specified: returns True only if specified (or larger) format exists.
' DataFormat,NodesCount,hTreeView,hImageList always returned
Public Function OLEIsMyFormat(Data As DataObject, _
                              Optional DataFormat As tvDataFormats = 0&, _
                              Optional NodesCount As Long, _
                              Optional hTreeView As Long, _
                              Optional hImageList As Long) As Boolean
   
   ' returns actual DataFormat
   OLEIsMyFormat = pOLEIsMyFormat(Data, DataFormat)
   
   If (DataFormat >= OLE_FORMAT_ID) Then
      If m_pbDrag Is Nothing Then
         Set m_pbDrag = New PropertyBag
      End If
      ' restore DataObject in propertybag for subsequent OLEGetDragInfo calls
      m_pbDrag.Contents = Data.GetData(DataFormat)
      
      hTreeView = m_pbDrag.ReadProperty("hTree", 0&)
      hImageList = m_pbDrag.ReadProperty("hIls", 0&)
      NodesCount = m_pbDrag.ReadProperty("Count", 0&)
      Debug.Assert (hTreeView <> 0&) And (NodesCount <> 0&)
   Else
      hTreeView = 0&
      hImageList = 0&
      NodesCount = 0&
      Set m_pbDrag = Nothing
   End If

End Function

Private Function pOLEIsMyFormat(Data As DataObject, _
                                Optional DataFormat As tvDataFormats = 0&) As Boolean
   Dim eDF  As tvDataFormats
   
   If Data.GetFormat(OLE_FORMAT_ID) Then
      eDF = OLE_FORMAT_ID
   ElseIf Data.GetFormat(OLE_FORMAT_ID1) Then
      eDF = OLE_FORMAT_ID1
   ElseIf Data.GetFormat(OLE_FORMAT_ID2) Then
      eDF = OLE_FORMAT_ID2
   ElseIf Data.GetFormat(OLE_FORMAT_ID3) Then
      eDF = OLE_FORMAT_ID3
   End If

   If (DataFormat = 0&) Then
      ' no DataFormat passed: return True if any of my formats exist.
      pOLEIsMyFormat = (eDF >= OLE_FORMAT_ID)
   Else
      ' DataFormat passed: return True only if specified (or larger) format exists.
      pOLEIsMyFormat = (eDF >= OLE_FORMAT_ID) And (eDF >= DataFormat)
   End If
   
   DataFormat = eDF

End Function

' usable with all DataFormats(tvDataFormats) (idxChild != 0)
' - sNode = "0" := drag node / CStr(1 to NodesCount-1) := multiselected nodes
' - get NodesCount from calling OLEIsMyFormat first
' DataFormats >= tvDataFormats.OLE_FORMAT_ID2:
' - idxChild = 0 := sNode / 1 to ChildCount(idxNode):= child x of sNode
' - get ChildCount(idxNode) from OLEGetDragInfo call with sNode, idxChild = 0 first
' - idxChild <> 0: pass sNode = sChild from OLEGetDragInfo call with sNode, idxChild = 0
Public Function OLEGetDragInfo(Data As DataObject, ByVal sNode As String, _
                               Optional ByVal idxChild As Long = 0&, _
                               Optional ByRef hNode As Long, _
                               Optional ByRef Key As String, _
                               Optional ByRef Tag As String, _
                               Optional ByRef ChildCount As Long, _
                               Optional ByRef sChild As String _
                               ) As Boolean
                                 
   Dim uNodeData  As NODE_DATA
   Dim eDF        As tvDataFormats
   Dim sIdxPB     As String
   Dim lRes       As Long
   
   hNode = 0&
   Key = vbNullString
   Tag = vbNullString
   ChildCount = 0&
   
   ' validate format
   If (idxChild = 0&) Then
      eDF = OLE_FORMAT_ID
   Else
      eDF = OLE_FORMAT_ID2
   End If
   
   If Not pOLEIsMyFormat(Data, eDF) Then
      Debug.Assert False
      Err.Raise 17
   End If
   
   Debug.Assert Not (m_pbDrag Is Nothing)
   
   ' validate sNode & idxChild
   lRes = pOLEValidate(sNode, idxChild, sIdxPB, ChildCount)
   ' idxChild <> 0: return child ID for subsequent calls / idxChild = 0: same as sNode
   sChild = sNode
   
   If (lRes = 1&) Then
      pOLERead m_pbDrag, uNodeData, sIdxPB, daMinimal
      
      With uNodeData
         hNode = .hNode
         Key = .sKey
         Tag = .sTag
      End With
   
      OLEGetDragInfo = True
   
   ElseIf (lRes = 0&) Then
      ' passed sNode is invalid
      Debug.Assert False
   ElseIf (lRes = -1&) Then
      ' passed idxChild is invalid
      Debug.Assert False
   End If

End Function

#If CUSTDRAW Then
' usable with DataFormats >= tvDataFormats.OLE_FORMAT_ID1 (idxChild != 0)
' - sNode = "0" := drag node / CStr(1 to NodesCount-1) := multiselected nodes
' - get NodesCount from calling OLEIsMyFormat first
' DataFormats >= tvDataFormats.OLE_FORMAT_ID2:
' - idxChild = 0 := sNode / 1 to ChildCount(idxNode):= child x of sNode
' - get ChildCount(idxNode) from OLEGetDragInfo call with sNode, idxChild = 0 first
' - idxChild <> 0: pass sNode = sChild from OLEGetDragInfo call with sNode, idxChild = 0
Public Function OLEGetDragInfoEx1(Data As DataObject, ByVal sNode As String, _
                                  Optional ByVal idxChild As Long = 0&, _
                                  Optional ByRef ForeColor As Long, _
                                  Optional ByRef BackColor As Long, _
                                  Optional ByRef Font As StdFont, _
                                  Optional ByRef TextIndent As Long, _
                                  Optional ByRef ItemData As Long, _
                                  Optional ByRef ExpandedImage As Long _
                                  ) As Boolean
   Dim uNodeData  As NODE_DATA
   Dim eDF        As tvDataFormats
   Dim sIdxPB     As String
   Dim lRes       As Long
   
   ForeColor = CLR_NONE
   BackColor = CLR_NONE
   TextIndent = 0&
   ItemData = 0&
   ExpandedImage = IMG_NONE
   
   ' validate format
   If (idxChild = 0&) Then
      eDF = OLE_FORMAT_ID1
   Else
      eDF = OLE_FORMAT_ID2
   End If
   
   If Not pOLEIsMyFormat(Data, eDF) Then
      Debug.Assert False
      Err.Raise 17
   End If
   
   Debug.Assert Not (m_pbDrag Is Nothing)
   
   ' validate sNode & idxChild
   lRes = pOLEValidate(sNode, idxChild, sIdxPB)
   
   If (lRes = 1&) Then
      pOLERead m_pbDrag, uNodeData, sIdxPB, daCustomData, Font
      
      With uNodeData
         ForeColor = .lForeColor
         BackColor = .lBackColor
         TextIndent = .lIndent
         ItemData = .lItemData
         ExpandedImage = .idxExpImg
      End With
   
      OLEGetDragInfoEx1 = True
   
   ElseIf (lRes = 0&) Then
      ' passed sNode is invalid
      Debug.Assert False
   ElseIf (lRes = -1&) Then
      ' passed idxChild is invalid
      Debug.Assert False
   End If
   
End Function

#End If ' CUSTDRAW

' usable only with DataFormat = tvDataFormats.OLE_FORMAT_ID3
' - sNode = "0" := drag node / CStr(1 to NodesCount-1) := multiselected nodes
' - get NodesCount from calling OLEIsMyFormat first
' - idxChild = 0 := sNode / 1 to ChildCount(idxNode):= child x of sNode
' - get ChildCount(idxNode) from OLEGetDragInfo call with sNode, idxChild = 0 first
' - idxChild <> 0: pass sNode = sChild from OLEGetDragInfo call with sNode, idxChild = 0
Public Function OLEGetDragInfoEx3(Data As DataObject, ByVal sNode As String, _
                                  Optional ByVal idxChild As Long = 0&, _
                                  Optional ByRef Text As String, _
                                  Optional ByRef Image As Long, _
                                  Optional ByRef SelectedImage As Long _
                                  ) As Boolean
   Dim eDF        As tvDataFormats
   Dim sIdxPB     As String
   Dim lRes       As Long
   
   Text = vbNullString
   Image = IMG_NONE
   SelectedImage = IMG_NONE
   
   ' validate format
   eDF = OLE_FORMAT_ID3
   If Not pOLEIsMyFormat(Data, eDF) Then
      Debug.Assert False
      Err.Raise 17
   End If
   
   Debug.Assert Not (m_pbDrag Is Nothing)
   
   ' validate sNode & idxChild
   lRes = pOLEValidate(sNode, idxChild, sIdxPB)
   
   If (lRes = 1&) Then
      Text = m_pbDrag.ReadProperty("-1" & sIdxPB, vbNullString)
      Image = m_pbDrag.ReadProperty("-2" & sIdxPB, IMG_NONE)
      SelectedImage = m_pbDrag.ReadProperty("-3" & sIdxPB, IMG_NONE)
   
      OLEGetDragInfoEx3 = True
   
   ElseIf (lRes = 0&) Then
      ' passed sNode is invalid
      Debug.Assert False
   ElseIf (lRes = -1&) Then
      ' passed idxChild is invalid
      Debug.Assert False
   End If
   
End Function

Private Function pOLEValidate(ByRef sNode As String, ByVal idxChild As Long, _
                              ByRef sIdxPB As String, _
                              Optional ByRef ChildCount As Long) As Long
   sIdxPB = DELIM_PB & sNode
   
   ' return ChildCount of node (idxChild = 0&)
   ChildCount = m_pbDrag.ReadProperty("Count" & sIdxPB, -1&)
   
   If (ChildCount <> -1&) Then
      ' sNode is valid
      pOLEValidate = 1&
      
      If (idxChild <> 0&) Then
         Select Case idxChild
            Case 1 To ChildCount
               sIdxPB = sIdxPB & DELIM_PB & idxChild
               ' return ChildCount of child
               ChildCount = m_pbDrag.ReadProperty("Count" & sIdxPB, 0&)
               
               ' return child ID for subsequent calls
               sNode = sNode & DELIM_PB & idxChild
            Case Else
               ' passed idxChild is invalid
               Debug.Assert False
               sNode = vbNullString
               pOLEValidate = -1&
         End Select
      End If
      
   Else
      ' passed sNode is invalid
      Debug.Assert False
      sNode = vbNullString
      ChildCount = 0&
    ' pOLEValidate = 0&
   End If
End Function

' # add/remove properties as fit to reduce transfered data #
Private Sub pOLERead(PB As PropertyBag, ByRef uNodeData As NODE_DATA, _
                     ByVal sIdxPB As String, ByVal DataOptions As tvDataOptions, _
                     Optional ByRef Font As StdFont)
   
   With uNodeData
      If (DataOptions And daMinimal) Then
         .hNode = PB.ReadProperty("#1" & sIdxPB, 0&)
         .sKey = PB.ReadProperty("#2" & sIdxPB, vbNullString)
         .sTag = PB.ReadProperty("#3" & sIdxPB, vbNullString)
      End If
#If CUSTDRAW Then
      If (DataOptions And daCustomData) Then
         .lForeColor = PB.ReadProperty("+1" & sIdxPB, CLR_NONE)
         .lBackColor = PB.ReadProperty("+2" & sIdxPB, CLR_NONE)
         .idxFont = PB.ReadProperty("+3" & sIdxPB, 0&)
         Set Font = PB.ReadProperty("+4" & sIdxPB, Nothing)
         .lItemData = PB.ReadProperty("+5" & sIdxPB, 0&)
         .idxExpImg = PB.ReadProperty("+6" & sIdxPB, IMG_NONE)
         .lIndent = PB.ReadProperty("+7" & sIdxPB, 0&)
      End If
#End If
   ' Debug.Print sIdxPB, "-", .sKey, NodeFullPath(.hNode)  ' Target ucTree as wrapper for source tree nodes
   End With
End Sub

#End If  ' OLEDD

'========================================================================================
' Drag & Drop Image
'========================================================================================

' # Version1: ImageList_BeginDrag,ImageList_DragMove,ImageList_EndDrag API's
' # - all OS
' # - OK inprocess. partial problems, if scrolling last node.
' # - severe artifacts for crossprocess dragging
' # - image should be below cursor

' # Version2: Moving layered transparent window
' # - >=Win2K only
' # - OK for all dragging modes
' # - cursor can be in image

' # Transparent window for Win98:
' # - WS_EX_TRANSPARENT AKA forget it:
' #   - MSDN: "Microsoft Windows does not support fully functional transparent windows."
' # - SetWindowRgn: propably OK if dragimage consists only of nodeimages,
' #   but slurred for longer label text.
' # == don't offer drag images for <Win2K

' # Problem with Version2: Drag image still visible with drag contextmenu                 #
' # Contextmenu invoked in target's OLEDragDrop, source's OLECompleteDrag is too late.    #
' # Tried solution: In pDragImageMove() monitor any changes in GetCapture() returns:      #
' # Did provide the indication of menu popup, but failed for crossprocess dragging.       #
' # Surprise, Surprise: Although we set capture to m_hTreeView when starting drag,        #
' # GetCapture reveals mouse input is directed to a window of the CLIPBRDWNDCLASS,        #
' # an undocumented OLE-managed window acting as clipboard owner for our app.             #
' # Subclassing (it's in-process!) this wnd for WM_CAPTURECHANGED yields the needed       #
' # indication to destroy drag image wnd before menu popup.                               #
' # Spy++ reveals wnd receives the missing WM_MOUSEMOVE & WM_RBUTTONUP in drag loop,      #
' # but sadly not trappable here ??? (use hook ?).Receives trappable clipboard notifs!!!  #                                         #

#If DDIMG Then

Private Sub pDragImageStart(ByVal hNode As Long)
   Dim lW         As Long
   Dim lH         As Long
   Dim crColor    As Long
   Dim hBmp       As Long
   Dim lExStyle   As Long
   Dim tWC        As WNDCLASS
   
   ' Degrade gracefully for <Win2K
   If Not pvIsFunctionExported("SetLayeredWindowAttributes", "user32") Then Exit Sub
   
   pDragImageComplete
   
   ' m_tMemDC(1) will hold drag image
   hBmp = DragImageBmp(hNode, m_tpDDOffset.X, m_tpDDOffset.Y, lW, lH, crColor, bKeepDC:=True)
   
   ' Register popup window class
   With tWC
      .style = CS_DBLCLKS Or CS_SAVEBITS
      .lpfnWndProc = sc_aSubData(1).nAddrSub    ' == zSubclass_Proc()
      .hInstance = App.hInstance
      .lpszClassName = DD_WNDCLS
    ' .hbrBackground = 0
   End With
   
   If RegisterClass(tWC) Then
   
      lExStyle = WS_EX_TOPMOST Or WS_EX_TOOLWINDOW Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
      
      '-- Create transparent topmost popup window
      ' WS_POPUP: no caption, no borders / WS_EX_TOOLWINDOW: doesn't show in taskbar
      ' WS_EX_TRANSPARENT: mouse input falls trough underlying windows
      m_hDDImg = CreateWindowEx(lExStyle, DD_WNDCLS, vbNullString, WS_POPUP Or WS_DISABLED, _
                                0&, 0&, lW, lH, 0&, 0&, App.hInstance, ByVal 0&)
   End If
   
   If m_hDDImg Then
      
      m_bInDrag = True
      
      SetLayeredWindowAttributes m_hDDImg, crColor, 128&, LWA_COLORKEY Or LWA_ALPHA
                                 
      pDragImageMove
      ShowWindow m_hDDImg, SW_SHOW
      SetFocus m_hTreeView
      
      ' create a timer to move drag image with cursor (zSubclass_Proc/m_hTreeView/WM_TIMER)
      m_idxTimer = 99&
      m_idxTimer = SetTimer(m_hTreeView, m_idxTimer, TIM_IVALL, 0&)
   
'      Debug.Print "pDragImageStart"
   End If
   
End Sub

Private Sub pDragImageMove()
   Dim tP      As POINTAPI
   
   If m_bInDrag Then

      GetCursorPos tP
      SetWindowPos m_hDDImg, 0&, tP.X - m_tpDDOffset.X, tP.Y - m_tpDDOffset.Y, _
                   0&, 0&, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
   End If
End Sub

Private Sub pDragImageComplete()
   ' m_bInDrag can be false for OLE Target
   
   If m_idxTimer Then
      KillTimer m_hTreeView, m_idxTimer
      m_idxTimer = 0&
   End If
   
   If m_hDDImg Then
      DestroyWindow m_hDDImg
      UnregisterClass DD_WNDCLS, App.hInstance
      m_hDDImg = 0&
   End If
   
   If UBound(m_tMemDC()) = 1 Then
      pDestroyDC 1&
      ReDim Preserve m_tMemDC(0)
   End If
   
   m_bInDrag = False
   
'   Debug.Print "pDragImageComplete"
End Sub

#End If  ' DDIMG

#If DDIMG Or FLDBR Then

' returns handle to drag image bitmap, owned by caller.
Public Property Get DragImageBmp(ByVal hNode As Long, _
                                 Optional ByRef xOffset As Long, _
                                 Optional ByRef yOffset As Long, _
                                 Optional ByRef DragImageW As Long, _
                                 Optional ByRef DragImageH As Long, _
                                 Optional ByRef crColorKey As Long, _
                                 Optional ByVal bKeepDC As Boolean _
                                 ) As Long
   Dim rcItem     As RECT2
   Dim rc         As RECT2
   Dim dx         As Long
   Dim tP         As POINTAPI
   
   ReDim Preserve m_tMemDC(0 To 1)
   
   If crColorKey = 0 Then
      ' no transparent color passed: use window background
      crColorKey = GetSysColor(COLOR_WINDOW)    ' m_lTreeColors(clrTreeBK)
   End If
   
   ' x offset from rcItem.X1: include additional parts of NodeRect (dx = 0 : NodeText only)
   dx = -m_lImageW - SPACE_IL    ' include NodeImage
   
#If MULSEL Then
   ' DragImage of multiple drag nodes
   Dim idx        As Long
   Dim rcBounds   As RECT2
   
   GetClientRect m_hTreeView, rc
   
   ' get bounding rect of DragImage (== size of DC)
   pTVItemRect hNode, rcBounds, OnlyText:=True
   rcBounds.X1 = rcBounds.X1 + dx
   
   For idx = 1& To SelectionCount
      pTVItemRect SelectionNode(idx), rcItem, OnlyText:=True
      rcItem.X1 = rcItem.X1 + dx
      
      ' # exclude invisible nodes / ? or condense in drag image ? #
      If rcItem.Y1 >= rc.Y1 And rcItem.Y2 <= rc.Y2 Then
         UnionRect rcBounds, rcBounds, rcItem
      End If
   Next
      
   ' create memory DC for DragImage
   DragImageW = rcBounds.X2 - rcBounds.X1
   DragImageH = rcBounds.Y2 - rcBounds.Y1
   pCreateDC 1&, DragImageW, DragImageH
      
   ' fill BK of DC: tree backcolor will be transparent in DragImage
   SetRect rcItem, 0&, 0&, DragImageW, DragImageH
   pFillRect m_tMemDC(1).hdc, rcItem, crColorKey
   
   For idx = 1& To SelectionCount
      
      pTVItemRect SelectionNode(idx), rcItem, OnlyText:=True
   
      ' # exclude invisible nodes #
      If rcItem.Y1 >= rc.Y1 And rcItem.Y2 <= rc.Y2 Then
   
         ' # Option1: extra pDragImageDraw() for drawing #
         pDragImageDraw m_tMemDC(1).hdc, SelectionNode(idx), rcBounds.X1, rcBounds.Y1

'         ' # Option2: reuse pTreeOwnerDraw(). drawback: NodeText in DragImage drawn selected #
'         ' this fills BufferDC m_tMemDC(0)
'         Refresh SelectionNode(idx)
'         DoEvents
'
'         ' copy wanted part of NodeRect (NodeImage & NodeText)
'         BitBlt m_tMemDC(1).hdc, _
'                rcItem.X1 - rcBounds.X1, rcItem.Y1 - rcBounds.Y1, _
'                rcItem.X2 - rcItem.X1 , rcItem.Y2 - rcItem.Y1, _
'                m_tMemDC(0).hdc, rcItem.X1 , 0&, vbSrcCopy
      End If
   Next
   
'   ' offset cursor to DragImage (negative offset == below cursor)
'   ' ! Keep image below cursor, otherwise artifacts remain,       !
'   ' ! when scrolling down and cursor is below NodeLastVisible.   !                                  !
'   pTVItemRect hNode, rcItem, OnlyText:=True
'   rcItem.X1 = rcItem.X1 + dx - 16&    ' 16:= width of cursor
'   xOffset = rcItem.X1 - rcBounds.X1
'   yOffset = rcItem.Y1 - rcBounds.Y1

   ' offset cursor to DragImage (cursor pos in drag node =! cursor pos in drag image)
   GetCursorPos tP
   ScreenToClient m_hTreeView, tP
   xOffset = tP.X - rcBounds.X1 + dx + 16&   ' 16:= width of cursor
   yOffset = tP.Y - rcBounds.Y1

#Else
   ' DragImage of single drag node
   
   pTVItemRect hNode, rcItem, OnlyText:=True
   rcItem.X1 = rcItem.X1 + dx
   
   ' create memory DC for DragImage
   DragImageW = rcItem.X2 - rcItem.X1
   DragImageH = rcItem.Y2 - rcItem.Y1
   pCreateDC 1&, DragImageW, DragImageH
   
   ' # Option1: extra pDragImageDraw() for drawing #
   SetRect rc, 0&, 0&, DragImageW, DragImageH
   pFillRect m_tMemDC(1).hdc, rc, crColorKey

   pDragImageDraw m_tMemDC(1).hdc, hNode, rcItem.X1, rcItem.Y1
   
'   ' # Option2: reuse pTreeOwnerDraw(). drawback: NodeText in DragImage drawn selected #
'   ' this fills BufferDC m_tMemDC(0)
'   Refresh hNode
'
'   ' copy wanted part of NodeRect
'   BitBlt m_tMemDC(1).hDC, 0&, 0&, DragImageW, DragImageH, _
'          m_tMemDC(0).hDC, rcItem.X1, 0&, vbSrcCopy

'   ' offset cursor to DragImage (negative offset == below cursor)
'   xOffset = -16&       ' 16:= width of cursor
'   yOffset = 0&
     
   ' offset cursor to DragImage (cursor pos in drag node =! cursor pos in drag image)
   GetCursorPos tP
   ScreenToClient m_hTreeView, tP
   xOffset = tP.X - rcItem.X1 + dx + 16&     ' 16:= width of cursor
   yOffset = tP.Y - rcItem.Y1
   
#End If  ' MULSEL

   DragImageBmp = m_tMemDC(1).hBmp
   
   If Not bKeepDC Then
   
      ' free bitmap hBmp
      With m_tMemDC(1)
         SelectObject .hdc, .hBmpOld
         .hBmpOld = 0&
      End With
         
      m_tMemDC(1).hBmp = 0
      
      pDestroyDC 1&
      ReDim Preserve m_tMemDC(0)
   
   End If
   
End Property

Private Sub pDragImageDraw(ByVal lhDC As Long, ByVal hNode As Long, _
                           Optional ByVal xOffset As Long = 0, _
                           Optional ByVal yOffset As Long = 0)
   Dim idxNode       As Long
   Dim rcItem        As RECT2
   Dim rcItemCorr    As RECT2
   Dim rcText        As RECT2
   Dim rcImg         As RECT2
   Dim hFontOld      As Long
   Dim clrText       As Long
   Dim xPos          As Long
   Dim yPos          As Long
   
   idxNode = pTVlParam(hNode)
   
   pTVItemRect hNode, rcItem, OnlyText:=True
   OffsetRect rcItem, -xOffset, -yOffset
   rcItemCorr = pGetItemRectReal(hNode, rcItem, rcText, idxNode)
   
   With m_uNodeData(idxNode)
      If .idxFont <> 0& And m_bCustomDraw Then
         hFontOld = SelectObject(lhDC, m_hFnt(.idxFont))
      Else
         hFontOld = SelectObject(lhDC, m_hFont)
      End If
      If m_bDrawColor Then
         pTreeCustomColor idxNode, False, False, clrText
      Else
         clrText = m_lTreeColors(clrTree)
      End If
      SetTextColor lhDC, clrText
      SetBkMode lhDC, TRANSPARENT
      
#If UNICODE Then
      DrawTextW lhDC, StrPtr(.sText), -1, rcText, _
                DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX Or DT_NOCLIP
#Else
      DrawTextA lhDC, .sText, Len(.sText), rcText, _
                DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX Or DT_NOCLIP
#End If
      
      SelectObject lhDC, hFontOld
'      SetBkMode lhDC, OPAQUE
   
      If (.idxImg <> IMG_NONE) Then
         
         xPos = rcItem.X1 - SPACE_IL - m_lImageW \ 2
         yPos = (rcItem.Y1 + rcItem.Y2) \ 2
         
         ' calculate rect of image (vertical centered)
         SetRect rcImg, xPos, yPos, xPos + 1&, yPos + 1&
         InflateRect rcImg, m_lImageW \ 2, m_lImageH \ 2
      
         ImageList_Draw m_hImageList(ilNormal), .idxImg, lhDC, _
                        rcImg.X1, rcImg.Y1, ILD_NORMAL
      End If
   End With

End Sub

#End If  ' DDIMG OR FLDBR

#If CUSTDRAW Then

'========================================================================================
' CustomDraw Tree properties
'========================================================================================

Public Property Get TreeColor(Optional ByVal eColor As tvTreeColors) As OLE_COLOR
   Debug.Assert eColor >= clrSelected And eColor <= clrHilitBK
   TreeColor = m_lTreeColors(eColor)
End Property
' Pass CLR_NONE to reset to system default color
Public Property Let TreeColor(Optional ByVal eColor As tvTreeColors, ByVal lTreeColor As OLE_COLOR)
   Dim lColor As Long
   Debug.Assert eColor >= clrSelected And eColor <= clrHilitBK
   lColor = pTranslateColor(lTreeColor)
   If lColor = CLR_NONE Then
      Select Case eColor
          Case clrSelected, clrHilit
            m_lTreeColors(eColor) = GetSysColor(COLOR_HIGHLIGHTTEXT)
          Case clrSelectedBK, clrHilitBK
            m_lTreeColors(eColor) = GetSysColor(COLOR_HIGHLIGHT)
          Case clrSelectedNoFocusBk
            m_lTreeColors(eColor) = GetSysColor(COLOR_3DFACE)
          Case clrHot
            If (GetSysColorBrush(COLOR_HOTLIGHT) <> 0&) Then
               m_lTreeColors(eColor) = GetSysColor(COLOR_HOTLIGHT)
            Else
               ' < Win2K
               m_lTreeColors(eColor) = GetSysColor(COLOR_HIGHLIGHT)
            End If
      End Select
   Else
      m_lTreeColors(eColor) = lColor
   End If
End Property

' OR'ed
Public Property Get DoCustomDraw() As tvCustomDraw
   DoCustomDraw = -cdColor * m_bDrawColor - cdFont * m_bDrawFont _
                  - cdExpandedImage * m_bDrawExpanded - cdMixNoImage * m_bMixNoImage _
                  - cdProject * m_bDrawProject - cdLabelIndent * m_bDrawLabelTI _
                  - cdLabel * (m_bDrawLabel And Not m_bDrawLabelTI)
End Property
Public Property Let DoCustomDraw(ByVal DoDraw As tvCustomDraw)
   m_bDrawColor = ((DoDraw And cdColor) = cdColor)
   m_bDrawFont = ((DoDraw And cdFont) = cdFont)
   If m_bDrawLabelTI = ((DoDraw And cdLabelIndent) = cdLabelIndent) Then
      m_bDrawLabelTI = ((DoDraw And cdLabelIndent) = cdLabelIndent)
   Else
      m_bDrawLabelTI = ((DoDraw And cdLabelIndent) = cdLabelIndent)
      ' toggled cdLabel/cdLabelIndent:
      ' comctl must update ItemRect's for correct horizontal scrollbar behaviour
      pRefreshNodeRects
   End If
   m_bDrawLabel = ((DoDraw And cdLabel) = cdLabel) Or m_bDrawLabelTI
   m_bDrawExpanded = ((DoDraw And cdExpandedImage) = cdExpandedImage)
   m_bMixNoImage = ((DoDraw And cdMixNoImage) = cdMixNoImage)
   m_bDrawProject = ((DoDraw And cdProject) = cdProject)
   m_bCustomDraw = m_bDrawColor Or m_bDrawFont Or m_bDrawLabel Or m_bDrawExpanded Or _
                   m_bMixNoImage Or m_bDrawProject
End Property

' see tvSelectionStyle enum
Public Property Get SelectionStyle() As tvSelectionStyle
   SelectionStyle = -m_bPriorityExpanded * (ssPriorityExpanded - ssPrioritySelected) + ssPrioritySelected
#If MULSEL Then
   SelectionStyle = SelectionStyle _
                    - m_bMultiSelectImage * (ssImageMultiSelected - ssImageSelected) + ssImageSelected
#End If
End Property
Public Property Let SelectionStyle(ByVal eStyle As tvSelectionStyle)
   m_bPriorityExpanded = ((eStyle And ssPriorityExpanded) = ssPriorityExpanded)
#If MULSEL Then
   m_bMultiSelectImage = ((eStyle And ssImageMultiSelected) = ssImageMultiSelected)
#End If
End Property

'========================================================================================
' CustomDraw Node properties
'========================================================================================

Public Property Get NodeForecolor(Optional ByVal hNode As Long, _
                                  Optional ByRef idxNode As Long) As OLE_COLOR
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      NodeForecolor = m_uNodeData(idxNode).lForeColor
   End If
End Property
Public Property Let NodeForecolor(Optional ByVal hNode As Long, Optional ByRef idxNode As Long, _
                                  ByVal New_Forecolor As OLE_COLOR)
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      m_uNodeData(idxNode).lForeColor = pTranslateColor(New_Forecolor)
   End If
End Property

Public Property Get NodeBackcolor(Optional ByVal hNode As Long, _
                                  Optional ByRef idxNode As Long) As OLE_COLOR
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      NodeBackcolor = m_uNodeData(idxNode).lBackColor
   End If
End Property
Public Property Let NodeBackcolor(Optional ByVal hNode As Long, Optional ByRef idxNode As Long, _
                                  ByVal New_Backcolor As OLE_COLOR)
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      m_uNodeData(idxNode).lBackColor = pTranslateColor(New_Backcolor)
   End If
End Property

' NodeBold is irrelevant,if NodeFont is used!
Public Property Get NodeFont(Optional ByVal hNode As Long, Optional ByRef idxNode As Long, _
                             Optional ByVal bGroup As Boolean) As StdFont
   Dim idxFont As Long
   
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      idxFont = m_uNodeData(idxNode).idxFont
      If idxFont <> 0& Then
         Set NodeFont = pvCreateFont(m_FntLF(idxFont))
      Else
         ' tree font
         Set NodeFont = m_oFont
      End If
   End If
End Property
' pass Nothing to reset to default tree font
' bGroup:= True : applies to all nodes with same font
Public Property Set NodeFont(Optional ByVal hNode As Long, Optional ByRef idxNode As Long, _
                             Optional ByVal bGroup As Boolean, ByVal oNodeFont As StdFont)
   Dim bNew       As Boolean
   
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      If bGroup Then
         pChangeGroupFont idxNode, oNodeFont, bNew
      Else
         m_uNodeData(idxNode).idxFont = pAddFontIfRequired(oNodeFont, bNew)
      End If
#If AUTOFNT Then
      If bNew And m_bAutoFont Then
         pAdjustFontIfRequired
      End If
#End If
      pCalculateRcFont idxNode
   End If
End Property

Public Property Get NodeExpandedImage(Optional ByVal hNode As Long, _
                                      Optional ByRef idxNode As Long) As Long
   Dim lIndex As Long

   If (m_hTreeView) Then
      pIndex hNode, idxNode
      lIndex = m_uNodeData(idxNode).idxExpImg
      If (lIndex > IMG_NONE And lIndex < ImageListCount) Then
            NodeExpandedImage = lIndex
         Else
            NodeExpandedImage = IMG_NONE
      End If
   End If
End Property
Public Property Let NodeExpandedImage(Optional ByVal hNode As Long, Optional ByRef idxNode As Long, _
                                      ByVal New_Index As Long)
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      If (New_Index > IMG_NONE And New_Index < ImageListCount) Then
            m_uNodeData(idxNode).idxExpImg = New_Index
         Else
            m_uNodeData(idxNode).idxExpImg = IMG_NONE
      End If
      Refresh hNode
   End If
End Property

' adds extra space between NodeImage and NodeText -> ie draw second image.
' use with cdLabelIndent, neglected with cdLabel  -> easy toggling.
' if negative, Text draws over (missing) image
' -> could use it for nonexpandable nodes without image (cdMixNoImage!=0).
Public Property Get NodeTextIndent(Optional ByVal hNode As Long, _
                                   Optional ByRef idxNode As Long) As Long
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      NodeTextIndent = m_uNodeData(idxNode).lIndent
   End If
End Property
Public Property Let NodeTextIndent(Optional ByVal hNode As Long, Optional ByRef idxNode As Long, _
                                   ByVal New_TextIndent As Long)
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      m_uNodeData(idxNode).lIndent = New_TextIndent
   End If
End Property

Public Property Get NodeItemdata(Optional ByVal hNode As Long, _
                                 Optional ByRef idxNode As Long) As Long
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      NodeItemdata = m_uNodeData(idxNode).lItemData
   End If
End Property
Public Property Let NodeItemdata(Optional ByVal hNode As Long, Optional ByRef idxNode As Long, _
                                 ByVal New_ItemData As Long)
   If (m_hTreeView) Then
      pIndex hNode, idxNode
      m_uNodeData(idxNode).lItemData = New_ItemData
   End If
End Property

' hNode is optional, if called directly after AddNode proc (speed)
Public Function CustomizeNode(Optional ByVal hNode As Long, _
                              Optional ByVal ForeColor As OLE_COLOR = CLR_NONE, _
                              Optional ByVal BackColor As OLE_COLOR = CLR_NONE, _
                              Optional ByVal Font As StdFont, _
                              Optional ByVal ExpandedImage As Long = IMG_NONE, _
                              Optional ByVal TextIndent As Long = 0, _
                              Optional ByVal ItemData As Long = 0) As Boolean
   Dim idxNode As Long
   Dim bNew    As Boolean

   If (m_hTreeView) Then

      If hNode = 0& Then
         idxNode = m_lNodeCount
         hNode = m_uNodeData(idxNode).hNode
      Else
         idxNode = pTVlParam(hNode)
      End If
      
      If idxNode Then
         With m_uNodeData(idxNode)
            .lForeColor = pTranslateColor(ForeColor)
            .lBackColor = pTranslateColor(BackColor)
            .idxFont = pAddFontIfRequired(Font, bNew)
#If AUTOFNT Then
            If bNew And m_bAutoFont Then
               pAdjustFontIfRequired
            End If
#End If
            If ExpandedImage <> IMG_NONE Then
               .idxExpImg = ExpandedImage
            Else
               .idxExpImg = IMG_NONE
            End If
            .lIndent = TextIndent
            .lItemData = ItemData
            
            pCalculateRcFont idxNode
         End With

         CustomizeNode = True
      Else
         Debug.Assert False
      End If
   End If
End Function

'========================================================================================
' NodeFont handling
'========================================================================================

Public Sub ClearFonts()
   Dim idxFont As Long
   For idxFont = 1& To m_iFontCount
      DeleteObject m_hFnt(idxFont)
   Next
   ReDim m_FntLF(0)
   ReDim m_hFnt(0)
   m_iFontCount = 0
End Sub

Private Function pChangeGroupFont(ByVal idxNode As Long, ByVal oFont As IFont, _
                                  Optional ByRef bNew As Boolean) As Long
   Dim idxFont    As Long
   Dim tLF        As LOGFONT
   
   idxFont = m_uNodeData(idxNode).idxFont
   If idxFont Then
      If DeleteObject(m_hFnt(idxFont)) = 0& Then
         Debug.Assert False
         Exit Function
      End If
      If Not (oFont Is Nothing) Then
         pvStdFontToLogFont oFont, tLF
         LSet m_FntLF(idxFont) = tLF
         m_hFnt(idxFont) = CreateFontIndirect(tLF)
      Else
         LSet m_FntLF(idxFont) = tLF   ' or ZeroMemory
         m_hFnt(idxFont) = 0&
         bNew = True
      End If
      
      pChangeGroupFont = idxFont
   End If
End Function

' reduce gdi font handles, by adding only unique fonts and reusing hFont otherwise
Private Function pAddFontIfRequired(ByVal oFont As IFont, _
                                    Optional ByRef bNew As Boolean) As Long
   Dim idxFont    As Long
   Dim tLF        As LOGFONT
   Static hAdded  As Long
   
   If oFont Is Nothing Then
      ' reset to Tree font
      pAddFontIfRequired = 0&
      ' AutoFont: delete unneeded font
      bNew = True
      Exit Function
   End If
   
   If oFont.hFont = hAdded Then  ' passed font is same as last added?
      pAddFontIfRequired = m_iFontCount
      bNew = False
      Exit Function
   End If
   
   pvStdFontToLogFont oFont, tLF
   
   ' compare if passed font is already in array (this takes time ..)
   For idxFont = 1& To m_iFontCount 'm_iFontCount To 1 Step -1
      If m_hFnt(idxFont) <> 0& Then
         If pAreLogFontsEqual(tLF, m_FntLF(idxFont)) Then
            ' font exists
            bNew = False
            pAddFontIfRequired = idxFont
            Exit Function
         End If
      End If
   Next idxFont
   
   ' add oFont as new font
   m_iFontCount = m_iFontCount + 1&
   ReDim Preserve m_FntLF(1 To m_iFontCount) As LOGFONT
   ReDim Preserve m_hFnt(1 To m_iFontCount) As Long

   LSet m_FntLF(m_iFontCount) = tLF
   m_hFnt(m_iFontCount) = CreateFontIndirect(tLF)
   
   hAdded = oFont.hFont
   bNew = True
   pAddFontIfRequired = m_iFontCount
   
#If FNT_DBG Then
   Debug.Print "FONTCOUNT: " & m_iFontCount
   For idxFont = 1& To m_iFontCount
      With m_FntLF(idxFont)
         If m_hFnt(idxFont) Then
#If UNICODE Then
            Debug.Print "Font(" & idxFont & "): " & .lfFaceName & " , " & (.lfWeight >= FW_BOLD) & " , " & CStr(-(.lfHeight * Screen.TwipsPerPixelY / 20))
#Else
            Debug.Print "Font(" & idxFont & "): " & StrConv(.lfFaceName, vbUnicode) & " , " & (.lfWeight >= FW_BOLD) & " , " & CStr(-(.lfHeight * Screen.TwipsPerPixelY / 20))
#End If
         Else
            Debug.Print "Font(" & idxFont & "): deleted"
         End If
      End With
   Next
#End If
End Function

Private Function pAreLogFontsEqual(tA As LOGFONT, tB As LOGFONT) As Boolean
   Dim lSize   As Long
   Dim a()     As Byte
   Dim b()     As Byte
   
   lSize = LenB(tA)
   ReDim a(lSize):               ReDim b(lSize)
   CopyMemory a(0), tA, lSize:   CopyMemory b(0), tB, lSize
   
   pAreLogFontsEqual = (InStrB(1&, a, b) = 1&)
End Function

' (re-)calculate ptFont member for changes in NodeText or NodeFont
' ptFont is invariant for Tree font,ItemHeight,TextIndent
Private Sub pCalculateRcFont(ByVal idxNode As Long, Optional sText As String)
   Dim hFontOld   As Long
   Dim rcText     As RECT2
   
   With m_uNodeData(idxNode)
   
      If .idxFont <> 0& Then
         If LenB(sText) = 0& Then sText = m_uNodeData(idxNode).sText
         
         hFontOld = SelectObject(m_HDC, m_hFnt(.idxFont))
#If UNICODE Then
         .ptFont.cY = DrawTextW(m_HDC, StrPtr(sText), -1, rcText, _
                               DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX)
#Else
         .ptFont.cY = DrawTextA(m_HDC, sText, Len(sText), rcText, _
                               DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX)
#End If
         .ptFont.cX = rcText.X2
         
         SelectObject m_HDC, hFontOld
         
      Else
         ' Node uses tree font
         .ptFont.cX = 0&
         .ptFont.cY = 0&
      End If
      
   End With
   
End Sub

Private Function pGetRcTreeFont(ByRef sText As String) As Long
   Dim hFontOld   As Long
   Dim rcText     As RECT2
   
   hFontOld = SelectObject(m_HDC, m_hFont)
#If UNICODE Then
   DrawTextW m_HDC, StrPtr(sText), -1, rcText, DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX
#Else
   DrawTextA m_HDC, sText, Len(sText), rcText, DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX
#End If
   SelectObject m_HDC, hFontOld
   
   pGetRcTreeFont = rcText.X2
End Function

Private Sub pUpdateRcTreeFonts()
   Dim hFontOld   As Long
   Dim rcText     As RECT2
   Dim idxNode    As Long
   
   hFontOld = SelectObject(m_HDC, m_hFont)
   
   For idxNode = 1 To m_lNodeCount
      With m_uNodeData(idxNode)
#If UNICODE Then
         DrawTextW m_HDC, StrPtr(.sText), -1, rcText, DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX
#Else
         DrawTextA m_HDC, .sText, Len(.sText), rcText, DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX
#End If
         .xFont = rcText.X2
      End With
   Next
   
   SelectObject m_HDC, hFontOld
End Sub

Private Function pSubstituteHotFont(ByRef tLF As LOGFONT) As Long
   Dim tLFHot As LOGFONT
   
   LSet tLFHot = tLF
   tLFHot.lfUnderline = 1&
   pSubstituteHotFont = CreateFontIndirect(tLFHot)
End Function

'========================================================================================
' AutoFont  handling
'========================================================================================

#If AUTOFNT Then

' Feasible,if NodeFonts are unknown at design-time (ie user-configurable).
' - destroys unneeded NodeFonts.
' - raises AdjustFont event with proposed ItemHeight, if it should be changed
'   to avoid clipping or meager appearance.

Public Property Get AutoFont() As Boolean
   AutoFont = m_bAutoFont
End Property
Public Property Let AutoFont(ByVal bAutoFont As Boolean)
   m_bAutoFont = bAutoFont
End Property

Private Sub pAdjustFontIfRequired()
   Dim lItemHeight   As Long
   Dim lH            As Long
   Dim bMax          As Boolean
   Dim bMin          As Boolean
   Dim bHeight       As Boolean
   Dim tMax          As SIZEAPI

   pDestroyUnneededFonts
   
   If pGetMaxFont(tMax) Then
      ' needed ItemHeight, derived from maximum font
      lH = tMax.cY + tMax.cY Mod 2              ' round up to next even
      If lH < m_lImageH Then lH = m_lImageH     ' image height as mininmum
      
      lItemHeight = Me.ItemHeight
      bMax = lH > lItemHeight
      bMin = (lItemHeight - lH) > 2             ' incremental ItemHeight is 2 pixels
      
      If bMax Or bMin Then
         
         RaiseEvent AdjustFont(lH, bHeight, bMax)
         
         If bHeight And lH <> lItemHeight Then
#If FNT_DBG Then
            Debug.Print "HEIGHT ADJUSTED", lItemHeight, "NEW: " & lH
#End If
            Me.ItemHeight = lH
         End If
      End If
   End If
   
End Sub

Private Function pDestroyUnneededFonts() As Boolean
   Dim idxFont    As Long
   Dim idxNode    As Long
   Dim bOk        As Boolean
   Dim tLF        As LOGFONT
   
   For idxFont = 1& To m_iFontCount
      bOk = False
      For idxNode = 1& To m_lNodeCount
         If m_uNodeData(idxNode).idxFont = idxFont Then
            bOk = True: Exit For
         End If
      Next
      If Not bOk Then
         DeleteObject m_hFnt(idxFont)
         m_hFnt(idxFont) = 0&
         LSet m_FntLF(idxFont) = tLF
         pDestroyUnneededFonts = True
      End If
   Next
   
End Function

' interested only in setting ItemHeight: -> tMax.cX discarded
Private Function pGetMaxFont(ByRef tMax As SIZEAPI) As Long
   Dim idxFont       As Long
   Dim hFont         As Long
   Dim hFontOld      As Long
   Dim tS            As SIZEAPI
   
   ' dimensions for current Tree font
   hFontOld = SelectObject(m_HDC, m_hFont)
#If UNICODE Then
   GetTextExtentPoint32W m_HDC, StrPtr(TESTEXTENT), Len(TESTEXTENT), tMax
#Else
   GetTextExtentPoint32A m_HDC, TESTEXTENT, Len(TESTEXTENT), tMax
#End If
   
   For idxFont = 1& To m_iFontCount
      hFont = m_hFnt(idxFont)
      
      If hFont <> 0& Then
         SelectObject m_HDC, hFont
#If UNICODE Then
         GetTextExtentPoint32W m_HDC, StrPtr(TESTEXTENT), Len(TESTEXTENT), tS
#Else
         GetTextExtentPoint32A m_HDC, TESTEXTENT, Len(TESTEXTENT), tS
#End If
'         If (tS.cX > tMax.cX) Then tMax.cX = tS.cX
         If (tS.cY > tMax.cY) Then tMax.cY = tS.cY:   pGetMaxFont = idxFont
      End If
   Next
   
   SelectObject m_HDC, hFontOld
   
End Function

#End If ' AUTOFNT

'========================================================================================
' OwnerDraw drawing routines
'========================================================================================

Private Function pTreeOwnerDraw(ByVal hwnd As Long, ByVal lParam As Long) As Long
   Dim NMTVCD        As NMTVCUSTOMDRAW
   Dim hNode         As Long
   Dim idxNode       As Long
   Dim idxImg        As Long
   Dim hFontHot      As Long
   Dim hFontOld      As Long
   Dim hBr           As Long
   Dim hBrOld        As Long
   Dim hPen          As Long
   Dim hPenOld       As Long
   Dim bExpanded     As Boolean
   Dim bDrawExpanded As Boolean
   Dim bSelected     As Boolean
   Dim bMulSelected  As Boolean
   Dim bFocus        As Boolean
   Dim bHot          As Boolean
   Dim bHasButton    As Boolean
   Dim rcItemCorr    As RECT2
   Dim rcItemBK      As RECT2
   Dim rcItem        As RECT2
   Dim rcText        As RECT2
   Dim rcImg         As RECT2
   Dim rcState       As RECT2
   Dim rcBtn         As RECT2
   Dim rc            As RECT2
   Dim tS            As SIZEAPI
   Dim tS2           As SIZEAPI
   Dim lhDC          As Long
   Dim hTheme        As Long
   Dim xPos          As Long
   Dim yPos          As Long
   Dim lImageW       As Long
   Dim lImageH       As Long
   Dim lItemIndent   As Long
   Dim lR            As Long
   
   ' Get the CustomDraw data (iLevel member for COMCTL>= 4.71 only)
   CopyMemory NMTVCD, ByVal lParam, Len(NMTVCD)
   
   ' First see what stage of painting:
   Select Case NMTVCD.NMCD.dwDrawStage

      Case CDDS_PREPAINT
         ' In WM_Paint cycle GetUpdateRegion returns nil,
         ' but NMTVCD.NMCD.rc holds (empty) UpdateRect.
         
         If (IsRectEmpty(NMTVCD.NMCD.rc) = 0&) Then
            ' ask for CDDS_ITEMPREPAINT notifications
            pTreeOwnerDraw = CDRF_NOTIFYITEMDRAW
            
#If DRAW_DBG Then
            Debug.Print "CDDS_PREPAINT ####################"
#End If
            
            ' Erase background below last node regardless of current update region
            pTVItemRect NodeLastVisible(bLastNode:=True), rcItem, OnlyText:=False
            GetClientRect m_hTreeView, rc
            If rc.Y2 > rcItem.Y2 Then
               rc.Y1 = rcItem.Y2
               pFillRect NMTVCD.NMCD.hdc, rc, m_lTreeColors(clrTreeBK)
#If DRAW_DBG Then
               Debug.Print "LastNode ERASEBK"
#End If
            End If
            
         Else
            ' skip CDDS_ITEMPREPAINT notifications if update region is empty
            ' -> fast expanding of nodes
            pTreeOwnerDraw = CDRF_DODEFAULT
            
         End If
         
      Case CDDS_ITEMPREPAINT
         ' An item is being drawn: we will do all painting.
         pTreeOwnerDraw = CDRF_SKIPDEFAULT

         If (RectVisible(NMTVCD.NMCD.hdc, NMTVCD.NMCD.rc) = 0&) Then
            ' node above the ones needing updating: unnecessary, because drawing is clipped
#If DRAW_DBG Then
            Debug.Print "CDDS_ITEMPREPAINT: ---SKIPPED---", NodeText(NMTVCD.NMCD.dwItemSpec)
#End If
            Exit Function
         End If

         hNode = NMTVCD.NMCD.dwItemSpec
         idxNode = NMTVCD.NMCD.lItemlParam   ' != pTVlParam(hNode)

#If DRAW_DBG Then
         Debug.Print "CDDS_ITEMPREPAINT: " & NMTVCD.iLevel, NodeText(NMTVCD.NMCD.dwItemSpec)
#End If

' ********* XP Visual Styles **************************************************************

         If (m_lComctlVersion >= 6&) Then
            hTheme = OpenThemeData(hwnd, StrPtr("Treeview"))
         End If
         
' ********* State *************************************************************************

         bSelected = ((NMTVCD.NMCD.uItemState And CDIS_SELECTED) = CDIS_SELECTED)
         bFocus = ((NMTVCD.NMCD.uItemState And CDIS_FOCUS) = CDIS_FOCUS)
         bHot = ((NMTVCD.NMCD.uItemState And CDIS_HOT) = CDIS_HOT)
         
         bExpanded = pTVState(hNode, TVIS_EXPANDED)
         ' When only child is deleted, parent's expanded/expandedonce states don't
         ' change(COMCTL default) -> for bExpanded check for children.
         If bExpanded Then
            bExpanded = pTVHasChildren(hNode)
         End If
      
#If MULSEL Then
         If m_bMultiSelect And Not bSelected Then
            bMulSelected = NodeSelected(hNode)
         End If
'#Else   bMulSelected = False
#End If
         
' ********* Geometry ********************************************************************
        
' x offset in DC for ButtonCenter:
' dX = ButtonSize -1 + Level * Indent
' y offset in DC for ButtonCenter:
' dY = ???

' All xy Pos & Rects are in Window Client coordinates (not DC!)

' NMTVCD.NMCD.rc is bounding rect of total row (like TVM_GETITEMRECT/False)
' rcItem is bounding rect of Label

' Horizontal axis of all knots of this node != const in this proc
' yPos = center of NodeRect = (NMTVCD.NMCD.rc.Y1 + NMTVCD.NMCD.rc.Y2) \ 2

' x offset in client area for ButtonCenter:
' xPos(Button) + Indent + lImageW \ 2 + m_lImageW + SPACE_IL != rcItem.X1

' Indent: distance between Button center and  NodeImage center
'    OR : distance between Button center and StateImage center (TVS_CHECKBOXES style)

' StateImage.Right = NodeImage.Left
            
' O.Level        knot: Root node(s)
'           xPos = xPos(Button) - i * Indent
'
' (i-j)L.     (4 + k). knots: vertical tree lines of nodes with lower level
'           xPos = xPos(Button) - j * Indent
'
' i.Level      prev siblings: vertical treelines intersect 4. knot
'
' i= NodeL.   4. knot: center of Button
'           xPos = rcItem.X1 - SPACE_IL - m_lImageW - lImageW \ 2 - Indent
'
' i.Level      next siblings: vertical treelines offspring downwards on 4. knot
'
'             3. knot: center of StateImage
'           xPos = rcItem.X1 - SPACE_IL - m_lImageW - lImageW \ 2
'              \
' (i+1)L.       - children: vertical treelines offspring downwards on 2. or 3. knot
'              /
'             2. knot: center of NodeImage
'           xPos = rcItem.X1 - SPACE_IL - m_lImageW \ 2
'
'             1. knot: begin of Label
'           xPos = rcItem.X1
            
         ' rcItem is bounding rect of Label
         rcItem.X1 = hNode
         SendMessage hwnd, TVM_GETITEMRECT, 1&, rcItem
         
         ' Buffer DC is used only for the current item, not the whole client area.
         With NMTVCD.NMCD.rc
            If (.X2 - .X1) > m_tMemDC(0).lWidth Or (.Y2 - .Y1) > m_tMemDC(0).lHeight Then
               pCreateDC 0&, (.X2 - .X1), (.Y2 - .Y1)
            End If
            ' transform rcItem and NMTVCD.NMCD.rc for memDC
            m_tpOffset.X = .X1
            m_tpOffset.Y = .Y1
            OffsetRect NMTVCD.NMCD.rc, -m_tpOffset.X, -m_tpOffset.Y
            OffsetRect rcItem, -m_tpOffset.X, -m_tpOffset.Y
         End With
         lhDC = m_tMemDC(0).hdc

         ' Horizontal axis of all knots of this node != const in this proc
         ' yPos: center of NodeRect
         yPos = (NMTVCD.NMCD.rc.Y1 + NMTVCD.NMCD.rc.Y2) \ 2
         
         ' get real Item width for NodeFont (rcItem is based on Tree font)
         rcItemCorr = pGetItemRectReal(hNode, rcItem, rcText, idxNode)
         
         ' Drawing order: Background,Label,TextIndent,Image,StateImage,Button,Lines
        
' ********* Color *************************************************************************
         
         If m_bDrawColor Then
            pTreeCustomColor idxNode, bSelected Or bMulSelected, bHot, _
                             NMTVCD.clrText, NMTVCD.clrTextBk
         End If

' ********* Background ********************************************************************

         Debug.Assert OPAQUE = SetBkMode(lhDC, OPAQUE)
         
         ' erase Background of complete item
         pFillRect lhDC, NMTVCD.NMCD.rc, m_lTreeColors(clrTreeBK)
         
         ' # Option1: height of item background ~ text height          #
         rcItemBK.X1 = rcItemCorr.X1
         rcItemBK.X2 = rcItemCorr.X2
         rcItemBK.Y1 = rcText.Y1 + (rcText.Y1 > rcItemCorr.Y1)
         rcItemBK.Y2 = rcText.Y2 - (rcText.Y2 < rcItemCorr.Y2)
         
'         ' # Option2: height of item background = item height = const #
'         LSet rcItemBK = rcItemCorr
         
         If NMTVCD.clrTextBk <> m_lTreeColors(clrTreeBK) Then
            ' - draw background for selected,multiselected,hilit node
            ' - draw NodeBackColor
            If Not (m_bFullRowSelect And Not m_bHasLines) Then
               pFillRect lhDC, rcItemBK, NMTVCD.clrTextBk
            Else
               ' FullRowSelect mode: draw full row background
               pFillRect lhDC, NMTVCD.NMCD.rc, NMTVCD.clrTextBk
            End If
         End If
            
            
' ********* Font **************************************************************************

         ' destroy hFontHot & hFontOld in CleanUp section
         
         If m_bDrawFont And (m_uNodeData(idxNode).idxFont <> 0&) Then
            ' NodeFont
            If bHot Then
               ' use underlined font
               hFontHot = pSubstituteHotFont(m_FntLF(m_uNodeData(idxNode).idxFont))
               hFontOld = SelectObject(lhDC, hFontHot)
            Else
               hFontOld = SelectObject(lhDC, m_hFnt(m_uNodeData(idxNode).idxFont))
            End If
         Else
            ' Tree font ' BUGFIX5
            If bHot Then
               ' use underlined font
               hFontHot = pSubstituteHotFont(m_iFontLF)
               hFontOld = SelectObject(lhDC, hFontHot)
            Else
               hFontOld = SelectObject(lhDC, m_hFont)
            End If
         End If

' ********* Label *************************************************************************

         ' 1. knot: begin of Label
         ' xPos = rcItem.X1

         SetTextColor lhDC, NMTVCD.clrText
         SetBkMode lhDC, TRANSPARENT
            
         ' we draw vertical centered, comctl draws at top
#If UNICODE Then
         DrawTextW lhDC, StrPtr(m_uNodeData(idxNode).sText), -1, _
                   rcText, DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX Or DT_NOCLIP
#Else
         DrawTextA lhDC, m_uNodeData(idxNode).sText, Len(m_uNodeData(idxNode).sText), _
                   rcText, DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX Or DT_NOCLIP
#End If
      
       '' DrawFocusRect can 't be scrolled.Using buffer DC coord transformation produces
       '' a full line orange focus rect. (? only on WinXP,Win2K is OK ?) -> draw afterwards
       ' If bFocus Then
       '    ' draw focus rectangle, unless in FullRowSelect mode
       '    If Not (m_bFullRowSelect And Not m_bHasLines) Then
       '       DrawFocusRect lhDC, rcItemBK
       '    End If
       ' End If

         SetBkMode lhDC, OPAQUE
            
' ********* TextIndent ********************************************************************

         ' Free space between begin of Label and begin of Text at your disposal.
         ' Width determined individually by NodeTextIndent.Valid only with cdLabelIndent.
            
'         If m_bDrawLabelTI Then
'            ' TextIndent rectangle rc
'            LSet rc = rcItem
'            rc.X2 = rcItemCorr.X1
'            Debug.Assert rc.X2 - rc.X1 = NodeTextIndent(hNode)
'            DrawFocusRect lhDC, rc
'         End If
            
' ********* rcItem.X1 & Space *************************************************************
         
         ' rcItem.X1 located here
         ' Standard 3 pixels space between Image and Label rcItem.X1
         ' Const SPACE_IL = 3& '(const for all sizes,versions & options)
         
' ********* Image *************************************************************************
              
         idxImg = IMG_NONE
         
         ' 2. knot: center of NodeImage
         xPos = rcItem.X1 - SPACE_IL - m_lImageW \ 2
         SetRect rcImg, xPos, yPos, xPos, yPos
         
         If (m_hImageList(ilNormal) <> 0&) Then
            
            If m_bDrawExpanded Then
               ' # need to refresh node on Collapse ! (no CustomDraw notification sent) #
               If bExpanded Then

                  If m_uNodeData(idxNode).idxExpImg <> IMG_NONE Then

                     If bSelected Or bMulSelected Then
                        ' selected: draw expanded only with ssPriorityExpanded style
                        bDrawExpanded = m_bPriorityExpanded
                     Else
                        ' not selected
                        bDrawExpanded = True
                     End If

                     If bDrawExpanded Then
                        ' draw expanded image
                        idxImg = m_uNodeData(idxNode).idxExpImg
                     End If

                  End If
                  
               End If
            End If   ' m_bDrawExpanded
               
            If Not bDrawExpanded Then
            
               If bSelected Then
                  ' draw selected image
                  idxImg = m_uNodeData(idxNode).idxSelImg
#If MULSEL Then
               ' only needed if all multiselected nodes display their selectedimages.
               ElseIf bMulSelected Then
                  If m_bMultiSelect And m_bMultiSelectImage Then
                     ' draw selected image
                     idxImg = m_uNodeData(idxNode).idxSelImg
                  End If
#End If
               End If   ' bSelected
            
            End If   ' Not bDrawExpanded
               
            If (idxImg = IMG_NONE) Then
               ' last chance: draw normal image
               idxImg = m_uNodeData(idxNode).idxImg
            End If
            
            If (idxImg <> IMG_NONE) Then
            
               ' calculate rect of image (vertical centered)
               SetRect rcImg, xPos, yPos, xPos + 1&, yPos + 1&
               InflateRect rcImg, m_lImageW \ 2, m_lImageH \ 2
               
               ' NMTVCD.NMCD.uItemState doesn't indicate ghosted icon.
               If pTVState(hNode, TVIS_CUT) = False Then
               
                  ImageList_Draw m_hImageList(ilNormal), idxImg, lhDC, _
                                 rcImg.X1, rcImg.Y1, ILD_TRANSPARENT _
                                 Or pvINDEXTOOVERLAYMASK(pTVOverlayImage(hNode))
               Else
                  ' BUGFIX6
                  ImageList_DrawEx m_hImageList(ilNormal), idxImg, lhDC, _
                                 rcImg.X1, rcImg.Y1, m_lImageW, m_lImageH, _
                                 CLR_NONE, CLR_NONE, _
                                 ILD_TRANSPARENT Or ILD_SELECTED _
                                 Or pvINDEXTOOVERLAYMASK(pTVOverlayImage(hNode))
                                 
                  ' # try other rgbBK,rgbFg : m_lTreeColors(clrTreeBK) , CLR_DEFAULT #
               End If
               
            End If
         
         End If   ' (m_hImageList(ilNormal) <> 0&)
            

' ********* State Image *******************************************************************

         ' comctl default:
         ' - State Images are drawn bottom aligned in ItemRect
         
         ' we center them

         If m_bCheckBoxes Then
            
            ' size of state image
            ImageList_GetIconSize m_hImageList(ilState), lImageW, lImageH
            Debug.Assert (m_hImageList(ilState) <> 0&)
            
            '  3. knot: center of StateImage
            xPos = rcItem.X1 - SPACE_IL - m_lImageW - lImageW \ 2
            
            ' calculate rect of stateimage (left of node image, vertical centered)
            SetRect rcState, xPos, yPos, xPos + 1&, yPos + 1&
            InflateRect rcState, lImageW \ 2, lImageH \ 2
            
            idxImg = NodeStateImage(hNode)
            
            If idxImg Then
               ' draw state image
               ImageList_Draw m_hImageList(ilState), idxImg, lhDC, _
                              rcState.X1, rcState.Y1, ILD_NORMAL
            End If
            
         Else
            '  3. knot: center of StateImage = 2. knot(center of NodeImage)
            Debug.Assert xPos = (rcItem.X1 - SPACE_IL - m_lImageW \ 2)
            ' set rect of stateimage
            SetRect rcState, xPos, yPos, xPos, yPos
         End If

            
' ********* Button ***********************************************************************

         ' comctl default:
         ' - Button size is derived from image size (not image height!! nor ItemHeight)
         '   -> you get a monstrous button for W >> H and ItemHeight > 16
         
         ' For all choices we keep width & height seperate in calculation and
         ' first choice is deriving size from state imagelist (Button as bitmap option).
         ' We stick to the formula, that button is a quarter of an image
         ' -> standard HitTest by comctl, otherwise customize it.
         
         lItemIndent = ItemIndent
         
          ' 1. knot : ButtonCenter xPos
         If Not m_bCheckBoxes Then
            xPos = rcItem.X1 - SPACE_IL - m_lImageW \ 2 - lItemIndent
         Else
            xPos = rcItem.X1 - SPACE_IL - m_lImageW - lImageW \ 2 - lItemIndent
         End If
         
         ' size of state image
         If (m_hImageList(ilState) = 0&) Then
            ' no state imagelist exists
            If (m_hImageList(ilNormal) <> 0&) Then
               ' use node image size
               tS.cX = m_lImageH:    tS.cY = m_lImageW
            Else
               ' no images: use default / # alternate related to ItemHeight #
               tS.cX = 16&:          tS.cY = tS.cX
            End If
         Else
            ImageList_GetIconSize m_hImageList(ilState), tS.cX, tS.cY
         End If
         
         ' Define the default button size, regardless whether drawn
         ' tS  ButtonSize: half of (state)image height + 1 pixel -> always uneven
         ' tS2 holds (ButtonSize -1) / 2
         tS.cX = tS.cX \ 2& + 1&:         tS.cY = tS.cY \ 2& + 1&
         tS2.cX = tS.cX \ 2&:             tS2.cY = tS.cY \ 2&
         
         ' button rect: xpos - W \ 2, ypos - H \ 2, xpos + W \ 2 + 1, ypos + H \ 2 + 1
         SetRect rcBtn, xPos, yPos, xPos, yPos
         
         If m_bHasButtons Then
            If (NMTVCD.iLevel <> 0) Then
               bHasButton = CBool(pTVcChildren(hNode))  ' == NodePlusMinusButton()
            Else
               ' comctl default: root nodes have button only when rootlines enabled
               If m_bHasRootLines Then
                  bHasButton = CBool(pTVcChildren(hNode))  ' == NodePlusMinusButton()
               End If
            End If
         End If
               
         If bHasButton Then
         
            SetRect rcBtn, xPos, yPos, xPos + 1&, yPos + 1&
            InflateRect rcBtn, tS2.cX, tS2.cY
            
            If (hTheme <> 0&) Then
               
'               ' dummy rect to get real button size (Luna 9x9)
'               LSet rcBtn = rcItem
'
'               GetThemePartSize hTheme, lHDC, TVP_GLYPH, _
'                                IIf(bExpanded, GLPS_OPENED, GLPS_CLOSED), _
'                                rcBtn, TS_DRAW, tS
'
'               ' button rect: xpos - W \ 2, ypos - H \ 2, xpos + W \ 2 + 1, ypos + H \ 2 + 1
'               tS2.cX = tS.cX \ 2&:    tS2.cY = tS.cY \ 2&
'               SetRect rcBtn, xPos, yPos, xPos + 1&, yPos + 1&
'               InflateRect rcBtn, tS2.cX, tS2.cY
                                             
               DrawThemeBackground hTheme, lhDC, TVP_GLYPH, _
                                   IIf(bExpanded, GLPS_OPENED, GLPS_CLOSED), _
                                   rcBtn, rcBtn
                                   
            ElseIf False Then
               ' Button as bitmap from (external) state imagelist
               ' Index Convention:
               ' - Collapsed (+) is image with highest index = ImageListCount(ilState) -1
               ' - Expanded  (-) is image before             = ImageListCount(ilState) -2
               ' - No Button                                 = 0
               ' Size Convention:
               ' - Button part should be centered in bitmap for standard HitTest behaviour.
               
               idxImg = ImageListCount(ilState) - 1& + CLng(bExpanded)
               
               If (idxImg > 0&) Then
               
                  SetRect rcState, xPos, yPos, xPos, yPos
                  InflateRect rcState, tS.cX - 1&, tS.cY - 1&
                  Debug.Assert lImageW = (rcState.X2 - rcState.X1)
                  Debug.Assert lImageH = (rcState.Y2 - rcState.Y1)
                  
                  ' draw button image
                  ImageList_Draw m_hImageList(ilState), idxImg, lhDC, _
                                 rcState.X1, rcState.Y1, ILD_TRANSPARENT
               End If
               
            Else
               ' Draw default windows button 16 to 48+ size
               
               ' Draw Button rectangle with LineColor
               hBr = CreateSolidBrush(LineColor)
               lR = FrameRect(lhDC, rcBtn, hBr)
               DeleteObject hBr:    hBr = 0&
               
               ' +/- Signs
               If (lImageH >= 32&) Then
                  ' outer parts as 3 pixels wide rectangles with tree ForeColor
                  hBr = CreateSolidBrush(m_lTreeColors(clrTree))
                  ' -
                  SetRect rc, xPos, yPos, xPos + 1&, yPos + 1&
                  InflateRect rc, tS2.cX \ 2& + 1&, 1&
                  lR = FrameRect(lhDC, rc, hBr)
                  If Not bExpanded Then
                     ' |
                     SetRect rc, xPos, yPos, xPos + 1&, yPos + 1&
                     InflateRect rc, 1&, tS2.cY \ 2& + 1&
                     lR = FrameRect(lhDC, rc, hBr)
                  End If
                  DeleteObject hBr
               End If
               
               ' Signs: 1 pixel lines with tree ForeColor
               ' Image >= 32: inner part with tree BackColor drawn over rectangles
               hPen = CreatePen(PS_SOLID, 0&, _
                                   m_lTreeColors(IIf(lImageH < 32&, clrTree, clrTreeBK)))
               hPenOld = SelectObject(lhDC, hPen)
               ' -
               MoveToExNull lhDC, xPos - tS2.cX \ 2&, yPos, 0&
               LineTo lhDC, xPos + tS2.cX \ 2& + 1, yPos
               If Not bExpanded Then
                  ' |
                  MoveToExNull lhDC, xPos, yPos - tS2.cY \ 2&, 0&
                  LineTo lhDC, xPos, yPos + tS2.cY \ 2& + 1
               End If
               DeleteObject SelectObject(lhDC, hPenOld)
               
            End If
            
         End If
            
' ********* Tree Lines *******************************************************************

         ' eRop = PATCOPY    : dotted line with LineColor
         ' eRop = PATINVERT  : dotted line with inverted LineColor, ButtonRect has LineColor
         ' eRop = DSTINVERT  : solid black line,                    ButtonRect has LineColor
         Const eRop As Long = PATCOPY

         If m_bHasLines Then
            
            Debug.Assert m_hBrDot
            Debug.Assert OPAQUE = SetBkMode(lhDC, OPAQUE)
            
            ' LineColor
            SetTextColor lhDC, LineColor
            SetBkColor lhDC, m_lTreeColors(clrTreeBK)
            hBrOld = SelectObject(lhDC, m_hBrDot)
            
            If lImageW Then
               ' Horizontal line: Button to StateImage
               PatBlt lhDC, rcBtn.X2, yPos, rcState.X1 - rcBtn.X2, 1&, eRop
               
               ' Vertical treeline of child, offspring from StateImage
               If bExpanded Then
                  PatBlt lhDC, (rcState.X1 + rcState.X2) \ 2, rcState.Y2, _
                         1, rcItem.Y2 - rcState.Y2, eRop
               End If
               
            Else
               ' Horizontal line: Button to NodeImage
               If idxImg <> IMG_NONE Then
                  PatBlt lhDC, rcBtn.X2, yPos, rcImg.X1 - rcBtn.X2, 1&, eRop
                  
                  ' Vertical treeline of child, offspring from NodeImage
                  If bExpanded Then
                     PatBlt lhDC, (rcImg.X1 + rcImg.X2) \ 2, rcImg.Y2, _
                            1&, rcItem.Y2 - rcImg.Y2, eRop
                  End If
               
               ElseIf (m_lImageW <> 0&) Then
                  ' Using images, but this node has none
                  Debug.Assert (rcImg.X1 = rcImg.X2) And (rcImg.Y1 = rcImg.Y2)
                  
                  If m_bMixNoImage Then
                     ' fill up missing treelines
                     PatBlt lhDC, rcBtn.X2, yPos, rcImg.X1 + m_lImageW \ 2 - rcBtn.X2, 1&, eRop
                     
                     ' Vertical treeline of child, offspring from NodeImage center
                     If bExpanded Then
                        PatBlt lhDC, (rcImg.X1 + rcImg.X2) \ 2, rcImg.Y2, _
                               1&, rcItem.Y2 - rcImg.Y2, eRop
                     End If
                     
                  Else
                     ' respect borders of missing NodeImage
                     PatBlt lhDC, rcBtn.X2, yPos, rcImg.X1 - m_lImageW \ 2 - rcBtn.X2, 1&, eRop
                     
                     ' Vertical treeline of child, offspring from NodeImage
                     If bExpanded Then
                        PatBlt lhDC, rcImg.X1, rcImg.Y2 + m_lImageH \ 2, _
                               1&, rcItem.Y2 - rcImg.Y2 - m_lImageH \ 2, eRop
                     End If
                  
                  End If
               End If
               
            End If
            
            If NMTVCD.iLevel <> 0 Then
               ' Vertical treelines of parent and siblings, offspring from Button
               
               ' upwards
               If NodeParent(hNode) Then
                  PatBlt lhDC, (rcBtn.X1 + rcBtn.X2) \ 2, rcItem.Y1, _
                         1&, rcBtn.Y1 - rcItem.Y1, eRop
               ElseIf NodePreviousSibling(hNode) Then
                  PatBlt lhDC, (rcBtn.X1 + rcBtn.X2) \ 2, rcItem.Y1, _
                         1&, rcBtn.Y1 - rcItem.Y1, eRop
               End If
               
               ' downwards
               If NodeNextSibling(hNode) Then
                  PatBlt lhDC, (rcBtn.X1 + rcBtn.X2) \ 2, rcBtn.Y2, _
                         1&, rcItem.Y2 - rcBtn.Y2, eRop
               End If
               
            Else
               ' Root node
               If m_bHasRootLines Then
                  ' Vertical rootlines of siblings, offspring from Button
                  ' upwards
                  If NodePreviousSibling(hNode) Then
                     PatBlt lhDC, (rcBtn.X1 + rcBtn.X2) \ 2, rcItem.Y1, _
                            1&, rcBtn.Y1 - rcItem.Y1, eRop
                  End If
                  ' downwards
                  If NodeNextSibling(hNode) Then
                     PatBlt lhDC, (rcBtn.X1 + rcBtn.X2) \ 2, rcBtn.Y2, _
                            1&, rcItem.Y2 - rcBtn.Y2, eRop
                  End If
               End If
            End If
            
            ' Vertical treelines of expanded nodes above
            ' start at 1. knot : ButtonCenter xPos
            Debug.Assert ((xPos = rcItem.X1 - SPACE_IL - m_lImageW \ 2 - lItemIndent) And Not m_bCheckBoxes) Or _
                         ((xPos = rcItem.X1 - SPACE_IL - m_lImageW - lImageW \ 2 - lItemIndent) And m_bCheckBoxes)

            Dim idxLevel   As Long
            Dim hParent    As Long
            Dim xPosLine   As Long
            
            idxLevel = NMTVCD.iLevel
            hParent = hNode
            xPosLine = xPos
            
            Do While idxLevel >= 0
            
               xPosLine = xPosLine - lItemIndent
               If xPosLine < 0 Then Exit Do
               
               hParent = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_PARENT, hParent)
               If hParent Then
                  ' Draw unless hParent is last sibling
                  If NodeNextSibling(hParent) <> 0 Then
                     PatBlt lhDC, xPosLine, rcItem.Y1, 1&, rcItem.Y2 - rcItem.Y1, eRop
                  End If
               End If
                  
               idxLevel = idxLevel - 1
            Loop
            
            SelectObject lhDC, hBrOld
         
         End If
         
' ********* project-specific customdraw routine ****************************************
            
         ' maybe refresh node on Collapse ! (no CustomDraw notification sent)

         If m_bDrawProject Then

            ' # simple example: draw overlayimage #
            If m_uNodeData(idxNode).lItemData <> 0& Then
               ImageList_Draw m_hImageList(ilNormal), m_uNodeData(idxNode).lItemData, lhDC, _
                              rcImg.X1, rcImg.Y1, ILD_TRANSPARENT
            End If
            
         End If   ' m_bDrawProject
         
         
' ********* BitBlt BufferDC ************************************************************

         BitBlt NMTVCD.NMCD.hdc, m_tpOffset.X, m_tpOffset.Y, _
                m_tMemDC(0).lWidth, m_tMemDC(0).lHeight, m_tMemDC(0).hdc, 0&, 0&, vbSrcCopy
                
         ' restore NMTVCD.NMCD.rc
         OffsetRect NMTVCD.NMCD.rc, m_tpOffset.X, m_tpOffset.Y
         
         If bFocus Then
            ' draw focus rectangle now, unless in FullRowSelect mode
            If Not (m_bFullRowSelect And Not m_bHasLines) Then
               OffsetRect rcItemBK, m_tpOffset.X, m_tpOffset.Y
               DrawFocusRect NMTVCD.NMCD.hdc, rcItemBK
            End If
         End If

' ********* CleanUp ********************************************************************
   
         If (hTheme <> 0&) Then
            CloseThemeData hTheme
         End If
         
         If hFontOld Then
            SelectObject lhDC, hFontOld
         End If
         If hFontHot Then
            DeleteObject hFontHot
         End If
               
' ********* PostPaint ******************************************************************
      
'      Case CDDS_POSTPAINT
'         ' faked drawstage in zSubclass_Proc/WM_PAINT : NMTVCD is empty!
'#If DRAW_DBG Then
'         Debug.Print "CDDS_POSTPAINT"
'#End If
         
   End Select  ' NMTVCD.NMCD.dwDrawStage

'With rcItem
'   Debug.Print "rcItem        ", .X1, .X2, .Y1, .Y2
'End With
'With NMTVCD.NMCD.rc
'   Debug.Print "NMTVCD.NMCD.rc", .X1, .X2, .Y1, .Y2
'End With

End Function

Private Sub pTreeCustomColor(ByVal idxNode As Long, _
                             ByVal bSelected As Boolean, ByVal bHot As Boolean, _
                             Optional ByRef clrText As Long, _
                             Optional ByRef clrTextBk As Long)
                             
   With m_uNodeData(idxNode)
   
      If bSelected Then
         ' bFocus!= False for bMulSelected
         If m_bFocus Or (m_eHideSelection = sfShowSelectionAlways) Then
            ' vbHighlightText, vbHighlight
            clrText = m_lTreeColors(clrSelected)
            clrTextBk = m_lTreeColors(clrSelectedBK)
            
         Else
            ' vbWindowText(CLR_NONE)
            If .lForeColor <> CLR_NONE Then clrText = .lForeColor
            
            If (m_eHideSelection = sfHideSelection) Then
               ' vbWindowText(CLR_NONE), vbWindowBackground(CLR_NONE)
               If .lBackColor <> CLR_NONE Then clrTextBk = .lBackColor
               
            Else
               ' vbWindowText(CLR_NONE), vbButtonFace
               clrTextBk = m_lTreeColors(clrSelectedNoFocusBk)
            End If
         End If

      Else

         If pTVState(.hNode, TVIS_DROPHILITED) Then  ' NodeHilit(.hNode)
            ' vbHighlightText, vbHighlight
            clrText = m_lTreeColors(clrHilit)
            clrTextBk = m_lTreeColors(clrHilitBK)
            
         ElseIf bHot Then
            ' COLOR_HOTLIGHT, vbWindowBackground(CLR_NONE)
            clrText = m_lTreeColors(clrHot)
            If .lBackColor <> CLR_NONE Then clrTextBk = .lBackColor
            
         Else
            ' vbWindowText(CLR_NONE), vbWindowBackground(CLR_NONE)
            If .lForeColor <> CLR_NONE Then clrText = .lForeColor
            If .lBackColor <> CLR_NONE Then clrTextBk = .lBackColor
         End If

      End If   ' bSelected
   End With
End Sub

Private Sub pTreeEraseBK()
   Dim tRD           As RGNDATA
   Dim hRgn          As Long
   Dim b()           As Byte
   Dim lBuffSize     As Long
   Dim lCount        As Long
   Dim i             As Long
   Const RCSIZE      As Long = 16      ' == LenB(rc(0))
   
   ' need dummy region
   hRgn = CreateRectRgn(0, 0, 10, 10)

   ' Retrieve the invalidated region
   If GetUpdateRgn(m_hTreeView, hRgn, 0&) > NULLREGION Then

      ' get required buffer size to retrieve all rectangles
      lBuffSize = GetRegionData(hRgn, 0&, ByVal 0&)

      If lBuffSize <> 0& Then

         ' Allocate a byte buffer
         ReDim b(0 To lBuffSize - 1)

         ' Get the data
         If GetRegionData(hRgn, lBuffSize, b(0)) <> 0 Then

            ' Buffer contains a RGNDATAHEADER struct and an unknown number of RECT's.
            ' number of invalidated rectangles
            lCount = (lBuffSize - LenB(tRD.rdh)) \ RCSIZE
            ReDim tRD.Buffer(1 To lCount)

            ' cast byte buffer as RGNDATA struct
            CopyMemory tRD.rdh, b(0), LenB(tRD.rdh)
            CopyMemory tRD.Buffer(1), b(LenB(tRD.rdh)), lCount * RCSIZE
            Erase b()

            Debug.Assert tRD.rdh.nCount = lCount

#If DRAW_DBG Then
            With tRD.rdh.rcBound
               Debug.Print "rcBound ", .X1, .Y1, .X2, .Y2, .X2 - .X1, .Y2 - .Y1
            End With
#End If

            ' The rectangles are sorted top to bottom, left to right.
            For i = 1 To lCount
               If (tRD.rdh.rcBound.Y2 - tRD.rdh.rcBound.Y1) >= ItemHeight Then
                  ' Use exact region to erase bk. Skip erasing partial rows.
                  pFillRect m_HDC, tRD.Buffer(i), m_lTreeColors(clrTreeBK)
               End If
#If DRAW_DBG Then
               With tRD.Buffer(i)
                  Debug.Print "InvalidRect " & i, .X1, .Y1, .X2, .Y2, .X2 - .X1, .Y2 - .Y1
               End With
#End If
            Next
         End If

         ' release RgnData
         Erase tRD.Buffer()

      End If

   End If

   ' Destroy the region
   DeleteObject (hRgn)

End Sub

' Display NoData text for Nodecount = 0
Private Sub pDrawNoDataText()
   Dim sText    As String
   Dim tR       As RECT2
   Dim hOldFnt  As Long
   Dim hFont    As Long
   Dim fnt      As IFont
   
   RaiseEvent NoDataText(sText, fnt)

   If LenB(sText) Then
      
      GetClientRect m_hTreeView, tR
      InflateRect tR, -16&, -16&

      ' ensure last drawn text is completely erased
      SendMessageLong m_hTreeView, WM_ERASEBKGND, m_HDC, 0&

      If (fnt Is Nothing) Then
         ' use treeview font for text
         hFont = m_hFont
      Else
         ' client supplied font
         hFont = fnt.hFont
      End If
      
      Debug.Assert hFont
      hOldFnt = SelectObject(m_HDC, hFont)
      SetBkMode m_HDC, TRANSPARENT

#If UNICODE Then
      DrawTextW m_HDC, StrPtr(sText), -1, tR, DT_WORDBREAK Or DT_NOPREFIX Or DT_EXPANDTABS
#Else
      DrawTextA m_HDC, sText, Len(sText), tR, DT_WORDBREAK Or DT_NOPREFIX Or DT_EXPANDTABS
#End If

      SetBkMode m_HDC, OPAQUE
      SelectObject m_HDC, hOldFnt

   End If

End Sub

'                  <--------- rcItem (TextOnly) ------------>
' -----------------------------------------------------------
'   |  rcImg  |3pix| TI |2pix|    rcText    |4pix|          |   TI: NodeTextIndent
' -----------------------------------------------------------
'                       <------- rcItemCorr------>

' get real Item width for NodeFont (rcItem is based on Tree font)
' call with original rcItem (TextOnly).
' returns correct rcItem, based on NodeFont & NodeTextIndent.
' rcItem & rcItemCorr have equal y1 & y2 members.
' rcText is bounding rectangle of vertically centered Text.
Private Function pGetItemRectReal(ByVal hNode As Long, rcItem As RECT2, rcText As RECT2, _
                                  Optional idxNode As Long) As RECT2
   Dim rcItemCorr    As RECT2
   Dim bFont         As Boolean

   pIndex hNode, idxNode
   
   With m_uNodeData(idxNode)
      bFont = m_bDrawFont And (.idxFont <> 0&)
      
      LSet rcItemCorr = rcItem
      If bFont Then
         rcItemCorr.X2 = rcItem.X1 + .ptFont.cX + 6&     ' 2 leading + 4 trailing pixels
      '  Else: for tree font rcItemCorr.X2 != rcItem.X2
      End If
      If m_bDrawLabelTI Then
         OffsetRect rcItemCorr, .lIndent, 0&             ' shift by TextIndent
      End If
      
      rcText.X1 = rcItemCorr.X1 + 2&                     ' add 2 leading pixels
      rcText.X2 = rcItemCorr.X2 - 4&                     ' subtract 4 trailing pixels
      
      If bFont Then
         ' rcText vertically centered in rcItemCorr
         rcText.Y1 = (rcItemCorr.Y1 + rcItemCorr.Y2) \ 2 - (.ptFont.cY \ 2)
         rcText.Y2 = rcText.Y1 + .ptFont.cY
      Else
         ' tree font
         rcText.Y1 = rcItemCorr.Y1
         rcText.Y2 = rcItemCorr.Y2
      End If
      
   End With
   
   pGetItemRectReal = rcItemCorr
End Function

Private Sub pFillRect(ByVal hdc As Long, ByRef rc As RECT2, ByVal lColor As Long)
   Dim hBr  As Long
   
   hBr = CreateSolidBrush(lColor)
   FillRect hdc, rc, hBr
   DeleteObject hBr
End Sub

Private Sub pCreateDotBrush()
   Dim tBM     As BITMAP
   Dim hBmp    As Long
   Dim i       As Long
   Dim lPattern(0 To 3) As Long
   
   For i = 0 To 3
      lPattern(i) = &HAAAA5555
   Next i

   pDestroyDotBrush
      
   ' Create a monochrome bitmap containing the desired pattern:
   tBM.bmType = 0
   tBM.bmWidth = 16
   tBM.bmHeight = 8
   tBM.bmWidthBytes = 2
   tBM.bmPlanes = 1
   tBM.bmBitsPixel = 1
   tBM.bmBits = VarPtr(lPattern(0))
   hBmp = CreateBitmapIndirect(tBM)

   ' Make a brush from the bitmap bits
   m_hBrDot = CreatePatternBrush(hBmp)
   ' Delete the useless bitmap
   DeleteObject hBmp
End Sub

Private Sub pDestroyDotBrush()
   If m_hBrDot <> 0 Then
      DeleteObject m_hBrDot
      m_hBrDot = 0
   End If
End Sub

Private Sub pRefreshNodeRects(Optional hNode As Long)
   
'   If hNode Then
'      ' # HowTo update a single ItemRect ??? #
'   Else
      ' This forces comctl to update it's ItemRect's by sending a TVN_GETDISPINFO for each node.
      TrackSelect = Not m_bTrackSelect
      TrackSelect = Not m_bTrackSelect
'   End If
End Sub

Private Sub pCreateDC(ByVal idxDC As Long, ByVal Width As Long, ByVal Height As Long)
   
   pDestroyDC idxDC
   
   With m_tMemDC(idxDC)
      
      .hdc = CreateCompatibleDC(UserControl.hdc)
      If (.hdc <> 0&) Then
         .hBmp = CreateCompatibleBitmap(UserControl.hdc, Width, Height)
         If (.hBmp <> 0&) Then
            .hBmpOld = SelectObject(.hdc, .hBmp)
            If (.hBmpOld <> 0&) Then
               .lWidth = Width
               .lHeight = Height
               Exit Sub
            End If
         End If
      End If
   End With
   
   pDestroyDC idxDC

End Sub

Private Sub pDestroyDC(ByVal idxDC As Long)

   With m_tMemDC(idxDC)
      If (.hBmpOld <> 0&) Then
         SelectObject .hdc, .hBmpOld
         .hBmpOld = 0&
      End If
      If (.hBmp <> 0&) Then
         DeleteObject .hBmp
         .hBmp = 0&
      End If
      If (.hdc <> 0&) Then
         DeleteDC .hdc
         .hdc = 0&
      End If
      .lWidth = 0&
      .lHeight = 0&
   End With
End Sub

#End If 'CUSTDRAW

#If MULSEL Then

'========================================================================================
' Multiple Selection
'========================================================================================

Public Property Get MultiSelect() As Boolean
   MultiSelect = m_bMultiSelect
End Property
Public Property Let MultiSelect(ByVal bMultiSelect As Boolean)
   m_bMultiSelect = bMultiSelect

   If m_bMultiSelect Then
      ' implementation needs CustomDraw with cdColor set
      Debug.Assert m_bDrawColor
      If (m_colSelected Is Nothing) Then
         Set m_colSelected = New Collection
      End If
      ' ensure SelectedNode is in collection
      If SelectedNode Then NodeSelected(SelectedNode) = True
   Else
      Set m_colSelected = Nothing
   End If

End Property

Public Property Get SelectionCount() As Long
   If m_bMultiSelect Then
      SelectionCount = m_colSelected.Count
   Else
      If SelectedNode Then SelectionCount = 1&
   End If
End Property

' 1 based: 1 to SelectionCount
Public Property Get SelectionNode(ByVal lIndex As Long) As Long
   If m_bMultiSelect Then
      On Error Resume Next
      SelectionNode = m_colSelected(lIndex)
   Else
      If (lIndex = 1&) And (SelectedNode > 0&) Then SelectionNode = SelectedNode
   End If
End Property

Public Property Get NodeSelected(ByVal hNode As Long) As Boolean
   If m_bMultiSelect Then
      On Error Resume Next
      NodeSelected = m_colSelected(CStr(hNode))
   Else
      NodeSelected = (hNode = SelectedNode)
   End If
End Property
Public Property Let NodeSelected(ByVal hNode As Long, bSelected As Boolean)
   If m_bMultiSelect Then
      On Error Resume Next
      If bSelected Then
         m_colSelected.Add hNode, CStr(hNode)
      Else
         m_colSelected.Remove CStr(hNode)
      End If
      ' redraw node
      Refresh hNode
   Else
      If bSelected Then
         SelectedNode = hNode
      Else
         If hNode = SelectedNode Then
            SelectedNode = 0&
         End If
      End If
   End If
End Property

Private Sub pRedrawSelection()
   Dim idxSel As Long
   
   If Not (m_colSelected Is Nothing) Then
      For idxSel = 1& To m_colSelected.Count
         Refresh m_colSelected(idxSel)
      Next
   End If
End Sub

Private Sub pSelectedNodeChanged(ByVal hSelNode As Long, ByVal lS As ShiftConstants, ByVal bMouseUp As Boolean)
   Dim bState     As Boolean
   Dim bClear     As Boolean

#If MULSEL_DBG Then
   Debug.Assert Not m_bInProc
   Debug.Assert hSelNode <> 0&
   Debug.Assert m_hEdit = 0& Or (Not bMouseUp And (lS = 0))
'   Debug.Assert m_hSelectionRoot <> 0
   
   Debug.Print "pSelectedNodeChanged", IIf(bMouseUp, "UP   ", "DOWN"), NodeText(hSelNode)
#End If

   If (lS And vbCtrlMask) Then
      If bMouseUp Then
         ' Delayed until MouseUp: Invert selection state for this node.
         bState = NodeSelected(hSelNode)
         NodeSelected(hSelNode) = Not bState
         If bState Then
            ' node in selection should be deselected, but tv selected it.
            ' Solution: deselect 'real' selected node by selecting other multiselected node.
            pDeselectNode hSelNode
         End If
         m_hSelectionRoot = SelectionNode(1&)                ' can be 0
      End If

   ElseIf (lS And vbShiftMask) And (m_hSelectionRoot <> hSelNode) Then
      If Not bMouseUp Then
         ' Immediately on MouseDown:
         ' Ensure all items between m_hSelectionRoot and SelectedNode
         ' are selected, and anything else is not selected.
#If MULSEL_DBG Then
         Debug.Print "ROOT before SHIFT: ", NodeText(m_hSelectionRoot)
#End If
         pSelectBetween hSelNode
      End If

   Else
      ' normal left click
      
      ' redraw nodes in previous selection, Clear selection.
      If bMouseUp Then
         ' MouseUp:     always
         bClear = True
      Else
         ' MouseDown:   if clicked outside of Selection or LabelEdit begin/termination
         bClear = Not NodeSelected(hSelNode) Or (m_hEdit <> 0&)
      End If
         
      If bClear Then
         If Not (m_colSelected Is Nothing) Then ' BUGFIX1
            With m_colSelected
               Do While .Count > 0
                  Refresh .Item(.Count)
                  .Remove (.Count)
               Loop
            End With
         Else: Debug.Assert False
         End If
      End If
      
      ' add 'real' selected node to selection on MouseUp
      If bMouseUp And (hSelNode = SelectedNode) Then
         NodeSelected(hSelNode) = True
         m_hSelectionRoot = hSelNode
      End If

   End If

#If MULSEL_DBG Then
   DebugSelection
#End If
End Sub

' triggers TVN_SELCHANGED : use m_bInProc to prevent
' - wrong NodeClick event (bNodeClick is False) in zSubclass_Proc.
' - recursion in pSelectedNodeChanged.
Private Function pDeselectNode(hSelNode As Long) As Boolean
   Dim hNewSelected  As Long
   Dim idxSel        As Long
   
   Debug.Assert hSelNode <> 0&
   Debug.Assert hSelNode = SelectedNode
   
   idxSel = SelectionCount

   If idxSel > 0& Then

      ' deselect 'real' selected node by selecting other multiselected node:
      ' the next last selected (= highest index in collection)
      Do While idxSel >= 1&
         hNewSelected = SelectionNode(idxSel)
         If hNewSelected <> hSelNode Then

            m_bInProc = True
            NodeSelected(hSelNode) = False
            SelectedNode = hNewSelected
            m_bInProc = False

            pDeselectNode = True
            Exit Do
         Else
            idxSel = idxSel - 1&
         End If
      Loop

   End If

   If Not pDeselectNode Then
      ' deselect single (='real') selected node
      m_bInProc = True
      NodeSelected(hSelNode) = False
      m_hSelectionRoot = 0&
      SelectedNode = 0&
      m_bInProc = False
      pDeselectNode = True
   End If

#If MULSEL_DBG Then
   If hNewSelected Then
      Debug.Print "MOVESELECTION", NodeText(hSelNode), NodeText(hNewSelected)
   Else
      Debug.Print "DESELECTION", NodeText(hSelNode)
   End If
   DebugSelection
#End If
End Function

Private Sub pSelectBetween(ByVal hSelNode As Long)
   Dim hStart     As Long
   Dim tR1 As RECT2, tR2 As RECT2

   hStart = m_uNodeData(1).hNode

   pTVItemRect m_hSelectionRoot, tR1, OnlyText:=False
   pTVItemRect hSelNode, tR2, OnlyText:=False
   ' compare item.top
   If tR2.Y1 < tR1.Y1 Then
      m_hSelectionRoot = pTopBottomInSelection(False)    ' bottom
   Else
      m_hSelectionRoot = pTopBottomInSelection(True)     ' top
   End If
   
   If m_hSelectionRoot Then
      pSelectBetweenIterate hStart, m_hSelectionRoot, hSelNode
   Else
      m_hSelectionRoot = hSelNode
   End If
#If MULSEL_DBG Then
   Debug.Print "BETWEEN", NodeText(hStart), NodeText(m_hSelectionRoot), NodeText(hSelNode)
#End If
End Sub

' http://vbaccelerator.com/home/VB/Code/Controls/TreeView/Multi-Select_TreeView/article.asp
' Bug:   selecting sibling(1) upwards to a child(2) of previous sibling(3),
'        selects all nodes below sibling(1) to bottom node.
' Cause: hNode1,hNode2 swapped and selection end not found
' Fix:   use hComp to store initial selection ending.
Private Sub pSelectBetweenIterate(ByVal hStart As Long, _
                                  ByVal hNode1 As Long, ByVal hNode2 As Long, _
                                  Optional ByRef bInSelection As Boolean = False, _
                                  Optional ByRef bNoneFoundYet As Boolean = True, _
                                  Optional ByRef hComp As Long = 0)
   Dim hSwap      As Long

   Do While hStart <> 0
      ' Debug.Print "hSTART", NodeText(hStart), NodeText(hNode1), NodeText(hNode2)
      If Not (bInSelection) Then
         If (bNoneFoundYet) Then
            If (hStart = hNode1) Then
               hComp = hNode2
               bInSelection = True
               bNoneFoundYet = False
            ElseIf (hStart = hNode2) Then
               hComp = hNode1
               hSwap = hNode2
               hNode2 = hNode1
               hNode1 = hSwap
               bInSelection = True
               bNoneFoundYet = False
            End If
         End If
      End If
      If NodeSelected(hStart) <> bInSelection Then
         NodeSelected(hStart) = bInSelection
         ' Debug.Print "SELECT", bInSelection, NodeText(hStart)
      End If
      If (bInSelection) Then
         If (hStart = hComp) Then
            bInSelection = False
         End If
      End If
      If (NodeExpanded(hStart) And NodeChildren(hStart) > 0) Then
         pSelectBetweenIterate NodeChild(hStart), hNode1, hNode2, bInSelection, bNoneFoundYet, hComp
      End If
      hStart = NodeNextSibling(hStart)
   Loop

End Sub

' bTop = False  Bottom
' bTop = True   Top
' Top,Bottom in selection in geometrical terms (ItemRect.Top) not index,level, hNode etc.
Private Function pTopBottomInSelection(ByVal bTop As Boolean) As Long
   Dim tR      As RECT2
   Dim lTop    As Long
   Dim hNode   As Long
   Dim hReturn As Long
   Dim idxSel  As Long
   Dim bComp   As Boolean

   If SelectionCount > 1 Then

      If bTop Then
         pTVItemRect SelectionNode(1), tR, OnlyText:=False
         lTop = tR.Y1 + 1
      End If

      For idxSel = 1 To SelectionCount
         hNode = SelectionNode(idxSel)
         pTVItemRect hNode, tR, OnlyText:=False

         bComp = (tR.Y1 > lTop)
         If bTop Then bComp = Not bComp

         If bComp Then
            lTop = tR.Y1
            hReturn = hNode
         End If
      Next

      Debug.Assert hReturn

   Else
      hReturn = SelectionNode(1) ' = 0 if SelectionCount = 0
   End If

   pTopBottomInSelection = hReturn

#If MULSEL_DBG Then
   Debug.Print IIf(bTop, "TOP_SELECTION    ", "BOTTOM_SELECTION"), NodeText(hReturn)
#End If
End Function

#If MULSEL_DBG Then
Private Sub DebugSelection()
   Dim i As Long
   Dim hNode As Long

   Debug.Assert m_bMultiSelect
   If SelectionCount = 1 Then Debug.Assert SelectedNode
   Debug.Print
   Debug.Print "SELECTED", SelectedNode, NodeText(SelectedNode)
   Debug.Print "ROOT", m_hSelectionRoot, NodeText(m_hSelectionRoot)
   For i = 1 To SelectionCount
      hNode = SelectionNode(i)
      Debug.Print "SEL" & i, hNode, NodeText(hNode)
   Next
   Debug.Print

End Sub
#End If 'MULSEL_DBG

#End If 'MULSEL

'========================================================================================
' Enums retain capitalization
'========================================================================================
#If False Then

Public sfNormal
Public sfHideSelection
Public sfShowSelectionAlways

Public ilNormal
Public ilState

Public daMinimal
Public daCustomData
Public daMultipleSelection
Public daCurrentChildren
Public daChildren
Public daInterProcess

Public OLE_FORMAT_ID
Public OLE_FORMAT_ID1
Public OLE_FORMAT_ID2
Public OLE_FORMAT_ID3

Public cdOff
Public cdColor
Public cdFont
Public cdExpandedImage
Public cdMixNoImage
Public cdLabel
Public cdLabelIndent
Public cdAll
Public cdProject

Public ssPrioritySelected
Public ssPriorityExpanded
Public ssImageSelected
Public ssImageMultiSelected

Public clrSelected
Public clrSelectedBK
Public clrSelectedNoFocusBk
Public clrHot
Public clrHilit
Public clrHilitBK

Public bsNone
Public bsFixedSingle

Public rLast
Public rFirst
Public rSort
Public rNext
Public rPrevious

Public drgNone
Public drgManual
Public drgAutomatic

Public drpNone
Public drpManual

Public disInsertMark
Public disDropHilite
Public disAutomatic

Public sHome
Public sPageUp
Public sUp
Public sDown
Public sPageDown
Public sEnd
Public sLeft
Public sPageLeft
Public sLineLeft
Public sLineRight
Public sPageRight
Public sRight

Public TVHT_NOWHERE
Public TVHT_ONITEMICON
Public TVHT_ONITEMLABEL
Public TVHT_ONITEMINDENT
Public TVHT_ONITEMBUTTON
Public TVHT_ONITEMRIGHT
Public TVHT_ONITEMSTATEICON
Public TVHT_ABOVE
Public TVHT_BELOW
Public TVHT_TORIGHT
Public TVHT_TOLEFT
Public TVHT_ONITEM
Public TVHT_ONITEMTEXTINDENT

#End If ' False






