VERSION 5.00
Begin VB.UserControl SComboBox 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2265
   KeyPreview      =   -1  'True
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   151
   ToolboxBitmap   =   "SComboBox.ctx":0000
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   390
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1965
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   105
      Top             =   600
   End
   Begin VB.TextBox txtCombo 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox picList 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   465
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
      Begin VB.VScrollBar scrollI 
         Height          =   1095
         Left            =   585
         Max             =   100
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   240
      End
   End
End
Attribute VB_Name = "SComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************'
'*        All rights Reserved © HACKPRO TM 2004       *'
'******************************************************'
'*                   Versión 1.0.0                    *'
'******************************************************'
'* Control:       SComboBox                           *'
'******************************************************'
'* Author:        Heriberto Mantilla Santamaría       *'
'******************************************************'
'* Collaboration: fred_cpp                            *'
'*                So many thanks for his contribution *'
'*                for this project, some styles and   *'
'*                Traduction to English of some       *'
'*                comments.                           *'
'******************************************************'
'* Description:   This usercontrol simulates a Combo- *'
'*                Box But adds new an great features  *'
'*                like:                               *'
'*                                                    *'
'*                - The first ComboBox On PSC that    *'
'*                  actually works in a single file   *'
'*                  control.                          *'
'*                - When the list is shown doesn't    *'
'*                  deactivate the parent form.       *'
'*                - More than 20 Visual Styles; no    *'
'*                  images Everything done by code.   *'
'*                - Some extra cool properties.       *'
'******************************************************'
'* Started on:    Friday, 11-jun-2004                 *'
'******************************************************'
'* Fixes:         - List.                  (18/06/04) *'
'*                - Control appearance.    (20/06/04) *'
'*                - Standard appearance.   (21/06/04) *'
'*                - MAC Appearance.        (24/06/04) *'
'*                - XP Appearance.         (25/06/04) *'
'*                - List Elements.         (27/06/04) *'
'*                - List Events.           (27/06/04) *'
'*                - Control properties.    (28/08/09) *'
'*                - Control properties.    (29/08/09) *'
'*                - List.                  (29/06/09) *'
'*                - List.                  (01/07/09) *'
'*                - JAVA Appearance.       (02/07/09) *'
'*                - List.                  (03/07/09) *'
'*                - Soft Style Appearance. (03/07/09) *'
'*                - Ardent Appearance.     (04/07/09) *'
'*                - List.                  (04/07/09) *'
'*                - MAC Appearance.        (04/07/04) *'
'******************************************************'
'*       Errors corrected after the publication       *'
'*                   Versión 1.0.1                    *'
'*                                                    *'
'*  - ScrollBar Slider.                    (08/07/04) *'
'*  - ListIndex Property.                  (08/07/04) *'
'*  - Down or Up list when press keys.     (09/07/04) *'
'*  - Drop down List.                      (09/07/04) *'
'*  - AddItem parameters.                  (09/07/04) *'
'*  - ChangeItem parameters.               (09/07/04) *'
'*  - SeparatorLine for Item.              (10/07/09) *'
'*  - Reorganize Code.                     (11/07/09) *'
'*  - ListIndex Property.                  (11/10/09) *'
'******************************************************'
'* Enhancements:  - Office Xp.             (13/06/04) *'
'*                - Win98.                 (13/06/04) *'
'*                - Control Properties.    (14/06/04) *'
'*                - Appearance.            (15/06/04) *'
'*                - WinXp.                 (16/06/04) *'
'*                - Office 2000.           (16/06/04) *'
'*                - Soft Style.            (16/06/04) *'
'*                - ItemTag Property.      (16/06/04) *'
'*                - JAVA.                  (17/06/04) *'
'*                - GradientV.             (18/06/04) *'
'*                - GradientH.             (18/06/04) *'
'*                - OrderList.             (18/06/04) *'
'*                - Color properties.      (19/06/04) *'
'*                - Explorer Bar.          (19/06/04) *'
'*                - Picture.               (19/06/04) *'
'*                - Mac.                   (21/06/04) *'
'*                - Special Border.        (22/06/04) *'
'*                - Rounded.               (22/06/04) *'
'*                - Search.                (23/06/04) *'
'*                - Style Arrow.           (25/06/04) *'
'*                - Light Blue.            (26/06/04) *'
'*                - KDE.                   (29/06/04) *'
'*                - Style Arrow.           (29/06/04) *'
'*                - NiaWBSS.               (30/06/04) *'
'*                - Rhombus.               (30/06/04) *'
'*                - Additional Xp.         (01/07/04) *'
'*                - Ardent.                (03/07/04) *'
'******************************************************'
'*          Enhancements after the publication        *'
'*                   Versión 1.0.1                    *'
'*                                                    *'
'*  - Drop down when you press F4.         (08/07/04) *'
'*  - Press Enter hidden list.             (08/07/04) *'
'*  - MouseIcon property.                  (08/07/04) *'
'*  - MousePointer property.               (08/07/04) *'
'*  - Down or Up list when press keys.     (08/07/04) *'
'*  - Set ListIndex when the text change.  (09/07/04) *'
'*  - AutoCompleteWord property.           (09/07/04) *'
'*  - Const VK_LBUTTON.                    (09/07/04) *'
'*  - Const VK_RBUTTON.                    (09/07/04) *'
'*  - New comments.                        (09/07/04) *'
'*  - Add SeparatorLine for item.          (09/07/04) *'           *'
'*  - Add MouseIcon for item.              (09/07/04) *'           *'
'******************************************************'
'* Release date:  Sunday, 04-07-2004                  *'
'******************************************************'
'* Note:     Comments, suggestions, doubts or bug     *'
'*           reports are wellcome to these e-mail     *'
'*           addresses:                               *'
'*                                                    *'
'*                  heri_05-hms@mixmail.com or        *'
'*                  hcammus@hotmail.com               *'
'*                                                    *'
'*    That lives the Soccer and the América of Cali   *'
'*             Of Colombia for the world.             *'
'******************************************************'
'*        All rights Reserved © HACKPRO TM 2004       *'
'******************************************************'
Option Explicit
 
 '****************************'
 '* English: Private Type.   *'
 '* Español: Tipos Privados. *'
 '****************************'
 Private Type GRADIENT_RECT
  UpperLeft   As Long
  LowerRight  As Long
 End Type
 
 Private Type RECT
  Left        As Long
  Top         As Long
  Right       As Long
  Bottom      As Long
 End Type
 
 Private Type RGB
  Red         As Integer
  Green       As Integer
  Blue        As Integer
 End Type
  
 Private Type POINTAPI
  X           As Long
  Y           As Long
 End Type

 '* English: Elements of the list.
 '* Español: Elementos de la lista.
 Private Type PropertyCombo
  Color         As OLE_COLOR   '* Color of Text.
  Enabled       As Boolean     '* Item Enabled or Disabled.
  Image         As StdPicture  '* Item image.
  Index         As Long        '* Index item.
  MouseIcon     As StdPicture  '* Set MouseIcon for each item.
  SeparatorLine As Boolean     '* Set SeparatorLine for each group that you consider necessary.
  Tag           As String      '* Extra Information only if is necessary.
  Text          As String      '* Text of the item.
  ToolTipText   As String      '* ToolTipText for item.
 End Type
  
 Private Type TRIVERTEX
  X             As Long
  Y             As Long
  Red           As Integer
  Green         As Integer
  Blue          As Integer
  Alpha         As Integer
 End Type
  
 '*********************************************'
 '* English: Public Enum of Control.          *'
 '* Español: Enumeración Publica del control. *'
 '*********************************************'
 
 '* English: Enum for the alignment of the text of the list.
 '* Español: Enum para la alineación del texto de la lista.
 Public Enum AlignTextCombo
  AlignLeft = 0
  AlignRight = 1
  AlignCenter = 2
 End Enum
 
 '* English: Appearance Combo.
 '* Español: Apariencias del Combo.
 Public Enum ComboAppearance
  [Office Xp] = 1       '* By HACKPRO TM.
  Win98 = 2             '* By fre_cpp.
  WinXp = 3             '* By fre_cpp & HACKPRO TM.
  [Office 2000] = 4     '* By fre_cpp & HACKPRO TM.
  [Soft Style] = 5      '* By fre_cpp.
  KDE = 6               '* By HACKPRO TM.
  Mac = 7               '* By fre_cpp & HACKPRO TM.
  JAVA = 8              '* By fre_cpp.
  [Explorer Bar] = 9    '* By HACKPRO TM.
  Picture = 10          '* By HACKPRO TM.
  [Special Borde] = 11  '* By HACKPRO TM.
  Circular = 12         '* By HACKPRO TM.
  [GradientV] = 13      '* By HACKPRO TM.
  [GradientH] = 14      '* By HACKPRO TM.
  [Light Blue] = 15     '* By HACKPRO TM.
  [Style Arrow] = 16    '* By HACKPRO TM.
  [NiaWBSS] = 17        '* By HACKPRO TM.
  [Rhombus] = 18        '* By HACKPRO TM.
  [Additional Xp] = 19  '* By HACKPRO TM.
  [Ardent] = 20         '* By HACKPRO TM.
 End Enum

 '* English: Type of Combo and behavior of the list.
 '* Español: Tipo de Combo y comportamiento de la lista.
 Public Enum ComboStyle
  [Dropdown Combo] = 0
  [Dropdown List] = 1
 End Enum
 
 '* English: Appearance standard style Xp.
 '* Español: Apariencias estándar estilo Xp.
 Public Enum ComboXpAppearance
  Aqua = 1              '* By HACKPRO TM.
  [Olive Green] = 2     '* By HACKPRO TM.
  Silver = 3            '* By HACKPRO TM.
  TasBlue = 4           '* By HACKPRO TM.
  Gold = 5              '* By HACKPRO TM.
  Blue = 6              '* By HACKPRO TM.
  CustomXP = 7          '* By HACKPRO TM.
 End Enum
  
 '* English: Enum for the type of text comparison.
 '* Español: Enum para el tipo de comparación de texto.
 Public Enum StringCompare
  None = 0
  ExactWord = 1
  CompleteWord = 2
 End Enum
  
 '********************************'
 '* English: Private variables.  *'
 '* Español: Variables privadas. *'
 '********************************'
 Private BigText                 As String
 Private ControlEnabled          As Boolean
 Private cValor                  As Long
 Private CurrentS                As Long
 Private First                   As Integer
 Private FirstView               As Integer
 Private g_Font                  As StdFont
 Private HighlightedItem         As Long
 Private iFor                    As Long
 Private IndexItemNow            As Long
 Private IsPicture               As Boolean
 Private ItemFocus               As Long
 Private KeyPos                  As Integer
 Private ListContents()          As PropertyCombo
 Private ListMaxL                As Long
 Private m_btnRect               As RECT
 Private m_bOver                 As Boolean
 Private myAlignCombo            As AlignTextCombo
 Private myAppearanceCombo       As ComboAppearance
 Private myArrowColor            As OLE_COLOR
 Private myAutoSel               As Boolean
 Private myBackColor             As OLE_COLOR
 Private myListColor             As OLE_COLOR
 Private myDisabledColor         As OLE_COLOR
 Private myDisabledPictureUser   As StdPicture
 Private myFocusPictureUser      As StdPicture
 Private myGradientColor1        As OLE_COLOR
 Private myGradientColor2        As OLE_COLOR
 Private myHighLightBorderColor  As OLE_COLOR
 Private myHighLightColorText    As OLE_COLOR
 Private myHighLightPictureUser  As StdPicture
 Private myItemsShow             As Long
 Private myNormalBorderColor     As OLE_COLOR
 Private myNormalColorText       As OLE_COLOR
 Private myNormalPictureUser     As StdPicture
 Private myMouseIcon             As StdPicture
 Private myMousePointer          As MousePointerConstants
 Private mySelectBorderColor     As OLE_COLOR
 Private mySelectListBorderColor As OLE_COLOR
 Private mySelectListColor       As OLE_COLOR
 Private myStyleCombo            As ComboStyle
 Private myText                  As String
 Private myXpAppearance          As ComboXpAppearance
 Private OrderListContents()     As PropertyCombo
 Private RGBColor                As RGB
 Private sumItem                 As Long
 Private tempBorderColor         As OLE_COLOR
 Private tmpC1                   As Long
 Private tmpC2                   As Long
 Private tmpC3                   As Long
 Private tmpColor                As Long
 
 '***************************************'
 '* English: Constant declares.         *'
 '* Español: declaración de Constantes. *'
 '***************************************'
 Private Const BDR_RAISEDINNER = &H4
 Private Const BDR_SUNKENOUTER = &H2
 Private Const BF_RECT = (&H1 Or &H2 Or &H4 Or &H8)
 Private Const COLOR_BTNFACE = 15
 Private Const COLOR_BTNSHADOW = 16
 Private Const COLOR_WINDOW = 5
 Private Const defAppearanceCombo = 1
 Private Const defArrowColor = &HC56A31
 Private Const defDisabledColor = &H808080
 Private Const defGradientColor1 = &HFF90FF
 Private Const defGradientColor2 = &HC0E0FF
 Private Const defHighLightBorderColor = &HC56A31
 Private Const defHighLightColorText = &HFFFFFF
 Private Const defNormalBorderColor = &HDEEDEF
 Private Const defNormalColorText = &HC56A31
 Private Const defListColor = &HFFFFFF
 Private Const defSelectBorderColor = &HC56A31
 Private Const defSelectListBorderColor = &H6B2408
 Private Const defSelectListColor = &HC56A31
 Private Const defStyleCombo = 0
 Private Const DSS_DISABLED = &H20
 Private Const DSS_NORMAL = &H0
 Private Const DST_BITMAP = &H3
 Private Const DST_COMPLEX = &H0
 Private Const DST_ICON = &H3
 Private Const DST_TEXT = &H2
 Private Const EDGE_RAISED = (&H1 Or &H4)
 Private Const EDGE_SUNKEN = (&H2 Or &H8)
 Private Const GRADIENT_FILL_RECT_H   As Long = &H0
 Private Const GRADIENT_FILL_RECT_V   As Long = &H1
 Private Const GWL_EXSTYLE = -20
 Private Const SWP_FRAMECHANGED = &H20
 Private Const SWP_NOMOVE = &H2
 Private Const SWP_NOSIZE = &H1
 Private Const WS_EX_TOOLWINDOW = &H80
 Private Const Version As String = "SComboBox 1.0.1 By HACKPRO TM"
 Private Const VK_LBUTTON = &H1
 Private Const VK_RBUTTON = &H2
 
 '******************************'
 '* English: Public Events.    *'
 '* Español: Eventos Públicos. *'
 '******************************'
 Public Event SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
   
 '**********************************'
 '* English: Calls to the API's.   *'
 '* Español: Llamadas a los API's. *'
 '**********************************'
 Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
 Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
 Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
 Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
 Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
 Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
 Private Declare Function DrawStateString Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lpString As String, ByVal cbStringLen As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
 Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
 Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
 Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
 Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
 Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
 Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
 Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
 Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
 Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
 Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
 Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
 Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
 Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
 Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
 Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
 
'***********************************************************'
'* English: Events of the controls and of the Usercontrol. *'
'* Español: Eventos de los controles y del Usercontrol.    *'
'***********************************************************'
Private Sub picList_Click()
 '* English: A Element has been selected or the control has been clicked
 '* Español: Establece el elemento donde se hizo clic.
On Error Resume Next
 If (ListContents(HighlightedItem + 1).Enabled = True) Then
  If (HighlightedItem + 1 >= ListCount1) Then HighlightedItem = HighlightedItem - 1
  ItemFocus = HighlightedItem + 1
  Text = ListContents(ItemFocus).Text
  Call DrawAppearance(myAppearanceCombo, 1)
  scrollI.Visible = False
  tmrFocus.Enabled = True
  Call ListIndex1
  RaiseEvent SelectionMade(ListContents(ListIndex1).Text, HighlightedItem + 1)
 End If
End Sub

Private Sub picList_KeyDown(KeyCode As Integer, Shift As Integer)
 Call UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub picList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '* English: The mouse has been moved over the list
 '* Español: Mueve el mouse por la lista.
 FirstView = 1
 HighlightedItem = Int(Y / 20)
 If (HighlightedItem + 1 + scrollI.Value > MaxListLength) Then Exit Sub
 IndexItemNow = HighlightedItem + 1
 If (ListContents(HighlightedItem + 1 + scrollI.Value).Enabled = True) Then
  HighlightedItem = HighlightedItem + scrollI.Value
  If (HighlightedItem + 1 > scrollI.Value + MaxListLength - 1) Then HighlightedItem = scrollI.Value + MaxListLength - 1
  If (HighlightedItem + 1 > ListCount1 - 1) Then HighlightedItem = ListCount1 - 1
  If (HighlightedItem + 1 < ListCount1) Then Call DrawList(scrollI.Value, NumberItemsToShow)
  picList.Refresh
 Else
  HighlightedItem = -1
 End If
End Sub

Private Sub scrollI_Change()
 FirstView = 1
 HighlightedItem = Abs(IndexItemNow - 1)
 tmrFocus.Enabled = False
 Call DrawList(scrollI.Value, NumberItemsToShow)
End Sub

Private Sub scrollI_Scroll()
 scrollI_Change
End Sub

Private Sub tmrFocus_Timer()
 If (InFocusControl(UserControl.hWnd) = True) And (picList.Visible = False) Then
  If (m_bOver = False) Then Call DrawAppearance(myAppearanceCombo, 2)
  m_bOver = True
 ElseIf (m_bOver = True) And (picList.Visible = False) Then
  Call DrawAppearance(myAppearanceCombo, 1)
  tmrFocus.Enabled = False
  m_bOver = False
 End If
End Sub

Private Sub txtCombo_Change()
 Dim sItem As Long, iLen As Long
 
 If (myAutoSel = False) Then
  sItem = FindItemText(txtCombo.Text, CompleteWord)
  If (sItem > 0) Then
   If (ListContents(sItem).Enabled = True) Then
    ItemFocus = sItem
    IndexItemNow = sItem
    If (IndexItemNow > NumberItemsToShow) Then
     iLen = (NumberItemsToShow + IndexItemNow) - IndexItemNow
    Else
     iLen = IndexItemNow - (NumberItemsToShow + IndexItemNow)
    End If
    If (iLen > scrollI.Max) Then
     scrollI.Value = scrollI.Max
    ElseIf (iLen < 0) Then
     scrollI.Value = 0
    Else
     scrollI.Value = scrollI.Max
    End If
    Call scrollI_Change
   End If
  Else
   ItemFocus = -1
  End If
 ElseIf (KeyPos <> 8) Then
  sItem = FindItemText(txtCombo.Text)
  If (sItem > 0) Then
   iLen = Len(txtCombo.Text)
   txtCombo.Text = txtCombo.Text & Mid$(ListContents(sItem).Text, iLen + 1, Len(ListContents(sItem).Text))
   txtCombo.SelStart = iLen
   txtCombo.SelLength = Len(txtCombo.Text)
   sItem = FindItemText(txtCombo.Text, CompleteWord)
   If (sItem > 0) Then
    If (ListContents(sItem).Enabled = True) Then
     ItemFocus = sItem
     IndexItemNow = sItem
    End If
   Else
    ItemFocus = -1
   End If
  Else
   ItemFocus = -1
  End If
 End If
End Sub

Private Sub txtCombo_GotFocus()
 txtCombo.SelStart = 0
 txtCombo.SelLength = Len(txtCombo.Text)
End Sub

Private Sub txtCombo_KeyDown(KeyCode As Integer, Shift As Integer)
 If (KeyCode = 115) Then Call UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub txtCombo_KeyPress(KeyAscii As Integer)
 KeyPos = KeyAscii
End Sub

Private Sub txtCombo_KeyUp(KeyCode As Integer, Shift As Integer)
 If (KeyCode = 115) Then Call UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub txtCombo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 tmrFocus.Enabled = True
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
 tmrFocus.Enabled = False
End Sub

Private Sub UserControl_ExitFocus()
 Call IsEnabled(ControlEnabled)
 tmrFocus.Enabled = True
End Sub

Private Sub UserControl_InitProperties()
 '* English: Setup properties values.
 '* Español: Establece propiedades iniciales.
 ControlEnabled = True
 ItemFocus = -1
 IsPicture = False
 ListIndex = -1
 ListMaxL = 10
 myAutoSel = False
 myAppearanceCombo = defAppearanceCombo
 myArrowColor = defArrowColor
 myBackColor = defListColor
 myDisabledColor = defDisabledColor
 myGradientColor1 = defGradientColor1
 myGradientColor2 = defGradientColor2
 myHighLightBorderColor = defHighLightBorderColor
 myHighLightColorText = defHighLightColorText
 myItemsShow = 7
 myListColor = defListColor
 myNormalBorderColor = defNormalBorderColor
 myNormalColorText = defNormalColorText
 mySelectBorderColor = defSelectBorderColor
 mySelectListBorderColor = defSelectListBorderColor
 mySelectListColor = defSelectListColor
 myStyleCombo = defStyleCombo
 myText = Ambient.DisplayName
 Text = myText
 myXpAppearance = 1
 Set g_Font = Ambient.Font
 sumItem = 0
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13 '* Enter.
   If (picList.Visible = True) Then Call UserControl_MouseDown(0, 0, 0, 0)
  Case 33 '* PageDown.
   If (IndexItemNow > NumberItemsToShow) Then
    IndexItemNow = IndexItemNow - NumberItemsToShow - 1
    If (IndexItemNow < 0) Then IndexItemNow = 1
    If (scrollI.Value - NumberItemsToShow - 1 > 0) Then scrollI.Value = scrollI.Value - NumberItemsToShow - 1 Else scrollI.Value = 0
   Else
    IndexItemNow = 1
    scrollI.Value = 0
   End If
   scrollI_Change
  Case 34 '* PageUp.
   If (IndexItemNow < sumItem) Then
    IndexItemNow = IndexItemNow + NumberItemsToShow - 1
    If (IndexItemNow > sumItem) Then IndexItemNow = sumItem
    If (scrollI.Value + NumberItemsToShow - 1 < scrollI.Max) Then scrollI.Value = scrollI.Value + NumberItemsToShow - 1 Else scrollI.Value = scrollI.Max
   Else
    IndexItemNow = sumItem
    scrollI.Value = scrollI.Max
   End If
   scrollI_Change
  Case 35 '* End.
   IndexItemNow = sumItem
   scrollI.Value = scrollI.Max
   scrollI_Change
  Case 36 '* Start.
   IndexItemNow = 1
   scrollI.Value = 0
   scrollI_Change
  Case 38 '* Up arrow.
   If (IndexItemNow > 0) Then
    IndexItemNow = IndexItemNow - 1
    If (scrollI.Value > 0) And (IndexItemNow - NumberItemsToShow < NumberItemsToShow) Then scrollI.Value = scrollI.Value - 1
    scrollI_Change
   End If
  Case 40 '* Down arrow.
   If (IndexItemNow < sumItem) Then
    IndexItemNow = IndexItemNow + 1
    If (scrollI.Value < scrollI.Max) And (IndexItemNow > NumberItemsToShow) Then scrollI.Value = scrollI.Value + 1
    scrollI_Change
   End If
  Case 115 '* Key F4.
   Call UserControl_MouseDown(1, 0, 0, 0)
 End Select
End Sub

Private Sub UserControl_LostFocus()
 Call UserControl_ExitFocus
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '* English: Show or hide the list.
 '* Español: Muestra la lista ó la oculta.
 If (Button = vbLeftButton) And (picList.Visible = False) Then
  Dim oRect As RECT
  
  First = 1
  HighlightedItem = -1
  IndexItemNow = ListIndex
  scrollI.Max = IIf(MaxListLength - NumberItemsToShow < 0, 0, MaxListLength - NumberItemsToShow)
  If (ListCount > NumberItemsToShow) And (ItemFocus > 1) And (ItemFocus < scrollI.Max) Then
   scrollI.Value = IIf(NumberItemsToShow < ItemFocus - 1, Abs(scrollI.Max - NumberItemsToShow), 1)
  ElseIf (ItemFocus > scrollI.Max) Then
   scrollI.Value = scrollI.Max
  Else
   scrollI.Value = 0
  End If
  FirstView = 0
  tmrFocus.Enabled = False
  Call GetWindowRect(hWnd, oRect)
  Call picList.Move(oRect.Left * Screen.TwipsPerPixelX, oRect.Bottom * Screen.TwipsPerPixelY + 21)
  Call SetWindowPos(UserControl.picList.hWnd, -1, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED)
  Call DrawAppearance(myAppearanceCombo, 3)
  If (myAppearanceCombo = 2) Or (myAppearanceCombo = 3) And (myXpAppearance <> 7) Or (myAppearanceCombo = 8) Then
   Call Espera(0.09)
   Call DrawAppearance(myAppearanceCombo, 1)
  End If
  scrollI.Top = 1
  scrollI.Left = ScaleWidth - 16
  If (ListCount > NumberItemsToShow) Then
   picList.Height = NumberItemsToShow * 300
  ElseIf (ListCount > 0) Then
   picList.Height = ListCount * 300
  Else
   picList.Height = 240
  End If
  If (NumberItemsToShow < MaxListLength) Then
   scrollI.Height = picList.ScaleHeight - 1
   scrollI.Visible = True
  Else
   scrollI.Visible = False
  End If
  Call DrawList(scrollI.Value, NumberItemsToShow)
  picList.Visible = True
 Else
  Call DrawAppearance(myAppearanceCombo, 2)
  picList.Visible = False
  First = 0
 End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 tmrFocus.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 tmrFocus.Enabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 Alignment = PropBag.ReadProperty("Alignment", 0)
 AppearanceCombo = PropBag.ReadProperty("AppearanceCombo", defAppearanceCombo)
 ArrowColor = PropBag.ReadProperty("ArrowColor", defArrowColor)
 AutoCompleteWord = PropBag.ReadProperty("AutoCompleteWord", False)
 BackColor = PropBag.ReadProperty("BackColor", defListColor)
 DisabledColor = PropBag.ReadProperty("DisabledColor", defDisabledColor)
 Set DisabledPictureUser = PropBag.ReadProperty("DisabledPictureUser", Nothing)
 Enabled = PropBag.ReadProperty("Enabled", True)
 GradientColor1 = PropBag.ReadProperty("GradientColor1", defGradientColor1)
 GradientColor2 = PropBag.ReadProperty("GradientColor2", defGradientColor2)
 Set FocusPictureUser = PropBag.ReadProperty("FocusPictureUser", Nothing)
 Set g_Font = PropBag.ReadProperty("Font", Ambient.Font)
 HighLightBorderColor = PropBag.ReadProperty("HighLightBorderColor", defHighLightBorderColor)
 HighLightColorText = PropBag.ReadProperty("HighLightColorText", defHighLightColorText)
 Set HighLightPictureUser = PropBag.ReadProperty("HighLightPictureUser", Nothing)
 ListColor = PropBag.ReadProperty("ListColor", defListColor)
 MaxListLength = PropBag.ReadProperty("MaxListLength", "10")
 Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
 MousePointer = PropBag.ReadProperty("MousePointer", 0)
 NormalBorderColor = PropBag.ReadProperty("NormalBorderColor", defNormalBorderColor)
 NormalColorText = PropBag.ReadProperty("NormalColorText", defNormalColorText)
 Set NormalPictureUser = PropBag.ReadProperty("NormalPictureUser", Nothing)
 NumberItemsToShow = PropBag.ReadProperty("NumberItemsToShow", "7")
 SelectBorderColor = PropBag.ReadProperty("SelectBorderColor", defSelectBorderColor)
 SelectListBorderColor = PropBag.ReadProperty("SelectListBorderColor", defSelectListBorderColor)
 SelectListColor = PropBag.ReadProperty("SelectListColor", defSelectListColor)
 Style = PropBag.ReadProperty("Style", defStyleCombo)
 Text = PropBag.ReadProperty("Text", Ambient.DisplayName)
 XpAppearance = PropBag.ReadProperty("XpAppearance", 1)
End Sub

Private Sub UserControl_Resize()
 Call IsEnabled(ControlEnabled)
 Call IsEnabled(ControlEnabled)
End Sub

Private Sub UserControl_Show()
 Dim lResult As Long
 
 lResult = GetWindowLong(picList.hWnd, GWL_EXSTYLE)
 Call SetWindowLong(picList.hWnd, GWL_EXSTYLE, lResult Or WS_EX_TOOLWINDOW)
 Call SetWindowPos(picList.hWnd, picList.hWnd, 0, 0, 0, 0, 39)
 Call SetWindowLong(picList.hWnd, -8, Parent.hWnd)
 Call SetParent(picList.hWnd, 0)
 tmrFocus.Enabled = False
 If (IsPicture = False) Then txtCombo.Left = 8
End Sub

Private Sub UserControl_Terminate()
 Erase ListContents
 tmrFocus.Enabled = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("Alignment", myAlignCombo, 0)
 Call PropBag.WriteProperty("AppearanceCombo", myAppearanceCombo, defAppearanceCombo)
 Call PropBag.WriteProperty("ArrowColor", myArrowColor, defArrowColor)
 Call PropBag.WriteProperty("AutoCompleteWord", myAutoSel, False)
 Call PropBag.WriteProperty("BackColor", myBackColor, defListColor)
 Call PropBag.WriteProperty("DisabledColor", myDisabledColor, defDisabledColor)
 Call PropBag.WriteProperty("DisabledPictureUser", myDisabledPictureUser, Nothing)
 Call PropBag.WriteProperty("Enabled", Enabled, True)
 Call PropBag.WriteProperty("FocusPictureUser", myFocusPictureUser, Nothing)
 Call PropBag.WriteProperty("Font", g_Font, Ambient.Font)
 Call PropBag.WriteProperty("GradientColor1", myGradientColor1, defGradientColor1)
 Call PropBag.WriteProperty("GradientColor2", myGradientColor2, defGradientColor2)
 Call PropBag.WriteProperty("HighLightBorderColor", myHighLightBorderColor, defHighLightBorderColor)
 Call PropBag.WriteProperty("HighLightColorText", myHighLightColorText, defHighLightColorText)
 Call PropBag.WriteProperty("HighLightPictureUser", myHighLightPictureUser, Nothing)
 Call PropBag.WriteProperty("ListColor", myListColor, defListColor)
 Call PropBag.WriteProperty("MaxListLength", ListMaxL, "10")
 Call PropBag.WriteProperty("MouseIcon", myMouseIcon, Nothing)
 Call PropBag.WriteProperty("MousePointer", myMousePointer, 0)
 Call PropBag.WriteProperty("NormalBorderColor", myNormalBorderColor, defNormalBorderColor)
 Call PropBag.WriteProperty("NormalColorText", myNormalColorText, defNormalColorText)
 Call PropBag.WriteProperty("NormalPictureUser", myNormalPictureUser, Nothing)
 Call PropBag.WriteProperty("NumberItemsToShow", myItemsShow, "7")
 Call PropBag.WriteProperty("SelectBorderColor", mySelectBorderColor, defSelectBorderColor)
 Call PropBag.WriteProperty("SelectListBorderColor", mySelectListBorderColor, defSelectListBorderColor)
 Call PropBag.WriteProperty("SelectListColor", mySelectListColor, defSelectListColor)
 Call PropBag.WriteProperty("Style", myStyleCombo, defStyleCombo)
 Call PropBag.WriteProperty("Text", myText, Ambient.DisplayName)
 Call PropBag.WriteProperty("XpAppearance", myXpAppearance, 1)
End Sub

'*******************************************'
'* English: Properties of the Usercontrol. *'
'* Español: Propiedades del Usercontrol.   *'
'*******************************************'
Public Property Get Alignment() As AlignTextCombo
Attribute Alignment.VB_Description = "Devuelve o establece la Alineación del texto en la lista."
 '* English: Sets/Gets alignment of the text in the list.
 '* Español: Devuelve o establece la alineación del texto en la lista.
 Alignment = myAlignCombo
End Property

Public Property Let Alignment(ByVal New_Align As AlignTextCombo)
 myAlignCombo = New_Align
 Call PropertyChanged("Alignment")
 Refresh
End Property

Public Property Get AppearanceCombo() As ComboAppearance
Attribute AppearanceCombo.VB_Description = "Devuelve o establece la apariencia que tomara el combo."
 '* English: Sets/Gets the style of the Combo.
 '* Español: Devuelve o establece el estilo del Combo.
 AppearanceCombo = myAppearanceCombo
End Property

Public Property Let AppearanceCombo(ByVal New_Style As ComboAppearance)
 myAppearanceCombo = IIf(New_Style <= 0, 1, New_Style)
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("AppearanceCombo")
 Refresh
End Property

Public Property Get ArrowColor() As OLE_COLOR
Attribute ArrowColor.VB_Description = "Devuelve o establece el color de la flecha."
 '* English: Sets/Gets the color of the arrow.
 '* Español: Devuelve o establece el color de la flecha.
 ArrowColor = myArrowColor
End Property

Public Property Let ArrowColor(ByVal New_Color As OLE_COLOR)
 myArrowColor = ConvertSystemColor(New_Color)
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("ArrowColor")
 Refresh
End Property

Public Property Get AutoCompleteWord() As Boolean
Attribute AutoCompleteWord.VB_Description = "Devuelve o establece si se completa la palabra con un elemento similar de la lista."
 '* English: Sets/Gets complete the word with a similar element of the list.
 '* Español: Devuelve o establece si se completa la palabra con un elemento similar de la lista.
 AutoCompleteWord = myAutoSel
End Property
'* Note: When this property this active one and the list _
         is shown, it is not tried to locate the element _
         in the list to make quicker the search of the _
         text to complete.
'* Nota: Cuando esta propiedad este activa y la lista se _
         muestre, no se intentara ubicar el elemento en la _
         lista para hacer más rápido la búsqueda del texto _
         a completar.

Public Property Let AutoCompleteWord(ByVal New_Value As Boolean)
 myAutoSel = New_Value
 Call PropertyChanged("AutoCompleteWord")
 Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo del combo."
 '* English: Sets/Gets the color of the Usercontrol.
 '* Español: Devuelve o establece el color del Usercontrol.
 BackColor = myBackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
 myBackColor = ConvertSystemColor(New_Color)
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("BackColor")
 Refresh
End Property

Public Property Get DisabledColor() As OLE_COLOR
Attribute DisabledColor.VB_Description = "Devuelve o establece el color del texto deshabilitado."
 '* English: Sets/Gets the color of the disabled text.
 '* Español: Devuelve o establece el color del texto deshabilitado.
 DisabledColor = ShiftColorOXP(myDisabledColor, 94)
End Property

Public Property Let DisabledColor(ByVal New_Color As OLE_COLOR)
 myDisabledColor = ConvertSystemColor(New_Color)
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("DisabledColor")
 Refresh
End Property

Public Property Get DisabledPictureUser() As StdPicture
Attribute DisabledPictureUser.VB_Description = "Devuelve o establece la imagen cuando el control este deshabilitado."
 '* English: Sets/Gets an image like topic of the Combo when the Object is not enabled.
 '* Español: Devuelve o establece una imagen como tema del combo cuando el Objeto este inactivo.
 Set DisabledPictureUser = myDisabledPictureUser
End Property

Public Property Set DisabledPictureUser(ByVal New_Picture As StdPicture)
 Set myDisabledPictureUser = New_Picture
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("DisabledPictureUser")
 Refresh
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
 '* English: Sets/Gets the Enabled property of the control.
 '* Español: Devuelve o establece si el Usercontrol esta habilitado ó deshabilitado.
 Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
 UserControl.Enabled = New_Enabled
 ControlEnabled = New_Enabled
 Call IsEnabled(New_Enabled)
 Call IsEnabled(New_Enabled)
 Call PropertyChanged("Enabled")
End Property

Public Property Get FocusPictureUser() As StdPicture
Attribute FocusPictureUser.VB_Description = "Devuelve o establece la imagen cuando el control tome el enfoque."
 '* English: Sets/Gets the image like topic of the Combo when It has the focus.
 '* Español: Devuelve o establece una imagen como tema del combo cuando se tiene el enfoque.
 Set FocusPictureUser = myFocusPictureUser
End Property

Public Property Set FocusPictureUser(ByVal New_Picture As StdPicture)
 Set myFocusPictureUser = New_Picture
 Call PropertyChanged("FocusPictureUser")
 Refresh
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Devuelve un Objeto Font."
 '* English: Sets/Gets the Font of the control.
 '* Español: Devuelve o establece el tipo de fuente del texto.
 Set Font = g_Font
End Property

Public Property Set Font(ByVal New_Font As StdFont)
On Error Resume Next
 With g_Font
  .Name = New_Font.Name
  .Size = New_Font.Size
  .Bold = New_Font.Bold
  .Italic = New_Font.Italic
  .Underline = New_Font.Underline
  .Strikethrough = New_Font.Strikethrough
 End With
 txtCombo.Font = New_Font
 Call PropertyChanged("Font")
 Refresh
End Property

Public Property Get GradientColor1() As OLE_COLOR
Attribute GradientColor1.VB_Description = "Devuelve o establece un color degradado."
 '* English: Sets/Gets the color First gradient color.
 '* Español: Devuelve o establece el color Gradient 1.
 GradientColor1 = myGradientColor1
End Property

Public Property Let GradientColor1(ByVal New_Color As OLE_COLOR)
 myGradientColor1 = ConvertSystemColor(New_Color)
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("GradientColor1")
 Refresh
End Property

Public Property Get GradientColor2() As OLE_COLOR
Attribute GradientColor2.VB_Description = "Devuelve o establece un color degradado."
 '* English: Sets/Gets the Second gradient color.
 '* Español: Devuelve o establece el color Gradient 2.
 GradientColor2 = myGradientColor2
End Property

Public Property Let GradientColor2(ByVal New_Color As OLE_COLOR)
 myGradientColor2 = ConvertSystemColor(New_Color)
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("GradientColor2")
 Refresh
End Property

Public Property Get HighLightBorderColor() As OLE_COLOR
Attribute HighLightBorderColor.VB_Description = "Devuelve o establece el color del borde  cuando el mouse pasa sobre el combo."
 '* English: Sets/Gets the color of the border of the control when the the control is highlighted.
 '* Español: Devuelve o establece el color del borde del control cuando el pasa sobre él.
 HighLightBorderColor = myHighLightBorderColor
End Property

Public Property Let HighLightBorderColor(ByVal New_Color As OLE_COLOR)
 myHighLightBorderColor = ConvertSystemColor(New_Color)
 Call PropertyChanged("HighLightBorderColor")
 Refresh
End Property

Public Property Get HighLightColorText() As OLE_COLOR
Attribute HighLightColorText.VB_Description = "Devuelve o establece el color del texto cuando el mouse se situe sobre él."
 '* English: Sets/Gets the color of the selection of the text.
 '* Español: Devuelve o establece el color de selección del texto.
 HighLightColorText = myHighLightColorText
End Property

Public Property Let HighLightColorText(ByVal New_Color As OLE_COLOR)
 myHighLightColorText = ConvertSystemColor(New_Color)
 Call PropertyChanged("HighLightColorText")
 Refresh
End Property

Public Property Get HighLightPictureUser() As StdPicture
Attribute HighLightPictureUser.VB_Description = "Devuelve o establece la imagen cuando el mouse pase sobre él objeto."
 '* English: Sets/Gets an image like topic of the Combo when the mouse is over the control.
 '* Español: Devuelve o establece una imagen como tema del combo cuando el mouse pasa por el Objeto.
 Set HighLightPictureUser = myHighLightPictureUser
End Property

Public Property Set HighLightPictureUser(ByVal New_Picture As StdPicture)
 Set myHighLightPictureUser = New_Picture
 Call PropertyChanged("HighLightPictureUser")
 Refresh
End Property

Public Property Get ItemTag(ByVal ListIndex As Long) As String
 '* English: Returns the tag of a specified item.
 '* Español: Selecciona el tag de Item.
 ItemTag = ""
On Error GoTo myErr:
 ItemTag = ListContents(ListIndex).Tag
 Exit Property
myErr:
 ItemTag = ""
End Property

Public Property Get ListColor() As OLE_COLOR
Attribute ListColor.VB_Description = "Devuelve o establece el color de la lista."
 '* English: Sets/Gets the color of the List.
 '* Español: Devuelve o establece el color de la lista.
 ListColor = myListColor
End Property

Public Property Let ListColor(ByVal New_Color As OLE_COLOR)
 myListColor = ConvertSystemColor(New_Color)
 picList.BackColor = myListColor
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("ListColor")
 Refresh
End Property

Public Property Get ListCount() As Long
 '* English: Returns the number of elements in the list.
 '* Español: Devuelve o establece el número de elementos de la lista.
 ListCount = ListCount1 - 1
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_MemberFlags = "400"
 '* English: Sets/Gets the selected item.
 '* Español: Devuelve o establece el item actual seleccionado.
 ListIndex = ListIndex1
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Long)
 Call ListIndex1(New_ListIndex)
End Property

Public Property Get MaxListLength() As Long
Attribute MaxListLength.VB_MemberFlags = "400"
 '* English: Sets/Gets the maximum size of the list.
 '* Español: Devuelve o establece el tamaño máximo de la lista.
 MaxListLength = IIf(ListMaxL < 0, ListCount, ListMaxL)
End Property

Public Property Let MaxListLength(ByVal ListMax As Long)
 If (ListMax > 0) And (ListMax < ListCount1) Then
  ListMaxL = ListMax
 Else
  ListMaxL = ListCount
 End If
 Call PropertyChanged("MaxListLength")
 Refresh
End Property

Public Property Get MouseIcon() As StdPicture
Attribute MouseIcon.VB_Description = "Establece un icono personalizado para el mouse."
 '* English: Sets a custom mouse icon.
 '* Español: Establece un icono escogido por el usuario.
 Set MouseIcon = myMouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
 Set myMouseIcon = New_MouseIcon
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Devuelve o establece el tipo de puntero del mouse mostrado al pasar por encima de un Objeto."
 '* English: Sets/Gets the type of mouse pointer displayed when over part of an object.
 '* Español: Devuelve o establece el tipo de puntero a mostrar cuando el mouse pase sobre el objeto.
 MousePointer = myMousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
 myMousePointer = New_MousePointer
End Property

Public Property Get NewIndex() As Long
 '* English: Sets/Gets the last Item added.
 '* Español: Devuelve o establece el último item agregado.
 If (sumItem <= 0) Then NewIndex = -1 Else NewIndex = sumItem
End Property

Public Property Get NormalBorderColor() As OLE_COLOR
Attribute NormalBorderColor.VB_Description = "Devuelve o establece el color normal del borde del combo."
 '* English: Sets/Gets the normal border color of the control.
 '* Español: Devuelve o establece el color normal del borde del control.
 NormalBorderColor = myNormalBorderColor
End Property

Public Property Let NormalBorderColor(ByVal New_Color As OLE_COLOR)
 myNormalBorderColor = ConvertSystemColor(New_Color)
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("NormalBorderColor")
 Refresh
End Property

Public Property Let NormalColorText(ByVal New_Color As OLE_COLOR)
Attribute NormalColorText.VB_Description = "Devuelve o establece el color normal del texto."
 myNormalColorText = ConvertSystemColor(New_Color)
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("NormalColorText")
 Refresh
End Property

Public Property Get NormalColorText() As OLE_COLOR
 '* English: Sets/Gets the normal text color in the control.
 '* Español: Devuelve o establece el color del texto normal.
 NormalColorText = myNormalColorText
End Property

Public Property Get NormalPictureUser() As StdPicture
Attribute NormalPictureUser.VB_Description = "Devuelve o establece la imagen normal del combo."
 '* English: Sets/Gets an image like topic of the Combo in normal state.
 '* Español: Devuelve o establece una imagen como tema del combo en estado normal.
 Set NormalPictureUser = myNormalPictureUser
End Property

Public Property Set NormalPictureUser(ByVal New_Picture As StdPicture)
 Set myNormalPictureUser = New_Picture
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("NormalPictureUser")
 Refresh
End Property

Public Property Get NumberItemsToShow() As Long
Attribute NumberItemsToShow.VB_MemberFlags = "400"
 '* English: Sets/Gets the number of items to show per time.
 '* Español: Devuelve o establece el número de items a mostrar por vez.
 If (myItemsShow < 0) Then myItemsShow = IIf(MaxListLength > 8, 7, MaxListLength)
 NumberItemsToShow = myItemsShow
End Property

Public Property Let NumberItemsToShow(ByVal ItemsShow As Long)
 If (ItemsShow <= 1) Or (ItemsShow >= MaxListLength) Then
  myItemsShow = IIf(MaxListLength > 8, MaxListLength - 8, ListCount)
 Else
  myItemsShow = ItemsShow
 End If
 Call PropertyChanged("NumberItemsToShow")
 Refresh
End Property

Public Property Get SelectBorderColor() As OLE_COLOR
Attribute SelectBorderColor.VB_Description = "Devuelve o establece el color del borde cuando se seleccione el combo."
 '* English: Sets/Gets the color of the border of the control when It has the focus.
 '* Español: Devuelve o establece el color del borde del control cuando el tenga el enfoque.
 SelectBorderColor = mySelectBorderColor
End Property

Public Property Let SelectBorderColor(ByVal New_Color As OLE_COLOR)
 mySelectBorderColor = ConvertSystemColor(New_Color)
 Call PropertyChanged("SelectBorderColor")
 Refresh
End Property

Public Property Get SelectListBorderColor() As OLE_COLOR
Attribute SelectListBorderColor.VB_Description = "Devuelve o establece el color del borde de la lista de un elemento seleccionado."
 '* English: Sets/Gets the border color of the item selected in the list.
 '* Español: Devuelve o establece el color del borde del item seleccionado en la lista.
 SelectListBorderColor = mySelectListBorderColor
End Property

Public Property Let SelectListBorderColor(ByVal New_Color As OLE_COLOR)
 mySelectListBorderColor = ConvertSystemColor(New_Color)
 Call PropertyChanged("SelectListBorderColor")
 Refresh
End Property

Public Property Get SelectListColor() As OLE_COLOR
Attribute SelectListColor.VB_Description = "Devuelve o establece el color de selección de un item de la lista."
 '* English: Sets/Gets the color of the item selected in the list.
 '* Español: Devuelve o establece el color del item seleccionado en la lista.
 SelectListColor = mySelectListColor
End Property

Public Property Let SelectListColor(ByVal New_Color As OLE_COLOR)
 mySelectListColor = ConvertSystemColor(New_Color)
 Call PropertyChanged("SelectListColor")
 Refresh
End Property

Public Property Get Style() As ComboStyle
Attribute Style.VB_Description = "Devuelve o establece un valor que determina el tipo de control y el comportamiento de su parte de cuadro de lista."
 '* English: Sets/Gets the style of the Combo.
 '* Español: Devuelve o establece el estilo del Combo.
 Style = myStyleCombo
End Property

Public Property Let Style(ByVal New_Style As ComboStyle)
 myStyleCombo = New_Style
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("Style")
 Refresh
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Devuelve o establece el texto contenido en el control."
 '* English: Sets/Gets the text of the selected item.
 '* Español: Devuelve o establece el texto del item seleccionado.
 Text = myText
End Property

Public Property Let Text(ByVal NewText As String)
 myText = NewText
 txtCombo.Text = myText
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("Text")
End Property

Public Property Get XpAppearance() As ComboXpAppearance
Attribute XpAppearance.VB_Description = "Devuelve o establece el tipo de apariencia de WinXp."
 '* English: Sets the appearance in Xp Mode.
 '* Español: Establece la apariencia en modo Xp.
 XpAppearance = myXpAppearance
End Property

Public Property Let XpAppearance(ByVal New_Style As ComboXpAppearance)
 myXpAppearance = New_Style
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("XpAppearance")
End Property

'********************************************************'
'* English: Subs and Functions of the Usercontrol.      *'
'* Español: Procedimientos y Funciones del Usercontrol. *'
'********************************************************'
Public Sub AddItem(ByVal Item As String, Optional ByVal ColorTextItem As OLE_COLOR = &HC56A31, Optional ByVal ImageItem As StdPicture = Nothing, Optional ByVal EnabledItem As Boolean = True, Optional ByVal ToolTipTextItem As String = "", Optional ByVal IndexItem As Long = -1, Optional ByVal ItemTag As String = "", Optional ByVal MouseIcon As StdPicture = Nothing, Optional ByVal SeparatorLine As Boolean = False)
 '* English: Add a new item to the list.
 '* Español: Agrega un nuevo item a la lista.
 sumItem = sumItem + 1
 ReDim Preserve ListContents(sumItem)
 If (IndexItem > 0) And (IndexItem < sumItem) And (NoFindIndex(IndexItem) = False) Then
  ListContents(sumItem).Index = IndexItem
 Else
  ListContents(sumItem).Index = sumItem
 End If
 ListContents(sumItem).Color = IIf(EnabledItem = True, ColorTextItem, DisabledColor)
 If (Len(Item) > Len(BigText)) Then BigText = Item
 ListContents(sumItem).Text = Item
 ListContents(sumItem).Enabled = EnabledItem
 ListContents(sumItem).Index = IndexItem
 ListContents(sumItem).ToolTipText = ToolTipTextItem
 ListContents(sumItem).Tag = ItemTag
 Set ListContents(sumItem).MouseIcon = MouseIcon
 ListContents(sumItem).SeparatorLine = SeparatorLine
 Set ListContents(sumItem).Image = ImageItem
 If Not (ImageItem Is Nothing) Then IsPicture = True
 MaxListLength = sumItem
End Sub

Private Sub APIFillRect(ByVal hDC As Long, ByRef RC As RECT, ByVal Color As Long)
 Dim NewBrush As Long
 
 '* English: The FillRect function fills a rectangle by using the specified brush. _
             This function includes the left and top borders, but excludes the right _
             and bottom borders of the rectangle.
 '* Español: Pinta el rectángulo de un objeto.
 NewBrush& = CreateSolidBrush(Color&)
 Call FillRect(hDC&, RC, NewBrush&)
 Call DeleteObject(NewBrush&)
End Sub

Private Sub APILine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal lColor As Long)
 Dim PT As POINTAPI, hPen As Long, hPenOld As Long
 
 '* English: Use the API LineTo for Fast Drawing.
 '* Español: Pinta líneas de forma sencilla y rápida.
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(UserControl.hDC, hPen)
 Call MoveToEx(UserControl.hDC, X1, Y1, PT)
 Call LineTo(UserControl.hDC, X2, Y2)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

Private Function APIRectangle(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Optional ByVal lColor As OLE_COLOR = -1) As Long
 Dim hPen As Long, hPenOld As Long
 Dim PT   As POINTAPI
 
 '* English: Paint a rectangle using API.
 '* Español: Pinta el rectángulo de un Objeto.
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(hDC, hPen)
 Call MoveToEx(hDC, X, Y, PT)
 Call LineTo(hDC, X + W, Y)
 Call LineTo(hDC, X + W, Y + H)
 Call LineTo(hDC, X, Y + H)
 Call LineTo(hDC, X, Y)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
End Function

Private Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long) As Long
 '* English: Blend two colors in a 50%.
 '* Español: Mezclar dos colores al 50%.
 BlendColors = RGB(((lColor1 And &HFF) + (lColor2 And &HFF)) / 2, (((lColor1 \ &H100) And &HFF) + ((lColor2 \ &H100) And &HFF)) / 2, (((lColor1 \ &H10000) And &HFF) + ((lColor2 \ &H10000) And &HFF)) / 2)
End Function
        
Public Sub ChangeItem(ByVal Index As Long, ByVal Item As String, Optional ByVal ColorTextItem As OLE_COLOR = &HC56A31, Optional ByVal ImageItem As StdPicture = Nothing, Optional ByVal EnabledItem As Boolean = True, Optional ByVal ToolTipTextItem As String = "", Optional ByVal IndexItem As Long = -1, Optional ByVal ItemTag As String = "", Optional ByVal MouseIcon As StdPicture = Nothing, Optional ByVal SeparatorLine As Boolean = False)
 '* English: Modifies an item of the list.
 '* Español: Modifica un item de la lista.
 ListContents(Index).Color = IIf(EnabledItem = True, ColorTextItem, ShiftColorOXP(DisabledColor))
 ListContents(Index).Text = Item
 ListContents(Index).Enabled = EnabledItem
 If (IndexItem > 0) And (IndexItem < sumItem) And (NoFindIndex(IndexItem) = False) Then ListContents(Index).Index = IndexItem
 Set ListContents(Index).MouseIcon = MouseIcon
 ListContents(Index).SeparatorLine = SeparatorLine
 ListContents(Index).ToolTipText = ToolTipTextItem
 ListContents(Index).Tag = ItemTag
 Set ListContents(Index).Image = ImageItem
 If Not (ImageItem Is Nothing) Then IsPicture = True
End Sub
        
Public Sub Clear()
 '* English: Clear the list.
 '* Español: Borra toda la lista.
 sumItem = 0
 ReDim ListContents(0)
 Refresh
 IsPicture = False
End Sub
        
Private Function ConvertSystemColor(ByVal theColor As Long) As Long
 '* English: Convert Long to System Color.
 '* Español: Convierte un long en un color del sistema.
 Call OleTranslateColor(theColor, 0, ConvertSystemColor)
End Function
        
Private Sub CreateImage(ByVal myPicture As StdPicture, ByRef WhatObject As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal Disabled As Boolean = False, Optional ByVal nHeight As Long = 16, Optional ByVal nWidth As Long = 16, Optional ByVal nColor As OLE_COLOR = "&HFFFFFF")
 Dim sTMPpathFName As String
 
 '* English: Draw the image in the Object.
 '* Español: Crea la imagen sobre el Objeto.
 Set picTemp.Picture = myPicture
 picTemp.BackColor = nColor
 If (Disabled = True) Then
  sTMPpathFName = App.Path + "\~ConvIcon2Bmp.tmp"
  SavePicture picTemp.Image, sTMPpathFName
  Set picTemp.Picture = LoadPicture(sTMPpathFName)
  Call Kill(sTMPpathFName)
  picTemp.Refresh
 Else
  Call PicDisabled(picTemp)
  sTMPpathFName = App.Path + "\~ConvIcon2Bmp.tmp"
  SavePicture picTemp.Image, sTMPpathFName
  Set picTemp.Picture = LoadPicture(sTMPpathFName)
  Call Kill(sTMPpathFName)
  picTemp.Refresh
 End If
 If (WhatObject = 1) Then
  Call StretchBlt(picList.hDC, X, Y, nWidth, nHeight, picTemp.hDC, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, vbSrcCopy)
 Else
  Call StretchBlt(UserControl.hDC, X, Y, nWidth, nHeight, picTemp.hDC, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, vbSrcCopy)
 End If
 If (Disabled = False) Then
  If (WhatObject = 1) Then
   Call StretchBlt(picList.hDC, X, Y, nWidth, nHeight, picTemp.hDC, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, vbSrcCopy)
  Else
   Call StretchBlt(UserControl.hDC, X, Y, nWidth, nHeight, picTemp.hDC, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, vbSrcCopy)
  End If
 End If
End Sub

Private Function CreateMacOSXRegion() As Long
 Dim pPoligon(8) As POINTAPI, lW As Long, lh As Long
 
 '* English: Create a nonrectangular region for the MAC OS X Style.
 '* Español: Crea el Estilo MAC OS X.
 lW = UserControl.ScaleWidth
 lh = UserControl.ScaleHeight
 pPoligon(0).X = 0:      pPoligon(0).Y = 2
 pPoligon(1).X = 2:      pPoligon(1).Y = 0
 pPoligon(2).X = lW - 2: pPoligon(2).Y = 0
 pPoligon(3).X = lW:     pPoligon(3).Y = 2
 pPoligon(4).X = lW:     pPoligon(4).Y = lh - 5
 pPoligon(5).X = lW - 6: pPoligon(5).Y = lh
 pPoligon(6).X = 3:      pPoligon(6).Y = lh
 pPoligon(7).X = 0:      pPoligon(7).Y = lh - 3
 CreateMacOSXRegion = CreatePolygonRgn(pPoligon(0), 8, 1)
End Function

Private Function CreatePicture(ByVal myIndex As Long, ByVal CurrentS As Long, Optional ByVal nColor As OLE_COLOR) As Boolean
 Dim xS As Long
 '* English: Set the picture of the list.
 '* Español: Crea la imagen sobre la lista.
 If (sumItem = 0) Or (myIndex > ListCount) Then Exit Function
On Error GoTo myErr
 CreatePicture = False
 If Not (ListContents(myIndex).Image Is Nothing) Then
  Call CreateImage(ListContents(myIndex).Image, 1, 4, CurrentS + xS + 3, ListContents(myIndex).Enabled, , , nColor)
  CreatePicture = True
 End If
 Exit Function
myErr:
 Debug.Print Err.Description
End Function

Private Sub CreateText(ByVal Counter As Long)
 Dim Msg As String
 
 '* English: Set the text of the list.
 '* Español: Crea el texto sobre el objeto.
On Error Resume Next
 With picList
  If (myAlignCombo = 0) Then
   '* English: Alignment to the left.
   '* Español: Alineación a la izquierda.
   If (IsPicture = True) Then
    Msg = Space$(8) & ListContents(Counter + 1).Text
   Else
    Msg = Space$(2) & ListContents(Counter + 1).Text
   End If
   .CurrentX = 11
  ElseIf (myAlignCombo = 1) Then
   '* Español: Alineación a la derecha.
   '* English: Alignment to the right.
   Msg = ListContents(Counter + 1).Text
   .CurrentX = Abs(ScaleWidth + .TextWidth(BigText) - 203)
  ElseIf (myAlignCombo = 2) Then
   '* English: Alignment to the Center.
   '* Español: Alineación en el centro.
   Msg = ListContents(Counter + 1).Text
   .CurrentX = ScaleWidth + (.TextWidth(BigText) / 2) + 32  '* Español: Calcula la mitad del ancho.
  End If
  picList.Print Msg
 End With
End Sub

Private Sub DrawAppearance(Optional ByVal Style As ComboAppearance = 1, Optional ByVal m_State As Integer = 1)
 Dim isText As String, m_lRegion As Long
 
 '* English: Draw appearance of the control.
 '* Español: Dibuja la apariencia del control.
 Cls
 AutoRedraw = True
 FillStyle = 1
 If (AppearanceCombo <> 7) Then UserControl.BackColor = myBackColor
 With txtCombo
  .Height = Abs(ScaleHeight / 2 - 7)
  .Top = Abs(ScaleHeight / 2 - 7)
  .BackColor = myBackColor
  .ForeColor = IIf(Enabled = True, myNormalColorText, ShiftColorOXP(DisabledColor))
 On Error Resume Next
  .Font = txtCombo.Font
  If (myStyleCombo = 1) Then
   .Visible = False
  Else
   .Visible = True
  End If
 End With
 If (Height < 300) And (Style <> 12) Then
  Height = 300
 ElseIf (Height < 310) And (Style = 12) Then
  Height = 310
 End If
 If (Width < 570) Then Width = 570
 If (m_State <> 3) Then picList.Visible = False
 Select Case Style
  Case 1
   Call DrawOfficeButton(m_State, 1)
  Case 2
   '* English: Style Windows 98.
   '* Español: Estilo Windows 98.
   Call DrawCtlEdge(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, EDGE_SUNKEN)
   Call APIFillRect(UserControl.hDC, m_btnRect, GetSysColor(COLOR_BTNFACE))
   tempBorderColor = GetSysColor(COLOR_BTNSHADOW)
   Call DrawCtlEdgeByRect(UserControl.hDC, m_btnRect, IIf(m_State = 3, EDGE_SUNKEN, EDGE_RAISED))
   Call DrawStandardArrow(m_btnRect, IIf(m_State = -1, ShiftColorOXP(ArrowColor, 66), ArrowColor))
  Case 3
   '* English: Style Windows Xp.
   '* Español: Estilo Windows Xp.
   If (myXpAppearance = 1) Then     '* Aqua.
    tmpColor = &HB99D7F
   ElseIf (myXpAppearance = 2) Then '* Olive Green.
    tmpColor = &H94CCBC
   ElseIf (myXpAppearance = 3) Then '* Silver.
    tmpColor = &HA29594
   ElseIf (myXpAppearance = 4) Then '* TasBlue.
    tmpColor = &HF09F5F
   ElseIf (myXpAppearance = 5) Then '* Gold.
    tmpColor = &HBFE7F0
   ElseIf (myXpAppearance = 6) Then '* Blue.
    tmpColor = ShiftColorOXP(&HA0672F, 123)
   ElseIf (myXpAppearance = 7) Then '* Custom.
    If (m_State = 1) Then
     tmpColor = NormalBorderColor
    ElseIf (m_State = 2) Then
     tmpColor = HighLightBorderColor
    ElseIf (m_State = 3) Then
     tmpColor = SelectBorderColor
    End If
   End If
   Call APIRectangle(UserControl.hDC, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, IIf(m_State <> -1, tmpColor, &HDEE7E7))
   Call APIRectangle(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, GetSysColor(COLOR_WINDOW))
   Call DrawWinXPButton(m_State, myXpAppearance)
   Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, 2, UserControl.BackColor)
   Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 2, UserControl.BackColor)
   Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, m_btnRect.Bottom - 1, UserControl.BackColor)
   Call SetPixel(UserControl.hDC, m_btnRect.Right - 1, UserControl.ScaleHeight - 3, UserControl.BackColor)
  Case 4
   Call DrawOfficeButton(m_State, 2)
  Case 5
   '* English: Style Soft.
   '* Español: Estilo Suavizado.
   Call DrawCtlEdge(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BDR_SUNKENOUTER)
   Call APIFillRect(UserControl.hDC, m_btnRect, IIf(m_State = -1, ShiftColorOXP(NormalBorderColor, 228), ShiftColorOXP(NormalBorderColor, 155)))
   tempBorderColor = GetSysColor(COLOR_BTNFACE)
   Call APIRectangle(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, GetSysColor(COLOR_BTNFACE))
   Call APILine(m_btnRect.Left - 1, m_btnRect.Top, m_btnRect.Left - 1, m_btnRect.Bottom, GetSysColor(COLOR_BTNFACE))
   Call DrawCtlEdgeByRect(UserControl.hDC, m_btnRect, IIf(m_State = 3, BDR_SUNKENOUTER, BDR_RAISEDINNER))
   Call DrawStandardArrow(m_btnRect, IIf(m_State = -1, ShiftColorOXP(ArrowColor, 106), ArrowColor))
  Case 6
   Call DrawKDEButton(m_State)
  Case 7
   '* English: Style MAC.
   '* Español: Estilo MAC.
   Call DrawMacOSXCombo(m_State)
  Case 8
   '* English: Style JAVA.
   '* Español: Estilo JAVA.
   tmpColor = ShiftColorOXP(NormalBorderColor, 52)
   tempBorderColor = GetSysColor(COLOR_BTNSHADOW)
   Call DrawJavaBorder(0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_WINDOW))
   Call APIFillRect(UserControl.hDC, m_btnRect, IIf(m_State = 2, tmpColor, IIf(m_State <> -1, NormalBorderColor, ShiftColorOXP(NormalBorderColor, 192))))
   Call DrawJavaBorder(m_btnRect.Left, m_btnRect.Top, m_btnRect.Right - m_btnRect.Left - 1, m_btnRect.Bottom - m_btnRect.Top - 1, GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_WINDOW))
   Call DrawStandardArrow(m_btnRect, IIf(m_State = -1, ShiftColorOXP(ArrowColor, 166), ArrowColor))
  Case 9
   Call DrawExplorerBarButton(m_State)
  Case 10
   Dim tempPict As StdPicture
   
   '* English: Style User Picture.
   '* Español: Estilo Imagen de Usuario.
   Set tempPict = Nothing
   If (m_State = 1) Then
    Set tempPict = myNormalPictureUser
    tmpColor = NormalBorderColor
   ElseIf (m_State = 2) Then
    Set tempPict = myHighLightPictureUser
    tmpColor = HighLightBorderColor
   ElseIf (m_State = 3) Then
    Set tempPict = myFocusPictureUser
    tmpColor = SelectBorderColor
    tempBorderColor = tmpColor
   Else
    Set tempPict = myDisabledPictureUser
    tmpColor = ShiftColorOXP(NormalBorderColor, 43)
   End If
   Call DrawRectangleBorder(ScaleWidth - 19, 0, ScaleWidth, ScaleHeight, Parent.BackColor, False)
   Call DrawRectangleBorder(0, 0, ScaleWidth - 19, ScaleHeight, IIf(m_State <> -1, tmpColor, ShiftColorOXP(DisabledColor, 145)), True)
   Call CreateImage(tempPict, 2, ScaleWidth - 18, Abs(Int(ScaleHeight / 2) - 11), True, 18, 17, Parent.BackColor)
  Case 11
   '* English: Special Style.
   '* Español: Estilo especial con borde recortado.
   If (m_State = 1) Then
    tmpColor = ShiftColorOXP(&HDCC6B4, 75)
   ElseIf (m_State = 2) Then
    tmpColor = ShiftColorOXP(&HDCC6B4, 45)
   ElseIf (m_State = 3) Then
    tmpColor = ShiftColorOXP(&HDCC6B4, 15)
   Else
    tmpColor = ShiftColorOXP(&H0&, 237)
   End If
   Call DrawRectangleBorder(UserControl.ScaleWidth - 17, 1, 16, UserControl.ScaleHeight - 2, ShiftColorOXP(tmpColor, 25), False)
   If (m_State = 1) Then
    tmpColor = ShiftColorOXP(&HC56A31, 143)
   ElseIf (m_State = 2) Or (m_State = 3) Then
    tmpColor = ShiftColorOXP(&HC56A31, 113)
    tempBorderColor = tmpColor
   Else
    tmpColor = ShiftColorOXP(&H0&)
   End If
   Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
   Call DrawRectangleBorder(UserControl.ScaleWidth - 17, 0, 17, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5), True)
   tmpC2 = 12
   For tmpC1 = 2 To 5
    tmpC2 = tmpC2 + 1
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left - 1, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, BackColor)
   Next
   tmpC2 = 17
   tmpC3 = -2
   For tmpC1 = 5 To 2 Step -1
    tmpC2 = tmpC2 - 1
    tmpC3 = tmpC3 + 1
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, tmpColor)
   Next
   Call DrawStandardArrow(m_btnRect, IIf(m_State = -1, ShiftColorOXP(&H404040, 166), ArrowColor))
  Case 12
   '* English: Rounded Style.
   '* Español: Estilo Circular.
   tempBorderColor = ShiftColorOXP(&H9F3000, 45)
   If (m_State = 1) Then
    iFor = &HCF989F
    tmpColor = &HA07F7F
    cValor = &HFFFFFF
   ElseIf (m_State = 2) Or (m_State = 3) Then
    iFor = &H9F3000
    tmpColor = &HAF572F
    cValor = &HFFFFFF
   Else
    tmpColor = ShiftColorOXP(&H404040, 166)
    iFor = &HFFF8FF
    cValor = ShiftColorOXP(&H404040, 16)
   End If
   FillStyle = 0
   FillColor = iFor
   UserControl.Circle (m_btnRect.Left + 7, CInt(UserControl.ScaleHeight / 2)), 8, tmpColor
   UserControl.Circle (m_btnRect.Left + 7, CInt(UserControl.ScaleHeight / 2)), 7, &HFFFFFF
   Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
   m_btnRect.Left = m_btnRect.Left - 5
   m_btnRect.Top = CInt(UserControl.ScaleHeight / 2) - 11
   UserControl.Line (m_btnRect.Left + 9, m_btnRect.Top + 8)-(m_btnRect.Left + 13, m_btnRect.Top + 12), IIf(m_State = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 10, m_btnRect.Top + 8)-(m_btnRect.Left + 13, m_btnRect.Top + 11), IIf(m_State = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 15, m_btnRect.Top + 8)-(m_btnRect.Left + 11, m_btnRect.Top + 12), IIf(m_State = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 14, m_btnRect.Top + 8)-(m_btnRect.Left + 11, m_btnRect.Top + 11), IIf(m_State = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 9, m_btnRect.Top + 12)-(m_btnRect.Left + 13, m_btnRect.Top + 16), IIf(m_State = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 10, m_btnRect.Top + 12)-(m_btnRect.Left + 13, m_btnRect.Top + 15), IIf(m_State = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 15, m_btnRect.Top + 12)-(m_btnRect.Left + 11, m_btnRect.Top + 16), IIf(m_State = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 14, m_btnRect.Top + 12)-(m_btnRect.Left + 11, m_btnRect.Top + 15), IIf(m_State = -1, ShiftColorOXP(&H404040, 166), cValor)
  Case 13
   Call DrawGradientButton(m_State, 1)
  Case 14
   Call DrawGradientButton(m_State, 2)
  Case 15
   Call DrawLightBlueButton(m_State)
  Case 16
   '* English: Arrow Style.
   '* Español: Estilo Flecha.
   If (m_State = 1) Then
    cValor = GetLngColor(GradientColor1)
    iFor = GetLngColor(GradientColor2)
    tmpColor = NormalBorderColor
   ElseIf (m_State = 2) Then
    cValor = GetLngColor(ShiftColorOXP(GradientColor1, 65))
    iFor = GetLngColor(ShiftColorOXP(GradientColor2, 65))
    tmpColor = HighLightBorderColor
   ElseIf (m_State = 3) Then
    cValor = GetLngColor(GradientColor1)
    iFor = GetLngColor(GradientColor2)
    tmpColor = SelectBorderColor
   Else
    cValor = GetLngColor(ShiftColorOXP(GradientColor1))
    iFor = GetLngColor(GradientColor2)
    tmpColor = ShiftColorOXP(&H0&)
   End If
   tempBorderColor = tmpColor
   Call DrawGradient(UserControl.hDC, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1, iFor, cValor, 1)
   Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
   Call DrawRectangleBorder(UserControl.ScaleWidth - 17, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5))
   cValor = UserControl.ScaleHeight / 2 + 1
   iFor = ArrowColor
   For tmpColor = 7 To -2 Step -1
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 6, cValor - (tmpColor / 2), UserControl.ScaleWidth - 7, cValor - (tmpColor / 2), IIf(m_State = -1, ShiftColorOXP(iFor, 166), iFor))
   Next
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 5, cValor, UserControl.ScaleWidth - 6, cValor, IIf(m_State = -1, ShiftColorOXP(iFor, 166), iFor))
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 6, cValor + 1, UserControl.ScaleWidth - 7, cValor + 1, IIf(m_State = -1, ShiftColorOXP(iFor, 166), iFor))
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 7, cValor + 2, UserControl.ScaleWidth - 8, cValor + 2, IIf(m_State = -1, ShiftColorOXP(iFor, 166), iFor))
  Case 17
   Call DrawNiaWBSSButton(m_State)
  Case 18
   Call DrawRhombusButton(m_State)
  Case 19
   Call DrawXpButton(m_State)
  Case 20
   '* English: Ardent Style.
   '* Español: Estilo Ardent.
   If (m_State = 1) Then
    tmpC1 = ShiftColorOXP(BlendColors(GradientColor1, GradientColor2), 24)
    cValor = tmpC1
    iFor = tmpC1
    tmpColor = NormalBorderColor
   ElseIf (m_State = 2) Then
    tmpC1 = ShiftColorOXP(BlendColors(GradientColor1, GradientColor2), 65)
    cValor = tmpC1
    iFor = tmpC1
    tmpColor = HighLightBorderColor
   ElseIf (m_State = 3) Then
    tmpC1 = ShiftColorOXP(BlendColors(GradientColor1, GradientColor2), 14)
    cValor = tmpC1
    iFor = tmpC1
    tmpColor = SelectBorderColor
    tempBorderColor = tmpColor
   Else
    tmpC1 = ShiftColorOXP(BlendColors(GradientColor1, GradientColor2))
    cValor = tmpC1
    iFor = tmpC1
    tmpColor = DisabledColor
   End If
   Call DrawVGradient(cValor, iFor, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1)
   Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
   Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 0, 19, UserControl.ScaleHeight, tmpColor)
   Call DrawRectangleBorder(1, 1, UserControl.ScaleWidth - 19, UserControl.ScaleHeight - 2, ShiftColorOXP(cValor, 85))
   tmpC1 = 7
   tmpC2 = 4
   tmpC3 = ScaleHeight / 2 + 1
   For tmpColor = 6 To 2 Step -1
    tmpC1 = tmpC1 - 1
    tmpC2 = tmpC2 - 1
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC1, tmpC3 + tmpC2, UserControl.ScaleWidth - (tmpColor + 2), tmpC3 + tmpC2, IIf(m_State = -1, ShiftColorOXP(ArrowColor, 146), ArrowColor))
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC1, tmpC3 - 2 + tmpC2 - 1, UserControl.ScaleWidth - (tmpColor + 2), tmpC3 - 2 + tmpC2 - 1, IIf(m_State = -1, ShiftColorOXP(ArrowColor, 146), ArrowColor))
   Next
 End Select
 Call SetRect(m_btnRect, UserControl.ScaleWidth - 18, 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2)
 If (Style = 7) Then
  If (m_lRegion <> 0) Then Call DeleteObject(m_lRegion)
  m_lRegion = CreateMacOSXRegion
  Call SetWindowRgn(UserControl.hWnd, m_lRegion, True)
 Else
  m_lRegion = CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
  Call SetWindowRgn(UserControl.hWnd, m_lRegion, True)
 End If
 If (ItemFocus > 0) Then
  '* English: Sets the image of the current item.
  '* Español: Establece la imagen del item actual.
  picTemp.BackColor = ListColor
  Call CreateImage(ListContents(ItemFocus).Image, 2, 5, Abs(Int(ScaleHeight / 2) - 7) - IIf(Style = 7, 1, 0), Enabled, , , BackColor)
  If Not (ListContents(ItemFocus).Image Is Nothing) Then
   cValor = 27
  Else
   cValor = 8
  End If
  isText = ListContents(ItemFocus).Text
 Else
  isText = Text
  cValor = 8
 End If
 txtCombo.Left = cValor
 If (myStyleCombo = 1) Then
  With UserControl
   .CurrentX = cValor
   .CurrentY = Int(UserControl.ScaleHeight / 2) - 6
   .Font = txtCombo.Font
   If (Enabled = False) Then
    Call SetTextColor(.hDC, DisabledColor)
   Else
    Call SetTextColor(.hDC, NormalColorText)
   End If
   Call DrawStateString(.hDC, 0, 0, isText, Len(isText), .CurrentX, .CurrentY, 0, 0, DST_TEXT Or DSS_NORMAL)
  End With
 End If
 txtCombo.Width = IIf(txtCombo.Left = 27, Abs(ScaleWidth - 48), Abs(ScaleWidth - 28))
 picList.Width = Width
End Sub

Private Sub DrawCtlEdge(ByVal hDC As Long, ByVal X As Single, ByVal Y As Single, ByVal W As Single, ByVal H As Single, Optional ByVal Style As Long = EDGE_RAISED, Optional ByVal flags As Long = BF_RECT)
 Dim R As RECT
 
 '* English: The DrawEdge function draws one or more edges of rectangle. _
             using the specified coords.
 '* Español: Dibuja uno ó más bordes del rectángulo.
 With R
  .Left = X
  .Top = Y
  .Right = X + W
  .Bottom = Y + H
 End With
 Call DrawEdge(hDC, R, Style, flags)
End Sub

Private Sub DrawCtlEdgeByRect(ByVal hDC As Long, ByRef RT As RECT, Optional ByVal Style As Long = EDGE_RAISED, Optional ByVal flags As Long = BF_RECT)
 '* English: Draws the edge in a rect.
 '* Español: Colorea uno ó más bordes del rectángulo del Control.
 Call DrawEdge(hDC, RT, Style, flags)
End Sub

Private Sub DrawExplorerBarButton(ByVal m_State As Long)
 '* English: Style ExplorerBar.
 '* Español: Estilo ExplorerBar.
 myBackColor = ShiftColorOXP(&HDEEAF0, 184)
 txtCombo.BackColor = myBackColor
 UserControl.BackColor = myBackColor
 Call DrawRectangleBorder(1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, &HEAF3F7)
 If (m_State = 1) Then
  cValor = ShiftColorOXP(&HB6BFC3, 91)
  iFor = &HEAF3F7
  tmpColor = ShiftColorOXP(&HB6BFC3, 162)
 ElseIf (m_State = 2) Then
  cValor = ShiftColorOXP(&HB6BFC3, 31)
  iFor = &HDCEBF1
  tmpColor = ShiftColorOXP(&HB6BFC3, 132)
 ElseIf (m_State = 3) Then
  cValor = ShiftColorOXP(&HB6BFC3, 21)
  iFor = &HCEE3EC
  tmpColor = ShiftColorOXP(&HB6BFC3, 112)
  tempBorderColor = ShiftColorOXP(&HB6BFC3, 21)
 Else
  UserControl.BackColor = ShiftColorOXP(&HEAF3F7, 124)
  txtCombo.BackColor = UserControl.BackColor
  cValor = ShiftColorOXP(&HB6BFC3, 84)
  tmpC1 = ShiftColorOXP(&HEAF3F7, 124)
  iFor = ShiftColorOXP(&HEAF3F7, 123)
  tmpColor = ShiftColorOXP(&HB6BFC3, 132)
 End If
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, cValor)
 If (m_State = -1) Then Call DrawRectangleBorder(1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, tmpC1)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 2, 16, UserControl.ScaleHeight - 4, iFor, False)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 2, 16, UserControl.ScaleHeight - 4, tmpColor)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, 2, UserControl.BackColor)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 2, UserControl.BackColor)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, m_btnRect.Bottom - 1, UserControl.BackColor)
 Call SetPixel(UserControl.hDC, m_btnRect.Right - 1, UserControl.ScaleHeight - 3, UserControl.BackColor)
 Call DrawStandardArrow(m_btnRect, IIf(m_State = -1, ShiftColorOXP(&H404040, 196), ArrowColor))
End Sub

Private Sub DrawGradient(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Direction As Integer)
 Dim Vert(1) As TRIVERTEX, gRect As GRADIENT_RECT

 '* English: Draw a gradient in the selected coords and hDC.
 '* Español: Dibuja el objeto en forma degradada.
 Call LongToRGB(Color1)
 With Vert(0)
  .X = X
  .Y = Y
  .Red = Val("&H" & Hex$(RGBColor.Red) & "00")
  .Green = Val("&H" & Hex$(RGBColor.Green) & "00")
  .Blue = Val("&H" & Hex$(RGBColor.Blue) & "00")
  .Alpha = 1
 End With
 Call LongToRGB(Color2)
 With Vert(1)
  .X = X1
  .Y = Y1
  .Red = Val("&H" & Hex$(RGBColor.Red) & "00")
  .Green = Val("&H" & Hex$(RGBColor.Green) & "00")
  .Blue = Val("&H" & Hex$(RGBColor.Blue) & "00")
  .Alpha = 0
 End With
 gRect.UpperLeft = 0
 gRect.LowerRight = 1
 If (Direction = 1) Then
  Call GradientFillRect(hDC, Vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V)
 Else
  Call GradientFillRect(hDC, Vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_H)
 End If
End Sub

Private Sub DrawGradientButton(ByVal m_State As Long, ByVal WhatGradient As Long)
 '* English: Draw a Vertical or Horizontal Gradient style appearance.
 '* Español: Dibuja la apariencia degradada bien sea vertical ó horizontal.
 If (m_State = 1) Then
  tmpColor = ShiftColorOXP(&HC56A31, 133)
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HD8CEC5)
 ElseIf (m_State = 2) Then
  tmpColor = ShiftColorOXP(&HC56A31, 113)
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HD6BEB5)
 ElseIf (m_State = 3) Then
  tmpColor = ShiftColorOXP(&HC56A31, 93)
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HB3A29B)
  tempBorderColor = tmpColor
 Else
  tmpColor = CLng(ShiftColorOXP(&H0&))
  cValor = GetLngColor(&HC0C0C0)
  iFor = GetLngColor(&HFFFFFF)
 End If
 Call DrawGradient(UserControl.hDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom, cValor, iFor, WhatGradient)
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 2, 16, UserControl.ScaleHeight - 4, tmpColor, True)
 Call DrawStandardArrow(m_btnRect, IIf(m_State = -1, ShiftColorOXP(&H404040, 166), ArrowColor))
End Sub

Private Sub DrawJavaBorder(ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, ByVal lColorShadow As Long, ByVal lColorLight As Long, ByVal lColorBack As Long)
 '* English: Draw the edge with a JAVA style.
 '* Español: Dibuja el borde estilo JAVA.
 Call APIRectangle(UserControl.hDC, X, Y, W - 1, H - 1, lColorShadow)
 Call APIRectangle(UserControl.hDC, X + 1, Y + 1, W - 1, H - 1, lColorLight)
 Call SetPixel(UserControl.hDC, X, Y + H, lColorBack)
 Call SetPixel(UserControl.hDC, X + W, Y, lColorBack)
 Call SetPixel(UserControl.hDC, X + 1, Y + H - 1, BlendColors(lColorLight, lColorShadow))
 Call SetPixel(UserControl.hDC, X + W - 1, Y + 1, BlendColors(lColorLight, lColorShadow))
End Sub

Private Sub DrawKDEButton(ByVal m_State As Long)
 '* English: Style KDE.
 '* Español: Estilo KDE.
 If (m_State = 1) Then
  tmpColor = NormalBorderColor
  cValor = GetLngColor(ShiftColorOXP(GradientColor1, 63))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2, 63))
 ElseIf (m_State = 2) Then
  tmpColor = HighLightBorderColor
  cValor = GetLngColor(ShiftColorOXP(GradientColor1, 127))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2, 127))
 ElseIf (m_State = 3) Then
  tmpColor = SelectBorderColor
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
 Else
  tmpColor = &HC0C0C0
  cValor = GetLngColor(&HFFFFFF)
  iFor = ShiftColorOXP(GetLngColor(&HC0C0C0), 45)
 End If
 tempBorderColor = tmpColor
 '* Español: Top Left.
 '* Español: Parte Superior Izquierda.
 Call DrawGradient(UserControl.hDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right - 8, m_btnRect.Bottom - 8, iFor, cValor, 1)
 '* Español: Top Right.
 '* Español: Parte Inferior Derecha.
 Call DrawGradient(UserControl.hDC, m_btnRect.Left + 8, m_btnRect.Top + 8, m_btnRect.Right, m_btnRect.Bottom, cValor, iFor, 1)
 '* Español: Bottom Right.
 '* Español: Parte Inferior Derecha.
 Call DrawGradient(UserControl.hDC, m_btnRect.Left + 8, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom - 8, iFor, cValor, 1)
 '* Español: Bottom Left.
 '* Español: Parte Inferior Izquierda.
 Call DrawGradient(UserControl.hDC, m_btnRect.Left, m_btnRect.Top + 8, m_btnRect.Right - 8, m_btnRect.Bottom, cValor, iFor, 1)
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 2, 16, UserControl.ScaleHeight - 4, tmpColor, True)
 Call DrawStandardArrow(m_btnRect, IIf(m_State = -1, ShiftColorOXP(&H404040, 166), ArrowColor))
End Sub

Private Sub DrawLightBlueButton(ByVal m_State As Long)
 Dim PT      As POINTAPI, cx As Long, cy As Long
 Dim hPenOld As Long, hPen   As Long
   
 '* English: Style LightBlue.
 '* Español: Estilo LightBlue.
 If (m_State = 1) Or (m_State = 3) Then
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HA87057)
  tmpColor = &HA69182
  tempBorderColor = tmpColor
 ElseIf (m_State = 2) Then
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HCFA090)
  tmpColor = &HAF9080
 Else
  cValor = GetLngColor(&HFFFFFF)
  iFor = ShiftColorOXP(GetLngColor(&HA87057))
  tmpColor = ShiftColorOXP(&HA69182, 146)
 End If
 Call DrawGradient(UserControl.hDC, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1, cValor, iFor, 1)
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 If (m_State = 2) Then
  Call DrawRectangleBorder(UserControl.ScaleWidth - 17, 0, 17, UserControl.ScaleHeight, &H53969F)
  Call DrawRectangleBorder(UserControl.ScaleWidth - 16, 1, 15, UserControl.ScaleHeight - 2, &H92C4D8)
  tmpColor = &H3EB4DE
 ElseIf (m_State <> -1) Then
  Call DrawRectangleBorder(UserControl.ScaleWidth - 17, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5))
  tmpColor = ArrowColor
 End If
 cx = m_btnRect.Left + (m_btnRect.Right - m_btnRect.Left) / 2 + 2
 cy = m_btnRect.Top + (m_btnRect.Bottom - m_btnRect.Top) / 2 + 2
 hPen = CreatePen(0, 1, IIf(m_State <> -1, tmpColor, ShiftColorOXP(&HC0C0C0, 97)))
 hPenOld = SelectObject(UserControl.hDC, hPen)
 Call MoveToEx(UserControl.hDC, cx - 3, cy - 1, PT)
 Call LineTo(UserControl.hDC, cx + 1, cy - 1)
 Call LineTo(UserControl.hDC, cx, cy)
 Call LineTo(UserControl.hDC, cx - 2, cy)
 Call LineTo(UserControl.hDC, cx, cy + 2)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
 hPen = CreatePen(0, 1, IIf(m_State <> -1, tmpColor, ShiftColorOXP(&HC0C0C0, 97)))
 hPenOld = SelectObject(UserControl.hDC, hPen)
 cx = m_btnRect.Left + (m_btnRect.Right - m_btnRect.Left) / 2 + 3
 Call MoveToEx(UserControl.hDC, cx - 4, cy - 3, PT)
 Call LineTo(UserControl.hDC, cx, cy - 3)
 Call LineTo(UserControl.hDC, cx - 2, cy - 5)
 Call LineTo(UserControl.hDC, cx - 3, cy - 4)
 Call LineTo(UserControl.hDC, cx - 1, cy - 3)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

Private Sub DrawList(ByVal TopItem As Integer, ByVal NumberOfItems As Integer)
 Dim Counter As Long, d As Long, Valor As Long, isFocus As Long
 
 '* English: Draw the list with the elements.
 '* Español: Crea la lista con los elementos guardados.
 picList.Cls
 picList.Line (0, 0)-(picList.ScaleWidth - 1, picList.ScaleHeight - 1), tempBorderColor, B
 If (ListCount < 0) Then Exit Sub
 With picList
  .AutoRedraw = True
  .ScaleMode = vbTwips
  Counter = TopItem - 1
  CurrentS = -20
  isFocus = -1
  Do Until (Counter = TopItem + NumberOfItems)
   CurrentS = CurrentS + 20
   d = (Counter - TopItem) * 255 + CurrentS * 2 + 27
  On Error Resume Next
   Valor = IIf((HighlightedItem + 1) < ListCount, HighlightedItem + 1, ListCount)
   If ((ListContents(Counter + 1).Enabled = True) And (Len(txtCombo.Text) > 0) And (ItemFocus = Counter + 1) And (ListContents(Counter + 1).Text = txtCombo.Text) And (FirstView = 0)) Or ((Counter = HighlightedItem) And (ListContents(Valor).Enabled = True) And (FirstView = 1)) Then
    If (scrollI.Visible = True) Then
     picList.Line (3, d - 45)-(picList.Width - 260, d + 245), SelectListColor, BF
    ElseIf (Abs(Counter - scrollI.Value + 1) = NumberItemsToShow) And (ListCount > 1) Then
     picList.Line (3, d - 45)-(picList.Width, d + Abs(picList.Height - (d + 20))), SelectListColor, BF
    Else
     picList.Line (3, d - 45)-(picList.Width, d + 255), SelectListColor, BF
    End If
    If (Abs(Counter - scrollI.Value + 1) = 1) And (ListCount > 1) Then
     picList.Line (3, d - 65)-(picList.Width - 15, d + 255), SelectListBorderColor, B
    ElseIf (Abs(Counter - scrollI.Value + 1) = NumberItemsToShow) And (ListCount > 1) Then
     picList.Line (3, d - 45)-(picList.Width - 15, d + Abs(picList.Height - (d + 20))), SelectListBorderColor, B
    ElseIf (ListCount = Counter) Then
     picList.Line (3, d - 45)-(picList.Width, d - 55), SelectListBorderColor, B
    ElseIf (ListCount = 1) Then
     picList.Line (0, 0)-(picList.ScaleWidth - 8, picList.ScaleHeight - 8), SelectListBorderColor, B
    Else
     picList.Line (3, d - 45)-(picList.Width - 15, d + IIf(Abs(picList.Height - (d + 21)) > 255, 255, Abs(picList.Height - (d + 21)))), SelectListBorderColor, B
    End If
    picList.ForeColor = HighLightColorText
    picList.CurrentY = d + 20 - IIf(ListCount = 1, 9, 0)
    Call CreatePicture(Counter + 1, Abs(CurrentS - 20), SelectBorderColor)
    Call CreateText(Counter)
    picList.ToolTipText = ListContents(Counter + 1).ToolTipText
    picList.MousePointer = vbCustom
    Set picList.MouseIcon = ListContents(Counter + 1).MouseIcon
    isFocus = Counter
   Else
    picList.ForeColor = ListContents(Counter + 1).Color
    If (Counter < sumItem) And (sumItem > 1) And (ListContents(Counter).SeparatorLine = True) And (Counter <> isFocus + 1) Then picList.Line (-3, picList.CurrentY + 32)-(picList.ScaleWidth - 11, picList.CurrentY + 32), vbButtonShadow, B
   End If
   picList.CurrentY = d + 20 - IIf(ListCount = 1, 9, 0)
   Call CreatePicture(Counter + 2, CurrentS, ListColor)
   Call CreateText(Counter)
   Counter = Counter + 1
  Loop
  .ScaleMode = vbPixels
 End With
 FirstView = 1
End Sub

Private Sub DrawMacOSXCombo(ByVal Mode As Integer)
 Dim PT      As POINTAPI, cy  As Long, cx     As Long, Color1 As Long, ColorG As Long
 Dim hPen    As Long, hPenOld As Long, Color2 As Long, Color3 As Long, ColorH As Long
 Dim Color4  As Long, Color5  As Long, Color6 As Long, Color7 As Long, ColorI As Long
 Dim Color8  As Long, Color9  As Long, ColorA As Long, ColorB As Long
 Dim ColorC  As Long, ColorD  As Long, ColorE As Long, ColorF As Long
 
 '* English: Draw the Mac OS X combo (this is a cool style!)
 '* Español: Dibujar el combo estilo Mac OS X
 m_btnRect.Left = m_btnRect.Left - 4
 tempBorderColor = GetSysColor(COLOR_BTNSHADOW)
 '* English: Button gradient top.
 ColorA = &HA0A0A0
 If (Mode = 1) Then
  Color1 = ShiftColorOXP(&HFDF2C3, 9)
  Color2 = ShiftColorOXP(&HDE8B45, 9)
  Color3 = ShiftColorOXP(&HDD873E, 9)
  Color4 = ShiftColorOXP(&HB33A01, 9)
  Color5 = ShiftColorOXP(&HE9BD96, 9)
  Color6 = ShiftColorOXP(&HB9B2AD, 9)
  Color7 = ShiftColorOXP(&H968A82, 9)
  Color8 = ShiftColorOXP(&HA25022, 9)
  Color9 = ShiftColorOXP(&HB8865E, 9)
  ColorB = ShiftColorOXP(&HDFBC86, 9)
  ColorC = ShiftColorOXP(&HFFBA77, 9)
  ColorD = ShiftColorOXP(&HE3D499, 9)
  ColorE = ShiftColorOXP(&HFFD996, 9)
  ColorF = ShiftColorOXP(&HE1A46D, 9)
  ColorG = ShiftColorOXP(&HCBA47B, 9)
  ColorH = ShiftColorOXP(&HDFDFDF, 9)
  ColorI = ShiftColorOXP(&HD0D0D0, 9)
 ElseIf (Mode = 2) Then
  Color1 = ShiftColorOXP(&HFDF2C3, 89)
  Color2 = ShiftColorOXP(&HDE8B45, 89)
  Color3 = ShiftColorOXP(&HDD873E, 89)
  Color4 = ShiftColorOXP(&HB33A01, 99)
  Color5 = ShiftColorOXP(&HE9BD96, 109)
  Color6 = ShiftColorOXP(&HB9B2AD, 109)
  Color7 = ShiftColorOXP(&H968A82, 109)
  Color8 = ShiftColorOXP(&HA25022, 109)
  Color9 = ShiftColorOXP(&HB8865E, 109)
  ColorB = ShiftColorOXP(&HDFBC86, 109)
  ColorC = ShiftColorOXP(&HFFBA77, 109)
  ColorD = ShiftColorOXP(&HE3D499, 109)
  ColorE = ShiftColorOXP(&HFFD996, 109)
  ColorF = ShiftColorOXP(&HE1A46D, 109)
  ColorG = ShiftColorOXP(&HCBA47B, 109)
  ColorH = ShiftColorOXP(&HDFDFDF, 109)
  ColorI = ShiftColorOXP(&HD0D0D0, 109)
 ElseIf (Mode = 3) Then
  Color1 = ShiftColorOXP(&HFDF2C3, 15)
  Color2 = ShiftColorOXP(&HDE8B45, 15)
  Color3 = ShiftColorOXP(&HDD873E, 15)
  Color4 = ShiftColorOXP(&HB33A01, 15)
  Color5 = ShiftColorOXP(&HE9BD96, 15)
  Color6 = ShiftColorOXP(&HB9B2AD, 15)
  Color7 = ShiftColorOXP(&H968A82, 15)
  Color8 = ShiftColorOXP(&HA25022, 15)
  Color9 = ShiftColorOXP(&HB8865E, 15)
  ColorB = ShiftColorOXP(&HDFBC86, 15)
  ColorC = ShiftColorOXP(&HFFBA77, 15)
  ColorD = ShiftColorOXP(&HE3D499, 15)
  ColorE = ShiftColorOXP(&HFFD996, 15)
  ColorF = ShiftColorOXP(&HE1A46D, 15)
  ColorG = ShiftColorOXP(&HCBA47B, 15)
  ColorH = ShiftColorOXP(&HDFDFDF, 15)
  ColorI = ShiftColorOXP(&HD0D0D0, 15)
 Else
  Color1 = ShiftColorOXP(&H808080, 195)
  Color2 = ShiftColorOXP(&H808080, 135)
  Color3 = ShiftColorOXP(&H808080, 135)
  Color4 = ShiftColorOXP(&H808080, 5)
  Color5 = Color1
  Color6 = Parent.BackColor
  Color7 = Color6
  Color8 = ShiftColorOXP(&H808080, 65)
  Color9 = Color6
  ColorA = Color6
  ColorB = Color4
  ColorC = Color4
  ColorD = Color4
  ColorE = Color4
  ColorF = Color4
  ColorG = Color4
  ColorH = Color6
  ColorI = Color6
 End If
 Call DrawVGradient(Color1, Color2, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, 1, UserControl.ScaleWidth - 1, UserControl.ScaleHeight / 3)
 '* English: Button gradient bottom.
 Call DrawVGradient(Color3, Color1, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, UserControl.ScaleHeight / 3, UserControl.ScaleWidth - 1, UserControl.ScaleHeight * 2 / 3 - 4)
 '* English: Lines for the text area
 Call APILine(2, 0, UserControl.ScaleWidth - 3, 0, &HA1A1A1)
 Call APILine(1, 0, 1, UserControl.ScaleHeight - 3, &HA1A1A1)
 '* English: Left shadow.
 If (Mode <> -1) Then
  Call DrawVGradient(ColorH, &HBBBBBB, 0, 0, 1, 3)
  Call DrawVGradient(&HBBBBBB, ColorA, 0, 4, 1, UserControl.ScaleHeight / 2 - 4)
  Call DrawVGradient(ColorA, &HBBBBBB, 0, UserControl.ScaleHeight / 2, 1, UserControl.ScaleHeight / 2 - 5)
  Call DrawVGradient(&HBBBBBB, ColorH, 0, UserControl.ScaleHeight - 5, 1, 2)
 Else
  Call DrawVGradient(ColorH, ColorH, 0, 0, 1, 3)
  Call DrawVGradient(ColorA, ColorA, 0, 4, 1, UserControl.ScaleHeight / 2 - 4)
  Call DrawVGradient(ColorA, ColorA, 0, UserControl.ScaleHeight / 2, 1, UserControl.ScaleHeight / 2 - 5)
  Call DrawVGradient(ColorH, ColorH, 0, UserControl.ScaleHeight - 5, 1, 2)
 End If
 '* English: Bottom shadows.
 Call APILine(1, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3, &H747474)
 Call APILine(1, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 2, &HA1A1A1)
 Call APILine(2, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 1, &HDDDDDD)
 '* English: Lines for the button area.
 Call DrawVGradient(ColorB, Color3, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, 1, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 1, UserControl.ScaleHeight / 3)
 Call DrawVGradient(Color3, ColorB, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, UserControl.ScaleHeight / 3, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 1, UserControl.ScaleHeight * 2 / 3 - 4)
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, 0, UserControl.ScaleWidth - 3, 0, Color4)
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 1, 1, UserControl.ScaleWidth - 4, 1, Color5)
 '* English: Right shadow.
 Call DrawVGradient(ColorH, ColorI, UserControl.ScaleWidth - 1, 2, UserControl.ScaleWidth, 3)
 Call DrawVGradient(ColorI, ColorA, UserControl.ScaleWidth - 1, 3, UserControl.ScaleWidth, UserControl.ScaleHeight / 2 - 6)
 Call DrawVGradient(ColorA, ColorI, UserControl.ScaleWidth - 1, UserControl.ScaleHeight / 2 - 2, UserControl.ScaleWidth, UserControl.ScaleHeight / 2 - 6)
 Call DrawVGradient(ColorI, ColorH, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 8, UserControl.ScaleWidth, 3)
 '* English: Layer1.
 Call DrawVGradient(Color4, Color3, UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 7)
 '* English: Layer2.
 Call DrawVGradient(Color4, ColorC, UserControl.ScaleWidth - 3, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 6)
 '* English: Doted Area / 1-Bottom.
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4, ColorG)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 4, Color7)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 5, ColorF)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 5, Color7)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 6, Color9)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 4, Color6)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, Color6)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 2, &HCACACA)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 5, UserControl.ScaleHeight - 2, &HBFBFBF)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 1, &HE4E4E4)
 '* English: Doted Area / 2-Botom
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 5, UserControl.ScaleHeight - 4, ColorD)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 5, ColorE)
 '* English: Doted Area / 3-Top
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, 0, IIf(Mode <> -1, &HA76E4A, ShiftColorOXP(&H808080, 55)))
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 0, Color6)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 1, Color8)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, 1, IIf(Mode <> -1, &HB3A49D, Parent.BackColor))
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 5, 1, Color9)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, 1, Color8)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, 2, Color9)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 3, Color8)
 '* English: Draw Twin Arrows.
 cx = m_btnRect.Left + (m_btnRect.Right - m_btnRect.Left) / 2 + 2
 cy = m_btnRect.Top + (m_btnRect.Bottom - m_btnRect.Top) / 2 - 1
 hPen = CreatePen(0, 1, IIf(Mode <> -1, &H0&, ShiftColorOXP(&H0&)))
 hPenOld = SelectObject(UserControl.hDC, hPen)
 '* English: Down Arrow.
 Call MoveToEx(UserControl.hDC, cx - 3, cy + 1, PT)
 Call LineTo(UserControl.hDC, cx + 1, cy + 1)
 Call LineTo(UserControl.hDC, cx, cy + 2)
 Call LineTo(UserControl.hDC, cx - 2, cy + 2)
 Call LineTo(UserControl.hDC, cx - 2, cy + 3)
 Call LineTo(UserControl.hDC, cx, cy + 3)
 Call LineTo(UserControl.hDC, cx - 1, cy + 4)
 Call LineTo(UserControl.hDC, cx - 1, cy + 6)
 '* English: Up Arrow.
 Call MoveToEx(UserControl.hDC, cx - 3, cy - 2, PT)
 Call LineTo(UserControl.hDC, cx + 1, cy - 2)
 Call LineTo(UserControl.hDC, cx, cy - 3)
 Call LineTo(UserControl.hDC, cx - 2, cy - 3)
 Call LineTo(UserControl.hDC, cx - 2, cy - 4)
 Call LineTo(UserControl.hDC, cx, cy - 4)
 Call LineTo(UserControl.hDC, cx - 1, cy - 5)
 Call LineTo(UserControl.hDC, cx - 1, cy - 7)
 '* English: Destroy PEN.
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
 '* English: Undo the offset.
 m_btnRect.Left = m_btnRect.Left + 4
End Sub

Private Sub DrawNiaWBSSButton(ByVal m_State As Long)
 '* English: NiaWBSS Style.
 '* Español: Estilo NiaWBSS.
 If (m_State = 1) Then
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
  tmpColor = NormalBorderColor
 ElseIf (m_State = 2) Then
  cValor = GetLngColor(ShiftColorOXP(GradientColor1, 65))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2, 65))
  tmpColor = HighLightBorderColor
 ElseIf (m_State = 3) Then
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
  tmpColor = SelectBorderColor
  tempBorderColor = tmpColor
 Else
  cValor = GetLngColor(ShiftColorOXP(GradientColor1))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2))
  tmpColor = ShiftColorOXP(DisabledColor, 156)
 End If
 Call DrawVGradient(cValor, iFor, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1)
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5))
 tmpC1 = UserControl.ScaleHeight / 2 - 2
 tmpC2 = ArrowColor
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 2, tmpC1 - 1, UserControl.ScaleWidth - 12, tmpC1 - 1, IIf(m_State = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 10, tmpC1 - 1, UserControl.ScaleWidth - 4, tmpC1 - 1, IIf(m_State = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 3, tmpC1, UserControl.ScaleWidth - 10, tmpC1, IIf(m_State = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 8, tmpC1, UserControl.ScaleWidth - 5, tmpC1, IIf(m_State = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 4, tmpC1 + 1, UserControl.ScaleWidth - 6, tmpC1 + 1, IIf(m_State = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 4, tmpC1 + 2, UserControl.ScaleWidth - 6, tmpC1 + 2, IIf(m_State = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 For tmpC3 = 3 To 6
  If (tmpC3 = 3) Or (tmpC3 = 4) Then
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 5, tmpC1 + tmpC3, UserControl.ScaleWidth - 7, tmpC1 + tmpC3, IIf(m_State = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
  Else
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 6, tmpC1 + tmpC3, UserControl.ScaleWidth - 8, tmpC1 + tmpC3, IIf(m_State = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
  End If
 Next
End Sub

Private Sub DrawOfficeButton(ByVal m_State As Long, ByVal WhatOffice As Integer)
 '* English: Draw Office Style appearance.
 '* Español: Dibuja la apariencia de Office.
 Select Case WhatOffice
  Case 1
   '* English: Style Office Xp, appearance default.
   '* Español: Estilo Office Xp, apariencia por defecto.
   If (m_State = 1) Then
    '* English: Normal Color.
    '* Español: Color Normal.
    tmpColor = NormalBorderColor
   ElseIf (m_State = 2) Then
    '* English: Highlight Color.
    '* Español: Color de Selección MouseMove.
    tmpColor = HighLightBorderColor
    cValor = 185
   ElseIf (m_State = 3) Then
    '* English: Down Color.
    '* Español: Color de Selección MouseDown.
    tmpColor = SelectBorderColor
    tempBorderColor = tmpColor
    cValor = 125
   Else
    '* English: Disabled Color.
    '* Español: Color deshabilitado.
    tmpColor = ConvertSystemColor(ShiftColorOXP(NormalBorderColor, 41))
   End If
   If (m_State > 1) Then
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), tmpColor, B
    UserControl.Line (UserControl.ScaleWidth - 2, 1)-(UserControl.ScaleWidth - 14, UserControl.ScaleHeight - 2), ShiftColorOXP(tmpColor, cValor), BF
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 15, UserControl.ScaleHeight - 1), tmpColor, B
   Else
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), tmpColor, B
    UserControl.Line (UserControl.ScaleWidth - 3, 2)-(UserControl.ScaleWidth - 13, UserControl.ScaleHeight - 3), tmpColor, BF
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 15, UserControl.ScaleHeight - 1), tmpColor, B
   End If
   With UserControl
    .CurrentX = UserControl.ScaleWidth - 13
    .CurrentY = Int(UserControl.ScaleHeight / 2) - 6
    .Font = "Marlett"
    .FontSize = 8
    If (Enabled = False) Then
     Call SetTextColor(.hDC, ShiftColorOXP(myArrowColor, 123))
    Else
     Call SetTextColor(.hDC, myArrowColor)
    End If
    Call DrawStateString(.hDC, 0, 0, "6", Len("6"), .CurrentX, .CurrentY, 0, 0, DST_TEXT Or DSS_NORMAL)
    .Font = txtCombo.Font
   End With
  Case 2
   '* English: Style Office 2000.
   '* Español: Estilo Office 2000.
   If (m_State = 1) Then
    '* English: Flat.
    '* Español: Normal.
    tmpColor = NormalBorderColor
    Call DrawRectangleBorder(UserControl.ScaleWidth - 13, 1, 12, UserControl.ScaleHeight - 2, ShiftColorOXP(tmpColor, 175), False)
   ElseIf (m_State = 2) Or (m_State = 3) Then
    '* English: Mouse Hover or Mouse Pushed.
    '* Español: Mouse presionado o MouseMove.
    If (m_State = 2) Then
     tmpColor = ShiftColorOXP(HighLightBorderColor)
    Else
     tmpColor = ShiftColorOXP(SelectBorderColor)
     tempBorderColor = tmpColor
    End If
    Call DrawCtlEdge(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BDR_SUNKENOUTER)
    m_btnRect.Left = m_btnRect.Left + 4
    Call APIFillRect(UserControl.hDC, m_btnRect, tmpColor)
    m_btnRect.Left = m_btnRect.Left - 1
    Call APIRectangle(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, tmpColor)
    Call APILine(m_btnRect.Left, m_btnRect.Top, m_btnRect.Left, m_btnRect.Bottom, tmpColor)
    If (m_State = 2) Then
     Call DrawCtlEdgeByRect(UserControl.hDC, m_btnRect, BDR_RAISEDINNER)
    Else
     Call DrawCtlEdgeByRect(UserControl.hDC, m_btnRect, BDR_SUNKENOUTER)
    End If
    m_btnRect.Left = m_btnRect.Left - 3
   Else
    '* English: Disabled control.
    '* Español: Control deshabilitado.
    tmpColor = ShiftColorOXP(NormalBorderColor, 193)
    Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 1, 17, UserControl.ScaleHeight - 2, ShiftColorOXP(tmpColor, 175), False)
   End If
   m_btnRect.Left = m_btnRect.Left + 4
   Call DrawStandardArrow(m_btnRect, IIf(m_State = -1, ShiftColorOXP(ArrowColor, 196), ArrowColor))
 End Select
End Sub

Private Sub DrawRectangleBorder(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional ByVal SetBorder As Boolean = True)
 Dim hBrush As Long, tempRect As RECT

 '* English: Draw a rectangle.
 '* Español: Crea el rectángulo.
 tempRect = m_btnRect
 m_btnRect.Left = X
 m_btnRect.Top = Y
 m_btnRect.Right = X + Width
 m_btnRect.Bottom = Y + Height
 hBrush = CreateSolidBrush(Color)
 If (SetBorder = True) Then
  Call FrameRect(UserControl.hDC, m_btnRect, hBrush)
 Else
  Call FillRect(UserControl.hDC, m_btnRect, hBrush)
 End If
 Call DeleteObject(hBrush)
 m_btnRect = tempRect
End Sub

Private Sub DrawRhombusButton(ByVal m_State As Long)
 '* English: Rhombus Style.
 '* Español: Estilo Rombo.
 If (m_State = 1) Then
  tmpColor = ShiftColorOXP(NormalBorderColor, 25)
 ElseIf (m_State = 2) Then
  tmpColor = ShiftColorOXP(HighLightBorderColor, 25)
 ElseIf (m_State = 3) Then
  tmpColor = ShiftColorOXP(SelectBorderColor, 25)
 Else
  tmpColor = ShiftColorOXP(&H0&, 237)
 End If
 Call DrawRectangleBorder(UserControl.ScaleWidth - 17, 1, 16, UserControl.ScaleHeight - 2, ShiftColorOXP(tmpColor, 25), False)
 If (m_State = 1) Then
  tmpColor = ShiftColorOXP(ArrowColor, 143)
 ElseIf (m_State = 2) Or (m_State = 3) Then
  tmpColor = ShiftColorOXP(ArrowColor, 113)
  tempBorderColor = tmpColor
 Else
  tmpColor = ShiftColorOXP(&H0&)
 End If
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 17, 0, 17, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5), True)
 '* English: Left top border.
 '* Español: Borde Superior Izquierdo.
 tmpC2 = 12
 For tmpC1 = 2 To 5
  tmpC2 = tmpC2 + 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left - 1, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, IIf(m_State = -1, BackColor, BackColor))
 Next
 tmpC2 = 17
 tmpC3 = -2
 For tmpC1 = 5 To 2 Step -1
  tmpC2 = tmpC2 - 1
  tmpC3 = tmpC3 + 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, tmpColor)
 Next
 '* English: Left bottom border
 '* Español: Borde Inferior Izquierdo.
 tmpC2 = 17
 For tmpC1 = 3 To 1 Step -1
  tmpC2 = tmpC2 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left - 1, UserControl.ScaleHeight - tmpC1 - 1, UserControl.ScaleWidth - tmpC2, UserControl.ScaleHeight - tmpC1 - 1, BackColor)
 Next
 tmpC2 = 12
 tmpC3 = 3
 For tmpC1 = 1 To 3
  tmpC2 = tmpC2 + 1
  tmpC3 = tmpC3 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, UserControl.ScaleHeight - tmpC1 - 1, UserControl.ScaleWidth - tmpC2, UserControl.ScaleHeight - tmpC1 - 1, tmpColor)
 Next
 '* English: Right top border
 '* Español: Borde Superior Derecho.
 tmpC2 = 0
 tmpC3 = 23
 For tmpC1 = 6 To 1 Step -1
  tmpC2 = tmpC2 + 1
  tmpC3 = tmpC3 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, Parent.BackColor)
 Next
 tmpC2 = 0
 tmpC3 = 17
 For tmpC1 = 6 To 1 Step -1
  tmpC2 = tmpC2 + 1
  tmpC3 = tmpC3 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, tmpColor)
 Next
 '* English: Right bottom border
 '* Español: Borde Inferior Derecho.
 tmpC2 = 6
 For tmpC1 = 0 To 3
  tmpC2 = tmpC2 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 19, UserControl.ScaleHeight - tmpC1 - 1, UserControl.ScaleWidth - tmpC2, UserControl.ScaleHeight - tmpC1 - 1, Parent.BackColor)
 Next
 tmpC2 = 1
 tmpC3 = 16
 For tmpC1 = 3 To 0 Step -1
  tmpC2 = tmpC2 + 1
  tmpC3 = tmpC3 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, UserControl.ScaleHeight - tmpC1 - 2, UserControl.ScaleWidth - tmpC2, UserControl.ScaleHeight - tmpC1 - 2, tmpColor)
 Next
 m_btnRect.Left = m_btnRect.Left + 1
 Call DrawStandardArrow(m_btnRect, IIf(m_State = -1, ShiftColorOXP(&H404040, 166), ArrowColor))
End Sub

Private Sub DrawStandardArrow(ByRef RT As RECT, ByVal lColor As Long)
 Dim PT   As POINTAPI, hPenOld As Long, cx As Long
 Dim hPen As Long, cy          As Long
 
 '* English: Draw the standard arrow in a Rect.
 '* Español: Dibuje la flecha normal en un Rect.
 If (AppearanceCombo <> 12) And (AppearanceCombo <> 5) Then hPen = 1 Else hPen = 0
 If (AppearanceCombo = 11) Then hPen = 2
 cx = RT.Left + (RT.Right - RT.Left) / 2 + hPen
 cy = RT.Top + (RT.Bottom - RT.Top) / 2
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(UserControl.hDC, hPen)
 Call MoveToEx(UserControl.hDC, cx - 3, cy - 1, PT)
 Call LineTo(UserControl.hDC, cx + 1, cy - 1)
 Call LineTo(UserControl.hDC, cx, cy)
 Call LineTo(UserControl.hDC, cx - 2, cy)
 Call LineTo(UserControl.hDC, cx, cy + 2)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

Private Sub DrawVGradient(ByVal lEndColor As Long, ByVal lStartcolor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long)
 Dim dR As Single, dG As Single, dB As Single, Ni As Long
 Dim sR As Single, sG As Single, sB As Single
 Dim eR As Single, eG As Single, eB As Single
 
 '* English: Draw a Vertical Gradient in the current hDC.
 '* Español: Dibuja un degradado en forma vertical.
 sR = (lStartcolor And &HFF)
 sG = (lStartcolor \ &H100) And &HFF
 sB = (lStartcolor And &HFF0000) / &H10000
 eR = (lEndColor And &HFF)
 eG = (lEndColor \ &H100) And &HFF
 eB = (lEndColor And &HFF0000) / &H10000
 dR = (sR - eR) / Y2
 dG = (sG - eG) / Y2
 dB = (sB - eB) / Y2
 For Ni = 0 To Y2
  Call APILine(X, Y + Ni, X2, Y + Ni, RGB(eR + (Ni * dR), eG + (Ni * dG), eB + (Ni * dB)))
 Next
End Sub

Private Sub DrawWinXPButton(ByVal Mode As Integer, ByRef XpAppearance As ComboXpAppearance)
 Dim lhDC  As Long, tempColor As Long, XpColor     As Long, tmpA As Long
 Dim lh    As Long, lW        As Long, lcH         As Long
 Dim lcW   As Long, Ni        As Single, tmpColorA As Long
  
 '* English: This Sub Draws the XpAppearance Button.
 '* Español: Este procedimiento dibuja el Botón estilo XP.
 lW = m_btnRect.Right - m_btnRect.Left
 lh = m_btnRect.Bottom - m_btnRect.Top
 lhDC = UserControl.hDC
 lcW = m_btnRect.Left + lW / 2 + 1
 lcH = m_btnRect.Top + lh / 2
 Select Case XpAppearance
  Case 1
   '* English: Style WinXp Aqua.
   '* Español: Estilo WinXp Aqua.
   tmpA = &H85614D
   tempBorderColor = &HC56A31
   tmpColorA = &HB99D7F
   If (Mode = 1) Then
    XpColor = &HF5C8B3
    tempColor = &HFFFFFF
   ElseIf (Mode = 2) Then
    XpColor = ShiftColorOXP(&HF5C8B3, 58)
    tempColor = &HFFFFFF
   ElseIf (Mode = 3) Then
    XpColor = &HF9A477
    tempColor = &HFFFFFF
   End If
  Case 2
   '* English: Style WinXp Olive Green.
   '* Español: Estilo WinXp Olive Green.
   tmpA = &HFFFFFF
   tempBorderColor = &H668C7D
   tmpColorA = &H94CCBC
   If (Mode = 1) Then
    XpColor = &H8BB4A4
    tempColor = &HFFFFFF
   ElseIf (Mode = 2) Then
    XpColor = &HA7D7CA
    tempColor = &HFFFFFF
   ElseIf (Mode = 3) Then
    XpColor = &H80AA98
    tempColor = &HFFFFFF
   End If
  Case 3
   '* English: Style WinXp Silver.
   '* Español: Estilo WinXp Silver.
   tempBorderColor = &HA29594
   tmpA = &H48483E
   tmpColorA = &HA29594
   If (Mode = 1) Then
    XpColor = &HDACCCB
    tempColor = &HFFFFFF
   ElseIf (Mode = 2) Then
    XpColor = ShiftColorOXP(&HDACCCB, 58)
    tempColor = &HFFFFFF
   ElseIf (Mode = 3) Then
    XpColor = &HE5D1CF
    tempColor = &HFFFFFF
   End If
  Case 4
   '* English: Style WinXp TasBlue.
   '* Español: Estilo WinXp TasBlue.
   tempBorderColor = &HF09F5F
   tmpA = ShiftColorOXP(&H703F00, 58)
   tmpColorA = &HF09F5F
   If (Mode = 1) Then
    XpColor = &HF0AF70
    tempColor = &HFFE7CF
   ElseIf (Mode = 2) Then
    XpColor = ShiftColorOXP(&HF0BF80, 58)
    tempColor = &HFFEFD0
   ElseIf (Mode = 3) Then
    XpColor = &HF09F5F
    tempColor = &HFFEFD0
   End If
  Case 5
   '* English: Style WinXp Gold.
   '* Español: Estilo WinXp Gold.
   tempBorderColor = &HBFE7F0
   tmpA = ShiftColorOXP(&H6F5820, 45)
   tmpColorA = &HBFE7F0
   If (Mode = 1) Then
    XpColor = ShiftColorOXP(&HCFFFFF, 54)
    tempColor = &HBFF0FF
   ElseIf (Mode = 2) Then
    XpColor = &HBFEFFF
    tempColor = ShiftColorOXP(&HCFFFFF, 58)
   ElseIf (Mode = 3) Then
    XpColor = &HCFFFFF
    tempColor = &HBFE8FF
   End If
  Case 6
   '* English: Style WinXp Blue.
   '* Español: Estilo WinXp Blue.
   tempBorderColor = ShiftColorOXP(&HA0672F, 123)
   tmpA = &H6F5820
   tmpColorA = ShiftColorOXP(&HA0672F, 123)
   If (Mode = 1) Then
    XpColor = &HEFF0F0
    tempColor = &HF0F7F0
   ElseIf (Mode = 2) Then
    XpColor = &HF0F8FF
    tempColor = &HF0F7F0
   ElseIf (Mode = 3) Then
    XpColor = &HF1946E
    tempColor = 15647412
   End If
  Case 7
   '* English: Style WinXp Custom.
   '* Español: Estilo WinXp Custom.
   tempBorderColor = SelectBorderColor
   tmpA = ArrowColor
   If (Mode = 1) Then
    XpColor = NormalBorderColor
    tempColor = &HFFFFFF
   ElseIf (Mode = 2) Then
    XpColor = HighLightBorderColor
    tempColor = &HFFFFFF
   ElseIf (Mode = 3) Then
    XpColor = SelectBorderColor
    tempColor = &HFFFFFF
   End If
   tmpColorA = XpColor
 End Select
 Call DrawGradient(UserControl.hDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom, tempColor, XpColor, 1)
 Call APIRectangle(lhDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right - m_btnRect.Left - 1, m_btnRect.Bottom - m_btnRect.Top - 1, tmpColorA)
 Call SetPixel(lhDC, m_btnRect.Left, m_btnRect.Top, ShiftColorOXP(tmpColorA, 32))
 Call SetPixel(lhDC, m_btnRect.Right - 1, m_btnRect.Top, ShiftColorOXP(tmpColorA, 32))
 Call SetPixel(lhDC, m_btnRect.Right - 1, m_btnRect.Bottom - 1, ShiftColorOXP(tmpColorA, 32))
 Call SetPixel(lhDC, m_btnRect.Left, m_btnRect.Bottom - 1, ShiftColorOXP(tmpColorA, 32))
 '* English: Draw The XP Style Arrow.
 '* Español: Dibuja la flecha estilo XP.
 If (Mode = -1) Then
  tempColor = &HE5ECEC
  For Ni = 3 To lh
   Call APILine(m_btnRect.Left + 1, lh - Ni + 3, m_btnRect.Right - 1, lh - Ni + 3, tempColor)
  Next
  Call APIRectangle(lhDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right - m_btnRect.Left - 1, m_btnRect.Bottom - m_btnRect.Top - 1, &HE5ECEC)
  Call SetPixel(lhDC, m_btnRect.Left, m_btnRect.Top, ShiftColorOXP(&HEED2C1, 32))
  Call SetPixel(lhDC, m_btnRect.Right - 1, m_btnRect.Top, ShiftColorOXP(&HEED2C1, 32))
  Call SetPixel(lhDC, m_btnRect.Right - 1, m_btnRect.Bottom - 1, ShiftColorOXP(&HEED2C1, 32))
  Call SetPixel(lhDC, m_btnRect.Left, m_btnRect.Bottom - 1, ShiftColorOXP(&HEED2C1, 32))
 End If
 tempColor = IIf(Mode = -1, &HC2C9C9, tmpA)
 Call APILine(lcW - 5, lcH - 2, lcW, lcH + 3, tempColor)
 Call APILine(lcW - 4, lcH - 2, lcW, lcH + 2, tempColor)
 Call APILine(lcW - 4, lcH - 3, lcW, lcH + 1, tempColor)
 Call APILine(lcW + 3, lcH - 2, lcW - 2, lcH + 3, tempColor)
 Call APILine(lcW + 2, lcH - 2, lcW - 2, lcH + 2, tempColor)
 Call APILine(lcW + 2, lcH - 3, lcW - 2, lcH + 1, tempColor)
End Sub

Private Sub DrawXpButton(ByVal m_State As Long)
 '* English: Additional Xp Style.
 '* Español: Estilo Xp Adicional.
 If (m_State = 1) Then
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
  tmpColor = NormalBorderColor
 ElseIf (m_State = 2) Then
  cValor = GetLngColor(ShiftColorOXP(GradientColor1, 65))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2, 65))
  tmpColor = HighLightBorderColor
 ElseIf (m_State = 3) Then
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
  tmpColor = SelectBorderColor
  tempBorderColor = tmpColor
 Else
  cValor = GetLngColor(ShiftColorOXP(GradientColor1))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2))
  tmpColor = DisabledColor
 End If
 Call DrawVGradient(cValor, iFor, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1)
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left - 3, 1, UserControl.ScaleWidth - 144, 1, IIf(m_State = -1, ShiftColorOXP(DisabledColor, 198), ShiftColorOXP(tmpColor, 168)))
 tmpC1 = m_btnRect.Right - m_btnRect.Left
 tmpC2 = m_btnRect.Bottom - m_btnRect.Top + 1
 tmpC1 = m_btnRect.Left + tmpC1 / 2 + 1
 tmpC2 = m_btnRect.Top + tmpC2 / 2
 tmpC3 = IIf(m_State = -1, &HC2C9C9, ArrowColor)
 Call APILine(tmpC1 - 5, tmpC2 - 2, tmpC1, tmpC2 + 3, tmpC3)
 Call APILine(tmpC1 - 4, tmpC2 - 2, tmpC1, tmpC2 + 2, tmpC3)
 Call APILine(tmpC1 - 4, tmpC2 - 3, tmpC1, tmpC2 + 1, tmpC3)
 Call APILine(tmpC1 + 3, tmpC2 - 2, tmpC1 - 2, tmpC2 + 3, tmpC3)
 Call APILine(tmpC1 + 2, tmpC2 - 2, tmpC1 - 2, tmpC2 + 2, tmpC3)
 Call APILine(tmpC1 + 2, tmpC2 - 3, tmpC1 - 2, tmpC2 + 1, tmpC3)
 tmpC2 = 15
 tmpC3 = 178
 iFor = 10
 For tmpC1 = 1 To 16
  tmpC2 = tmpC2 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC2, 1, UserControl.ScaleWidth - tmpC1, 1, IIf(m_State = -1, ShiftColorOXP(GradientColor1, tmpC3), ShiftColorOXP(GradientColor1, tmpC3)))
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - tmpC1, UserControl.ScaleHeight - 2, IIf(m_State = -1, ShiftColorOXP(&H646464, tmpC3), ShiftColorOXP(&H646464, iFor)))
  tmpC3 = tmpC3 - 5
  iFor = iFor + 5
 Next
 m_btnRect.Bottom = m_btnRect.Bottom - 11
 tmpC3 = 128
 iFor = 70
 For tmpC1 = 0 To 12
  Call APILine(m_btnRect.Left + 1, m_btnRect.Top + tmpC1 - 1, m_btnRect.Left + 1, m_btnRect.Bottom + tmpC1 - 1, IIf(m_State = -1, ShiftColorOXP(GradientColor1, tmpC3), ShiftColorOXP(GradientColor1, tmpC3)))
  Call APILine(UserControl.ScaleWidth - 2, m_btnRect.Top + tmpC1 - 1, UserControl.ScaleWidth - 2, m_btnRect.Bottom + tmpC1 - 1, IIf(m_State = -1, ShiftColorOXP(&H646464, tmpC3), ShiftColorOXP(&H646464, iFor)))
  tmpC3 = tmpC3 + 5
  iFor = iFor - 5
 Next
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left - 3, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 144, UserControl.ScaleHeight - 2, IIf(m_State = -1, ShiftColorOXP(DisabledColor, 198), ShiftColorOXP(tmpColor, 168)))
End Sub

Private Sub Espera(ByVal Segundos As Single)
 Dim ComienzoSeg As Single, FinSeg As Single
 
 '* English: Wait a certain time.
 '* Español: Esperar un determinado tiempo.
 ComienzoSeg = Timer
 FinSeg = ComienzoSeg + Segundos
 Do While FinSeg > Timer
  DoEvents
  If (ComienzoSeg > Timer) Then FinSeg = FinSeg - 24 * 60 * 60
 Loop
End Sub

Public Function FindItemText(ByVal Text As String, Optional ByVal Compare As StringCompare = 0) As Long
 Dim i As Long
 
 '* English: Search Text in the list and return the position.
 '* Español: Busca una cadena dentro de la lista y devuelve su posición en la misma.
 FindItemText = -1
 If (Text = "") Or (Compare < 0) Or (Compare > 2) Then Exit Function
 For i = 1 To sumItem
  If (Compare = 0) Then
   If (InStr(1, UCase$(ListContents(i).Text), UCase$(Text), vbBinaryCompare) <> 0) Then
    FindItemText = i
    Exit For
   End If
  ElseIf (Compare = 1) Then
   If (UCase$(Text) = UCase$(ListContents(i).Text)) Then
    FindItemText = i
    Exit For
   End If
  Else
   If (Text = ListContents(i).Text) Then
    FindItemText = i
    Exit For
   End If
  End If
 Next
End Function

Public Function GetControlVersion() As String
 '* English: Control Version.
 '* Español: Versión del Control.
 GetControlVersion = Version & " © " & Year(Now) & "."
End Function

Private Function GetLngColor(ByVal Color As Long) As Long
 '* English: The GetSysColor function retrieves the current color of the specified display element. Display elements are the parts of a window and the Windows display that appear on the system display screen.
 '* Español: Recupera el color actual del elemento de despliegue especificado.
 If (Color And &H80000000) Then
  GetLngColor = GetSysColor(Color And &H7FFFFFFF)
 Else
  GetLngColor = Color
 End If
End Function

Private Function InFocusControl(ByVal ObjecthWnd As Long) As Boolean
 Dim mPos  As POINTAPI, KeyLeft As Boolean
 Dim oRect As RECT, KeyRight    As Boolean
 
 '* English: Verifies if the mouse is on the object or if one makes clic outside of him.
 '* Español: Verifica si el mouse se encuentra sobre el objeto ó si se hace clic fuera de él.
 Call GetCursorPos(mPos)
 Call GetWindowRect(ObjecthWnd, oRect)
 KeyLeft = GetAsyncKeyState(VK_LBUTTON)
 KeyRight = GetAsyncKeyState(VK_RBUTTON)
 UserControl.MousePointer = myMousePointer
 '* English: Set MouseIcon only drop down list.
 '* Español: Coloca el icono del mouse únicamente donde se expande ó retrae la lista.
 If (mPos.X > oRect.Left + (UserControl.ScaleWidth - 18)) And (mPos.X < oRect.Right) Then
  Set UserControl.MouseIcon = myMouseIcon
 Else
  Set UserControl.MouseIcon = Nothing
 End If
 If (mPos.X >= oRect.Left) And (mPos.X <= oRect.Right) And (mPos.Y >= oRect.Top) And (mPos.Y <= oRect.Bottom) Then
  InFocusControl = True
  First = 0
 ElseIf (KeyLeft = True) Or (KeyRight = True) And (First = 0) Then
  If (HighlightedItem > -1) And (FirstView <> 1) Then
   If (mPos.X < oRect.Left) Or (mPos.X > oRect.Right) Or (mPos.Y < oRect.Top) Or (mPos.Y > oRect.Bottom) Then
    InFocusControl = False
    First = 1
    picList.Visible = False
   End If
  End If
 End If
End Function

Private Sub IsEnabled(ByVal isTrue As Boolean)
 '* English: Shows the state of Enabled or Disabled of the Control.
 '* Español: Muestra el estado de Habilitado ó Deshabilitado del Control.
 If (isTrue = True) Then
  Call DrawAppearance(myAppearanceCombo, 1)
 Else
  Call DrawAppearance(myAppearanceCombo, -1)
 End If
End Sub

Public Sub ItemEnabled(ByVal ListIndex As Long, ByVal ItemEnabled As Boolean)
 '* English: Sets the Enabled/disabled property in an Item
 '* Español: Habilita o Deshabilita un Item.
On Error GoTo myErr:
 ListContents(ListIndex).Enabled = ItemEnabled
 Exit Sub
myErr:
End Sub

Public Function List(ByVal ListIndex As Long) As String
 '* English: Show one item of the list.
 '* Español: Muestra un elemento de la lista.
 HighlightedItem = ListIndex
 ItemFocus = ListIndex
 List = ListContents(ListIndex).Text
 Call IsEnabled(ControlEnabled)
End Function

Private Function ListCount1() As Long
On Error Resume Next
 '* English: Total of elements of the list.
 '* Español: Total de elementos de la lista.
 If (sumItem = 0) Then Exit Function
 ListCount1 = UBound(ListContents) + 1
End Function

Private Function ListIndex1(Optional ByVal Item As Long = -1) As Long
 '* English: Function to know the position of the selected index of the list.
 '* Español: Función para saber la posición del index seleccionado de la lista.
 If (Item = -1) And (ListCount1 > 0) Then
  ListIndex1 = IIf(ItemFocus = 0, -1, ItemFocus)
 Else
  ListIndex1 = Item
 End If
 If (ListIndex1 > 0) Then
  HighlightedItem = ListIndex1
  ItemFocus = ListIndex1
  Text = ListContents(ListIndex1).Text
 End If
End Function

Private Sub LongToRGB(ByVal lColor As Long)
 '* English: Convert a Long to RGB format.
 '* Español: Convierte un Long en formato RGB.
 RGBColor.Red = lColor And &HFF
 RGBColor.Green = (lColor \ &H100) And &HFF
 RGBColor.Blue = (lColor \ &H10000) And &HFF
End Sub

Private Function NoFindIndex(ByVal Index As Long) As Boolean
 Dim i As Long
 
 '* English: Search if the Index has not been assigned.
 '* Español: Busca si ya no se ha asignado este Index.
 NoFindIndex = False
 For i = 1 To sumItem
  If (ListContents(i).Index = Index) Then NoFindIndex = True: Exit For
 Next
End Function

Public Sub OrderList(Optional ByVal Order As Integer = 1)
 Dim N As Long, i As Long, j As Long
 
 '* English: Order the list with the search method (I Exchange).
 '* Español: Ordena la lista con el método de búsqueda (Intercambio).
 If (Order <> 1) And (Order <> 2) Then Exit Sub
 ReDim OrderListContents(0)
 N = UBound(ListContents)
 For i = 1 To N
  ReDim Preserve OrderListContents(i)
  OrderListContents(i).Color = ListContents(i).Color
  OrderListContents(i).Enabled = ListContents(i).Enabled
  Set OrderListContents(i).Image = ListContents(i).Image
  OrderListContents(i).Index = ListContents(i).Index
  Set OrderListContents(i).MouseIcon = ListContents(i).MouseIcon
  OrderListContents(i).SeparatorLine = ListContents(i).SeparatorLine
  OrderListContents(i).Tag = ListContents(i).Tag
  OrderListContents(i).Text = ListContents(i).Text
  OrderListContents(i).ToolTipText = ListContents(i).ToolTipText
 Next
 i = 1
 For i = 1 To N
  For j = (i + 1) To N
   Select Case Order
    Case 1: If (OrderListContents(j).Text < OrderListContents(i).Text) Then Call SetInfo(i, j)
    Case 2: If (OrderListContents(j).Text > OrderListContents(i).Text) Then Call SetInfo(i, j)
   End Select
  Next
 Next
 ReDim ListContents(0)
 For i = 1 To N
  ReDim Preserve ListContents(i)
  ListContents(i).Color = OrderListContents(i).Color
  ListContents(i).Enabled = OrderListContents(i).Enabled
  Set ListContents(i).Image = OrderListContents(i).Image
  ListContents(i).Index = OrderListContents(i).Index
  Set ListContents(i).MouseIcon = OrderListContents(i).MouseIcon
  ListContents(i).SeparatorLine = OrderListContents(i).SeparatorLine
  ListContents(i).Tag = OrderListContents(i).Tag
  ListContents(i).Text = OrderListContents(i).Text
  ListContents(i).ToolTipText = OrderListContents(i).ToolTipText
 Next
 ReDim OrderListContents(0)
End Sub

Private Sub PicDisabled(ByRef picTo As Object)
 Dim sTMPpathFName As String, lFlags As Long
 
 '* English: Disables a image
 '* Español: Deshabilita la imagen.
 Select Case picTo.Picture.Type
  Case vbPicTypeBitmap
   lFlags = DST_BITMAP
  Case vbPicTypeIcon
   lFlags = DST_ICON
  Case Else
   lFlags = DST_COMPLEX
 End Select
 If Not (picTo.Picture Is Nothing) Then
  Call DrawState(picTo.hDC, 0, 0, picTo.Picture, 0, 0, 0, picTo.ScaleWidth, picTo.ScaleHeight, lFlags Or DSS_DISABLED)
  sTMPpathFName = App.Path + "\~ConvIcon2Bmp.tmp"
  SavePicture picTo.Image, sTMPpathFName
  Set picTo.Picture = LoadPicture(sTMPpathFName)
  Call Kill(sTMPpathFName)
  picTo.Refresh
 End If
End Sub

Public Sub RemoveItem(ByVal Index As Long)
 Dim TempList() As PropertyCombo, sCount As Long
 Dim Count      As Long, TempCount       As Long
 
 '* English: Delete a Item from the list
 '* Español: Elimina un elemento de la lista.
On Error GoTo myErr
 If (ListCount = 0) Then Exit Sub
 If (sumItem > 0) Then sumItem = Abs(sumItem - 1)
 For Count = 1 To ListCount1 - 1
  If (Index <> Count) Then
   sCount = sCount + 1
   ReDim Preserve TempList(sCount)
   TempList(sCount).Color = ListContents(Count).Color
   TempList(sCount).Enabled = ListContents(Count).Enabled
   Set TempList(sCount).Image = ListContents(Count).Image
   TempList(sCount).Index = sCount
   TempList(sCount).Tag = ListContents(Count).Tag
   TempList(sCount).Text = ListContents(Count).Text
   TempList(sCount).ToolTipText = ListContents(Count).ToolTipText
  End If
 Next
 TempCount = Abs(Count - 2)
 sCount = 0
 ReDim ListContents(0)
 For Count = 1 To TempCount
  sCount = sCount + 1
  ReDim Preserve ListContents(sCount)
  ListContents(sCount).Color = TempList(Count).Color
  ListContents(sCount).Enabled = TempList(Count).Enabled
  Set ListContents(sCount).Image = TempList(Count).Image
  ListContents(sCount).Index = TempList(Count).Index
  ListContents(sCount).Tag = TempList(Count).Tag
  ListContents(sCount).Text = TempList(Count).Text
  ListContents(sCount).ToolTipText = TempList(Count).ToolTipText
 Next
 ReDim Preserve ListContents(TempCount)
 Refresh
 MaxListLength = Abs(MaxListLength - 1)
 If (myText = ListContents(MaxListLength + 1).Text) Then
  ListIndex = -1
  ItemFocus = -1
  Text = ""
  Call IsEnabled(ControlEnabled)
 End If
 Exit Sub
myErr:
End Sub

Private Sub SetInfo(ByVal i As Long, ByVal j As Long)
 Dim Temp As Variant
 
 '* English: Reorders the values.
 '* Español: Reordena los valores.
 Temp = OrderListContents(i).Color
 OrderListContents(i).Color = OrderListContents(j).Color
 OrderListContents(j).Color = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).Enabled
 OrderListContents(i).Enabled = OrderListContents(j).Enabled
 OrderListContents(j).Enabled = Temp
 '*******************************************************************'
 Set Temp = OrderListContents(i).Image
 Set OrderListContents(i).Image = OrderListContents(j).Image
 Set OrderListContents(j).Image = Temp
 '*******************************************************************'
 Set Temp = OrderListContents(i).MouseIcon
 Set OrderListContents(i).MouseIcon = OrderListContents(j).MouseIcon
 Set OrderListContents(j).MouseIcon = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).SeparatorLine
 OrderListContents(i).SeparatorLine = OrderListContents(j).SeparatorLine
 OrderListContents(j).SeparatorLine = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).Tag
 OrderListContents(i).Tag = OrderListContents(j).Tag
 OrderListContents(j).Tag = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).Text
 OrderListContents(i).Text = OrderListContents(j).Text
 OrderListContents(j).Text = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).ToolTipText
 OrderListContents(i).ToolTipText = OrderListContents(j).ToolTipText
 OrderListContents(j).ToolTipText = Temp
End Sub

Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
 Dim cRed   As Long, cBlue  As Long
 Dim Delta  As Long, cGreen As Long

 '* English: Shift a color
 '* Español: Devuelve un Color con menos intensidad.
 cBlue = ((theColor \ &H10000) Mod &H100)
 cGreen = ((theColor \ &H100) Mod &H100)
 cRed = (theColor And &HFF)
 Delta = &HFF - Base
 cBlue = Base + cBlue * Delta \ &HFF
 cGreen = Base + cGreen * Delta \ &HFF
 cRed = Base + cRed * Delta \ &HFF
 If (cRed > 255) Then cRed = 255
 If (cGreen > 255) Then cGreen = 255
 If (cBlue > 255) Then cBlue = 255
 ShiftColorOXP = cRed + 256& * cGreen + 65536 * cBlue
End Function
