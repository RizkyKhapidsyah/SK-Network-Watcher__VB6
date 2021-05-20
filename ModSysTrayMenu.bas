Attribute VB_Name = "ModSysTrayMenu"
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

Declare Function CreatePopupMenu Lib "user32" () As Long

Declare Function TrackPopupMenu Lib "user32" _
        (ByVal hMenu As Long, _
        ByVal wFlags As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nReserved As Long, _
        ByVal hwnd As Long, _
        ByVal lprc As Any) As Long
        
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" _
        (ByVal hMenu As Long, _
        ByVal wFlags As Long, _
        ByVal wIDNewItem As Long, _
        ByVal lpNewItem As Any) As Long
        
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" _
        (ByVal hMenu As Long, _
        ByVal nPosition As Long, _
        ByVal wFlags As Long, _
        ByVal wIDNewItem As Long, _
        ByVal lpString As Any) As Long
        
Public Declare Function DestroyMenu Lib "user32" _
        (ByVal hMenu As Long) As Long
        
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Declare Function BitBlt Lib "gdi32" _
        (ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Declare Function SetRect Lib "user32" (lpRect As RECT, _
        ByVal X1 As Long, _
        ByVal Y1 As Long, _
        ByVal X2 As Long, _
        ByVal Y2 As Long) As Long

Declare Function DrawCaption Lib "user32" _
        (ByVal hwnd As Long, _
        ByVal hdc As Long, _
        pcRect As RECT, _
        ByVal un As Long) As Long
        
Declare Function GetMenuItemRect Lib "user32" _
        (ByVal hwnd As Long, ByVal hMenu As Long, _
        ByVal uItem As Long, _
        lprcItem As RECT) As Long

Declare Function GetMenuItemCount Lib "user32" _
        (ByVal hMenu As Long) As Long

Declare Function GetPixel Lib "gdi32" _
        (ByVal hdc As Long, _
        ByVal X As Long, _
        ByVal Y As Long) As Long
        
Declare Function SetPixel Lib "gdi32" _
        (ByVal hdc As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal crColor As Long) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    itemData As Long
End Type

Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    itemData As Long
End Type

Public Type PointAPI
    X As Long
    Y As Long
End Type

Const MF_APPEND = &H100&
Const MF_BYCOMMAND = &H0&
Const MF_BYPOSITION = &H400&
Const MF_DEFAULT = &H1000&
Const MF_DISABLED = &H2&
Const MF_ENABLED = &H0&
Const MF_GRAYED = &H1&
Const MF_MENUBARBREAK = &H20&
Const MF_OWNERDRAW = &H100&
Const MF_POPUP = &H10&
Const MF_REMOVE = &H1000&
Const MF_SEPARATOR = &H800&
Const MF_STRING = &H0&
Const MF_UNCHECKED = &H0&
Const MF_BITMAP = &H4&
Const MF_USECHECKBITMAPS = &H200&

Public Const MF_CHECKED = &H8&
Public Const MFT_RADIOCHECK = &H200&

Const TPM_RETURNCMD = &H100&

Const DC_GRADIENT = &H20
Const DC_ACTIVE = &H1
Const DC_ICON = &H4
Const DC_SMALLCAP = &H2
Const DC_TEXT = &H8

Public hMenu As Long
Public hSubmenu As Long
Public chkMnuFlags(2) As Long
Public MP As PointAPI, sMenu As Long
Public mnuHeight As Single

Public Skin(0 To 20) As Boolean
Public SkinText(0 To 20) As String * 20           ' used for the popupwindow..

Public RectPopup As RECT

Public Sub MeasureMenu(ByRef lP As Long)
    Dim MIS As MEASUREITEMSTRUCT
    
    CopyMemory MIS, ByVal lP, Len(MIS)
        MIS.itemWidth = 5   '(18 - 1) - 12. I don't know where the 12 comes
        '
    CopyMemory ByVal lP, MIS, Len(MIS)
End Sub

Public Sub DrawMenu(ByRef lP As Long)
    
    Dim DIS As DRAWITEMSTRUCT, rct As RECT, lRslt As Long
    
    CopyMemory DIS, ByVal lP, Len(DIS)
    
    With AppForm
       mnuHeight = 0
        'String Menus
        GetMenuItemRect .hwnd, hMenu, 1, rct
        mnuHeight = (rct.Bottom - rct.Top) * ((AdaptorPopup * 4) + 2)  '(GetMenuItemCount(hMenu) - GetMenuItemCount(hSubmenu))
        
        RectPopup.Bottom = mnuHeight
        
        'Separators
        GetMenuItemRect .hwnd, hMenu, 3, rct
        mnuHeight = (mnuHeight + (rct.Bottom - rct.Top) * (AdaptorPopup + 1)) - (AdaptorPopup * 12) '
        RectPopup.Bottom = mnuHeight
        
        'set the size of our sidebar
        SetRect rct, 0, 0, mnuHeight, 18
        DrawCaption .hwnd, .hdc, rct, DC_SMALLCAP Or DC_ACTIVE Or DC_TEXT Or DC_GRADIENT
        
        Dim X As Single, Y As Single
        Dim nColor As Long
        For X = 0 To mnuHeight
            For Y = 0 To 17
                nColor = GetPixel(.hdc, X, Y)
                SetPixel DIS.hdc, Y, mnuHeight - X, nColor
            Next Y
        Next X
        'remove the caption picture from the user form
        .Cls
        'Hopefully this operation was so fast that you did'nt see it happen.
     End With
RectPopup = rct
End Sub

Public Sub MenuPopUp()
Dim X As Integer
Dim Displ As Byte
Dim cnt As Integer
    'create the menu
    hMenu = CreatePopupMenu()
    hSubmenu = CreatePopupMenu()
    
       AppendMenu hMenu, MF_OWNERDRAW Or MF_DISABLED, 1000, 0& 'SideBar
    AppendMenu hMenu, MF_MENUBARBREAK, 1000, "About"
    'AppendMenu hMenu, 0&, 1000, "About"
    
    For cnt = 1 To modInfoAdaptor.m_AdaptorCnt
    'take per ..block 100
    'check first if the ip not 0.0.0.0 is..
    If Mid(modInfoAdaptor.Xmenu(cnt).Caption, 1, 1) <> "0" Then
    X = 3100 + ((cnt - 1) * 100)
    AppendMenu hMenu, MF_SEPARATOR, 0&, 0&
    
    AppendMenu hMenu, chkMnuFlags(1 + CheckVisible(cnt)), X, modInfoAdaptor.Xmenu(cnt).Caption
    AppendMenu hMenu, 0& + (1 - CheckVisible(cnt)), X + 1, "Change Scale"
    AppendMenu hMenu, 0& + (1 - CheckVisible(cnt)), X + 2, "View More"
    AppendMenu hMenu, 0& + (1 - CheckVisible(cnt)), X + 3, SkinText(cnt)
    End If
    Next
    
     AppendMenu hMenu, MF_SEPARATOR, 0&, 0&
    AppendMenu hMenu, 0&, 2000, "Exit"
 'tmrPopup.enable = True
 End Sub

Public Sub MenuTrack(frm As Form)

Dim AdaptorChoose As Integer    'value adaptor..
Dim strNr  As String   ' counter to check the value of sMenu.. index.
Dim locVal As Integer  'local value..
Dim SkinCaption As String  ' Display analoge or digital meters.
    GetCursorPos MP

'Send a sendmessage to the popupbox..

   frmMenu.tmrPopup.Enabled = True
    'Check if the cursor  isn't outside the popubox..
    DoEvents
    sMenu = TrackPopupMenu(hMenu, TPM_RETURNCMD, MP.X, MP.Y, 0, frm.hwnd, 0&)
    'store the sMenu value local.
   locVal = sMenu
            
            '***************** ABOUT *********************
     If sMenu = 1000 Then
            frmAbout.Show
            frmMenu.tmrPopup.Enabled = False
     End If
            '***************** EXIT ??? *********************
     If sMenu = 2000 Then
     frmMenu.tmrPopup.Enabled = False
            UnHook
            Unload AppForm
            Exit Sub
    End If
    ' **************** if > 3000 then the user clicked something on the adaptors ******
    If sMenu >= 3000 Then
        sMenu = sMenu - 3000  ' example sMenu = 3102
        'Convert byte to string to get the first number..
        strNr = CStr(sMenu)
        Adaptor = CInt(Mid(strNr, 1, 1))
        sMenu = sMenu - (100 * Adaptor)
       frmMenu.tmrPopup.Enabled = False
        Select Case sMenu
                Case 0
                    modInfoAdaptor.Xmenu(Adaptor).ShowThisWindow
                    ModifyMenu hMenu, locVal, chkMnuFlags(1 + CheckVisible(Adaptor)), locVal, modInfoAdaptor.Xmenu(Adaptor).Caption
                    'Enable or disable the View more - change scale ect..
                    ModifyMenu hMenu, locVal + 1, (0& + 1 - CheckVisible(Adaptor)), locVal + 1, "Change Scale"
                    ModifyMenu hMenu, locVal + 2, (0& + 1 - CheckVisible(Adaptor)), locVal + 2, "View More"
                    ModifyMenu hMenu, locVal + 3, (0& + 1 - CheckVisible(Adaptor)), locVal + 3, "Show Digital"
                    
                Case 1
                    modInfoAdaptor.Xmenu(Adaptor).ScaleSettings_Click
                    
                Case 2
                    modInfoAdaptor.Xmenu(Adaptor).View_more_Click
                    
                    
                Case 3
                    SkinCaption = ChangeSkin(Adaptor)
                    ModifyMenu hMenu, locVal, 0&, locVal, SkinCaption
                    modInfoAdaptor.Xmenu(Adaptor).ShowAnalog_Digital (Skin(Adaptor))
                    
        End Select
               
    End If
End Sub

'Check if the form is visible..
Function CheckVisible(ByVal menu As Integer) As Byte

If modInfoAdaptor.Xmenu(menu).Visible = True Then
        CheckVisible = 1
     Else
        CheckVisible = 0
     End If
End Function

Function ChangeSkin(ByVal adapter As Integer) As String
'Skin = Not Skin
If Skin(adapter) = True Then
    'Show analoge meters..
    ChangeSkin = "Show Digital "
Else
    'Show digital meters
    ChangeSkin = "Show Analog"
End If
    Skin(adapter) = Not Skin(adapter)
End Function

'
