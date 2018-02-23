Attribute VB_Name = "mdlLog"
Option Explicit


Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType     As Long
    CtlID       As Long
    itemID      As Long
    itemAction  As Long
    itemState   As Long
    hwndItem    As Long
    hdc         As Long
    rcItem      As RECT
    ItemData    As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const COLOR_HIGHLIGHT       As Integer = 13
Private Const COLOR_HIGHLIGHTTEXT   As Integer = 14
Private Const COLOR_WINDOW          As Integer = 5
Private Const COLOR_WINDOWTEXT      As Integer = 8
Private Const LB_GETTEXT            As Integer = &H189
Private Const WM_DRAWITEM           As Integer = &H2B
Private Const ODS_FOCUS             As Integer = &H10
Private Const ODT_LISTBOX           As Integer = 2

Public Const GWL_WNDPROC            As Integer = (-4)

Public lPrevWndProc                 As Long

Public Function SubClassedList(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tItem As DRAWITEMSTRUCT
    Dim sBuff As String * 255
    Dim sItem As String
    Dim lBack As Long
    
    If Msg = WM_DRAWITEM Then
        'Redraw the listbox
        'This function only passes the Address of the DrawItem Structure, so we need to
        'use the CopyMemory API to Get a Copy into the Variable we setup:
        Call CopyMemory(tItem, ByVal lParam, Len(tItem))
        'Make sure we're dealing with a Listbox
        If tItem.CtlType = ODT_LISTBOX Then
            'Get the Item Text
            Call SendMessage(tItem.hwndItem, LB_GETTEXT, tItem.itemID, ByVal sBuff)
            sItem = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
            If (tItem.itemState And ODS_FOCUS) Then
                'Item has Focus, Highlight it, I'm using the Default Focus
                'Colors for this example.
                lBack = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
                Call FillRect(tItem.hdc, tItem.rcItem, lBack)
                Call SetBkColor(tItem.hdc, GetSysColor(COLOR_HIGHLIGHT))
                Call SetTextColor(tItem.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT))
                TextOut tItem.hdc, tItem.rcItem.Left, tItem.rcItem.Top, ByVal sItem, Len(sItem)
                DrawFocusRect tItem.hdc, tItem.rcItem
            Else
                'Item Doesn't Have Focus, Draw it's Colored Background
                'Create a Brush using the Color we stored in ItemData
                lBack = CreateSolidBrush(tItem.ItemData)
                'Paint the Item Area
'                Call FillRect(tItem.hdc, tItem.rcItem, lBack)
                
                Call FillRect(tItem.hdc, tItem.rcItem, vbWhite)
                'Set the Text Colors
                'Call SetBkColor(tItem.hdc, tItem.ItemData)
                Call SetBkColor(tItem.hdc, vbWhite)
                Call SetTextColor(tItem.hdc, IIf(tItem.ItemData = vbBlue, vbBlue, vbBlack))
                
                'Display the Item Text
                TextOut tItem.hdc, tItem.rcItem.Left, tItem.rcItem.Top, ByVal sItem, Len(sItem)
            End If
            Call DeleteObject(lBack)
            'Don't Need to Pass a Value on as we've just handled the Message ourselves
            SubClassedList = 0
            Exit Function
                    
        End If
            
    End If
    SubClassedList = CallWindowProc(lPrevWndProc, hwnd, Msg, wParam, lParam)
End Function

Public Function OnLog(ByVal lpStr As String)
    Dim LogFormList As ListBox
    
    On Error Resume Next
    
    Select Case nNowForm
        Case FM_RUN: Set LogFormList = frmRun.lstMsg
        Case FM_SETUP: Set LogFormList = frmSetup.lstSetupMsg
        Case FM_SYSTEM: Set LogFormList = frmSystem.lstSystemMsg
        Case FM_SPLASH: Set LogFormList = frmSplash.lstMsg
        Case Else: Exit Function
    End Select
    
    If lpStr = "" Then
        LogFormList.Clear
    Else
        LogFormList.AddItem (lpStr)
    End If
    
    If (lpStr Like UCase("*Error*")) Then
        LogFormList.ItemData(LogFormList.NewIndex) = vbBlue
    Else
        LogFormList.ItemData(LogFormList.NewIndex) = vbWhite
    End If
    
    If LogFormList.ListIndex < 30000 Then
        LogFormList.ListIndex = LogFormList.ListCount - 1
    Else
        LogFormList.Clear
    End If
End Function

