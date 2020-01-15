Attribute VB_Name = "Utility"
Option Explicit

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, Pts As Any, ByVal cPoints As Long) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Any) As Long
Private Declare Function ReleaseCapture& Lib "user32" ()
Private Declare Function DrawTextW Lib "user32" (ByVal hDC&, ByVal lpUniCode&, ByVal CharCount&, lpRect As Any, ByVal wFormat&) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, Optional ByVal pPoint As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal SrcColor&, ByVal hPalette&, DstColor&) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As Any, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Any, pClipRect As Any) As Long

Public Declare Function ExcludeClipRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function RemoveClipRgn Lib "gdi32" Alias "SelectClipRgn" (ByVal hDC As Long, Optional ByVal hRgn As Long) As Long
Public Declare Function GetTickCount64 Lib "kernel32" () As Currency
'Public Declare Function GetClientRect Lib "user32" (ByVal hwnd&, Rct As Any) As Long
Public Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long

Public Const BP_GROUPBOX = 4, TABP_TABITEM = 1, TABP_TABITEMLEFTEDGE = 2, TABP_TABITEMRIGHTEDGE = 3, TABP_PANE = 9, TABP_BODY = 10
Public Const WM_MOVE = &H3, WM_MOUSEACTIVATE = &H21, MA_NOACTIVATEANDEAT = &H4, WM_LBUTTONDOWN = &H201
Public Const DT_CENTER = &H1, DT_RIGHT = &H2, DT_SINGLELINE = &H20, DT_WORD_ELLIPSIS = &H40000, DT_RTLREADING = &H20000, DT_WORDBREAK = &H10, DT_CALCRECT = &H400

'DrawEdge Consts for rendering "old style"
Private Const BF_LEFT = &H1, BF_TOP = &H2, BF_RIGHT = &H4, BF_BOTTOM = &H8
Private Const BDR_RAISEDOUTER = &H1, BDR_SUNKENOUTER = &H2, BDR_RAISEDINNER = &H4, BDR_SUNKENINNER = &H8
 
'拷贝容器背景
Public Sub CopyContainerBG(hWndCont As Long, hwnd As Long, hDC As Long)
    Dim Pt(0 To 1) As Long
    
    MapWindowPoints hWndCont, hwnd, Pt(0), 1
    
    Const WM_PAINT = &HF
    SetViewportOrgEx hDC, Pt(0), Pt(1)
    SendMessageW hWndCont, WM_PAINT, hDC, ByVal 0&
    SetViewportOrgEx hDC, 0, 0
    
End Sub
 
Public Function WindowUnderMouse(Optional x, Optional y) As Long
    Dim Pt(0 To 1) As Long, hwnd As Long
    
    GetCursorPos Pt(0)
    hwnd = WindowFromPoint(Pt(0), Pt(1))
    MapWindowPoints 0, hwnd, Pt(0), 1
    x = Pt(0)
    y = Pt(1)
    WindowUnderMouse = hwnd
    
End Function

Public Function TranslateColor(ByVal SysColorID As SystemColorConstants) As Long
    
    OleTranslateColor SysColorID, 0, TranslateColor
    
End Function

Function GetTextExtent(hDC As Long, Text As String, ByVal DTFlags As Long, R() As Long, Optional TH As Long) As Long
    Dim Ext() As Long
    
    Ext = R
    DTFlags = DTFlags Or DT_CALCRECT
    DrawTextW hDC, StrPtr(Text), Len(Text), Ext(0), DTFlags
 
    GetTextExtent = Ext(2) - Ext(0)
    TH = Ext(3) - Ext(1)
    
End Function

Public Sub DrawText(hDC As Long, Text As String, ForeColor As Long, DTFlags As Long, Rct() As Long)

    DrawTextW hDC, StrPtr(Text), Len(Text), Rct(0), DTFlags   '...to be able to fallback to DrawTextW on XP
    
End Sub

Public Sub DrawThemeBGFrame(hTheme As Long, hDC As Long, R() As Long, Optional ByVal PartID As Long = BP_GROUPBOX, Optional ByVal State As Long = 1)

    If hTheme Then
        DrawThemeBackground hTheme, hDC, PartID, State, R(0), R(0)
    Else
        DrawEdge hDC, R(0), BDR_SUNKENOUTER Or BDR_RAISEDINNER, BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
    End If

End Sub

Public Sub DrawThemeBGTab(hTheme As Long, hDC As Long, R() As Long, Optional ByVal PartID As Long = TABP_BODY, Optional ByVal State As Long = 1, Optional ByVal BackColor As Long = -1)

    If hTheme <> 0 And BackColor <> -1 Then
        DrawThemeBackground hTheme, hDC, PartID, State, R(0), R(0)
    Else
        DrawEdge hDC, R(0), BDR_RAISEDINNER, BF_TOP Or IIf(PartID = TABP_TABITEMRIGHTEDGE, 0, BF_LEFT) Or IIf(PartID <> TABP_TABITEMRIGHTEDGE, 0, BF_RIGHT)
    End If

End Sub

Public Function IsXP() As Boolean
    On Error Resume Next

    GetTickCount64
    IsXP = Err

End Function

Public Sub StartMoving(hwnd As Long)

    ReleaseCapture
    SendMessageW hwnd, &H112, &HF012&, 0&

End Sub

'对象是否存在于集合中
Public Function ObjectInCol( _
    ByVal strKey As String, _
    ByRef oCol As Collection) As Boolean
    On Error GoTo RunErr
    
    ObjectInCol = VarType(oCol.Item(strKey))
    ObjectInCol = True
    
    Exit Function
RunErr:
    
End Function

'从集合中删除对象
Public Function RemoveObjectFromCol( _
    ByVal strKey As String, _
    ByRef oCol As Collection) As Boolean
    On Error GoTo RunErr
    
    oCol.Remove strKey
    RemoveObjectFromCol = True
    
    Exit Function
RunErr:
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@Function Name:    GetNoXString
'@Main Func:        获得第X处的字符串
'@Author:           denglf
'@Last Modify:      2008-05-08
'@Param:            strText,待分析的字符串文本
'@Param:            intX,第X处
'@Param:            strDelimiter,分隔符号
'@Returns:          返回获取到的字符串
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNoXString( _
    ByVal strText As String, _
    Optional ByVal intX As Integer, _
    Optional ByVal strDelimiter As String = " ") As String
    On Error GoTo RunErr
    Dim strTemp() As String
    Dim i As Integer
    Dim intDims As Integer
    
    strTemp = Split(strText, strDelimiter)
    intDims = UBound(strTemp)
    If intX >= (intDims + 1) Then
        intX = intDims
        GetNoXString = Trim$(strTemp(intDims))
    ElseIf intX <= 0 Then
        intDims = 0
        GetNoXString = Trim$(strTemp(intDims))
    Else
        For i = 0 To intDims
            If (i + 1) = intX Then
                GetNoXString = Trim$(strTemp(i))
                Exit For
            End If
        Next
    End If
    
    Exit Function
RunErr:
    
End Function
