VERSION 5.00
Begin VB.UserControl SSTab 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   3075
   ScaleWidth      =   4350
   Begin VB.Timer tmrHover 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   60
   End
End
Attribute VB_Name = "SSTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@Module Name:  SSTab(.ctl)
'@Main Func:    多页Tab
'@Author:       denglf
'@Last Modify:  2018-09-03
'@Notes:        替代 Microsoft SSTab
'@Interface:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Event Click(PreviousTab As Long)

Private Const cMinPos As Long = -15000
Private Const cMaxPos As Long = -75000

Private Enum TabStateConstants
    Normal = 0
    Hot = 1
    Pressed = 2
    Disabled = 3
    Defaulted = 4
End Enum

Private Type TArea
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TTab
    Caption As String
    Area As TArea
    Enabled As Boolean
    Visible As Boolean
    Hovered As Boolean
    Picture As StdPicture
    Controls As New Collection
End Type

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private WithEvents mclsSubclass As TSubclass.Subclass
Attribute mclsSubclass.VB_VarHelpID = -1
Private mclsGdi As TGdi.WinGdi

Private mTabs() As TTab
Attribute mTabs.VB_VarHelpID = -1
Private mlngTabIndex As Long                'Tab索引
Private mdblTabHeight As Single             'Tab高度
Private mdblTabMaxWidth As Long             'Tab最大宽度
Private mlngTabInnerSpace As Long           'Tab间距
Private mblnFocused As Boolean              'Tab控件是否获得焦点

Private mclrForeColor As OLE_COLOR          '前景色
Private mclrBackColor As OLE_COLOR          '背景色
Private mclrDisabledColor As OLE_COLOR      '不可用色
Private mclrBorderColor As OLE_COLOR        '边框色
Private mclrTabActiveColor As OLE_COLOR     'Tab活动色
Private mclrTabInActiveColor As OLE_COLOR   'Tab非活动色
Private mclrTabBorderColor As OLE_COLOR     'Tab边框色
Private mclrTabBackColor As OLE_COLOR       'Tab背景色

Private mblnEnabled As Boolean              '是否可用
Private mblnWordWrap As Boolean             '是否自动换行

Private Sub UserControl_Initialize()
    
    Set mclsGdi = New TGdi.WinGdi
    UserControl.AutoRedraw = True
    UserControl.ScaleMode = vbPixels
    
    Dim i As Long
    ReDim mTabs(0 To 2)
    For i = 0 To 2
        mTabs(i).Enabled = True
        mTabs(i).Visible = True
        mTabs(i).Caption = "Tab " & i
    Next
    mlngTabInnerSpace = 4
    mdblTabHeight = 330 / Screen.TwipsPerPixelY
    mlngTabIndex = 0
    
    mclrForeColor = vbButtonText
    mclrBackColor = vbWindowBackground
    mclrDisabledColor = vbGrayText
    mclrBorderColor = vb3DShadow
    mclrTabActiveColor = vbWindowBackground
    mclrTabInActiveColor = vbButtonFace
    mclrTabBorderColor = vbActiveBorder
    mclrTabBackColor = vbWindowBackground
    
    mblnEnabled = True
    mblnWordWrap = False
    mblnFocused = False
    
End Sub

Private Sub UserControl_Terminate()
    
    Set mclsGdi = Nothing
    
End Sub

Private Function mclsSubclass_OnMessage(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean) As Long
    
    Select Case uMsg
'    Case WM_MOUSEACTIVATE
'        Dim X As Long
'        Dim Y As Long
'        If Ambient.UserMode And WindowUnderMouse(X, Y) = hWnd Then
'            If Y > mdblTabHeight Then
'                'Result = MA_NOACTIVATEANDEAT
'                'Exit Sub 'override the default and leave
'                bEatIt = True
'            End If
'        End If
        
    Case WM_LBUTTONDOWN
        pSwitchTab lParam And &HFFFF&, lParam \ 65536
        
    End Select
    
End Function

'窗口句柄
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'是否可用
Public Property Get Enabled() As Boolean
    Enabled = mblnEnabled
End Property
Public Property Let Enabled(ByVal blnValue As Boolean)
    mblnEnabled = blnValue
    UserControl.Enabled = blnValue
    Draw
    PropertyChanged "Enabled"
End Property

'自动换行
Public Property Get WordWrap() As Boolean
    WordWrap = mblnWordWrap
End Property
Public Property Let WordWrap(ByVal blnValue As Boolean)
    mblnWordWrap = blnValue
    Draw
    PropertyChanged "WordWrap"
End Property

'前景色
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mclrForeColor
End Property
Public Property Let ForeColor(ByVal clrValue As OLE_COLOR)
    mclrForeColor = clrValue
    Draw
    PropertyChanged "ForeColor"
End Property

'背景色
Public Property Get BackColor() As OLE_COLOR
    BackColor = mclrBackColor
End Property
Public Property Let BackColor(ByVal clrValue As OLE_COLOR)
    mclrBackColor = clrValue
    Refresh
    PropertyChanged "BackColor"
End Property

'不可用色
Public Property Get DisabledColor() As OLE_COLOR
    DisabledColor = mclrDisabledColor
End Property
Public Property Let DisabledColor(ByVal clrValue As OLE_COLOR)
    mclrDisabledColor = clrValue
    Refresh
    PropertyChanged "DisabledColor"
End Property

'边框色
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mclrBorderColor
End Property
Public Property Let BorderColor(ByVal clrValue As OLE_COLOR)
    mclrBorderColor = clrValue
    Refresh
    PropertyChanged "BorderColor"
End Property

'Tab活动色
Public Property Get TabActiveColor() As OLE_COLOR
    TabActiveColor = mclrTabActiveColor
End Property
Public Property Let TabActiveColor(ByVal clrValue As OLE_COLOR)
    mclrTabActiveColor = clrValue
    Refresh
    PropertyChanged "TabActiveColor"
End Property

'Tab非活动色
Public Property Get TabInActiveColor() As OLE_COLOR
    TabInActiveColor = mclrTabInActiveColor
End Property
Public Property Let TabInActiveColor(ByVal clrValue As OLE_COLOR)
    mclrTabInActiveColor = clrValue
    Refresh
    PropertyChanged "TabInActiveColor"
End Property

'Tab边框色
Public Property Get TabBorderColor() As OLE_COLOR
    TabBorderColor = mclrTabBorderColor
End Property
Public Property Let TabBorderColor(ByVal clrValue As OLE_COLOR)
    mclrTabBorderColor = clrValue
    Refresh
    PropertyChanged "TabBorderColor"
End Property

'Tab背景色
Public Property Get TabBackColor() As OLE_COLOR
    TabBackColor = mclrTabBackColor
End Property
Public Property Let TabBackColor(ByVal clrValue As OLE_COLOR)
    mclrTabBackColor = clrValue
    Refresh
    PropertyChanged "TabBackColor"
End Property

'Tab数量
Public Property Get Tabs() As Long
    Tabs = UBound(mTabs) + 1
End Property
Public Property Let Tabs(ByVal lngValue As Long)

    '最多64个Tab
    If lngValue > 64 Then
        Err.Raise 5
    End If
    
    If lngValue < 1 Then
        lngValue = 1
    End If
    
    Dim i As Long
    Dim j As Long
    If Tabs < lngValue Then
        j = Tabs
        ReDim Preserve mTabs(0 To lngValue - 1)
        For i = j To lngValue - 1
            mTabs(i).Caption = "Tab " & i
            mTabs(i).Enabled = True
            mTabs(i).Visible = True
        Next
        
    ElseIf Tabs > lngValue Then
        For i = lngValue To Tabs - 1
            If mTabs(i).Controls.Count Then
                Err.Raise 5
            End If
        Next
        ReDim Preserve mTabs(0 To lngValue - 1)
        TabIndex = 0
    End If
    
    Draw
    
    PropertyChanged "Tabs"
    
End Property

'Tab索引
Public Property Get TabIndex() As Long
    TabIndex = mlngTabIndex
End Property
Public Property Let TabIndex(ByVal lngValue As Long)
    
    If mlngTabIndex = lngValue Or lngValue < 0 Or lngValue >= Tabs Then
        Exit Property
    End If
    
    Dim lngPrevTabIndex As Long
    lngPrevTabIndex = mlngTabIndex
    mlngTabIndex = lngValue
    pClearControl
    pSaveControl lngPrevTabIndex
    pShowControl lngValue
    RaiseEvent Click(lngPrevTabIndex)
    Refresh
    
    PropertyChanged "Tab"
    
End Property

'Tab标题
Public Property Get Caption() As String
    Caption = mTabs(mlngTabIndex).Caption
End Property
Public Property Let Caption(ByVal strValue As String)
    TabCaption(mlngTabIndex) = strValue
End Property

'指定Tab标题
Public Property Get TabCaption(ByVal lngTabIndex As Long) As String
    TabCaption = mTabs(lngTabIndex).Caption
End Property
Public Property Let TabCaption(ByVal lngTabIndex As Long, ByVal strValue As String)
    
    mTabs(lngTabIndex).Caption = strValue
    Draw
    PropertyChanged "TabCaption(" & lngTabIndex & ")"
    
End Property

''Tab图像Key
'Public Property Get ImageKey() As String
'    ImageKey = mTabs(mlngTabIndex).ImageKey
'End Property
'Public Property Let ImageKey(ByVal strValue As String)
'    TabImageKey(mlngTabIndex) = strValue
'End Property
'
''指定Tab图像Key
'Public Property Get TabImageKey(ByVal lngTabIndex As Long) As String
'    TabImageKey = mTabs(lngTabIndex).ImageKey
'End Property
'Public Property Let TabImageKey(ByVal lngTabIndex As Long, ByVal strValue As String)
'    mTabs(lngTabIndex).ImageKey = strValue
'    Draw
'    PropertyChanged "TabImageKey(" & lngTabIndex & ")"
'End Property

'Tab图像
Public Property Get Picture() As StdPicture
    Set Picture = mTabs(mlngTabIndex).Picture
End Property
Public Property Set Picture(ByVal picValue As StdPicture)
    Set TabPicture(mlngTabIndex) = picValue
End Property

'指定Tab图像
Public Property Get TabPicture(ByVal lngTabIndex As Long) As StdPicture
    Set TabPicture = mTabs(lngTabIndex).Picture
End Property
Public Property Let TabPicture(ByVal lngTabIndex As Long, ByVal picValue As StdPicture)
    Set TabPicture(lngTabIndex) = picValue
End Property
Public Property Set TabPicture(ByVal lngTabIndex As Long, ByVal picValue As StdPicture)
    Set mTabs(lngTabIndex).Picture = picValue
    Draw
    PropertyChanged "TabPicture(" & lngTabIndex & ")"
End Property

'Tab高度
Public Property Get TabHeight() As Single
    TabHeight = mdblTabHeight * Screen.TwipsPerPixelY
End Property
Public Property Let TabHeight(ByVal dblValue As Single)
    If mdblTabHeight = dblValue / Screen.TwipsPerPixelY Then
        Exit Property
    Else
        mdblTabHeight = dblValue / Screen.TwipsPerPixelY
    End If
    Draw
    PropertyChanged "TabHeight"
End Property

'Tab宽度
Public Property Get TabMaxWidth() As Single
    TabMaxWidth = mdblTabMaxWidth * Screen.TwipsPerPixelY
End Property
Public Property Let TabMaxWidth(ByVal dblValue As Single)
    If mdblTabMaxWidth = dblValue / Screen.TwipsPerPixelY Then
        Exit Property
    Else
        mdblTabMaxWidth = dblValue / Screen.TwipsPerPixelY
    End If
    Draw
    PropertyChanged "TabMaxWidth"
End Property

'Tab是否可用
Public Property Get TabEnabled(ByVal lngTabIndex As Long) As Boolean
    TabEnabled = mTabs(lngTabIndex).Enabled
End Property
Public Property Let TabEnabled(ByVal lngTabIndex As Long, ByVal blnValue As Boolean)
    mTabs(lngTabIndex).Enabled = blnValue
    Draw
End Property

'Tab是否可见
Public Property Get TabVisible(ByVal lngTabIndex As Long) As Boolean
    TabVisible = mTabs(lngTabIndex).Visible
End Property
Public Property Let TabVisible(ByVal lngTabIndex As Long, ByVal blnValue As Boolean)
    
    If Tabs = 1 And blnValue = False Then
        Exit Property
    End If
    
    Dim i As Long
    mTabs(lngTabIndex).Visible = blnValue
    If blnValue = False And mlngTabIndex = lngTabIndex Then
        For i = mlngTabIndex + 1 To UBound(mTabs)
            If mTabs(i).Visible Then
                TabIndex = i
                Exit Property
            End If
        Next
        For i = mlngTabIndex - 1 To 0 Step -1
            If mTabs(i).Visible Then
                TabIndex = i
                Exit Property
            End If
        Next
    End If
    
    Draw
    
End Property

'刷新
Public Sub Refresh()
    Dim Ctl As Object
    
    Draw
    
    For Each Ctl In UserControl.ContainedControls
        If TypeOf Ctl Is TSSTab.SSTab Then
            Ctl.Refresh
        End If
    Next
    
End Sub

Private Sub tmrHover_Timer()
    Dim x As Long
    Dim y As Long
    Dim i As Long
    
    If WindowUnderMouse(x, y) = hwnd Then
        If pSwitchTab(x, y, True) = tmrHover.Tag Then
            Exit Sub
        End If
    End If
    
    For i = 0 To UBound(mTabs)
        mTabs(i).Hovered = False
    Next
    
    tmrHover.Enabled = False
    
    Draw
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = vbKeyLeft Then
        
        KeyCode = 0
        For i = mlngTabIndex - 1 To 0 Step -1
            If mTabs(i).Visible And mTabs(i).Enabled Then
                TabIndex = i
                Exit Sub
            End If
        Next
        For i = UBound(mTabs) To mlngTabIndex + 1 Step -1
            If mTabs(i).Visible And mTabs(i).Enabled Then
                TabIndex = i
                Exit Sub
            End If
        Next
      
    ElseIf KeyCode = vbKeyRight Then
        
        KeyCode = 0
        For i = mlngTabIndex + 1 To UBound(mTabs)
            If mTabs(i).Visible And mTabs(i).Enabled Then
                TabIndex = i
                Exit Sub
            End If
        Next
        For i = 0 To mlngTabIndex - 1
            If mTabs(i).Visible And mTabs(i).Enabled Then
                TabIndex = i
                Exit Sub
            End If
        Next
        
    End If
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    pSwitchTab ScaleX(x, ScaleMode, vbPixels), ScaleY(y, ScaleMode, vbPixels)
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngTabIndex As Long
    Dim lngX As Long
    Dim lngY As Long
    
    lngX = UserControl.ScaleX(x, UserControl.ScaleMode, vbPixels)
    lngY = UserControl.ScaleY(y, UserControl.ScaleMode, vbPixels)
    lngTabIndex = pSwitchTab(lngX, lngY, True)
    If lngTabIndex >= 0 And tmrHover.Enabled = False Then
        mTabs(lngTabIndex).Hovered = True
        tmrHover.Tag = lngTabIndex
        tmrHover.Enabled = True
        Draw
    End If
    
End Sub

Private Sub UserControl_Show()

    '显示控件
    If Ambient.UserMode Then
        pShowControl mlngTabIndex
    End If
    
    '设计模式才需要子类处理，目的是单击鼠标左键能切换选项卡
    If Not Ambient.UserMode Then
        Set mclsSubclass = New TSubclass.Subclass
        mclsSubclass.hwnd = UserControl.hwnd
    End If
    
    Draw
    
End Sub

Private Sub UserControl_Hide()
    
    If Not mclsSubclass Is Nothing Then
        Set mclsSubclass = Nothing
    End If
    
End Sub

Private Sub UserControl_Resize()
    
    Refresh
    
End Sub

Private Sub UserControl_EnterFocus()
    
    mblnFocused = True
    Draw
    
End Sub

Private Sub UserControl_ExitFocus()
    
    'mblnFocused = False
    Draw
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    mlngTabIndex = PropBag.ReadProperty("Tab", 0)
    mblnEnabled = PropBag.ReadProperty("Enabled", True)
    mblnWordWrap = PropBag.ReadProperty("WordWrap", False)
    mclrForeColor = PropBag.ReadProperty("ForeColor", vbButtonText)
    mclrBackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    mclrDisabledColor = PropBag.ReadProperty("DisabledColor", vbGrayText)
    mclrBorderColor = PropBag.ReadProperty("BorderColor", vb3DShadow)
    mclrTabActiveColor = PropBag.ReadProperty("TabActiveColor", vbWindowBackground)
    mclrTabInActiveColor = PropBag.ReadProperty("TabInActiveColor", vbButtonFace)
    mclrTabBorderColor = PropBag.ReadProperty("TabBorderColor", vbActiveBorder)
    mclrTabBackColor = PropBag.ReadProperty("TabBackColor", vbWindowBackground)
    Tabs = PropBag.ReadProperty("Tabs", 3)
    TabHeight = UserControl.ScaleY(PropBag.ReadProperty("TabHeight", 582), vbHimetric, vbTwips)
    TabMaxWidth = UserControl.ScaleY(PropBag.ReadProperty("TabMaxWidth", 0), vbHimetric, vbTwips)
    
    pReadTabs PropBag
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "Enabled", mblnEnabled, True
    PropBag.WriteProperty "WordWrap", mblnWordWrap, False
    PropBag.WriteProperty "ForeColor", mclrForeColor, vbButtonText
    PropBag.WriteProperty "BackColor", mclrBackColor, vbWindowBackground
    PropBag.WriteProperty "DisabledColor", mclrDisabledColor, vbGrayText
    PropBag.WriteProperty "BorderColor", mclrBorderColor, vb3DShadow
    PropBag.WriteProperty "TabActiveColor", mclrTabActiveColor, vbWindowBackground
    PropBag.WriteProperty "TabInActiveColor", mclrTabInActiveColor, vbButtonFace
    PropBag.WriteProperty "TabBorderColor", mclrTabBorderColor, vbActiveBorder
    PropBag.WriteProperty "TabBackColor", mclrTabBackColor, vbWindowBackground
    PropBag.WriteProperty "Tabs", Tabs, 3
    PropBag.WriteProperty "Tab", TabIndex, 0
    PropBag.WriteProperty "TabHeight", CLng(UserControl.ScaleY(TabHeight, vbTwips, vbHimetric)), 582
    PropBag.WriteProperty "TabMaxWidth", CLng(UserControl.ScaleY(TabMaxWidth, vbTwips, vbHimetric)), 0
    
    pWriteTabs PropBag
    
End Sub

'读Tab信息
Private Sub pReadTabs(PropBag As PropertyBag)
    Dim strItem As String
    Dim strValue As String
    Dim strCtlName As String
    Dim lngCount As Long
    Dim i As Long
    Dim j As Long
    
    For i = 0 To Tabs - 1
        
        '标题
        mTabs(i).Caption = PropBag.ReadProperty("TabCaption(" & i & ")", "Tab " & i)
        
        '图片
        Set mTabs(i).Picture = PropBag.ReadProperty("TabPicture(" & i & ")", Nothing)
        If Not mTabs(i).Picture Is Nothing Then
            If mTabs(i).Picture.Handle = 0 Then
                Set mTabs(i).Picture = Nothing
            End If
        End If
        
        '控件
        lngCount = PropBag.ReadProperty("Tab(" & i & ").ControlCount", 0)
        For j = 1 To lngCount
            strItem = "Tab(" & i & ").Control(" & j & ")"
            strValue = LCase$(PropBag.ReadProperty(strItem, ""))
            If strValue <> "" Then
                strCtlName = GetNoXString(strValue, 1, Chr(1))
                If Not ObjectInCol(strCtlName, mTabs(i).Controls) Then
                    mTabs(i).Controls.Add strValue, strCtlName
                End If
            End If
        Next
        
    Next
    
End Sub

'写Tab信息
Private Sub pWriteTabs(PropBag As PropertyBag)
    Dim oCtl As Object
    Dim strItem As String
    Dim strValue As String
    Dim i As Long
    Dim j As Long
    
    '设计环境时，保存Tab控件，避免当前Tab中增删控件后，没有切换Tab导致控件没有保存
    pSaveControl mlngTabIndex, False
    
    '清理冗余控件
    pClearControl
    
    '保存控件信息
    For i = 0 To Tabs - 1
        
        '标题
        PropBag.WriteProperty "TabCaption(" & i & ")", mTabs(i).Caption, "Tab " & i
        
        '图片
        PropBag.WriteProperty "TabPicture(" & i & ")", mTabs(i).Picture, Nothing
        
        '控件数量
        PropBag.WriteProperty "Tab(" & i & ").ControlCount", mTabs(i).Controls.Count
        
        '控件状态
        For j = 1 To mTabs(i).Controls.Count
            strItem = "Tab(" & i & ").Control(" & j & ")"
            strValue = LCase$(mTabs(i).Controls.Item(j))
            PropBag.WriteProperty strItem, strValue, ""
        Next
        
    Next
    
End Sub

'清理冗余控件
Private Sub pClearControl()
    Dim colCtls As Collection
    Dim oCtl As Object
    Dim strValue As String
    Dim strCtlName As String
    Dim blnExist As Boolean
    Dim i As Long
    Dim j As Long
 
    If Not Ambient.UserMode Then
        '设计环境
        
        '查找不存在的控件
        Set colCtls = New Collection
        For i = 0 To Tabs - 1
            For j = 1 To mTabs(i).Controls.Count
                blnExist = False
                strValue = mTabs(i).Controls.Item(j)
                strCtlName = GetNoXString(strValue, 1, Chr(1))
                For Each oCtl In UserControl.ContainedControls
                    If StrComp(strCtlName, pGetCtlName(oCtl), vbTextCompare) = 0 Then
                        blnExist = True
                        Exit For
                    End If
                Next
                If Not blnExist Then
                    colCtls.Add strCtlName
                End If
            Next
        Next
        
        '清理
        For i = 1 To colCtls.Count
            strCtlName = LCase$(colCtls.Item(i))
            For j = 0 To Tabs - 1
                RemoveObjectFromCol strCtlName, mTabs(j).Controls
            Next
        Next
        Set colCtls = Nothing
        
    End If
    
End Sub

'存储控件
'lngPrevTabIndex： 上一个Tab的索引
'算法： 当切换Tab时，缓存上一Tab中控件的状态
Private Sub pSaveControl(ByVal lngPrevTabIndex As Long, Optional ByVal blnInVisible As Boolean = True)
    Dim oCtl As Object
    Dim varValue As Variant
    Dim strValue As String
    Dim strCtlName As String
    Dim strCtlPos As String
    Dim lngCtlPos As Long
    Dim lngCtlTab As Long

    If Not Ambient.UserMode Then
        '设计环境
        
        '缓存上一Tab中的控件
        For Each oCtl In UserControl.ContainedControls
            strCtlName = pGetCtlName(oCtl)
            lngCtlPos = pGetCtlPos(oCtl, strCtlPos)
            lngCtlTab = pGetCtlTab(oCtl)
            
            If lngCtlPos > cMinPos Then
                strValue = strCtlName & Chr(1) & lngCtlPos & Chr(1) & lngCtlTab
                If ObjectInCol(strCtlName, mTabs(lngPrevTabIndex).Controls) Then
                    '修改
                    RemoveObjectFromCol strCtlName, mTabs(lngPrevTabIndex).Controls
                    mTabs(lngPrevTabIndex).Controls.Add strValue, strCtlName
                Else
                    '缓存
                    mTabs(lngPrevTabIndex).Controls.Add strValue, strCtlName
                End If
                
                '让控件不可见
                If blnInVisible Then
                    pSetCtlVisible oCtl, False
                    pSetCtlTab oCtl, False
                End If
                
            End If
        Next
    
    Else
        '运行环境
        
        '让控件不可见,便于按下Tab键时,上一Tab中的控件不能获得焦点
        For Each varValue In mTabs(lngPrevTabIndex).Controls
            strCtlName = GetNoXString(varValue, 1, Chr(1))
            Set oCtl = pGetCtlObj(strCtlName)
            If ObjPtr(oCtl) > 0 Then
                pSetCtlVisible oCtl, False
                pSetCtlTab oCtl, False
            End If
        Next
        
    End If
    
End Sub

'显示控件
'lngTabIndex： 当前Tab的索引
Private Sub pShowControl(ByVal lngTabIndex As Long)
    'On Error Resume Next
    Dim oCtl As Object
    Dim varValue As Variant
    Dim strCtlName As String
    Dim strCtlPos As String
    Dim lngCtlTab As Long
    
    '显示当前Tab中的控件
    For Each varValue In mTabs(lngTabIndex).Controls
        strCtlName = GetNoXString(varValue, 1, Chr(1))
        Set oCtl = pGetCtlObj(strCtlName)
        If ObjPtr(oCtl) > 0 Then
            
            '恢复控件的位置
            strCtlPos = GetNoXString(varValue, 2, Chr(1))
            pSetCtlPos oCtl, strCtlPos
        
            '恢复控件的可用状态
            lngCtlTab = Val(GetNoXString(varValue, 3, Chr(1)))
            pSetCtlTab oCtl, lngCtlTab
            
        End If
    Next
    
End Sub

'获得控件名称
Private Function pGetCtlName(Ctl As Object) As String
    On Error GoTo RunErr
    Dim strCtlName As String
    Dim lngCtlIndex As Long
    
    pGetCtlName = LCase$(Ctl.Name)
    lngCtlIndex = Ctl.Index
    If lngCtlIndex >= 0 Then '控件数组
        pGetCtlName = LCase$(Ctl.Name) & "(" & lngCtlIndex & ")"
    End If
    
RunErr:
    
End Function

'获得控件对象
Private Function pGetCtlObj(ByVal strCtlName As String) As Object
    Dim oCtl As Object
    
    Set pGetCtlObj = Nothing
    For Each oCtl In UserControl.ContainedControls
        If StrComp(strCtlName, pGetCtlName(oCtl), vbTextCompare) = 0 Then
            Set pGetCtlObj = oCtl
            Exit For
        End If
    Next
    
End Function

'获取控件位置
Private Function pGetCtlPos(oCtl As Object, Optional ByRef strCtlPos As String = "") As Long
    On Error GoTo RunErr
    Dim lngLeft As Long
    Dim lngRight As Long
    
    If TypeOf oCtl Is Line Then
        lngLeft = oCtl.X1
        lngRight = oCtl.X2
        strCtlPos = lngLeft & "-" & lngRight
    Else
        lngLeft = oCtl.Left
        strCtlPos = lngLeft
    End If
    
    pGetCtlPos = lngLeft
    
    Exit Function
RunErr:
    
    Debug.Assert False
    
End Function

'设置控件位置
Private Sub pSetCtlPos(oCtl As Object, ByVal strCtlPos As String)
    On Error GoTo RunErr
    Dim lngLeft As Long
    Dim lngRight As Long
    
    If TypeOf oCtl Is Line Then
        lngLeft = Val(GetNoXString(strCtlPos, 1, "-"))
        lngRight = Val(GetNoXString(strCtlPos, 2, "-"))
        oCtl.X1 = lngLeft
        oCtl.X2 = lngRight
    Else
        lngLeft = Val(strCtlPos)
        oCtl.Left = lngLeft
    End If
    
    Exit Sub
RunErr:
    
    Debug.Assert False
    
End Sub

'设置控件是否可见
Private Sub pSetCtlVisible(oCtl As Object, ByVal blnVisible As Boolean)
    On Error GoTo RunErr
    Dim lngLeft As Long
    
    If blnVisible Then
        '向右平移 75000 距离 显示
        lngLeft = -cMaxPos
    Else
        '向左平移 75000 距离 隐藏
        lngLeft = cMaxPos
    End If
    
    If TypeOf oCtl Is Line Then
        oCtl.X1 = oCtl.X1 + lngLeft
        oCtl.X2 = oCtl.X2 + lngLeft
    Else
        oCtl.Left = oCtl.Left + lngLeft
    End If
    
    Exit Sub
RunErr:
    
    Debug.Assert False
    
End Sub

'获取控件是否支持Tab键
Private Function pGetCtlTab(oCtl As Object) As Long
    On Error GoTo RunErr
    Dim lngEnabled As Long
     
    lngEnabled = IIf(oCtl.TabStop, 1, 0)
    pGetCtlTab = lngEnabled
    
    Exit Function
RunErr:
    
    'Debug.Assert False
    Debug.Print TypeName(oCtl) & "不支持TabStop"
    
End Function

'设置控件是否支持Tab键
Private Sub pSetCtlTab(oCtl As Object, ByVal lngEnabled As Long)
    On Error GoTo RunErr
    
    oCtl.TabStop = lngEnabled
    
    Exit Sub
RunErr:
    
    'Debug.Assert False
    Debug.Print TypeName(oCtl) & "不支持TabStop"
    
End Sub

'切换Tab
Private Function pSwitchTab(ByVal x As Long, ByVal y As Long, Optional ByVal blnCalcOnly As Boolean) As Long
    Dim i As Long
    
    pSwitchTab = -1
    
    If y < mdblTabHeight Then
        For i = 0 To Tabs - 1
            If x >= mTabs(i).Area.Left And x <= mTabs(i).Area.Right And mTabs(i).Enabled Then
                pSwitchTab = i
                If Not blnCalcOnly Then
                    TabIndex = i
                End If
                Exit For
            End If
        Next
    End If
    
End Function

Private Sub Draw()
    Dim tR As RECT
    
    '清除AutoRedraw设置为True时创建的持久图形，避免调整窗口尺寸时GDI Bitmap泄漏
    DeleteObject UserControl.Image.Handle
    
    CalcTabRects
    'CopyContainerBG ContainerHwnd, hwnd, hDC
    GetClientRect hwnd, tR
    DrawHead tR
    DrawInactiveTabs
    DrawBody tR
    DrawActiveTab
    
    '刷新并显示
    '当设置UserControl.AutoRedraw = True时，必须调用UserControl.Refresh，将结果从内存DC输出到显示DC
    UserControl.Refresh
    
End Sub

'计算Tab区域
Private Sub CalcTabRects()
    Dim ImgSize As Long
    Dim i As Long
    
    For i = 0 To Tabs - 1
        With mTabs(i)
            .Area.Bottom = mdblTabHeight
            ImgSize = GetTabImageSize(i)
            If i = 0 Then
                .Area.Left = 3
            Else
                .Area.Left = mTabs(i - 1).Area.Right
            End If
            .Area.Right = .Area.Left + GetTabTextWidth(i) + 2 * mlngTabInnerSpace + ImgSize + 2
            If Not .Visible Then
                .Area.Right = .Area.Left
            End If
        End With
    Next
    
End Sub

Private Sub DrawHead(tR As RECT)
    Dim ttR As TArea
 
    LSet ttR = tR
    ttR.Bottom = mdblTabHeight
    
    Line (ttR.Left, ttR.Top)-(ttR.Right, ttR.Bottom), mclrTabBackColor, BF
    
End Sub

Private Sub DrawBody(tR As RECT)
    Dim ttR As TArea
    Dim clrFrameColor As OLE_COLOR
    
    LSet ttR = tR
    ttR.Top = mdblTabHeight

    '画背景
    Line (ttR.Left + 1, ttR.Top + 1)-(ttR.Right - 1, ttR.Bottom - 1), mclrBackColor, BF
    
    '画边框
    clrFrameColor = IIf(mblnFocused, mclrBorderColor, mclrTabBorderColor)
    Line (ttR.Left, ttR.Top)-(ttR.Left, ttR.Bottom - 1), clrFrameColor
    Line (ttR.Left, ttR.Bottom - 1)-(ttR.Right - 1, ttR.Bottom - 1), clrFrameColor
    Line (ttR.Right - 1, ttR.Bottom - 1)-(ttR.Right - 1, ttR.Top), clrFrameColor
    
End Sub

'画非活动Tab
Private Sub DrawInactiveTabs()
    Dim tR As RECT
    Dim eState As TabStateConstants
    Dim i As Long
    
    For i = 0 To Tabs - 1
        If i <> mlngTabIndex And mTabs(i).Visible Then
            LSet tR = mTabs(i).Area
            tR.Top = tR.Top + 3 'topOffset
            
            If mTabs(i).Hovered Then
                eState = Hot
            Else
                eState = Normal
            End If
            If mTabs(i).Enabled = False Then
                eState = Disabled
            End If

            '画非活动背景
            Line (tR.Left + 1, tR.Top + 1)-(tR.Right - 1, tR.Bottom - 1), mclrTabInActiveColor, BF
            
            '画非活动区域
            If eState = Hot Then
                'mclsGdi.DrawGradientAlpha hDC, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, vbRed, mclrBackColor, Top_To_Bottom, 30
                mclsGdi.DrawArea hDC, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, mclrTabActiveColor
            End If
            
            '画边框
            Line (tR.Left, tR.Bottom)-(tR.Left, tR.Top), mclrTabBorderColor
            Line (tR.Left, tR.Top)-(tR.Right, tR.Top), mclrTabBorderColor
            If i = Tabs - 1 Then
                Line (tR.Right - 1, tR.Top)-(tR.Right - 1, tR.Bottom), mclrTabBorderColor
            End If
            
            '画图片和文本
            DrawPictureAndText i, IIf(eState = Disabled, mclrDisabledColor, mclrForeColor)
            
        End If
    Next
    
End Sub

'画活动Tab
Private Sub DrawActiveTab()
    Dim tR As RECT
    Dim clrFrameColor As OLE_COLOR
    
    LSet tR = mTabs(mlngTabIndex).Area
    tR.Left = tR.Left - 3
    tR.Right = tR.Right + 1
    tR.Bottom = tR.Bottom + 1

    '画活动背景
    Line (tR.Left + 1, tR.Top + 1)-(tR.Right - 1, tR.Bottom - 1), mclrBackColor, BF
    
    '画边框
    clrFrameColor = IIf(mblnFocused, mclrBorderColor, mclrTabBorderColor)
    Line (tR.Left, tR.Bottom - 1)-(tR.Left, tR.Top), clrFrameColor
    Line (tR.Left, tR.Top)-(tR.Right - 1, tR.Top), clrFrameColor
    Line (tR.Right - 1, tR.Top)-(tR.Right - 1, tR.Bottom - 1), clrFrameColor
    Line (0, tR.Bottom - 1)-(tR.Left, tR.Bottom - 1), clrFrameColor
    Line (tR.Right - 1, tR.Bottom - 1)-(UserControl.ScaleWidth, tR.Bottom - 1), clrFrameColor
    
    '画图片和文本
    DrawPictureAndText mlngTabIndex, mclrForeColor
    
End Sub

'画图片和文本
Private Sub DrawPictureAndText(ByVal lngTabIndex As Long, ByVal TextColor As Long)
    Dim tR As RECT
    Dim lngSize As Long

    tR = GetTabTextRect(lngTabIndex)
    lngSize = GetTabImageSize(lngTabIndex)
    With mTabs(lngTabIndex)
        
        '画图片
        If Not .Picture Is Nothing Then
            mclsGdi.DrawImage hDC, .Picture, tR.Left - lngSize - 3, tR.Top, lngSize, lngSize, ilStretch
        End If
        
        '画文本
        If mblnWordWrap Then
            mclsGdi.DrawText hDC, .Caption, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, vbButtonText, dtWordBreak
        Else
            mclsGdi.DrawText hDC, .Caption, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, vbButtonText, dtSingleLine Or dtWordEllipsis
        End If
        
    End With
    
End Sub

'获取Tab图片尺寸
Private Function GetTabImageSize(ByVal lngTabIndex As Long) As Long
    
    If mTabs(lngTabIndex).Picture Is Nothing Then
        Exit Function
    End If
    
    Dim tR As RECT
    Dim lngSize As Long

    LSet tR = mTabs(lngTabIndex).Area
    If lngTabIndex <> mlngTabIndex Then
        tR.Top = tR.Top + 3 'inactive Tabs have a slightly lowered topOffset
    Else
        tR.Bottom = tR.Bottom + 1 'active Tabs have their Bottom-lines shifted 1 pixel downwards
    End If
    
    lngSize = (tR.Bottom - tR.Top) * 0.8 - 1
    If lngSize < 0 Then
        lngSize = 0
    End If
    
    If lngSize > 1200 / Screen.TwipsPerPixelY Then
        lngSize = 1200 / Screen.TwipsPerPixelY
    End If
    
    GetTabImageSize = lngSize
    
End Function

'获取Tab文本区域
Private Function GetTabTextRect(ByVal lngTabIndex As Long) As RECT
    Dim tR As RECT
    Dim ImgSize As Long
    
    LSet tR = mTabs(lngTabIndex).Area
    ImgSize = GetTabImageSize(lngTabIndex)
    tR.Left = tR.Left + mlngTabInnerSpace + ImgSize + IIf(ImgSize, 3, 2) + IIf(lngTabIndex = mlngTabIndex, -3, -1)
    tR.Top = tR.Top + mlngTabInnerSpace - 1 + IIf(lngTabIndex = mlngTabIndex, 0, 3)
    tR.Right = tR.Right - mlngTabInnerSpace
    GetTabTextRect = tR
    
End Function

'获取Tab文本宽度
Private Function GetTabTextWidth(ByVal lngTabIndex As Long) As Long
    Dim tR As RECT
  
    tR = GetTabTextRect(lngTabIndex)
    If mblnWordWrap Then
        If mdblTabMaxWidth = 0 Then
            mdblTabMaxWidth = 1800 / Screen.TwipsPerPixelX
        End If
        tR.Right = tR.Left + mdblTabMaxWidth
    End If
    GetTabTextWidth = mclsGdi.TextWidth(hDC, mTabs(lngTabIndex).Caption)
    If mdblTabMaxWidth > 0 Then
        GetTabTextWidth = mdblTabMaxWidth
    End If
    
End Function
