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
'@Main Func:    ��ҳTab
'@Author:       denglf
'@Last Modify:  2018-09-03
'@Notes:        ��� Microsoft SSTab
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
Private mlngTabIndex As Long                'Tab����
Private mdblTabHeight As Single             'Tab�߶�
Private mdblTabMaxWidth As Long             'Tab�����
Private mlngTabInnerSpace As Long           'Tab���
Private mblnFocused As Boolean              'Tab�ؼ��Ƿ��ý���

Private mclrForeColor As OLE_COLOR          'ǰ��ɫ
Private mclrBackColor As OLE_COLOR          '����ɫ
Private mclrDisabledColor As OLE_COLOR      '������ɫ
Private mclrBorderColor As OLE_COLOR        '�߿�ɫ
Private mclrTabActiveColor As OLE_COLOR     'Tab�ɫ
Private mclrTabInActiveColor As OLE_COLOR   'Tab�ǻɫ
Private mclrTabBorderColor As OLE_COLOR     'Tab�߿�ɫ
Private mclrTabBackColor As OLE_COLOR       'Tab����ɫ

Private mblnEnabled As Boolean              '�Ƿ����
Private mblnWordWrap As Boolean             '�Ƿ��Զ�����

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

'���ھ��
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'�Ƿ����
Public Property Get Enabled() As Boolean
    Enabled = mblnEnabled
End Property
Public Property Let Enabled(ByVal blnValue As Boolean)
    mblnEnabled = blnValue
    UserControl.Enabled = blnValue
    Draw
    PropertyChanged "Enabled"
End Property

'�Զ�����
Public Property Get WordWrap() As Boolean
    WordWrap = mblnWordWrap
End Property
Public Property Let WordWrap(ByVal blnValue As Boolean)
    mblnWordWrap = blnValue
    Draw
    PropertyChanged "WordWrap"
End Property

'ǰ��ɫ
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mclrForeColor
End Property
Public Property Let ForeColor(ByVal clrValue As OLE_COLOR)
    mclrForeColor = clrValue
    Draw
    PropertyChanged "ForeColor"
End Property

'����ɫ
Public Property Get BackColor() As OLE_COLOR
    BackColor = mclrBackColor
End Property
Public Property Let BackColor(ByVal clrValue As OLE_COLOR)
    mclrBackColor = clrValue
    Refresh
    PropertyChanged "BackColor"
End Property

'������ɫ
Public Property Get DisabledColor() As OLE_COLOR
    DisabledColor = mclrDisabledColor
End Property
Public Property Let DisabledColor(ByVal clrValue As OLE_COLOR)
    mclrDisabledColor = clrValue
    Refresh
    PropertyChanged "DisabledColor"
End Property

'�߿�ɫ
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mclrBorderColor
End Property
Public Property Let BorderColor(ByVal clrValue As OLE_COLOR)
    mclrBorderColor = clrValue
    Refresh
    PropertyChanged "BorderColor"
End Property

'Tab�ɫ
Public Property Get TabActiveColor() As OLE_COLOR
    TabActiveColor = mclrTabActiveColor
End Property
Public Property Let TabActiveColor(ByVal clrValue As OLE_COLOR)
    mclrTabActiveColor = clrValue
    Refresh
    PropertyChanged "TabActiveColor"
End Property

'Tab�ǻɫ
Public Property Get TabInActiveColor() As OLE_COLOR
    TabInActiveColor = mclrTabInActiveColor
End Property
Public Property Let TabInActiveColor(ByVal clrValue As OLE_COLOR)
    mclrTabInActiveColor = clrValue
    Refresh
    PropertyChanged "TabInActiveColor"
End Property

'Tab�߿�ɫ
Public Property Get TabBorderColor() As OLE_COLOR
    TabBorderColor = mclrTabBorderColor
End Property
Public Property Let TabBorderColor(ByVal clrValue As OLE_COLOR)
    mclrTabBorderColor = clrValue
    Refresh
    PropertyChanged "TabBorderColor"
End Property

'Tab����ɫ
Public Property Get TabBackColor() As OLE_COLOR
    TabBackColor = mclrTabBackColor
End Property
Public Property Let TabBackColor(ByVal clrValue As OLE_COLOR)
    mclrTabBackColor = clrValue
    Refresh
    PropertyChanged "TabBackColor"
End Property

'Tab����
Public Property Get Tabs() As Long
    Tabs = UBound(mTabs) + 1
End Property
Public Property Let Tabs(ByVal lngValue As Long)

    '���64��Tab
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

'Tab����
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

'Tab����
Public Property Get Caption() As String
    Caption = mTabs(mlngTabIndex).Caption
End Property
Public Property Let Caption(ByVal strValue As String)
    TabCaption(mlngTabIndex) = strValue
End Property

'ָ��Tab����
Public Property Get TabCaption(ByVal lngTabIndex As Long) As String
    TabCaption = mTabs(lngTabIndex).Caption
End Property
Public Property Let TabCaption(ByVal lngTabIndex As Long, ByVal strValue As String)
    
    mTabs(lngTabIndex).Caption = strValue
    Draw
    PropertyChanged "TabCaption(" & lngTabIndex & ")"
    
End Property

''Tabͼ��Key
'Public Property Get ImageKey() As String
'    ImageKey = mTabs(mlngTabIndex).ImageKey
'End Property
'Public Property Let ImageKey(ByVal strValue As String)
'    TabImageKey(mlngTabIndex) = strValue
'End Property
'
''ָ��Tabͼ��Key
'Public Property Get TabImageKey(ByVal lngTabIndex As Long) As String
'    TabImageKey = mTabs(lngTabIndex).ImageKey
'End Property
'Public Property Let TabImageKey(ByVal lngTabIndex As Long, ByVal strValue As String)
'    mTabs(lngTabIndex).ImageKey = strValue
'    Draw
'    PropertyChanged "TabImageKey(" & lngTabIndex & ")"
'End Property

'Tabͼ��
Public Property Get Picture() As StdPicture
    Set Picture = mTabs(mlngTabIndex).Picture
End Property
Public Property Set Picture(ByVal picValue As StdPicture)
    Set TabPicture(mlngTabIndex) = picValue
End Property

'ָ��Tabͼ��
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

'Tab�߶�
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

'Tab���
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

'Tab�Ƿ����
Public Property Get TabEnabled(ByVal lngTabIndex As Long) As Boolean
    TabEnabled = mTabs(lngTabIndex).Enabled
End Property
Public Property Let TabEnabled(ByVal lngTabIndex As Long, ByVal blnValue As Boolean)
    mTabs(lngTabIndex).Enabled = blnValue
    Draw
End Property

'Tab�Ƿ�ɼ�
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

'ˢ��
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

    '��ʾ�ؼ�
    If Ambient.UserMode Then
        pShowControl mlngTabIndex
    End If
    
    '���ģʽ����Ҫ���ദ��Ŀ���ǵ������������л�ѡ�
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

'��Tab��Ϣ
Private Sub pReadTabs(PropBag As PropertyBag)
    Dim strItem As String
    Dim strValue As String
    Dim strCtlName As String
    Dim lngCount As Long
    Dim i As Long
    Dim j As Long
    
    For i = 0 To Tabs - 1
        
        '����
        mTabs(i).Caption = PropBag.ReadProperty("TabCaption(" & i & ")", "Tab " & i)
        
        'ͼƬ
        Set mTabs(i).Picture = PropBag.ReadProperty("TabPicture(" & i & ")", Nothing)
        If Not mTabs(i).Picture Is Nothing Then
            If mTabs(i).Picture.Handle = 0 Then
                Set mTabs(i).Picture = Nothing
            End If
        End If
        
        '�ؼ�
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

'дTab��Ϣ
Private Sub pWriteTabs(PropBag As PropertyBag)
    Dim oCtl As Object
    Dim strItem As String
    Dim strValue As String
    Dim i As Long
    Dim j As Long
    
    '��ƻ���ʱ������Tab�ؼ������⵱ǰTab����ɾ�ؼ���û���л�Tab���¿ؼ�û�б���
    pSaveControl mlngTabIndex, False
    
    '��������ؼ�
    pClearControl
    
    '����ؼ���Ϣ
    For i = 0 To Tabs - 1
        
        '����
        PropBag.WriteProperty "TabCaption(" & i & ")", mTabs(i).Caption, "Tab " & i
        
        'ͼƬ
        PropBag.WriteProperty "TabPicture(" & i & ")", mTabs(i).Picture, Nothing
        
        '�ؼ�����
        PropBag.WriteProperty "Tab(" & i & ").ControlCount", mTabs(i).Controls.Count
        
        '�ؼ�״̬
        For j = 1 To mTabs(i).Controls.Count
            strItem = "Tab(" & i & ").Control(" & j & ")"
            strValue = LCase$(mTabs(i).Controls.Item(j))
            PropBag.WriteProperty strItem, strValue, ""
        Next
        
    Next
    
End Sub

'��������ؼ�
Private Sub pClearControl()
    Dim colCtls As Collection
    Dim oCtl As Object
    Dim strValue As String
    Dim strCtlName As String
    Dim blnExist As Boolean
    Dim i As Long
    Dim j As Long
 
    If Not Ambient.UserMode Then
        '��ƻ���
        
        '���Ҳ����ڵĿؼ�
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
        
        '����
        For i = 1 To colCtls.Count
            strCtlName = LCase$(colCtls.Item(i))
            For j = 0 To Tabs - 1
                RemoveObjectFromCol strCtlName, mTabs(j).Controls
            Next
        Next
        Set colCtls = Nothing
        
    End If
    
End Sub

'�洢�ؼ�
'lngPrevTabIndex�� ��һ��Tab������
'�㷨�� ���л�Tabʱ��������һTab�пؼ���״̬
Private Sub pSaveControl(ByVal lngPrevTabIndex As Long, Optional ByVal blnInVisible As Boolean = True)
    Dim oCtl As Object
    Dim varValue As Variant
    Dim strValue As String
    Dim strCtlName As String
    Dim strCtlPos As String
    Dim lngCtlPos As Long
    Dim lngCtlTab As Long

    If Not Ambient.UserMode Then
        '��ƻ���
        
        '������һTab�еĿؼ�
        For Each oCtl In UserControl.ContainedControls
            strCtlName = pGetCtlName(oCtl)
            lngCtlPos = pGetCtlPos(oCtl, strCtlPos)
            lngCtlTab = pGetCtlTab(oCtl)
            
            If lngCtlPos > cMinPos Then
                strValue = strCtlName & Chr(1) & lngCtlPos & Chr(1) & lngCtlTab
                If ObjectInCol(strCtlName, mTabs(lngPrevTabIndex).Controls) Then
                    '�޸�
                    RemoveObjectFromCol strCtlName, mTabs(lngPrevTabIndex).Controls
                    mTabs(lngPrevTabIndex).Controls.Add strValue, strCtlName
                Else
                    '����
                    mTabs(lngPrevTabIndex).Controls.Add strValue, strCtlName
                End If
                
                '�ÿؼ����ɼ�
                If blnInVisible Then
                    pSetCtlVisible oCtl, False
                    pSetCtlTab oCtl, False
                End If
                
            End If
        Next
    
    Else
        '���л���
        
        '�ÿؼ����ɼ�,���ڰ���Tab��ʱ,��һTab�еĿؼ����ܻ�ý���
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

'��ʾ�ؼ�
'lngTabIndex�� ��ǰTab������
Private Sub pShowControl(ByVal lngTabIndex As Long)
    'On Error Resume Next
    Dim oCtl As Object
    Dim varValue As Variant
    Dim strCtlName As String
    Dim strCtlPos As String
    Dim lngCtlTab As Long
    
    '��ʾ��ǰTab�еĿؼ�
    For Each varValue In mTabs(lngTabIndex).Controls
        strCtlName = GetNoXString(varValue, 1, Chr(1))
        Set oCtl = pGetCtlObj(strCtlName)
        If ObjPtr(oCtl) > 0 Then
            
            '�ָ��ؼ���λ��
            strCtlPos = GetNoXString(varValue, 2, Chr(1))
            pSetCtlPos oCtl, strCtlPos
        
            '�ָ��ؼ��Ŀ���״̬
            lngCtlTab = Val(GetNoXString(varValue, 3, Chr(1)))
            pSetCtlTab oCtl, lngCtlTab
            
        End If
    Next
    
End Sub

'��ÿؼ�����
Private Function pGetCtlName(Ctl As Object) As String
    On Error GoTo RunErr
    Dim strCtlName As String
    Dim lngCtlIndex As Long
    
    pGetCtlName = LCase$(Ctl.Name)
    lngCtlIndex = Ctl.Index
    If lngCtlIndex >= 0 Then '�ؼ�����
        pGetCtlName = LCase$(Ctl.Name) & "(" & lngCtlIndex & ")"
    End If
    
RunErr:
    
End Function

'��ÿؼ�����
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

'��ȡ�ؼ�λ��
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

'���ÿؼ�λ��
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

'���ÿؼ��Ƿ�ɼ�
Private Sub pSetCtlVisible(oCtl As Object, ByVal blnVisible As Boolean)
    On Error GoTo RunErr
    Dim lngLeft As Long
    
    If blnVisible Then
        '����ƽ�� 75000 ���� ��ʾ
        lngLeft = -cMaxPos
    Else
        '����ƽ�� 75000 ���� ����
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

'��ȡ�ؼ��Ƿ�֧��Tab��
Private Function pGetCtlTab(oCtl As Object) As Long
    On Error GoTo RunErr
    Dim lngEnabled As Long
     
    lngEnabled = IIf(oCtl.TabStop, 1, 0)
    pGetCtlTab = lngEnabled
    
    Exit Function
RunErr:
    
    'Debug.Assert False
    Debug.Print TypeName(oCtl) & "��֧��TabStop"
    
End Function

'���ÿؼ��Ƿ�֧��Tab��
Private Sub pSetCtlTab(oCtl As Object, ByVal lngEnabled As Long)
    On Error GoTo RunErr
    
    oCtl.TabStop = lngEnabled
    
    Exit Sub
RunErr:
    
    'Debug.Assert False
    Debug.Print TypeName(oCtl) & "��֧��TabStop"
    
End Sub

'�л�Tab
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
    
    '���AutoRedraw����ΪTrueʱ�����ĳ־�ͼ�Σ�����������ڳߴ�ʱGDI Bitmapй©
    DeleteObject UserControl.Image.Handle
    
    CalcTabRects
    'CopyContainerBG ContainerHwnd, hwnd, hDC
    GetClientRect hwnd, tR
    DrawHead tR
    DrawInactiveTabs
    DrawBody tR
    DrawActiveTab
    
    'ˢ�²���ʾ
    '������UserControl.AutoRedraw = Trueʱ���������UserControl.Refresh����������ڴ�DC�������ʾDC
    UserControl.Refresh
    
End Sub

'����Tab����
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

    '������
    Line (ttR.Left + 1, ttR.Top + 1)-(ttR.Right - 1, ttR.Bottom - 1), mclrBackColor, BF
    
    '���߿�
    clrFrameColor = IIf(mblnFocused, mclrBorderColor, mclrTabBorderColor)
    Line (ttR.Left, ttR.Top)-(ttR.Left, ttR.Bottom - 1), clrFrameColor
    Line (ttR.Left, ttR.Bottom - 1)-(ttR.Right - 1, ttR.Bottom - 1), clrFrameColor
    Line (ttR.Right - 1, ttR.Bottom - 1)-(ttR.Right - 1, ttR.Top), clrFrameColor
    
End Sub

'���ǻTab
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

            '���ǻ����
            Line (tR.Left + 1, tR.Top + 1)-(tR.Right - 1, tR.Bottom - 1), mclrTabInActiveColor, BF
            
            '���ǻ����
            If eState = Hot Then
                'mclsGdi.DrawGradientAlpha hDC, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, vbRed, mclrBackColor, Top_To_Bottom, 30
                mclsGdi.DrawArea hDC, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, mclrTabActiveColor
            End If
            
            '���߿�
            Line (tR.Left, tR.Bottom)-(tR.Left, tR.Top), mclrTabBorderColor
            Line (tR.Left, tR.Top)-(tR.Right, tR.Top), mclrTabBorderColor
            If i = Tabs - 1 Then
                Line (tR.Right - 1, tR.Top)-(tR.Right - 1, tR.Bottom), mclrTabBorderColor
            End If
            
            '��ͼƬ���ı�
            DrawPictureAndText i, IIf(eState = Disabled, mclrDisabledColor, mclrForeColor)
            
        End If
    Next
    
End Sub

'���Tab
Private Sub DrawActiveTab()
    Dim tR As RECT
    Dim clrFrameColor As OLE_COLOR
    
    LSet tR = mTabs(mlngTabIndex).Area
    tR.Left = tR.Left - 3
    tR.Right = tR.Right + 1
    tR.Bottom = tR.Bottom + 1

    '�������
    Line (tR.Left + 1, tR.Top + 1)-(tR.Right - 1, tR.Bottom - 1), mclrBackColor, BF
    
    '���߿�
    clrFrameColor = IIf(mblnFocused, mclrBorderColor, mclrTabBorderColor)
    Line (tR.Left, tR.Bottom - 1)-(tR.Left, tR.Top), clrFrameColor
    Line (tR.Left, tR.Top)-(tR.Right - 1, tR.Top), clrFrameColor
    Line (tR.Right - 1, tR.Top)-(tR.Right - 1, tR.Bottom - 1), clrFrameColor
    Line (0, tR.Bottom - 1)-(tR.Left, tR.Bottom - 1), clrFrameColor
    Line (tR.Right - 1, tR.Bottom - 1)-(UserControl.ScaleWidth, tR.Bottom - 1), clrFrameColor
    
    '��ͼƬ���ı�
    DrawPictureAndText mlngTabIndex, mclrForeColor
    
End Sub

'��ͼƬ���ı�
Private Sub DrawPictureAndText(ByVal lngTabIndex As Long, ByVal TextColor As Long)
    Dim tR As RECT
    Dim lngSize As Long

    tR = GetTabTextRect(lngTabIndex)
    lngSize = GetTabImageSize(lngTabIndex)
    With mTabs(lngTabIndex)
        
        '��ͼƬ
        If Not .Picture Is Nothing Then
            mclsGdi.DrawImage hDC, .Picture, tR.Left - lngSize - 3, tR.Top, lngSize, lngSize, ilStretch
        End If
        
        '���ı�
        If mblnWordWrap Then
            mclsGdi.DrawText hDC, .Caption, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, vbButtonText, dtWordBreak
        Else
            mclsGdi.DrawText hDC, .Caption, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, vbButtonText, dtSingleLine Or dtWordEllipsis
        End If
        
    End With
    
End Sub

'��ȡTabͼƬ�ߴ�
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

'��ȡTab�ı�����
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

'��ȡTab�ı����
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
