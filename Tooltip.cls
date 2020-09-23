VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

''Windows API Functions
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

''Windows API Constants
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1

''Windows API Types
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

''Tooltip Window Constants
Private Const TTS_NOPREFIX = &H2
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

''Tooltip Window Types
Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    lHwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type


'local variable(s) to hold property value(s)
Private mvarBackColor As Long 'local copy
Private mvarTitle As String 'local copy
Private mvarForeColor As Long 'local copy
Private mvarParentControl As Object 'local copy
Private mvarIcon As ttIconType 'local copy
Private mvarCentered As Boolean 'local copy
Private mvarStyle As ttStyleEnum 'local copy

Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum

Public Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum

'private data
Private lHwnd As Long
Private ti As TOOLINFO
Public Property Let Style(ByVal vData As ttStyleEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Style = 5
    mvarStyle = vData
End Property


Public Property Get Style() As ttStyleEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Style
    Set Style = mvarStyle
End Property




Public Property Let Centered(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Centered = 5
    mvarCentered = vData
End Property


Public Property Get Centered() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Centered
    Centered = mvarCentered
End Property



Public Function Create() As Boolean
    Dim lpRect As RECT
    Dim lWinStyle As Long
    
    If lHwnd <> 0 Then
        DestroyWindow lHwnd
    End If
    
    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    
    ''create baloon style if desired
    If mvarStyle = TTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON
    
    ''the parent control has to have been set first
    If Not mvarParentControl Is Nothing Then
        lHwnd = CreateWindowEx(0&, _
                    TOOLTIPS_CLASSA, _
                    vbNullString, _
                    lWinStyle, _
                    CW_USEDEFAULT, _
                    CW_USEDEFAULT, _
                    CW_USEDEFAULT, _
                    CW_USEDEFAULT, _
                    mvarParentControl.hwnd, _
                    0&, _
                    App.hInstance, _
                    0&)
                    
        ''make our tooltip window a topmost window
        SetWindowPos lHwnd, _
            HWND_TOPMOST, _
            0&, _
            0&, _
            0&, _
            0&, _
            SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
                    
        ''get the rect of the parent control
        GetClientRect mvarParentControl.hwnd, lpRect
        
        ''now set our tooltip info structure
        With ti
            ''if we want it centered, then set that flag
            If mvarCentered Then
                .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
            Else
                .lFlags = TTF_SUBCLASS
            End If
            
            ''set the hwnd prop to our parent control's hwnd
            .lHwnd = mvarParentControl.hwnd
            .lId = 0
            .hInstance = App.hInstance
            '.lpstr = ALREADY SET
            .lpRect = lpRect
        End With
        
        ''add the tooltip structure
        SendMessage lHwnd, TTM_ADDTOOLA, 0&, ti
        
        ''if we want a title or we want an icon
        If mvarTitle <> vbNullString Or mvarIcon <> TTNoIcon Then
            SendMessage lHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
        End If
        
        If mvarForeColor <> Empty Then
            SendMessage lHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
        End If
        
        If mvarBackColor <> Empty Then
            SendMessage lHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
        End If
        
    End If
End Function

Public Property Set ParentControl(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ParentControl = Form1
    Set mvarParentControl = vData
End Property


Public Property Get ParentControl() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ParentControl
    Set ParentControl = mvarParentControl
End Property




Public Property Let Icon(ByVal vData As ttIconType)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Icon = 5
    mvarIcon = vData
    If lHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
        SendMessage lHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
    End If

End Property


Public Property Get Icon() As ttIconType
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Icon
    Set Icon = mvarIcon
End Property
Public Property Let ForeColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ForeColor = 5
    mvarForeColor = vData
    If lHwnd <> 0 Then
        SendMessage lHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
    End If

End Property


Public Property Get ForeColor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ForeColor
    ForeColor = mvarForeColor
End Property



Public Property Let Title(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarTitle = vData
    If lHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
        SendMessage lHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
    End If
End Property


Public Property Get Title() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Title
    Title = ti.lpStr
End Property



Public Property Let BackColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BackColor = 5
    mvarBackColor = vData
    If lHwnd <> 0 Then
        SendMessage lHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
    End If
    
End Property


Public Property Get BackColor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BackColor
    BackColor = mvarBackColor
End Property



Public Property Let TipText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TipText = 5
    ti.lpStr = vData
    If lHwnd <> 0 Then
        SendMessage lHwnd, TTM_UPDATETIPTEXTA, 0&, ti
    End If
    
End Property


Public Property Get TipText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TipText
    TipText = mvarTipText
End Property



Private Sub Class_Terminate()
    If lHwnd <> 0 Then
        DestroyWindow lHwnd
    End If
    
End Sub


