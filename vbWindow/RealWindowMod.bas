Attribute VB_Name = "RealWindowMod"
' This is a module to create a window from scratch in Visual Basic.
' I always wanted to do this ever sence i left VB for C++ 2 years back.
' You can do anything you want to this window like you would in Win32 C++.
' The sky is the limit.

' Author: Andrew Heinlein [Mouse]
' Web: www.mouseindustries.com
' Email: mouse@mouseindustries.com

' WARNING: If your brave and decide to debug this, besure to save your work
' before doing so.  You are now `subclassing` and Visual Basic wasnt meant to subclass.
' Also, if you get the "Failed to register window" message, just change the class
' name in the AppMain.BAS

Option Explicit

Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PAINTSTRUCT
    hdc As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(32) As Byte
End Type

Private Type WNDCLASSEX
    cbSize As Long
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
    hIconSm As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Const CS_HREDRAW = &H2
Private Const CS_VREDRAW = &H1
Private Const CS_PARENTDC = &H80
Private Const WS_OVERLAPPEDWINDOW = &HCF0000
Private Const WS_EX_APPWINDOW = &H40000
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const IDC_ARROW = &H7F00
Private Const COLOR_WINDOW = &H5
Private Const SW_SHOW = &H5
Private Const WM_DESTROY = &H2
Private Const WM_PAINT = &HF
Private Const DT_CENTER = &H1
Private Const CW_USEDEFAULT = &H80000000

Private Function MainWndProc(ByVal hwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    ' i put this here so that you can see how to process a message.
    ' this is the WM_PAINT message where it repaints the window.
    ' lets put "Hello World!" at the top of it like they do on
    ' the win32 C++ pre-made projects
    If message = WM_PAINT Then
        Dim rt As RECT
        Dim hdc As Long
        Dim ps As PAINTSTRUCT
        
        GetClientRect hwnd, rt
        
        hdc = BeginPaint(hwnd, ps)
        DrawText hdc, "Hello World!", Len("Hello World!"), rt, DT_CENTER
        EndPaint hwnd, ps
        
        ' since we handled this message, return 0. dont let the
        ' DefWindowProc handle it
        MainWndProc = 0
        Exit Function
    End If

    ' watch for WM_DESTROY message, if its sent, then let the GetMessage loop in
    ' CreateNewWindow know so it breaks out of the GetMessage loop
    If message = WM_DESTROY Then
        PostQuitMessage 0
        MainWndProc = 0
        Exit Function
    End If
        
    MainWndProc = DefWindowProc(hwnd, message, wParam, lParam)
End Function

Private Function CreateNewWindow(ByVal MyWndProc As Long, ByVal szWindowClass As String, ByVal szWindowTitle As String, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long) As Long
    ' Register a class
    Dim wcex As WNDCLASSEX
    wcex.cbSize = LenB(wcex)
    wcex.style = CS_HREDRAW Or CS_VREDRAW Or CS_PARENTDC
    wcex.lpfnWndProc = MyWndProc
    wcex.cbClsExtra = 0
    wcex.cbWndExtra = 0
    wcex.hInstance = App.hInstance
    wcex.hIcon = 0
    wcex.hCursor = LoadCursor(0, IDC_ARROW)
    wcex.hbrBackground = COLOR_WINDOW + 1
    wcex.lpszMenuName = vbNullString
    wcex.lpszClassName = szWindowClass
    wcex.hIconSm = 0

    If RegisterClassEx(wcex) = 0 Then
        MsgBox "Failed to register window!"
        CreateNewWindow = -1
        Exit Function
    End If
    
    ' create the window
    Dim vbWindow As Long
    vbWindow = CreateWindowEx(WS_EX_APPWINDOW Or WS_EX_WINDOWEDGE, _
                              szWindowClass, _
                              szWindowTitle, _
                              WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or WS_OVERLAPPEDWINDOW, _
                              x, y, cx, cy, 0, 0, App.hInstance, 0)
                              
    If vbWindow = 0 Then
        MsgBox "Failed to create the window!"
        UnregisterClass szWindowClass, App.hInstance
        CreateNewWindow = -1
        Exit Function
    End If
    
    ' show the window
    UpdateWindow vbWindow
    ShowWindow vbWindow, SW_SHOW
    
    ' message loop to process window messages
    Dim myMsg As MSG
    While GetMessage(myMsg, 0, 0, 0) <> 0 ' waiting for PostQuitMessage to be called to break out
        TranslateMessage myMsg
        DispatchMessage myMsg
    Wend
    
    ' done with window.. clean up what we created
    DestroyWindow vbWindow
    UnregisterClass szWindowClass, App.hInstance
    
    ' return exit code
    CreateNewWindow = myMsg.wParam
End Function

Public Function doWindow(ByVal szWindowTitle As String, ByVal szWindowClass As String, Optional ByVal x As Long = CW_USEDEFAULT, Optional ByVal y As Long = CW_USEDEFAULT, Optional ByVal cx As Long = CW_USEDEFAULT, Optional ByVal cy As Long = CW_USEDEFAULT) As Long
    doWindow = CreateNewWindow(AddressOf MainWndProc, szWindowClass, szWindowTitle, x, y, cx, cy)
End Function

