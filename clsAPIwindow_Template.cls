VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAPIwindow_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class is designed to be used to register/unregister
' custom window classes (i.e., ALL - API Windows). And also
' can be used to create windows using CreateWindowEx

' The class is also designed in a way that you can instantiate
' it just to create something or to destroy something, then
' terminate the class so you don't need to keep it alive.
' The only exception is that if you choose to register the
' custom window class assigning this class as its window procedure,
' then you would need to keep this class alive until the
' custom class was unregistered.

' NOTE: MANY OF THESE ROUTINES ARE IDENTICAL OR VERY SIMILAR TO THOSE
' FOR SELF-SUBCLASSING. HOWEVER, DON'T ASSUME THEY ARE THE SAME.
' Some key differences include:
'   scc_AddMsg & scc_DelMsg is by ClassName not hWnd
'   scc_lParamUser property is by ClassName not hWnd
'   zMap_VFunction uses ClassName not hWnd

' This class inspired by Paul Caton's APIWindow ASM project found
' on vbAccelerator. However, the ASM is based off his self-sub ASM
' code not that from vbAccelerator

' Routines overview. See the comments in those routines for more information
' ==========================================================================
' scc_CreateWindow - creates a window based off a passed window class name
' scc_RegClassProc - registers a new window class and assigns a window procedure for that class
' scc_UnRegisterClassProc - unregisters a window class and frees memory associated with it
' scc_AddMsg - adds a window message to be processed within the class window procedure
' scc_DelMsg - removes a window message from being processed within the procedure
' scc_lParamUser (let/get property) - sets/gets custom parameter assigned when calling scc_RegClassProc
' zALL_API_WndProc - routine prototype to be used for the window procedure

Private Declare Function CreateWindowExA Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function CreateWindowExW Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long

' Following are possible window and class styles
Public Enum eWindowStyle
        WS_OVERLAPPED = &H0
        WS_POPUP = &H80000000
        WS_CHILD = &H40000000
        WS_MINIMIZE = &H20000000
        WS_VISIBLE = &H10000000
        WS_DISABLED = &H8000000
        WS_CLIPSIBLINGS = &H4000000
        WS_CLIPCHILDREN = &H2000000
        WS_MAXIMIZE = &H1000000
        WS_CAPTION = &HC00000
        WS_BORDER = &H800000
        WS_DLGFRAME = &H400000
        WS_VSCROLL = &H200000
        WS_HSCROLL = &H100000
        WS_SYSMENU = &H80000
        WS_THICKFRAME = &H40000
        WS_GROUP = &H20000
        WS_TABSTOP = &H10000
        WS_MINIMIZEBOX = &H20000
        WS_MAXIMIZEBOX = &H10000
        WS_TILED = &H0
        WS_ICONIC = &H20000000
        WS_SIZEBOX = &H40000
        WS_OVERLAPPEDWINDOW = &HCF0000
        WS_POPUPWINDOW = &H80880000
End Enum
Public Enum eWindowStyleEx
        WS_EX_DLGMODALFRAME = &H1
        WS_EX_NOPARENTNOTIFY = &H4
        WS_EX_TOPMOST = &H8
        WS_EX_ACCEPTFILES = &H10
        WS_EX_TRANSPARENT = &H20
        WS_EX_MDICHILD = &H40
        WS_EX_TOOLWINDOW = &H80
        WS_EX_WINDOWEDGE = &H100
        WS_EX_CLIENTEDGE = &H200
        WS_EX_CONTEXTHELP = &H400
        WS_EX_RIGHT = &H1000
        WS_EX_LEFT = &H0
        WS_EX_RTLREADING = &H2000
        WS_EX_LTRREADING = &H0
        WS_EX_LEFTSCROLLBAR = &H4000
        WS_EX_RIGHTSCROLLBAR = &H0
        WS_EX_CONTROLPARENT = &H10000
        WS_EX_STATICEDGE = &H20000
        WS_EX_APPWINDOW = &H40000
        WS_EX_OVERLAPPEDWINDOW = &H300
        WS_EX_PALETTEWINDOW = &H188
        WS_EX_LAYERED = &H80000
End Enum
Public Enum eClassStyle
        CS_VREDRAW = 1
        CS_HREDRAW = 2
        CS_KEYCVTWINDOW = 4
        CS_DBLCLKS = 8
        CS_OWNDC = 32
        CS_CLASSDC = 64
        CS_PARENTDC = 128
        CS_NOKEYCVT = 256
        CS_NOCLOSE = 512
        CS_SAVEBITS = 2048
        CS_BYTEALIGNCLIENT = 4096
        CS_BYTEALIGNWINDOW = 8192
        CS_PUBLICCLASS = 16384
End Enum

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function GetClassInfoExA Lib "user32.dll" (ByVal hInstance As Long, ByVal lpcstr As String, ByRef lpwndclassexa As WNDCLASSEX) As Long
Private Declare Function GetClassInfoExW Lib "user32.dll" (ByVal hInstance As Long, ByVal lpcstr As Long, ByRef lpwndclassexa As WNDCLASSEX) As Long
Private Declare Function RegisterClassExA Lib "user32.dll" (ByRef pcWndClassEx As WNDCLASSEX) As Long
Private Declare Function RegisterClassExW Lib "user32.dll" (ByRef pcWndClassEx As WNDCLASSEX) As Long
Private Declare Function UnregisterClassA Lib "user32.dll" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function UnregisterClassW Lib "user32.dll" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long

Private Type WNDCLASSEX
    cbSize As Long
    Style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String   'not generally used. Add code for it if needed.
    lpszClassName As Long    'pointer
    hIconSm As Long
End Type

    
' ***** THUNK APIs and VARIABLES
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
    
Private z_ScMem             As Long        'Thunk base address

Private Const IDX_CALLBACK   As Long = 7   'Thunk data index of the custom class window proc
Private Const IDX_BTABLE    As Long = 8    'Thunk data index of the Before table for messages
Private Const IDX_ATABLE    As Long = 9    'Thunk data index of the After table for messages
Private Const IDX_PARMUSER  As Long = 10   'Thunk data index of the user-defined parameter
Private Const IDX_EBX       As Long = 13   'Thunk code patch index of the thunk data
Private Const IDX_UNICODE   As Long = 64   'Must be UBound(thunk data)+1, index of Unicode usage
Private Const WNDPROC_OFF   As Long = &H2C 'Offset where window proc starts from z_ScMem
Private Const MSG_ENTRIES   As Long = 32   'Number of msg table entries. Set to 1 if using ALL_MESSAGES for all custom windows

Public Enum sccMsgWhen                            'When to callback
  MSG_BEFORE = 1                                'Callback before the original WndProc
  MSG_AFTER = 2                                 'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER    'Callback before and after the original WndProc
End Enum
Public Enum sccALLMessages
    ALL_MESSAGES = -1                           'All messages will callback
End Enum

Public Function scc_CreateWindow(ByVal sClass As String, _
                        ByVal Style As eWindowStyle, _
                        Optional ByVal StyleEx As eWindowStyleEx, _
                        Optional ByVal sCaption As String = vbNullString, _
                        Optional ByVal Width As Long = 0, _
                        Optional ByVal Height As Long = 0, _
                        Optional ByVal Left As Long = 0, _
                        Optional ByVal Top As Long = 0, _
                        Optional ByVal ParentHWnd As Long = 0, _
                        Optional ByVal MenuHandle As Long = 0, _
                        Optional ByRef bUnicode As Boolean) As Long
    
    ' Function added as a courtesy. Contains everything you need to create an API window
    '*************************************************************************************************
    '* sClass - The class to create the window from
    '* Style - a combination of valid window styles.
    '* StyleEx - a combination of valid extended window styles
    '* sCaption - the window caption
    '* Width - the window's width (pixels)
    '* Height - the window's height (pixels)
    '* Left - the left position of the window in screen or client coords (pixels)
    '* Top - the top position of the window in screen or client coords (pixels)
    '* ParentHWnd - handle to the window's parent if window is a child window
    '* MenuHandle - menubar handle for the window
    '* bUnicode - Optional, if True, Unicode API calls should be made to the window vs ANSI calls
    '*            Parameter is byRef and its return value should be checked to know if ANSI to be used or not
    '*************************************************************************************************
                        
    If bUnicode Then bUnicode = (IsWindowUnicode(GetDesktopWindow) <> 0)
    If bUnicode Then
        scc_CreateWindow = CreateWindowExW(StyleEx, StrPtr(sClass), StrPtr(sCaption), Style, _
                Left, Top, Width, Height, ParentHWnd, MenuHandle, App.hInstance, ByVal 0&)
    Else
        scc_CreateWindow = CreateWindowExA(StyleEx, sClass, sCaption, Style, _
                Left, Top, Width, Height, ParentHWnd, MenuHandle, App.hInstance, ByVal 0&)
    End If
                        
End Function

Public Function scc_RegClassProc(ByVal sClass As String, _
                    Optional ByVal BackColor As Long = vbButtonFace, _
                    Optional ByVal Style As eClassStyle = 0, _
                    Optional ByVal hCursor As Long = 0, _
                    Optional ByVal hIcon As Long = 0, _
                    Optional ByVal hIconSm As Long = 0, _
                    Optional ByVal cbClassExtra As Long = 0, _
                    Optional ByVal cbWndExtra As Long = 0, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True, _
                    Optional ByRef bUnicode As Boolean = False) As Boolean

    '*************************************************************************************************
    '* Following are class creation properties
    '* sClass - A unicode or non-unicode class name. If unicode, and unicode class creation
    '           is not used, string is converted from unicode. Vice versa applies too
    '* BackColor - either system color or RGB color for the default window bkg color of the class
    '* Style - a combination of valid class styles.
    '* hCursor - a default cursor to use when in client area: Use LoadCursor and pass handle
    '* hIcon - a 32x32 icon to use. Use LoadIcon and pass handle
    '* hIconSm - a 16x16 icon to use (system menu). Use LoadIcon and pass handle
    '* cbClassExtra - Set/Retrieved via GetClassLong API
    '* cbWndExtra - Set/Retrieved via GetWindowLong API
    '
    '* Following are used to regsiter the class and create its window procedure
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures
    '* bUnicode - Optional, if True, Unicode API calls should be made to the class vs ANSI calls
    '*            Parameter is byRef and its return value should be checked to know whether ANSI used or not
    '*************************************************************************************************
    ' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

    Dim z_Sc(0 To IDX_UNICODE) As Long                 'Thunk machine-code initialised here
    Dim CODE_LEN      As Long                          'Thunk length in bytes
    Dim MEM_LEN       As Long                          'Bytes to allocate per thunk, data + code + msg tables
    
    Const PAGE_RWX      As Long = &H40&                'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&              'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&              'Release allocated memory flag
    Const IDX_NOPROCEDURE As Long = 1                  'Occurs when IDE stops while class is still active
    Const IDX_EBMODE    As Long = 2                    'Thunk data index of the EbMode function address
    Const IDX_DWP       As Long = 3                    'Thunk data index of the DefWindowProc function address
    Const IDX_DW        As Long = 4                    'Thunk data index of the DestroyWindow function address
    Const IDX_BADPTR    As Long = 5                    'Thunk data index of the IsBadCodePtr function address
    Const IDX_OWNER     As Long = 6                    'Thunk data index of the Owner object's vTable address
    Const SUB_NAME      As String = "scc_Subclass"     'This routine's name
    
    Dim nAddr         As Long
    Dim aClass()      As Byte
    Dim bIDE          As Boolean
    Dim WC As WNDCLASSEX

    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    
    nAddr = zAddressOf(oCallback, nOrdinal)             'Get the address of the specified ordinal method
    If nAddr = 0 Then                                   'Ensure that we've found the ordinal method
      zError SUB_NAME, "Callback method not found"
      Exit Function
    End If
        
    If zMap_VFunction(sClass, True) = 0& Then           ' See if this class is already registered
        zError SUB_NAME, "Class is not one registered by this application. Cannot modify it"
        Exit Function
    ElseIf Not z_ScMem = -1 Then                        ' class exists, try to unregister it first
        If scc_UnRegisterClassProc(sClass) = True Then z_ScMem = -1
    End If
    
    If z_ScMem = -1 Then                                ' class not previously registered
    
        ' Validate unicode request
        If bUnicode Then bUnicode = (IsWindowUnicode(GetDesktopWindow) <> 0)
        ' Add class name to array for use in registration
        If bUnicode Then
            CODE_LEN = Len(sClass) * 2 + 2
            ReDim aClass(0 To CODE_LEN - 1)
            RtlMoveMemory VarPtr(aClass(0)), StrPtr(sClass), CODE_LEN - 2
        Else
            aClass = StrConv(sClass & vbNullChar, vbFromUnicode)
        End If
        
        ' calculate memory size requirements
        CODE_LEN = IDX_UNICODE * 4 + 4
        MEM_LEN = CODE_LEN + (8 * MSG_ENTRIES + 8)
            
        z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
        If z_ScMem <> 0 Then                                     'Ensure the allocation succeeded
        
            If z_ScMem < WNDPROC_OFF Then
                
                zError SUB_NAME, "Negative memory address. Class not designed to handle these. Report this to author"
                GoTo ReleaseMemory
                
            Else
          
              ' actual thunk data
              z_Sc(11) = &HD231C031: z_Sc(12) = &HBBE58960: z_Sc(13) = &H12345678: z_Sc(14) = &H4339F631: z_Sc(15) = &H902D7508: z_Sc(16) = &H207B8B4A: z_Sc(17) = &HE81C7589: z_Sc(18) = &H6C&: z_Sc(19) = &H75147539: z_Sc(20) = &H90909017: z_Sc(21) = &H37E8&: z_Sc(22) = &H90909000: z_Sc(23) = &H7B8BD231: z_Sc(24) = &H52E824: z_Sc(25) = &H90900000: z_Sc(26) = &H10C261: z_Sc(27) = &H74047339: z_Sc(28) = &H1AE807: z_Sc(29) = &HF0EB0000: z_Sc(30) = &H3D0853FF: z_Sc(31) = &H1&: z_Sc(32) = &HC085BE74: z_Sc(33) = &H5E82274: z_Sc(34) = &HEB000000: z_Sc(35) = &H909090DB: z_Sc(36) = &HFF3075FF: z_Sc(37) = &H75FF2C75: z_Sc(38) = &H2475FF28: z_Sc(39) = &H890C53FF: z_Sc(40) = &H55E91C45: z_Sc(41) = &H90000000: z_Sc(42) = &H10443C7: z_Sc(43) = &HFF000000: z_Sc(44) = &H53FF2475: z_Sc(45) = &H90B1EB10: z_Sc(46) = &H40E30F8B: z_Sc(47) = &H1078C985: z_Sc(48) = &H4C781: z_Sc(49) = &H458B0000
              z_Sc(50) = &H75AFF228: z_Sc(51) = &H9090902F: z_Sc(52) = &H1C73FF52: z_Sc(53) = &H5A1453FF: z_Sc(54) = &H438D2275: z_Sc(55) = &H144D8D28: z_Sc(56) = &H1C458D50: z_Sc(57) = &HFF3075FF: z_Sc(58) = &H75FF2C75: z_Sc(59) = &H2475FF28: z_Sc(60) = &HFF525150: z_Sc(61) = &H53FF1873: z_Sc(62) = &H9090901C: z_Sc(63) = &HC3&
              
              z_Sc(IDX_EBX) = z_ScMem                                       'Patch the thunk data address
              z_Sc(IDX_BTABLE) = z_Sc(IDX_EBX) + CODE_LEN                   'Store the address of the before table in the thunk data
              z_Sc(IDX_ATABLE) = z_Sc(IDX_BTABLE) + (MSG_ENTRIES * 4) + 4   'Store the address of the after table in the thunk data
              z_Sc(IDX_OWNER) = ObjPtr(oCallback)                           'Store the callback owner's object address in the thunk data
              z_Sc(IDX_CALLBACK) = nAddr                                    'Store the callback address in the thunk data
              z_Sc(IDX_PARMUSER) = lParamUser                               'Store the lParamUser callback parameter in the thunk data
              z_Sc(IDX_UNICODE) = Abs(bUnicode)                             'Store whether unicode is used or not
              z_Sc(0) = z_Sc(IDX_EBX)                                       'Used to help validate this is our custom class
              
              z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode) 'Store the IsBadCodePtr function address in the thunk data
              z_Sc(IDX_DW) = zFnAddr("user32", "DestroyWindow", bUnicode)   'Store the DestroyWindow function address in the thunk data
              
              If bIdeSafety = True Then                                     'If the user wants IDE protection
                  Debug.Assert zInIDE(bIDE)
                  If bIDE = True Then z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode) 'Store the EbMode function address in the thunk data
              End If
              
              If bUnicode Then                                               'Store the DefWindowProc function address in the thunk data
                  z_Sc(IDX_DWP) = zFnAddr("user32", "DefWindowProcW", bUnicode)
              Else
                  z_Sc(IDX_DWP) = zFnAddr("user32", "DefWindowProcA", bUnicode)
              End If
              RtlMoveMemory z_Sc(IDX_EBX), VarPtr(z_Sc(0)), CODE_LEN         'Copy the thunk code/data to the allocated memory
          
             'fill in the window class structure
             With WC
                 .cbSize = Len(WC)
                 If cbClassExtra > 0 Then .cbClsExtra = cbClassExtra
                 If cbWndExtra > 0 Then .cbWndExtra = cbWndExtra
                 .hCursor = hCursor
                 .hIcon = hIcon
                 .hIconSm = hIconSm
                 .hInstance = App.hInstance
                 .lpszClassName = VarPtr(aClass(0))
                 .Style = Style
                 .lpfnWndProc = z_Sc(IDX_EBX) + WNDPROC_OFF
                  If (BackColor \ &H80FFFFFF) Then
                    ' convert vb system color constant to windows system color constant + 1
                      .hbrBackground = (BackColor And &HFF&) + 1
                  Else
                      .hbrBackground = CreateSolidBrush(BackColor)
                  End If
             End With
             
             'Attempt to register the window
             If bUnicode Then
               If RegisterClassExW(WC) = 0 Then GoTo ReleaseMemory
             Else
               If RegisterClassExA(WC) = 0 Then GoTo ReleaseMemory
            End If
            scc_RegClassProc = True
          End If
          
        Else
          zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
        End If
    
    Else
        ' should this happen. An attempt is being made to register a pre-existing
        ' class which cannot be unregistered (usually means a window created with
        ' that class still exists). We can update two key procedure attributes:
        
        zData(IDX_CALLBACK) = nAddr             'Set/change the window procedure
        zData(IDX_NOPROCEDURE) = 0              'Ensure procedure override is false
        scc_RegClassProc = True
        
    End If
    
Exit Function                                   'Exit
ReleaseMemory:
      VirtualFree z_ScMem, 0, MEM_RELEASE       'scc_Subclass has failed after memory allocation, so release the memory
End Function

Public Function scc_UnRegisterClassProc(ByVal sClass As String) As Boolean
    ' function unregisters a class if possible.
    ' If any windows are currently not destroyed that were created from the passed
    ' class, then unregistering will fail. Regardless, the class is always unregistered
    ' by windows when the app's instance is terminated. Likewise, since the class'
    ' window procedure is virtual memory, virtual memory is also destroyed at that
    ' point too. Bottom line, can't cause memory leaks whether unregistered or not
    If IsBadCodePtr(zMap_VFunction(sClass, False)) = 0 Then
        If zData(IDX_UNICODE) Then
            scc_UnRegisterClassProc = UnregisterClassW(StrPtr(sClass), App.hInstance)
        Else
            scc_UnRegisterClassProc = UnregisterClassA(sClass, App.hInstance)
        End If
        ' if we unregistered the class, release its memory too
        If scc_UnRegisterClassProc = True Then
            Const MEM_RELEASE As Long = &H8000&             'Release allocated memory flag
            VirtualFree z_ScMem, 0, MEM_RELEASE
        End If
    End If
End Function

'Add the message value to the window handle's specified callback table
Public Sub scc_AddMsg(ByVal ClassName As String, ByVal When As sccMsgWhen, ParamArray Messages() As Variant)
    
    If IsBadCodePtr(zMap_VFunction(ClassName, False)) = 0 Then              'Ensure that the thunk hasn't already released its memory
      Dim M As Long
      For M = LBound(Messages) To UBound(Messages)                          ' ensure no strings, arrays, doubles, objects, etc are passed
        Select Case VarType(Messages(M))
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                                     'If the message is to be added to the before original WndProc table...
              zAddMsg Messages(M), IDX_BTABLE                               'Add the message to the before table
            End If
            If When And MSG_AFTER Then                                      'If message is to be added to the after original WndProc table...
              zAddMsg Messages(M), IDX_ATABLE                               'Add the message to the after table
            End If
        End Select
      Next
    End If

End Sub

'Delete the message value from the window handle's specified callback table
Public Sub scc_DelMsg(ByVal ClassName As String, ByVal When As sccMsgWhen, ParamArray Messages() As Variant)
    
    If IsBadCodePtr(zMap_VFunction(ClassName, False)) = 0 Then              'Ensure that the thunk hasn't already released its memory
      Dim M As Long
      For M = LBound(Messages) To UBound(Messages)                          ' ensure no strings, arrays, doubles, objects, etc are passed
        Select Case VarType(Messages(M))
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                                     'If the message is to be added to the before original WndProc table...
              If zAddMsg(Messages(M), IDX_BTABLE) = False Then              'Add the message to the before table
                When = When And Not MSG_BEFORE
              End If
            End If
            If When And MSG_AFTER Then                                      'If message is to be added to the after original WndProc table...
              If zAddMsg(Messages(M), IDX_ATABLE) = False Then              'Add the message to the after table
                When = When And Not MSG_AFTER
              End If
            End If
        End Select
      Next
    End If

End Sub

'Get the subclasser lParamUser callback parameter
Public Property Get scc_lParamUser(ByVal ClassName As String) As Long
    If IsBadCodePtr(zMap_VFunction(ClassName, False)) = 0 Then        'Ensure that the thunk hasn't already released its memory
      scc_lParamUser = zData(IDX_PARMUSER)                        'Get the lParamUser callback parameter
    End If
End Property

'Let the subclasser lParamUser callback parameter
Public Property Let scc_lParamUser(ByVal ClassName As String, ByVal newValue As Long)
    If IsBadCodePtr(zMap_VFunction(ClassName, False)) = 0 Then        'Ensure that the thunk hasn't already released its memory
      zData(IDX_PARMUSER) = newValue                              'Set the lParamUser callback parameter
    End If
End Property

'Add the message to the specified table of the window handle
Private Function zAddMsg(ByVal uMsg As Long, ByVal nTable As Long) As Boolean
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long                                                        'Remember z_ScMem
      Dim I      As Long                                                        'Loop index
    
      zAddMsg = True
      nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
      z_ScMem = zData(nTable)                                                   'Map zData() to the specified table
      
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
      Else
        
        nCount = zData(0)                                                       'Get the current table entry count
        For I = 1 To nCount                                                     'Loop through the table entries
          If zData(I) = 0 Then                                                  'If the element is free...
            zData(I) = uMsg                                                     'Use this element
            GoTo Bail                                                           'Bail
          ElseIf zData(I) = uMsg Then                                           'If the message is already in the table...
            GoTo Bail                                                           'Bail
          End If
        Next I                                                                  'Next message table entry
    
        nCount = I                                                             'On drop through: i = nCount + 1, the new table entry count
        If nCount > MSG_ENTRIES Then                                           'Check for message table overflow
          zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
          zAddMsg = False
          GoTo Bail
        End If
        
        zData(nCount) = uMsg                                                    'Store the message in the appended table entry
      End If
    
      zData(0) = nCount                                                         'Store the new table entry count
Bail:
      z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Function

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
      Dim nCount As Long                                          'Table entry count
      Dim nBase  As Long                                          'Remember z_ScMem
      Dim I      As Long                                          'Loop index
    
      nBase = z_ScMem                                             'Remember z_ScMem so that we can restore its value on exit
      z_ScMem = zData(nTable)                                     'Map zData() to the specified table
    
      If uMsg = ALL_MESSAGES Then                                 'If ALL_MESSAGES are being deleted from the table...
        zData(0) = 0                                              'Zero the table entry count
      Else
        nCount = zData(0)                                         'Get the table entry count
        
        For I = 1 To nCount                                       'Loop through the table entries
          If zData(I) = uMsg Then                                 'If the message is found...
            zData(I) = 0                                          'Null the msg value -- also frees the element for re-use
            GoTo Bail                                             'Bail
          End If
        Next I                                                    'Next message table entry
        
        ' zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
      End If
      
Bail:
      z_ScMem = nBase                                             'Restore the value of z_ScMem
End Sub

Private Property Get zData(ByVal nIndex As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'Error handler
Private Sub zError(ByRef sRoutine As String, ByVal sMsg As String)
  ' \\LaVolpe -  Note. These two lines can be rem'd out if you so desire. But don't remove the routine
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long
  ' \\LaVolpe - Use ANSI calls for non-unicode usage, else use WideChar calls
  If asUnicode Then
    zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)         'Get the specified procedure address
  Else
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                 'Get the specified procedure address
  End If
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
  ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    ' Note: used both in subclassing and hooking routines
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim I     As Long                                                         'Loop index
  Dim J     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, I, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, I, bSub) Then                              'Probe for a Form method
      ' \\LaVolpe - Added propertypage offset
      If Not zProbe(nAddr + &H710, I, bSub) Then                            'Probe for a PropertyPage method
        If Not zProbe(nAddr + &H7A4, I, bSub) Then                          'Probe for a UserControl method
            Exit Function                                                   'Bail...
        End If
      End If
    End If
  End If
  
  I = I + 4                                                                 'Bump to the next entry
  J = I + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While I < J
    RtlMoveMemory VarPtr(nAddr), I, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    I = I + 4                                                               'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Do                                                             'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Function zInIDE(ByRef bIDE As Boolean) As Boolean
    ' only called when in IDE, never called when compiled
    bIDE = True
    zInIDE = bIDE
End Function

Private Function zMap_VFunction(ByRef ClassName As String, ByVal bReturnUnRegClass As Boolean) As Long
    ' Function maps a custom window class, registered using this class, to a memory pointer
    Dim WC As WNDCLASSEX
    
    z_ScMem = 0&                                    ' reset, zero indicates failure
    WC.cbSize = Len(WC)
    ' depending if RegisterClassExA or W was used, using
    ' the incorrect GetClassInfoExA or W will return invalid UDT values
    If GetClassInfoExA(App.hInstance, ClassName, WC) Then
        If WC.lpfnWndProc < WNDPROC_OFF Then        ' negative values are invalid
            If GetClassInfoExW(App.hInstance, StrPtr(ClassName), WC) < WNDPROC_OFF Then Exit Function
        End If
        If IsBadCodePtr(WC.lpfnWndProc - WNDPROC_OFF) = 0 Then
            ' most likely a class we registered, set the z_ScMem pointer
            z_ScMem = WC.lpfnWndProc - WNDPROC_OFF
            ' validate. Our class will have the memory pointer value
            ' in both the IDX_EBX and 0 offsets to that pointer
            If zData(IDX_EBX) = z_ScMem Then
                If zData(0) <> z_ScMem Then z_ScMem = 0&
            Else
                z_ScMem = 0&
            End If
        End If
    Else    ' class is not registered at all
        If bReturnUnRegClass Then z_ScMem = -1 ' special meaning during class registration
    End If
    zMap_VFunction = z_ScMem
    
End Function

'- callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zALL_API_WndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************

    
                        ' YOUR CODE HERE
    
' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub

