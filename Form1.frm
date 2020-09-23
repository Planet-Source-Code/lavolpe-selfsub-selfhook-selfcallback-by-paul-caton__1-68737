VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Self-Subclass, Self-Hook, Self-Callback, CDECL Calls"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCDECL 
      Caption         =   "Show CDECL Example"
      Height          =   420
      Left            =   3750
      TabIndex        =   24
      Top             =   4590
      Width           =   2115
   End
   Begin VB.CommandButton cmdAPI2 
      Caption         =   "Example: Auo-Destroy API Window"
      Height          =   510
      Left            =   3300
      TabIndex        =   23
      Top             =   6075
      Width           =   2940
   End
   Begin VB.CheckBox chkTimer 
      Caption         =   "Timer Example, updates label >"
      Height          =   270
      Left            =   180
      TabIndex        =   21
      Top             =   105
      Width           =   2760
   End
   Begin VB.CommandButton cmdCrossTalk 
      Caption         =   "Example: Call private function of another form"
      Height          =   510
      Left            =   165
      TabIndex        =   20
      Top             =   6075
      Width           =   2940
   End
   Begin VB.CommandButton cmdDialogExamples 
      Caption         =   "Example: Center ColorDialog"
      Height          =   495
      Index           =   1
      Left            =   3300
      TabIndex        =   17
      Top             =   6615
      Width           =   2970
   End
   Begin VB.CommandButton cmdDialogExamples 
      Caption         =   "Example: Center MessageBox"
      Height          =   495
      Index           =   0
      Left            =   150
      TabIndex        =   16
      Top             =   6615
      Width           =   2970
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5985
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkCreateAPI 
      Caption         =   "Create custom class && window"
      Height          =   285
      Left            =   3315
      TabIndex        =   14
      Top             =   5205
      Width           =   2700
   End
   Begin VB.TextBox txtAPIhWnd 
      Enabled         =   0   'False
      Height          =   960
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   5055
      Width           =   3015
   End
   Begin VB.CommandButton cmdEnumHwnd 
      Caption         =   "Enumerate Fonts"
      Height          =   435
      Index           =   1
      Left            =   4875
      TabIndex        =   12
      Top             =   4080
      Width           =   1425
   End
   Begin VB.CommandButton cmdEnumHwnd 
      Caption         =   "Enumerate hWnds"
      Height          =   435
      Index           =   0
      Left            =   3285
      TabIndex        =   10
      Top             =   4080
      Width           =   1560
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   135
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   3765
      Width           =   3045
   End
   Begin VB.TextBox txtSubclassed 
      Enabled         =   0   'False
      Height          =   915
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   450
      Width           =   3015
   End
   Begin VB.CheckBox chkHookMouse 
      Caption         =   "Hook the mouse"
      Height          =   495
      Left            =   3315
      TabIndex        =   4
      Top             =   2685
      Width           =   1995
   End
   Begin VB.TextBox txtMouseHook 
      Enabled         =   0   'False
      Height          =   1080
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2595
      Width           =   3015
   End
   Begin VB.CheckBox chkHookKeybd 
      Caption         =   "Hook the keybaord"
      Height          =   495
      Left            =   3315
      TabIndex        =   2
      Top             =   1440
      Width           =   1890
   End
   Begin VB.TextBox txtKeybdHook 
      Enabled         =   0   'False
      Height          =   1080
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
   End
   Begin VB.CheckBox chkSubclassed 
      Caption         =   "Subclass this form"
      Height          =   495
      Left            =   3315
      TabIndex        =   0
      Top             =   450
      Width           =   1995
   End
   Begin VB.CheckBox chkNoHook 
      Caption         =   "Run Example without hooking"
      Height          =   240
      Index           =   0
      Left            =   165
      TabIndex        =   18
      Top             =   7125
      Width           =   2970
   End
   Begin VB.CheckBox chkNoHook 
      Caption         =   "Run Example without hooking"
      Height          =   240
      Index           =   1
      Left            =   3315
      TabIndex        =   19
      Top             =   7125
      Width           =   2970
   End
   Begin VB.Label lblTick 
      Caption         =   "Tick Count"
      Height          =   255
      Left            =   3195
      TabIndex        =   22
      Top             =   135
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Play with that window to test"
      Height          =   225
      Index           =   4
      Left            =   3315
      TabIndex        =   15
      Top             =   5580
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Example of callback for Enum APIs"
      Height          =   225
      Index           =   3
      Left            =   3315
      TabIndex        =   11
      Top             =   3750
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Simply move mouse or click to test"
      Height          =   225
      Index           =   2
      Left            =   3315
      TabIndex        =   8
      Top             =   3225
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Simply type on to the form to test"
      Height          =   225
      Index           =   1
      Left            =   3315
      TabIndex        =   7
      Top             =   2070
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Simply move mouse, resize, etc to test"
      Height          =   225
      Index           =   0
      Left            =   3315
      TabIndex        =   6
      Top             =   1005
      Width           =   2835
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' cSelfSubHookCallback. Class with extra code for sample project. Created from the clsSelfSubHookCallBk_Template class
' clsSelfSubHookCallBk_Template. A class combining subclassing, hooking, callbacks that can be used to create custom classes/modules
' clsAPIWindow_Template. A class designed specifically for creating custom window classes
' See top remarks in the classes


' This demo only put together to quickly demonstrate the template classes. You should
' already be familiar with Paul Caton's subclassing/hooking routines and if not there
' are links given in the class module so you can play with the original projects
' to gain that familiarity

' The flexibility of the thunks is fantastic. Send a callback procedure
' to virtually anywhere in your project: class, form, usercontrol, property page.

' The flexibility displayed in this project include:
' 1. Using window/hook procedures within a class and also in this form
' 2. Using hooking & subclassing together to accomplish a task: centering dialog windows
' 3. Stopping hooking/subclassing within the hook/subclass procedure
' 4. Creating an all-API window and trapping its messages
' 5. Calling a private function of another object and sending it a UDT
' 6. Using a single subclass procedure, simultaneously, for 3 different purposes. See myWndProc
' 7. Example of creating your own API timer
' 8. Showing how to interact with a class that is used for callbacks
' 9. How the clsSelfSubHookCallBk_Template can be used and modified to include only
'       the functions you need, and also how making some routines Public can be beneficial
' Short of global subclassing and global hooking, I don't think there are many limitations

' following used for the "All-API" window example
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Const WM_DESTROY As Long = &H2
Private hWnd_APIwindow As Long

' following usd for the "Center Dialog Window" examples
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Type CBT_CREATEWND
    lpcs As Long
    hWndInsertAfter As Long ' pointer to CreateStruct UDT
End Type
Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    Style As Long
    lpszName As Long    ' pointer to window title
    lpszClass As Long   ' atom or pointer to class name (always numeric)
    ExStyle As Long
End Type
Private Type WINDOWPOS
    hwnd As Long
    hWndInsertAfter As Long
    x As Long
    y As Long
    cx As Long
    cy As Long
    flags As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const WM_WINDOWPOSCHANGING As Long = &H46
Private Const SPI_GETWORKAREA As Long = 48
Private Const WM_SHOWWINDOW As Long = &H18
Private Const HCBT_CREATEWND As Long = 3

' Following used for the "CrossTalk" example
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Type uTest
    sText As String
    aByteArray() As Byte
End Type

' Following used for the 3 window procedure examples
Private Const WM_CTLCOLORSTATIC = &H138
Private Const WM_PAINT = &HF
Private Enum eExamples
    exThisForm = 1
    exAllAPI = 2
    exDialogCenter = 3
End Enum

' Following used for the API timer example
Private Declare Function KillTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private m_TimerID As Long


Private WithEvents cSubclasserHooker As cSelfSubHookCallback
Attribute cSubclasserHooker.VB_VarHelpID = -1

Private Sub Form_Load()
    
    Set cSubclasserHooker = New cSelfSubHookCallback
    Randomize   ' used in the cmdCrossTalk routine

End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    ' kill our timer if it is still active
    If m_TimerID Then chkTimer.Value = 0
    ' destroy our API window if it still alive
    If hWnd_APIwindow Then chkCreateAPI.Value = 0
    ' clear all subclassing, hoooking, callback thunks
    Set cSubclasserHooker = Nothing

End Sub

Private Sub cmdCDECL_Click()
    frmCDECL.Show
End Sub

Private Sub cmdAPI2_Click()

    Dim lWnd As Long
    Dim cAPI As New clsAPIwindow_Template   ' only used because it has all the CreateWindowEx stuff I need for this example
    Dim cCrashTest As cSelfSubHookCallback
    
    ' WARNING:  THIS ROUTINE WILL USE THE END STATEMENT, SO SAVE YOUR WORK JUST IN CASE
    
    ' When creating an API window from an existing Windows system class (i.e., static, edit, etc), if you do not specify it as a child window,
    ' then when END occurs in IDE, the window will remain and closing it could be very difficult if the window is hidden.
    
    ' By providing the parameter bIsAPIwindow=True, the thunks will destroy the window for you when IDE is shut down.
    
    ' Example:
    '   1. An API window created without a parent hWnd. The window is hidden and is a message relayer between menus & a project
    '   2. The API window is also the hWnd for a system tray icon
    '   3. When app starts, systray appears as expected. It is removed when app is closed normally (not END)
    '   4. But when user hits END, the VB IDE project closes, but the systray remains
    '   5. Do this 4 times and you have 4 systray objects
    
    ' Other ways you can ensure your API window closes without using the bIsAPIwindow parameter?
    ' 1. Never hit END & ensure API windows are destroyed in clean up routines (Unload/Terminate Events)
    ' 2. Assign the API window as a child to an existing VB window during creation. When parent is destroyed, children are too
    ' 3. If your API window is visible and has another method of closing then us that, i.e., X button, menu, etc
    ' 4. Use task manager if the window shows up in it
    ' 5. Close VB completely, when the thread dies, so do any windows created in it
    
    lWnd = cAPI.scc_CreateWindow("Edit", WS_OVERLAPPEDWINDOW Or WS_VISIBLE Or WS_SYSMENU, WS_EX_TOPMOST, "API Crash Test", 200, 100, 50, 50)
    If lWnd Then
        Select Case MsgBox("The API window appearing near top left corner of the screen, when running in IDE, normally will not " & vbCrLf & _
            "close if END is executed. To verify this click 'Yes'" & vbCrLf & vbCrLf & _
            "Using an optional parameter when subclassing, we can tell the thunks to destroy a window if IDE ended and the effect " & vbCrLf & _
            "will be immediate or when window gets its next message. Click 'No' to test this.", vbYesNoCancel + vbQuestion, "END Test")
            
        Case vbCancel
            DestroyWindow lWnd
            
        Case vbYes
            Set cCrashTest = New cSelfSubHookCallback
            cCrashTest.ssc_Subclass lWnd, , , , , , False
            MsgBox "When project returns to IDE as a result of END, the window will still be visible. Close it by clicking the X button", vbInformation + vbOKOnly, "API Crash Test"
            End
            
        Case vbNo
            Set cCrashTest = New cSelfSubHookCallback
            cCrashTest.ssc_Subclass lWnd, , , , , , True
            MsgBox "When project returns to IDE as a result of END, the window should be destroyed too. If not, just mouse over it", vbInformation + vbOKOnly, "API Crash Test"
            End
            
        End Select
    End If
    
End Sub

Private Sub cmdCrossTalk_Click()
    ' example of calling another form, class, usercontrol's private member.
    
    ' This method must use CallWindowProc API, therefore, some restrictions apply:
    '   1) The destination function must have 4 parameters.
    '           Whether those parameters are ByRef or ByVal is up to you
    '   2) To pass strings, objects, etc as one of the 4 parameters,
    '       you would use VarPtr( ) when you call CallWindowProc and ByRef in target function
    '       Arrays are a tad more complicated, see bottom of this routine
    '   3) The hWnd, uMsg, wParam, & lParam of the CallWindowProc are the 4 parameters
    '       any or all of the parameters can be used to pass information
    '   4) The address you use must be obtained by a call to scb_SetCallbackAddr
    '           passing the destination object and its ordinal
    '   5) Obviously, the calling routine must know the ordinal of the destination
    '       routine and also know what to pass as the parameter values
    '   Because you are passing pointers to strings, method is also unicode friendly
    
    Dim lAddr As Long
    
    Load Form2
    ' FYI: we don't really need to Load Form2 because when we reference it
    ' in the next line, it will be loaded by VB automatically.
    
    ' now get the address of the last Private member of that form; it will have 4 parameters
    lAddr = cSubclasserHooker.scb_SetCallbackAddr(4, 1, Form2)
    
    If lAddr = 0 Then   ' validate we got it
        Unload Form2
        MsgBox "Something is wrong, callback address was not received as expected", vbInformation + vbOKOnly
        
    Else
        Dim x As Long, myTestUDT As uTest
        
        ' fill in the test UDT with some stuff
        myTestUDT.sText = "A UDT with string and byte array members"
        x = CLng(Rnd * 20) + 1
        ReDim myTestUDT.aByteArray(1 To x)
        For x = 1 To x
            myTestUDT.aByteArray(x) = CByte(Rnd * 250) + 1
        Next
        
        ' show the target so we can see the call worked
        Form2.Show
        
        ' now send the UDT to the target, notes above explain why VarPtr used
        ' The other 2 parameters are not used, so we just supply zeroes
        CallWindowProc lAddr, VarPtr(myTestUDT), 0&, 0&, VarPtr(x)
        
        ' in the target function, X was changed and its new value will show here because of the
        '   ByRef declaration in the target function.  Obviously we could have supplied a
        '   return value for CallWindowProc and use that, but I wanted to show that parameter
        '   content can be changed just like any VB function.
        
        MsgBox "This message box is from the main form." & vbCrLf & _
            "The value you entered from the other form is " & x & vbCrLf & vbCrLf & _
            "Review comments in cmdCrossTalk_Click for significance.", vbInformation + vbOKOnly, "Cross-Talk Example"
        
        ' don't need this any longer for our example, so we will remove it.
        ' If we were going to call this many times, we would want to cache
        ' the lAddr value and not release this until our target closed or
        ' we closed. What if we already closed Form2, how could we release
        ' the thunk since we won't have Form2 as a parameter to pass? The
        ' answer is that we can't; however, the thunk will be released if
        ' scb_TerminateCallbacks is called when the class terminates or when
        ' the VB instance eventually terminated (compiled app closes or IDE closes)
        cSubclasserHooker.scb_ReleaseCallback 1, Form2
        
    End If
    
' ARRAYS: HOW TO PASS THEM
'
'     You can't use VarPtr(myArray) on an array to pass it, VB will give you an error.
'     Using VarPtr(myArray(n)) would pass the pointer to array member n, not the array.
'
'     There are two options you can use to pass arrays
'       Note that the UDT array example above is another option, however, it requires the UDT to be
'       declared in both the source & target whereas these options do not use UDTs
'
'     Option 1: Use API VarPtrArray(): Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
'     Example:
'           Dim myArray(1 To 50) As Long    ' array can be dynamic too
'           pass from this routine:  CallWindowProc targetAddr, VarPtrArray(myArray), 0&, 0&, 0&
'           in target routine:  Private Function MyCrossTalk(ByRef tArray() As Long, ByVal dummy1 As Long, ByVal dummy2 As Long, ByVal dummy3 As Long) As Long
'           now from your target routine, simply access your array from LBound(tArray) to UBound(tArray)
'
'     Option 2: Abuse Variants and literally pass anything and any number of "items"
'     Example:
'           Dim myArray(1 To 50) As Long    ' array can be dynamic too
'           Dim vArray(1 To 5) As Variant   ' number of items we will be passing
'           Dim v As Variant                ' something we can use VarPtr() on
'           :: items we will pass ::
'           Set vArray(1) = Me              ' Form1        (form object)
'           Set vArray(2) = Me.Icon         ' Form1's icon (stdPicture)
'           Set vArray(3) = Me.Font         ' Form1's font (stdFont)
'           vArray(4) = Me.Caption          ' Form1's caption (string)
'           vArray(5) = myArray             ' Long array
'           v = vArray                      ' assign variant array to non-arrayed variant so we can use VarPtr
'
'           pass from this routine: CallWindowProc targetAddr, VarPtr(v), 0&, 0&, 0&
'           in target routine: Private Function MyCrossTalk(ByRef myStuff As Variant, ByVal dummy1 As Long, ByVal dummy2 As Long, ByVal dummy3 As Long) As Long
'           to access the "stuff" in myStuff... You can also check LBound/UBound of myStuff if it can be a dynamic array
'               Debug.Print ObjPtr(myStuff(1)) which would be same as ObjPtr(Form1)
'               Debug.Print myStuff(2).Handle which would be same as Form1.Icon.Handle & myStuff(1).Icon.Handle
'               Debug.Print myStuff(3).Name which would be same as Form1.Font.Name & myStuff(1).Font.Name
'               Debug.Print myStuff(4) which would be same as Form1.Caption & myStuff(1).Caption
'               Debug.Print myStuff(5)(1) which is same value as myArray(1)
'                   Note that arrays can be referenced this way: myStuff(5)(LBound(myStuff(5)) To myStuff(5)(UBound(myStuff(5))
'                    or: Dim myArray() As Long : myArray = myStuff(5) : then simply reference myArray

    
End Sub

Private Sub chkCreateAPI_Click()
    ' example of creating an all-API window
    
    Dim cAllAPIwindow As clsAPIwindow_Template
    Dim bUnicode As Boolean
    
    If chkCreateAPI Then
        
        Set cAllAPIwindow = New clsAPIwindow_Template
        With cAllAPIwindow
            bUnicode = False ' change to true to test unicode creation
            
            ' register the class and assign a window procedure to the 1st ordinal in our form (myWndProc)
            ' We also include the optional parameter to let the window procedure know messages are for this example
            If .scc_RegClassProc("MyTestClass", vbInfoBackground, , , Me.Icon.Handle, , , , ByVal exAllAPI, 1, Me, , bUnicode) Then
                
                ' if we create an All-API class, we probably want both before & after messages
                .scc_AddMsg "MyTestClass", sccMsgWhen.MSG_BEFORE_AFTER, sccALLMessages.ALL_MESSAGES
                
                txtAPIhWnd.Text = "Test Class Registered Successfully"
                
                ' now create a window from our newly registered class
                ' The window's procedure is the class procedure
                hWnd_APIwindow = .scc_CreateWindow("MyTestClass", WS_POPUPWINDOW Or WS_CAPTION Or WS_VISIBLE Or WS_THICKFRAME, , "ALL API Window", 200, 200, 200, 100, , , bUnicode)
                                
                If hWnd_APIwindow = 0& Then
                    txtAPIhWnd.Text = "API Window failed to be created"
                End If
            
            Else
                
                txtAPIhWnd.Text = "Failed to Register Class"
            
            End If
        End With
    
    Else
        ' destroy the window if it hasn't already been closed
        If hWnd_APIwindow Then
            DestroyWindow hWnd_APIwindow
            hWnd_APIwindow = 0&             ' reset handle (used in Form1.Unload)
        End If
        ' unregister the class
        Set cAllAPIwindow = New clsAPIwindow_Template
        If cAllAPIwindow.scc_UnRegisterClassProc("MyTestClass") Then
            txtAPIhWnd.Text = "Unregistered Class"
        Else
            txtAPIhWnd.Text = "Failed to Unregister Class"
        End If
    End If

End Sub

Private Sub chkTimer_Click()

    If chkTimer.Value Then
        ' We set the last optional parameter to true since this will be a timer callback.
        ' That last parameter provides extra protection when we call SetTimer without passing that
        '   API an hWnd parameter as shown below.  Why is that a good thing?  Whenever a timer is
        '   set using an hWnd, the timer will be destroyed when the window is destroyed. But if no
        '   hWnd is passed, the timer remains alive until KillTimer is called. If you hit END in
        '   your app/IDE, then clean-up code is not executed and any KillTimer you may have set up
        '   does not get triggered. Prior to thunks, the timer would be using a callback that doesn't
        '   exist and crashes can occur.  The thunks will recognize this situation and simply prevent
        '   the timer from calling back but if the last parameter isn't set to true, then the timer
        '   will continue on forever until VB is closed. If set to true, the timer will be destroyed.
        
        ' Windowless-timer call
        ' this example has the callback procedure located at ordinal #3 in our form & has 4 parameters
        m_TimerID = SetTimer(0&, 0&, 128, cSubclasserHooker.scb_SetCallbackAddr(4, 3, Me, , True))
        
    Else
    
        ' Any item that continually calls back to your thunks must be terminated in normal clean-up
        ' code before a call to scb_ReleaseCallback is made. Otherwise, scb_ReleaseCallback will destroy
        ' the virtual memory containing your thunk and when Windows sends the next event to the thunk,
        ' a crash can occur because the thunk was erased.
        
        ' Interesting enough, the reverse is just the opposite. If scb_ReleaseCallback was not called,
        ' then no crash would occur when the next event is forwarded because the thunk memory is still
        ' alive, but then the memory will not be destroyed until VB or the compiled app closes. Your choice.
        
        KillTimer 0&, m_TimerID     ' stop active timer
        m_TimerID = 0&              ' reset ID (used in Form1.Unload)
        lblTick = "Tick Count"
        
    End If
        
End Sub

Private Sub lblTick_Click()
    chkTimer.Value = Abs(chkTimer.Value - 1)
End Sub

Private Sub chkHookKeybd_Click()
    ' example of hooking the keygboard
    
    If chkHookKeybd Then
        ' this example has the hook procedure located at ordinal #2 in the class, not this form
        If cSubclasserHooker.shk_SetHook(WH_KEYBOARD, , eMsgWhen.MSG_AFTER, , 2) Then
            txtKeybdHook.Text = "Hook successful"
        Else
            txtKeybdHook.Text = "HOOK FAILED!"
        End If
    Else
        cSubclasserHooker.shk_UnHook WH_KEYBOARD
    End If
End Sub

Private Sub chkHookMouse_Click()
    ' example of hooking the mouse
    
    If chkHookMouse Then
        ' this example has the hook procedure located at ordinal #2 in the class, not this form
        If cSubclasserHooker.shk_SetHook(WH_MOUSE, , eMsgWhen.MSG_AFTER, , 2) Then
            txtMouseHook.Text = "Hook successful"
        Else
            txtMouseHook.Text = "HOOK FAILED!"
        End If
    Else
        cSubclasserHooker.shk_UnHook WH_MOUSE
    End If
End Sub

Private Sub chkSubclassed_Click()
    ' example of subclassing an existing window
    
    If chkSubclassed Then
        ' this example has the window procedure located at ordinal #1 in this form (myWndProc)
        ' We also include the optional parameter to let the window procedure know messages are for this example
        If cSubclasserHooker.ssc_Subclass(Me.hwnd, ByVal exThisForm, 1, Me) Then
            cSubclasserHooker.ssc_AddMsg Me.hwnd, eMsgWhen.MSG_AFTER, eAllMessages.ALL_MESSAGES
        End If
    Else
        cSubclasserHooker.ssc_UnSubclass Me.hwnd
    End If
End Sub

Private Sub cmdEnumHwnd_Click(Index As Integer)
    ' example of calling an enum API
    ' The calls to scb_SetCallbackAddr are in those routines...
    List1.Clear
    If Index = 0 Then
        Call cSubclasserHooker.EnumWindowHandles
    Else
        Call cSubclasserHooker.EnumFonts(Me.hDC)
    End If
End Sub

Private Sub cmdDialogExamples_Click(Index As Integer)
    ' example of centering dialog boxes
    
    ' start a hooking session, unless unwanted.
    
    ' We want the hook procedure to be the 2nd ordinal in our form
    If chkNoHook(Index) = 0 Then cSubclasserHooker.shk_SetHook WH_CBT, , , , 2, Me
    
    ' create the dialog window
    
    If Index = 0 Then ' message box example
        If chkNoHook(Index) = 0 Then
            MsgBox "This message box should be centered on the form", vbQuestion + vbOKOnly, "Centering Example"
        Else
            MsgBox "This message box is probably centered on the screen", vbInformation + vbOKOnly, "Centering Example"
        End If
    Else              ' color dialog example; change to ShowOpen, ShowFont, ShowPrinter to test other dialogs
        CommonDialog1.ShowColor
    End If
    
    ' Release our hook if not already unhooked. We don't want it to continue indefinitely
    If chkNoHook(Index) = 0 Then cSubclasserHooker.shk_UnHook WH_CBT
    
End Sub

Private Function PositionDialog(dlgWidth As Long, dlgHeight As Long, dlgLeft As Long, dlgTop As Long) As Boolean
    
    ' Example of centering dialog boxes
    ' The myWndProc & myHookProc, near end of this module, calls this procedure,
    ' passing the dialog window's width, height, left & top.
    ' We simply modify the left,top coords
    
    ' Remember that APIs use pixels, therefore we need to convert
    ' VB's vbTwips to vbPixels in order to provide accurate coords.
    If dlgWidth > 0 And dlgHeight > 0 Then
        ' when centering, check for width & height because they could be zero, believe it or not
        
        Dim wRect As RECT
        SystemParametersInfo SPI_GETWORKAREA, 0&, wRect, 0&
        
        dlgLeft = ((Me.Width \ Screen.TwipsPerPixelX) - dlgWidth) \ 2 + Me.Left \ Screen.TwipsPerPixelX
        ' simple check to prevent dialog from displaying off the screen (horizontally)
        If (dlgLeft + dlgWidth) > wRect.Right Then dlgLeft = wRect.Right - dlgWidth
        If dlgLeft < wRect.Left Then dlgLeft = wRect.Left
        
        dlgTop = ((Me.Height \ Screen.TwipsPerPixelY) - dlgHeight) \ 2 + Me.Top \ Screen.TwipsPerPixelY
        ' simple check to prevent dialog from displaying off the screen (vertically)
        If (dlgTop + dlgHeight) > wRect.Bottom Then dlgTop = wRect.Bottom - dlgHeight
        If dlgTop < wRect.Top Then dlgTop = wRect.Top
        
        PositionDialog = True
    
    End If

End Function


Private Sub cSubclasserHooker_EnumCallback(EnumValue As Variant)
    ' one of the class's Public Events
    List1.AddItem EnumValue
End Sub

Private Sub cSubclasserHooker_HookMessage(ByVal HookType As eHookType, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long)
    
    ' one of the class's Public Events
    If nCode = 0 Then 'HC_ACTION
        Dim sMessage As String
        sMessage = "nCode: &H" & Hex(nCode) & vbNewLine & "wParam: &H" & Hex(wParam)
        
        If HookType = eHookType.WH_KEYBOARD Then
            Select Case wParam
            Case vbKeyA To vbKeyZ
                sMessage = sMessage & "   " & Chr$(wParam)
            Case vbKey0 To vbKey9
                sMessage = sMessage & "   " & Chr$(wParam)
            Case vbKeyF1 To vbKeyF16
                sMessage = sMessage & "    F" & CStr(wParam - vbKeyF1 + 1)
            Case vbKeyHome
                sMessage = sMessage & "    Home"
            Case vbKeyEnd
                sMessage = sMessage & "    End"
            Case vbKeyPageDown
                sMessage = sMessage & "    PgDn"
            Case vbKeyPageUp
                sMessage = sMessage & "    PgUp"
            Case vbKeyInsert
                sMessage = sMessage & "    Insert"
            Case vbKeyDelete
                sMessage = sMessage & "    Delete"
            Case vbKeyDecimal
                sMessage = sMessage & "    Decimal"
            Case vbKeyDivide
                sMessage = sMessage & "    Divide"
            Case vbKeyMultiply
                sMessage = sMessage & "    Multiply"
            Case vbKeyAdd
                sMessage = sMessage & "    Add +"
            Case vbKeySubtract
                sMessage = sMessage & "    Subtract -"
            Case vbKeyDown
                sMessage = sMessage & "    Down Arrow"
            Case vbKeyLeft
                sMessage = sMessage & "    Left Arrow"
            Case vbKeyRight
                sMessage = sMessage & "    Right Arrow"
            Case vbKeyUp
                sMessage = sMessage & "    Up Arrow"
            Case vbKeyMenu
                sMessage = sMessage & "    ALT Key"
            Case vbKeyControl
                sMessage = sMessage & "    CTRL Key"
            Case vbKeyShift
                sMessage = sMessage & "    Shift"
            Case vbKeyReturn
                sMessage = sMessage & "    Enter"
            Case vbKeyEscape
                sMessage = sMessage & "    Escape"
            Case vbKeyTab
                sMessage = sMessage & "    Tab"
            Case vbKeyBack
                sMessage = sMessage & "    Backspace"
            Case vbKeySpace
                sMessage = sMessage & "    Space"
            Case vbKeyCapital
                sMessage = sMessage & "    Caps Lock"
            Case vbKeyNumlock
                sMessage = sMessage & "    Num Lock"
            Case vbKeyNumpad0 To vbKeyNumpad9
                sMessage = sMessage & "    Numpad" & wParam - vbKeyNumpad0
            Case &H5B, &H5C
                sMessage = sMessage & "    Flying Window Key"
            Case &H5D
                sMessage = sMessage & "    Context Menu Key"
            Case Else
                sMessage = sMessage & "    Other"
            End Select
        
            sMessage = sMessage & vbNewLine & "lParam: &H" & Hex(lParam)
            If lParam < 0 Then          'wParam are vbKey... variables
                sMessage = sMessage & vbCrLf & "Note: Key is up"
            Else
                If (lParam \ &H40000000) Then
                    sMessage = sMessage & vbCrLf & "Note: Key held down"
                Else
                    sMessage = sMessage & vbCrLf & "Note: Key is down"
                End If
                If ((lParam \ &H20000000) And 1) Then
                    sMessage = sMessage & vbCrLf & "Additional: ALT Key is also down"
                End If
            End If
        
        Else
            sMessage = sMessage & vbNewLine & "lParam: &H" & Hex(lParam)
        End If
        
        
        If HookType = WH_MOUSE Then
            txtMouseHook = sMessage
        Else
            txtKeybdHook = sMessage
        End If

    End If
End Sub

' ordinal #3 - Example of a timer procedure callback used with scb_SetCallbackAddr
' http://msdn2.microsoft.com/en-us/library/ms644907.aspx
Private Function myTimerProc(ByVal lng_hWnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long
    
    ' note: although a timer procedure, per MSDN, does not return a value,
    ' the function that is used for callbacks must return a value therefore
    ' all callback routines must be functions, even if the return value is not used
                        
    ' Note that the scb_SetCallbackAddr has an optional parameter for timers created by
    ' SetTimer. Set that parameter to true to provide extra protection while in IDE.
    ' See scb_SetCallbackAddr & scb_ReleaseCallback
    
    lblTick.Caption = "Tick Count: " & tickCount

End Function

' ordinal #2 ' Example of a hook procedure used to center a dialog window
Private Sub myHookProc(ByVal bBefore As Boolean, _
                        ByRef bHandled As Boolean, _
                        ByRef lReturn As Long, _
                        ByVal nCode As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long, _
                        ByVal lHookType As eHookType, _
                        ByRef lParamUser As Long)
'*************************************************************************************************
' http://msdn2.microsoft.com/en-us/library/ms644990.aspx
'* bBefore    - Indicates whether the callback is before or after the next hook in chain.
'* bHandled   - In a before next hook in chain callback, setting bHandled to True will prevent the
'*              message being passed to the next hook in chain and (if set to do so).
'* lReturn    - Return value. For Before messages, set per the MSDN documentation for the hook type
'* nCode      - A code the hook procedure uses to determine how to process the message
'* wParam     - Message related data, hook type specific
'* lParam     - Message related data, hook type specific
'* lHookType  - Type of hook calling this callback
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************
    
'http://msdn2.microsoft.com/en-us/library/ms644977.aspx
    If lHookType = WH_CBT Then
        If nCode = HCBT_CREATEWND Then  ' flag indicating window is being created
        
            Dim wcw As CREATESTRUCT
            Dim hcw As CBT_CREATEWND
            ' get the hcbt_createwnd structure
            CopyMemory hcw, ByVal lParam, Len(hcw)
            
            If hcw.lpcs Then    ' pointer to a createstruct
                CopyMemory wcw, ByVal hcw.lpcs, Len(wcw)   ' get that structure
                
                If wcw.lpszClass = 32770 Then ' dialog class name atom
                    ' not all dialogs are created equal :)
                    ' messageboxes can be positioned here, while other dialogs
                    ' can change position coords when this is received
                    ' and when it is finally shown....
                    
                    ' by trying to adjust size here and also subclassing
                    ' the window, we can catch it either way
                    
                    ' call local procedure to position dialog & save results
                    If PositionDialog(wcw.cx, wcw.cy, wcw.x, wcw.y) Then
                        CopyMemory ByVal hcw.lpcs, wcw, Len(wcw)
                    End If
                    
                    ' start subclassing the dialog window. wParam is the handle
                    ' we want the subclass procedure at ordinal #1 in our form (myWndProc)
                    ' We also include the optional parameter to let the window procedure know messages are for this example
                    If cSubclasserHooker.ssc_Subclass(wParam, ByVal exDialogCenter, 1, Me) Then
                        cSubclasserHooker.ssc_AddMsg wParam, eMsgWhen.MSG_AFTER, WM_SHOWWINDOW, WM_WINDOWPOSCHANGING
                    End If
                    
                    ' we can unhook now
                    cSubclasserHooker.shk_UnHook lHookType
                End If
                
            End If
            
        End If
    End If
End Sub

'- callback, usually ordinal #1, the last method in this source file----------------------
Private Sub myWndProc(ByVal bBefore As Boolean, _
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

    Select Case lParamUser

    Case exThisForm: ' subclassing this form
    
        If Not uMsg = WM_PAINT Then
            ' if timer is firing, we get lots of paint events as the label is constantly updated; let's not display paint messages
        
            If Not uMsg = WM_CTLCOLORSTATIC Then ' Else updating textbox fires another WM_CTLCOLORSTATIC; infinite loop
                ' if textbox was enabled, WM_CTLCOLOREDIT would need to be bypassed instead
            
                txtSubclassed.Text = "Message: &H" & Hex(uMsg) & vbNewLine & _
                    "wParam: &H" & Hex(wParam) & vbNewLine & _
                    "lParam: &H" & Hex(lParam)
    
    '            Debug.Print "Message: "; uMsg; wParam; lParam
                    
                ' Note that in real life, you wouldn't want to do this normally.
                ' Why? Well if we are shutting down, txtSubclassed control gets unloaded
                ' before main form. We are subclassing the main form. Subclassing continues
                ' until main form closes, and this routine will expect the textbox but it
                ' was unloaded when we closed the main form. Error "Object was Unloaded"
                ' There are 2 easy ways we can prevent this:
                '   1) Don't reference a control in a callback event
                '   2) Stop the subclassing on Form_Unload
                ' using On Error Resume, will just reload the form but not show it
                
                ' The same comments can apply to hooks too
            End If
            
        End If
        
    Case exAllAPI: ' all-API window example

        ' This example uses the clsAPIwindow_Template to register
        ' a custom window class and tell windows that the window
        ' procedure for the class is here. All windows created
        ' from the class will have their window procedure come
        ' here unless the window is subclassed elsewhere.

        If bBefore = False Then
            ' note that we asked for both Before & After messages.
            ' Since that is the case, in order to know if this message
            ' occurs before the subclassed window got it or after,
            ' we test the bBefore parameter

            txtAPIhWnd.Text = "Message: &H" & Hex(uMsg) & vbNewLine & _
                "wParam: &H" & Hex(wParam) & vbNewLine & _
                "lParam: &H" & Hex(lParam) & vbNewLine & _
                "hWnd: &H" & Hex(lng_hWnd)

            If uMsg = WM_DESTROY Then
                hWnd_APIwindow = 0  ' reset our variable (used in Form1.Unload)
            End If

        End If


        ' Note that in real life, you wouldn't want to do this normally.
        ' Why? We are subclassing a window outside of this form.
        ' Subclassing continues even though this form is closing. If this routine
        ' is called somewhere after this form is made invisible but before it is
        ' closed (somewhere between Unload & Terminate), the callback can prevent
        ' VB from closing the form. So this form becomes invisible, but remains loaded.

        ' There are 2 easy ways we can prevent this:
        '   1) Don't reference a control in a callback event
        '   2) Destroy the subclassed window in the Unload event



    Case exDialogCenter: ' dialog centering example

        ' this procedure is used for the subclassed dialog window
        ' See myHookProc above
        If uMsg = WM_WINDOWPOSCHANGING Then

            Dim wPos As WINDOWPOS
            ' window is resizing; get its dimensions
            CopyMemory wPos, ByVal lParam, Len(wPos)

            ' call local procedure to position dialog & save results
            If PositionDialog(wPos.cx, wPos.cy, wPos.x, wPos.y) Then
                ' save the new position
                CopyMemory ByVal lParam, wPos, Len(wPos)
            End If

        ElseIf uMsg = WM_SHOWWINDOW Then

            ' when window is shown, release subclassing
            ' otherwise user won't be able to move the window manually,
            ' because the WM_WindowPosChanging trap above will constantly
            ' try to recenter it on the form
            cSubclasserHooker.ssc_UnSubclass lng_hWnd

        End If

    End Select
    
' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub
