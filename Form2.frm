VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkNoMove 
      Caption         =   "Don't allow the form to be moved or minimized"
      Height          =   345
      Left            =   225
      TabIndex        =   5
      Top             =   3990
      Width           =   3870
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Prevent resizing method by preventing the  mouse from registering border hits."
      Height          =   450
      Index           =   1
      Left            =   225
      TabIndex        =   3
      Top             =   3420
      Width           =   3870
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Keep form size at 330 x 330 (Size Restriction Example)"
      Height          =   300
      Index           =   0
      Left            =   225
      TabIndex        =   2
      Top             =   3105
      Width           =   4275
   End
   Begin VB.TextBox Text1 
      Height          =   1110
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1590
      Width           =   4320
   End
   Begin VB.Label Label2 
      Caption         =   "Some more fun...."
      Height          =   285
      Left            =   195
      TabIndex        =   4
      Top             =   2775
      Width           =   4245
   End
   Begin VB.Label Label1 
      Height          =   1140
      Left            =   225
      TabIndex        =   1
      Top             =   315
      Width           =   4245
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' used for the CrossTalk example
Private Type uTest
    sText As String
    aByteArray() As Byte
End Type


' all of this is used for the example of restricting size
' and toggling ability to move
Private Const WM_SYSCOMMAND As Long = &H112
Private Const SC_MOVE As Long = &HF010&
Private Const SC_MINIMIZE As Long = &HF020&

Private Const WM_NCHITTEST As Long = &H84
Private Const HTCAPTION As Long = 2
Private Const HTBOTTOM As Long = 15
Private Const HTBOTTOMLEFT As Long = 16
Private Const HTBOTTOMRIGHT As Long = 17
Private Const HTLEFT As Long = 10
Private Const HTRIGHT As Long = 11
Private Const HTTOP As Long = 12
Private Const HTTOPLEFT As Long = 13
Private Const HTTOPRIGHT As Long = 14

Private Const WM_GETMINMAXINFO As Long = &H24
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const m_Static_CxCy As Long = 330
Private cSubclasser As cSelfSubHookCallback

Private Sub Form_Load()
    
    Label1.Caption = "Using a form doesn't do this example justice.  We can use callbacks to send a message to a private function within a form, usercontrol, and/or class.  And we can pass whatever we want, including UDTs.  In this example, the main form passed this form a custom UDT."
    
    Set cSubclasser = New cSelfSubHookCallback
    ' our subclass procedure will be this form, the 2nd ordinal
    If cSubclasser.ssc_Subclass(Me.hwnd, , 2, Me) = False Then
        Set cSubclasser = Nothing
        chkNoMove.Enabled = False
        Check1(0).Enabled = False
        Check1(1).Enabled = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cSubclasser = Nothing
End Sub

Private Sub chkNoMove_Click()

    If chkNoMove Then
        ' add the WM_SYSCOMMAND to be trapped, must be Before, not After for what we want
        cSubclasser.ssc_AddMsg Me.hwnd, eMsgWhen.MSG_BEFORE, WM_SYSCOMMAND
        
    Else
        ' remove the WM_SYSCOMMAND from the trapped message list
        cSubclasser.ssc_DelMsg Me.hwnd, eMsgWhen.MSG_BEFORE, WM_SYSCOMMAND
    End If

End Sub

Private Sub Check1_Click(Index As Integer)
    
    If Check1(Index).Value Then ' only one checked "No Size" option please
        
        ' if other checkbox is not checked, call event to remove previous message list
        If Check1(Abs(Index - 1)).Value = 0 Then
            Call Check1_Click(Abs(Index - 1))
        Else ' otherwise, set it to unchecked and event will be called for us
            Check1(Abs(Index - 1)).Value = 0
        End If
        
        If Index = 0 Then ' restrict the typical way: trap WM_GETMINMAXINFO
            ' WM_GETMINMAXINFO used to prevent manual resizing; Before or After - doesn't matter
            cSubclasser.ssc_AddMsg Me.hwnd, eMsgWhen.MSG_AFTER, WM_GETMINMAXINFO
            
            ' Start with a m_Static_CxCy by m_Static_CxCy window
            Me.Move Me.Left, Me.Top, m_Static_CxCy * Screen.TwipsPerPixelX, m_Static_CxCy * Screen.TwipsPerPixelY
                
        Else              ' restrict non-standard way: prevent mouse down on borders
            ' WM_NCHITTEST must be trapped before, not after
            cSubclasser.ssc_AddMsg Me.hwnd, eMsgWhen.MSG_BEFORE, WM_NCHITTEST
        
        End If
            
    Else
        
        ' unchecked item, remove message from message list as needed
        If Index = 0 Then
            cSubclasser.ssc_DelMsg Me.hwnd, eMsgWhen.MSG_AFTER, WM_GETMINMAXINFO
        Else
            cSubclasser.ssc_DelMsg Me.hwnd, eMsgWhen.MSG_BEFORE, WM_NCHITTEST
        End If
        
    End If
    
End Sub

'- ordinal #2
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
'*              message being passed to the original WndProc.
'*              Not applicable with After messages
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************
    
        ' in this example, twe are only trapping on a couple of messages
        
        Select Case uMsg
        
        Case WM_GETMINMAXINFO
        
            If Check1(0).Value Then ' restrict resizing this way
            
                Dim mmi As MINMAXINFO
                
                CopyMemory mmi, ByVal lParam, Len(mmi)   ' get suggested min/max data
                
                mmi.ptMaxTrackSize.x = m_Static_CxCy    ' set our max, unmaximized, size
                mmi.ptMaxTrackSize.y = m_Static_CxCy
                mmi.ptMinTrackSize = mmi.ptMaxTrackSize  ' make it our minimal size too
                
                CopyMemory ByVal lParam, mmi, Len(mmi)   ' override
            
                bHandled = True
                lReturn = 1
        
            End If
        
        Case WM_NCHITTEST
            
            If Check1(1).Value Then ' prevent resizing by not allowing border hits
                
                ' determine where on the form, the mouse is
                lReturn = DefWindowProc(lng_hWnd, uMsg, wParam, lParam)
                
                Select Case lReturn
                    ' if it is on one of these borders, then
                    ' we trick window into thinking it is NoWhere (value=0)
                    Case HTBOTTOM, HTBOTTOMLEFT, HTBOTTOMRIGHT, HTLEFT
                        lReturn = 0
                    Case HTTOP, HTTOPLEFT, HTTOPRIGHT, HTRIGHT
                        lReturn = 0
                End Select
                ' regardless, we have the real hit test or the faked one
                ' send the result back and don't forward message any further
                bHandled = True
            End If
        
        Case WM_SYSCOMMAND  ' only trapped when chkNoMove is checked
        
            Select Case wParam
            
            Case SC_MOVE, (SC_MOVE Or HTCAPTION)
            '   FYI: SC_Move occurs when Move selected from system menu
            '        SC_Move or HTCaption occurs when mouse is dragging caption bar
                
                ' prevent moving at all
                bHandled = True
                lReturn = 1
                
            Case SC_MINIMIZE
                ' prevent minimizing too
                bHandled = True
                lReturn = 1
                
            ' since the form has the MaxButton disabled, don't have to trap for SC_MAXIMIZE
            
            End Select
            
        End Select
            
        
End Sub

'- ordinal #1
Private Function TestUDTPassing(ByRef myCustomUDT As uTest, ByVal dummy1 As Long, ByVal dummy2 As Long, ByRef SomeValue As Long) As Long

    ' this function will be called by the main form.
    ' Since this will be used to be called from another window/class/etc, it must follow
    ' specific restrictions: 4 parameters and must be a function, not a sub
    ' The calling procedure must know in advance what to pass as the procedure's parameters,
    ' what ordinal this function exist at, and this procedure chooses whether the
    ' values are ByRef or ByVal
    
    ' Why is this special?  It isn't really, not with a form anyway.
    ' But lets say you want to use this in a usercontrol or class?
    ' Scenario:
    '   uc/class creates child classes that must be able to talk back to their parent object
    ' :: workarounds in the child classes, generally requiring the parent callback routine to be Public or Friend)
    '   1. Cache instance of parent (i.e., Set m_Parent = [whatever]) so you can call back to it.
    '       of course you now have circular references which are not wise to have
    '   2. Use system of implemented interfaces to talk back to parent. Doable, can be complicated
    '   3. Use soft references (i.e., CopyMemory pObject, ByVal m_Parent, 4& )
    '       problem here is that if not used correctly, crashes can occur
    ' :: with this method, only one work around is needed and parent callback routine can remain Private
    '   1. child class is passed the function address to call back to
    '       the parent calls scb_SetCallbackAddr for its address and then passes that for the child to use
    
    '   In fact, you can even do it in reverse too, if needed.
    '   Another scenario:
    '       Your usercontrol creates child classes. And you want to pass the child classes
    '       some data but you don't want their class' routine, the one used to receive the data,
    '       to be made Public or Friend?
    '   1. From parent, call scb_SetCallbackAddr passing the child class as the object
    '           Now use CallWindowProc to send the child class the information needed (including UDTs)

    Dim lSum As Long, x As Long
    For x = LBound(myCustomUDT.aByteArray) To UBound(myCustomUDT.aByteArray)
        lSum = lSum + myCustomUDT.aByteArray(x)
    Next
    
    Text1.Text = "UDT.sText = '" & myCustomUDT.sText & "'" & vbCrLf & vbCrLf
    Text1.Text = Text1.Text & "UDT.aByteArray UBound = " & UBound(myCustomUDT.aByteArray) & vbCrLf
    Text1.Text = Text1.Text & "Sum of UDT.aByteArray() Items: " & lSum
    
    SomeValue = Val(InputBox("Enter any non-decimal number below", "Cross-Talk", CLng(Rnd * vbWhite)))
    
    TestUDTPassing = True

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END FUNCTION" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Function

