VERSION 5.00
Begin VB.Form frmCDECL 
   Caption         =   "CDECL"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2730
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   255
      Width           =   1545
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1425
      TabIndex        =   2
      Text            =   "155"
      Top             =   255
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   180
      TabIndex        =   1
      Text            =   "500"
      Top             =   255
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run Code"
      Height          =   495
      Left            =   1515
      TabIndex        =   0
      Top             =   1650
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   795
      Left            =   180
      TabIndex        =   4
      Top             =   810
      Width           =   4125
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2340
      X2              =   2550
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2340
      X2              =   2550
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1110
      X2              =   1320
      Y1              =   450
      Y2              =   450
   End
End
Attribute VB_Name = "frmCDECL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Dim c_code(0 To 4) As Long
Dim c As New clsCDECL

' ASM code to mimic a function within a C dll
c_code(0) = &H81E58955: c_code(1) = &H4EC&: c_code(2) = &H8458B00: c_code(3) = &H890C452B: c_code(4) = &HC35DEC
' The simple ASM code follows:
'    push ebp
'    mov ebp, esp
'    sub esp, 4             ;create function stack frame
'    mov eax, [ebp + 8]     ;move 1st param into eax
'    sub eax, [ebp + 12]    ;subtract 2nd param from 1st param
'    mov esp, ebp           ;restore stack
'    pop ebp
'    ret                    ;return is eax, CDECL does not clean up stack else it would look like: ret 8
'                           ;CDECL and STDCALL calling conventions also reverse the order of the parameters
'                           ;But Paul Caton's CDECL routine does that for us, so we supply parameters in the order we are familiar with

On Error Resume Next

    Text3 = c.CallFunc(vbNullString, VarPtr(c_code(0)), Val(Text1), Val(Text2))

    If Err Then
        MsgBox Err.Description, vbExclamation + vbOKOnly
        Err.Clear
    End If

'    ' if you had a real C API declaration the call might look like the following
'    Private Declare Function Zcrc321 Lib "zlib1.dll" Alias "crc32" (ByVal crc As Long, ByRef buf As Any, ByVal Length As Long) As Long
'
'    Dim cCfunction As clsCDECL
'    Dim lReturn As Long
'
'    Set cCfunction = New clsCDECL
'    cCfunction.DllLoad "zlib1"
'    lReturn = clsCfunction.CallFunc("crc32", 0&, byteArray(0), lenByteArray)
'    cCfunction.DllUnload

End Sub


Private Sub Form_Load()
    Label1.Caption = "The clsCDECL allows you to call a C++ DLL or function. Really simple C++ function. Subtracts values and displays result."
End Sub
