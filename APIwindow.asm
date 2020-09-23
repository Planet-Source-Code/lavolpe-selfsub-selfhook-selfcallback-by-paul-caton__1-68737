;*****************************************************************************************
;** APIWindow.asm - subclassing thunk. Assemble with nasm.
;**
;** Paul_Caton@hotmail.com
;** Copyright free, use and abuse as you see fit.
;**
;** v1.0 LaVolpe: Created from Subclass ASM						... 20070630
;*****************************************************************************************

;***************
;API definitions
M_RELEASE	equ	8000h	    ;VirtualFree memory release flag

;******************************
;Stack frame access definitions
%define lParam          [ebp + 48]  ;WndProc lParam
%define wParam          [ebp + 44]  ;WndProc wParam
%define uMsg            [ebp + 40]  ;WndProc uMsg
%define hWnd            [ebp + 36]  ;WndProc hWnd
%define lRetAddr        [ebp + 32]  ;Return address of the code that called us
%define lReturn         [ebp + 28]  ;lReturn local, restored to eax after popad
%define bHandled        [ebp + 20]  ;bHandled local, restored to edx after popad

;***********************
;Data access definitions
%define nCallStack      [ebx]       ;validation flag
%define bShutdown       [ebx + 4]  	;Shutdown flag
%define fnEbMode        [ebx + 8]  	;EbMode function address
%define fnDefWinProc   	[ebx + 12]  ;DefWindowProc function address
%define fnDestroyWin    [ebx + 16]  ;DestroyWindow function address
%define fnIsBadCodePtr  [ebx + 20]  ;IsBadCodePtr function address
%define objOwner        [ebx + 24]  ;Owner object address
%define addrCallback    [ebx + 28]  ;Callback address
%define addrTableB      [ebx + 32]  ;Address of before original WndProc message table
%define addrTableA      [ebx + 36]  ;Address of after original WndProc message table
%define lParamUser      [ebx + 40]  ;User defined callback parameter

use32
;************
;Data storage
    dd_nCallStack		dd 0	    ;validation flag
    dd_bShutdown		dd 0	    ;Shutdown flag
    dd_fnEbMode 		dd 0	    ;EbMode function address
    dd_fnDefWinProc   	dd 0	    ;CallWindowProc function address
    dd_fnDestroyWin    	dd 0	    ;SetWindowsLong function address
    dd_fnIsBadCodePtr	dd 0	    ;ISBadCodePtr function address
    dd_objOwner 		dd 0	    ;Owner object address
    dd_addrCallback	dd 0	    ;Callback address
    dd_addrTableB		dd 0	    ;Address of before original WndProc message table
    dd_addrTableA		dd 0	    ;Address of after original WndProc message table
    dd_lParamUser		dd 0	    ;User defined callback parameter
    
;***********
;Thunk start    
    xor     eax, eax		    	;Zero eax, lReturn in the ebp stack frame
    xor     edx, edx		    	;Zero edx, bHandled in the ebp stack frame
    pushad			    		;Push all the cpu registers on to the stack
    mov     ebp, esp		    	;Setup the ebp stack frame
    mov     ebx, 012345678h	    	;dummy Address of the data, patched from VB
 
    xor     esi, esi		    	;Zero esi

    cmp     fnEbMode, eax	    	;Check if the EbMode address is set
    jnz     _ide_state		    	;Running in the VB IDE

Align 4
_before:			    		;Before the original WndProc
    dec     edx 		    		;edx <> 0, bBefore callback parameter = True
    mov     edi, addrTableB	    	;Get the before message table
    mov	lReturn, esi
    call    _callback		    	;Attempt the VB callback
    
_before_handled:
    cmp     bHandled, esi	    	;If bHandled <> False
    jne     _return		    	;The callback entirely handled the message

Align 4
_original_wndproc:
    call    _wndproc		    	;Call the original WndProc
    
Align 4
_after: 			    		;After the original WndProc
    xor     edx, edx		    	;Zero edx, bBefore callback parameter = False
    mov     edi, addrTableA	    	;Get the after message table
    call    _callback		    	;Attempt the VB callback

Align 4
_return:			    		;Clean up and return to caller
    popad			    		;Pop all registers. lReturn is popped into eax
    ret     16			    	;Return with a 16 byte stack release

Align 4
_ide_state:			    		;Running under the VB IDE
    cmp	bShutdown, esi
    jz	_check_state

    call	_wndproc
    jmp	_return

Align 4
_check_state:
    call    near fnEbMode	    	;Determine the IDE state

    cmp     eax, dword 1	    	;If EbMode = 1
    je	_before		    	;Running normally
    
    test    eax, eax		    	;If EbMode = 0
    jz	_shutdown			;Ended, shutdown

    call    _wndproc		    	;EbMode = 2, breakpoint... call original WndProc
    jmp     _return		    	;Return

Align 4
_wndproc:			    		
    push    dword lParam	    	;ByVal lParam
    push    dword wParam	    	;ByVal wParam
    push    dword uMsg		    	;ByVal uMsg
    push    dword hWnd		    	;ByVal hWnd
    call    near fnDefWinProc   	;Call DefWindowProc
    mov     lReturn, eax	    	;Save the return value
    jmp     _generic_ret	    	;return

Align 4
_shutdown:			    		;Destroy window
    mov	bShutdown, dword 1
    push    dword hWnd		    	;Push the window handle
    call    fnDestroyWin   		;Call DestroyWindow - only gets here when IDE has stopped
    jmp	_return	

Align 4
_callback:			    		;Validate the callback
    mov     ecx, [edi]		    	;ecx = table entry count
    jecxz   _generic_ret	    	;ecx = 0, table is empty

    test    ecx, ecx		    	;Set the flags as per ecx
    js	_call		    		;Table entry count is negative, all messages callback
    add     edi, 4		    	;Inc edi to point to the start of the callback table
    mov     eax, uMsg		    	;eax = the value to scan for
    repne   scasd		    		;Scan the callback table for uMsg
    jne     _generic_ret	    	;uMsg not in the callback table

Align 4
_call:				    	;Callback required, do it...
    push    edx 		    		;Preserve edx (bBefore)
    push    dword addrCallback	;Push the callback address
    call    dword fnIsBadCodePtr    ;Check the code is live
    pop     edx 		    		;Preserve edx (bBefore)
    jnz     _generic_ret	    	;If not, skip callback

    lea     eax, lParamUser	    	;Address of lParamUser
    lea     ecx, bHandled	    	;Address of bHandled
    push    eax 		    		;ByRef lParamUser
    lea     eax, lReturn	    	;Address of lReturn
    push    dword lParam	    	;ByVal lParam
    push    dword wParam	    	;ByVal wParam
    push    dword uMsg		    	;ByVal uMsg
    push    dword hWnd		    	;ByVal hWnd
    push    eax 		    		;ByRef lReturn
    push    ecx 		    		;ByRef bHandled
    push    edx 		    		;ByVal bBefore
    push    dword objOwner	    	;ByVal the owner object
    call    near addrCallback	    	;Call the zWndProc callback procedure

Align 4
_generic_ret:			    	;Shared return (local only)
    ret
